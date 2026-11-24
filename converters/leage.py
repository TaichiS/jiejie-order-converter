"""
converters/leage.py
樂齡網訂單轉換器。
輸入：PDF 採購單（多張）
輸出：鼎新-樂齡信{MMDD}.xlsx + 黑貓託運單{MMDD}.csv
"""
from __future__ import annotations

import csv
import re
from datetime import datetime, timedelta
from difflib import get_close_matches
from pathlib import Path

import pdfplumber
from openpyxl import Workbook

from converters.base import BaseConverter, ConversionResult, RowError
from app.utils import output_path

SENDER = {
    "name":    "賦寶股份有限公司",
    "tel":     "05-5870993",
    "mobile":  "0909-870-993",
    "address": "雲林縣莿桐鄉大美村溪美22-3號1樓",
    "temp":    "3",
    "size":    "2",
    "deliver": "4",
}

# 樂齡採購單的廠商貨號 → 鼎新品號。
SUPPLIER_SKU_ALIASES = {
    "N1HF00003340000": "F30100063",
    "N1HF00003350000": "F30100064",
    "N1HF00003360000": "F30100065",
    "N1HF00003370000": "F30100066",
    "N1HF00003340999": "E52990008",
}


class LeageConverter(BaseConverter):

    @property
    def source_type(self) -> str:
        return "leage"

    def _validate(self, input_files: list[Path]) -> list[RowError]:
        pdfs = [f for f in input_files if f.suffix.lower() == ".pdf"]
        if not pdfs:
            return [RowError(0, "file", "", "找不到 PDF 採購單")]
        return []

    def _process(self, input_files: list[Path]) -> ConversionResult:
        pdfs = sorted(f for f in input_files if f.suffix.lower() == ".pdf")

        sku_map = {
            str(product.get("品號", "")).strip(): product
            for product in self._product_map.values()
            if str(product.get("品號", "")).strip()
        }

        orders: list[dict]      = []
        errors: list[RowError]  = []
        first_delivery_date     = ""

        for row_idx, pdf_path in enumerate(pdfs, start=1):
            try:
                order = _parse_pdf(pdf_path)
            except Exception as e:
                errors.append(RowError(row_idx, "file", pdf_path.name, f"PDF 解析失敗：{e}",
                                       source_file=pdf_path.name))
                continue

            if not order:
                errors.append(RowError(row_idx, "content", pdf_path.name, "無法從 PDF 擷取有效內容",
                                       source_file=pdf_path.name))
                continue

            matched_items = []
            for item_no, item in enumerate(order["items"], start=1):
                norm_name = _strip_prefix(item["raw_product_name"])
                target_sku = SUPPLIER_SKU_ALIASES.get(item["supplier_sku"], "")
                product = sku_map.get(target_sku) if target_sku else None
                if not product:
                    product = self._match_leage_product(norm_name)
                if not product:
                    errors.append(RowError(
                        row_idx, f"品項{item_no}", item["raw_product_name"],
                        f"品號資料中找不到「{norm_name}」",
                        source_file=pdf_path.name,
                    ))
                    continue
                matched_items.append({
                    **item,
                    "product": product,
                    "prod_name": norm_name,
                })

            if not matched_items:
                continue
            order["items"] = matched_items
            if not first_delivery_date:
                first_delivery_date = order.get("delivery_date", "")
            orders.append(order)

        if not orders:
            return ConversionResult(source_type=self.source_type,
                                    fail_count=len(errors), errors=errors,
                                    status="failed")

        # 依 PO 尾碼排序
        orders.sort(key=lambda o: _po_suffix(o.get("po_number", "")))

        # 檔名與訂單編號日期使用實際上傳／執行轉換的當天日期，
        # 不受 PDF 檔名或交貨日期影響。
        delivery_date = first_delivery_date  # YYYY/MM/DD
        dt = datetime.today()
        mmdd_str = dt.strftime("%m%d")
        order_date_str = dt.strftime("%Y%m%d")

        # 指派訂單編號
        for seq, order in enumerate(orders, start=1):
            order["order_id"] = f"jjseniorfood{mmdd_str}{seq:03d}"

        # 儲存
        output_files = []
        date_obj     = dt

        p = _save_leage_xlsx(
            self.output_dir.parent, self.source_type, date_obj,
            f"鼎新-樂齡信{mmdd_str}.xlsx", orders, delivery_date
        )
        output_files.append(p)

        p = _save_leage_csv(
            self.output_dir.parent, self.source_type, date_obj,
            f"黑貓託運單{mmdd_str}.csv", orders
        )
        output_files.append(p)

        success_count = len(orders)
        fail_count    = len(errors)
        status = "completed" if fail_count == 0 else ("partial" if success_count > 0 else "failed")

        return ConversionResult(
            source_type   = self.source_type,
            success_count = success_count,
            fail_count    = fail_count,
            order_date    = order_date_str,
            output_files  = output_files,
            errors        = errors,
            status        = status,
        )

    def _match_leage_product(self, norm_name: str) -> dict | None:
        """
        精確比對或模糊比對樂齡網品名。
        只在樂齡品號（F30開頭）中比對。
        """
        leage_prods = {k: v for k, v in self._product_map.items()
                       if str(v.get("品號", "")).startswith("F")}
        leage_names = list(leage_prods.keys())

        # 精確比對
        if norm_name in leage_prods:
            return leage_prods[norm_name]

        # 模糊比對
        matches = get_close_matches(norm_name, leage_names, n=1, cutoff=0.5)
        if matches:
            return leage_prods[matches[0]]

        # 關鍵字比對（針對「粥彭湃/粥澎派」等拼字差異）
        for name, prod in leage_prods.items():
            # 簡化品名（只取漢字部分）
            n1 = re.sub(r"[^\u4e00-\u9fff]", "", norm_name)
            n2 = re.sub(r"[^\u4e00-\u9fff]", "", name)
            if n1 and n2 and (n1 in n2 or n2 in n1):
                return prod

        return None


# ── PDF 解析 ──────────────────────────────────────────────────────────────

def _parse_pdf(path: Path) -> dict | None:
    """從一張樂齡網 PDF 採購單中擷取關鍵欄位"""
    with pdfplumber.open(path) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    # 訂單單號（有空格，如 P O 2 6 0 3 2 3 0 0 0 0 2 5）
    po_match = re.search(r"訂\s*單\s*單\s*號\s*[:：]?\s*(P\s*O[\d\s]+)", text)
    po_number = re.sub(r"\s+", "", po_match.group(1)).upper() if po_match else ""
    # 確保格式為 PO + digits
    po_clean = re.sub(r"[^A-Z0-9]", "", po_number)

    # 預計交貨日
    delivery_match = re.search(r"預計交貨日[：:]\s*(\d{4}/\d{2}/\d{2})", text)
    delivery_date  = delivery_match.group(1) if delivery_match else ""

    items = _extract_items(text)

    # 備註（格式：郵遞區號地址/姓名/手機）
    remark_match = re.search(r"備註[:：]\s*(.+)", text)
    remark = remark_match.group(1).strip() if remark_match else ""
    address, name, phone = _parse_remark(remark)

    if not (po_clean and delivery_date and items and address):
        return None

    return {
        "po_number":         po_clean,
        "delivery_date":     delivery_date,
        "items":              items,
        "address":           address,
        "name":              name,
        "phone":             phone,
    }


def _strip_prefix(raw_name: str) -> str:
    """去除「捷捷樂齡食品-」前綴"""
    return re.sub(r"^捷捷樂齡食品-", "", raw_name).strip()


def _extract_items(text: str) -> list[dict]:
    """逐列擷取品項貨號、品名與該品項自己的採購量。"""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    items: list[dict] = []
    row_pattern = re.compile(
        r"^(\d+)\s+(N1HF\d+)\s+(N1HF\d+)\s+(\d+(?:\.\d+)?)\s+\S+\s+"
        r"[\d,]+(?:\.\d+)?\s+[\d,]+(?:\.\d+)?(?:\s+[\d,]+(?:\.\d+)?)?$"
    )
    for index, line in enumerate(lines):
        match = row_pattern.match(line)
        if not match or index + 1 >= len(lines):
            continue
        name = lines[index + 1]
        # 下一行應為品名，不能是另一筆品項或結算列。
        if row_pattern.match(name) or re.match(r"^(稅金|含稅總額|備註)\b", name):
            continue
        items.append({
            "supplier_sku": match.group(2),
            "raw_product_name": name,
            "qty": int(float(match.group(4))),
        })
    return items


def _parse_remark(remark: str) -> tuple[str, str, str]:
    """
    解析備註欄：郵遞區號+地址/姓名/手機
    回傳 (address, name, phone)
    """
    # 以 / 分割
    parts = [p.strip() for p in remark.split("/")]
    if len(parts) >= 3:
        address = parts[0].strip()
        name    = parts[1].strip()
        phone   = parts[2].strip()
        return address, name, phone
    if len(parts) == 2:
        return parts[0], parts[1], ""
    return remark, "", ""


def _po_suffix(po: str) -> int:
    """取 PO 號碼尾碼數字，用於排序"""
    m = re.search(r"(\d+)$", po)
    return int(m.group(1)) if m else 0


# ── 輸出 ──────────────────────────────────────────────────────────────────

def _save_leage_xlsx(base_dir: Path, source_type: str, date_obj: datetime,
                     filename: str, orders: list, delivery_date: str) -> Path:
    from app.utils import output_path as op
    p = op(base_dir, source_type, date_obj, filename)

    # 工作表名稱：YYYY-MM-DD
    try:
        sheet_name = datetime.strptime(delivery_date, "%Y/%m/%d").strftime("%Y-%m-%d")
    except ValueError:
        sheet_name = delivery_date or date_obj.strftime("%Y-%m-%d")

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]   # Excel 工作表名稱最長 31 字元

    ws.append(["訂單編號", "收件人", "地址", "電話",
               "產品編號", "產品名稱", "數量", "單價", "備註"])
    for order in orders:
        for item in order["items"]:
            prod = item["product"]
            ws.append([
                order["order_id"],
                order["name"],
                order["address"],
                order["phone"],
                prod["品號"],
                item["prod_name"],
                item["qty"],
                int(float(prod.get("商品結帳價") or 0)),
                order["po_number"],
            ])
    wb.save(p)
    return p


def _save_leage_csv(base_dir: Path, source_type: str, date_obj: datetime,
                    filename: str, orders: list) -> Path:
    from app.utils import output_path as op
    p = op(base_dir, source_type, date_obj, filename)

    csv_header = [
        "收件人姓名", "收件人電話", "收件人手機", "收件人地址",
        "代收金額或到付", "件數", "品名(詳參數表)", "備註", "訂單編號",
        "希望配達時間((詳參數表))", "出貨日期(YYYY/MM/DD)", "預定配達日期(YYYY/MM/DD)",
        "溫層((詳參數表))", "尺寸((詳參數表))",
        "寄件人姓名", "寄件人電話", "寄件人手機", "寄件人地址",
        "保值金額(20001~10萬之間)-會產生額外費用", "品名說明", "是否列印",
        "是否捐贈", "統一編號", "手機載具", "愛心碼", "可刷卡(Y/N)", "手機支付(Y/N)",
    ]

    with open(p, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(csv_header)
        # 黑貓託運單日期一律以實際轉單當天為出貨日，隔天為預定配達日。
        ship_date_fmt = date_obj.strftime("%Y/%m/%d")
        ship_date_next = (date_obj + timedelta(days=1)).strftime("%Y/%m/%d")
        for order in orders:
            # 黑貓手機欄：去掉 leading 0
            phone_raw = str(order["phone"]).lstrip("0")
            writer.writerow([
                order["name"],
                "",                 # 收件人電話（空）
                phone_raw,          # 收件人手機
                order["address"],
                "",                 # 代收金額
                1,                  # 每張採購單固定一件託運，不跟商品數量連動
                1,                  # 品名
                1,                  # 備註
                order["order_id"],
                SENDER["deliver"],
                ship_date_fmt,
                ship_date_next,
                SENDER["temp"],
                SENDER["size"],
                SENDER["name"],
                SENDER["tel"],
                SENDER["mobile"],
                SENDER["address"],
                "", "", "", "", "", "", "", "", "",
            ])
    return p


def _fmt_date_yyyymmdd(date_str: str) -> str:
    """YYYY/MM/DD → YYYYMMDD"""
    return date_str.replace("/", "") if date_str else ""


def _date_plus1(yyyymmdd: str) -> str:
    try:
        d = datetime.strptime(yyyymmdd, "%Y%m%d") + timedelta(days=1)
        return d.strftime("%Y%m%d")
    except ValueError:
        return ""
