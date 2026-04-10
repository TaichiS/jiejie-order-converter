"""
converters/a1baby.py
婦幼展訂單轉換器。
輸入：MMDD.xlsx（主檔）+ MMDD-1.xlsx（明細）
輸出：全家取貨_MMDD.xlsx / 黑貓配送_MMDD.xlsx / 黑貓託運單_MMDD.xlsx / 黑貓印製托運單_MMDD.csv
"""
from __future__ import annotations

import csv
import io
import re
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl import Workbook

from converters.base import BaseConverter, ConversionResult, RowError
from app.utils import output_path

SENDER = {
    "name":    "濬詮股份有限公司",
    "tel":     "05-5870993",
    "mobile":  "0909-870-993",
    "address": "雲林縣莿桐鄉大美村溪美22-3號1樓",
    "temp":    "3",    # 冷凍
    "size":    "2",
    "deliver": "4",    # 不指定
}

# 略過的 POS 品名關鍵字（折扣組合、運費、贈品等）
SKIP_KEYWORDS = (
    "自選", "組合", "黑貓", "宅配運費", "外島", "本島",
    "滿", "贈", "冰箱磁鐵", "感謝",
)

# 2-D 產品：以商品結帳價為 key
D_PRICE_MAP = {119: "2-D01", 129: "2-D02", 149: "2-D03"}

# 品名 key 擷取 regex（順序重要：2-M\d+ 必須在 2-M[S]? 之前）
KEY_PATTERN = re.compile(
    r"(1P-\d+|0-\d+|1-\d+|2-S\d+|2-M\d+|2-M[S]?|2-F|2-D0[123]|2-D|2-\d+|3-3|E52\d+)"
)


class A1BabyConverter(BaseConverter):

    @property
    def source_type(self) -> str:
        return "a1baby"

    def _validate(self, input_files: list[Path]) -> list[RowError]:
        errors = []
        xlsx = [f for f in input_files if f.suffix.lower() == ".xlsx"]
        main_f, detail_f = _find_pair(xlsx)
        if not main_f:
            errors.append(RowError(0, "file", "", "找不到 MMDD.xlsx 主檔"))
        if not detail_f:
            errors.append(RowError(0, "file", "", "找不到 MMDD-1.xlsx 明細檔"))
        return errors

    def _process(self, input_files: list[Path]) -> ConversionResult:
        xlsx = [f for f in input_files if f.suffix.lower() == ".xlsx"]
        main_f, detail_f = _find_pair(xlsx)
        mmdd = re.search(r"(\d{4})", main_f.stem)
        mmdd_str = mmdd.group(1) if mmdd else "0000"

        # 載入主檔
        wb_main = openpyxl.load_workbook(main_f, read_only=True, data_only=True)
        ws_main = wb_main.active
        main_rows = list(ws_main.iter_rows(values_only=True))
        main_header = list(main_rows[0])

        # 欄位索引（動態偵測）
        hi = {v: i for i, v in enumerate(main_header) if v}
        IDX_INVOICE  = hi.get("發票號碼", 1)
        IDX_CHECKOUT = hi.get("結帳時間", 3)
        IDX_RAW_NO   = hi.get("原始單號", 4)
        IDX_PAYMENT  = hi.get("支付模組", 13)
        IDX_NAME     = hi.get("顧客姓名", 18)
        IDX_PHONE    = hi.get("顧客電話", 19)
        IDX_REMARK   = hi.get("訂單標籤與備註", 20)
        IDX_TOTAL    = hi.get("發票金額", 12)

        # 載入明細檔
        wb_det = openpyxl.load_workbook(detail_f, read_only=True, data_only=True)
        ws_det = wb_det.active
        det_rows = list(ws_det.iter_rows(values_only=True))
        det_header = list(det_rows[0])
        dh = {v: i for i, v in enumerate(det_header) if v}
        D_NAME    = dh.get("商品名稱", 0)
        D_INVOICE = dh.get("發票號碼", 2)
        D_AMOUNT  = dh.get("發票金額", 9)

        # 建立品號 key 對照表
        # _code_map 由 base._load_reference() 建立（_CODE_RE 已涵蓋所有代碼前綴）

        # 建立 invoice → [detail rows]
        details: dict[str, list] = defaultdict(list)
        for row in det_rows[1:]:
            inv = str(row[D_INVOICE]).strip() if row[D_INVOICE] else ""
            if inv:
                details[inv].append(row)

        # 取訂單日期（從第一列有效資料）
        order_date = ""
        for r in main_rows[1:]:
            checkout = str(r[IDX_CHECKOUT]) if r[IDX_CHECKOUT] else ""
            if checkout and checkout != "--":
                order_date = checkout[:10].replace("/", "").replace("-", "")
                break

        # 輸出容器
        rows_cat  = []   # 全家取貨
        rows_black= []   # 黑貓配送
        rows_slip = []   # 黑貓託運單（一訂單一列）
        errors: list[RowError] = []
        success_count = fail_count = manual_count = 0

        for row_idx, main_row in enumerate(main_rows[1:], start=2):
            name = str(main_row[IDX_NAME]).strip() if main_row[IDX_NAME] else ""
            if name == "--":
                continue   # 現場零售，跳過

            invoice  = str(main_row[IDX_INVOICE]).strip() if main_row[IDX_INVOICE] else ""
            raw_no   = main_row[IDX_RAW_NO] or ""
            phone    = str(main_row[IDX_PHONE] or "").strip()
            remark   = str(main_row[IDX_REMARK] or "").strip()
            total    = float(main_row[IDX_TOTAL] or 0)
            payment  = _clean_payment(str(main_row[IDX_PAYMENT] or ""))

            ship_date, address = _parse_remark(remark)
            is_family = "全家" in remark

            # 整理出貨備註
            note = _build_note(remark)

            # 計算折扣總額（明細檔中負數金額的絕對值加總）
            discount_total = int(sum(
                abs(float(r[D_AMOUNT] or 0))
                for r in details.get(invoice, [])
                if float(r[D_AMOUNT] or 0) < 0
            ))

            # 明細品項
            order_items = []
            order_errors = []
            is_first_item = True
            for det_row in details.get(invoice, []):
                det_name   = str(det_row[D_NAME] or "").strip()
                det_amount = float(det_row[D_AMOUNT] or 0)

                # 略過折扣、運費、贈品
                if det_amount <= 0:
                    continue
                if any(k in det_name for k in SKIP_KEYWORDS):
                    continue
                # 已取品項：計入但標記需人工確認
                already_taken = det_name.startswith("【已取】")
                if already_taken:
                    det_name = det_name[4:].strip()  # 【已取】= 4字元
                    manual_count += 1

                # 查品號
                product, matched_key = self._match_product(det_name, det_amount)
                if not product:
                    _, candidates = self._lookup_by_name(det_name, amount=det_amount)
                    order_errors.append(RowError(
                        row_idx, "商品名稱", det_name,
                        f"找不到「{det_name}」（金額：{det_amount}）",
                        candidates=candidates,
                        source_file=detail_f.name,
                    ))
                    continue

                checkout_price = float(product.get("商品結帳價") or 1)
                if checkout_price <= 0:
                    order_errors.append(RowError(row_idx, "商品結帳價", product["品名"],
                                                  "商品結帳價為 0，無法計算數量",
                                                  source_file=detail_f.name))
                    continue

                qty = det_amount / checkout_price
                if qty != int(qty):
                    manual_count += 1
                    qty = round(qty)
                else:
                    qty = int(qty)

                order_items.append({
                    "invoice":      invoice,
                    "order_no":     raw_no,
                    "name":         name,
                    "phone":        str(phone).zfill(0),
                    "address":      address,
                    "prod_no":      product["品號"],
                    "prod_name":    product["品名"],
                    "qty":          qty,
                    "unit_price":   int(checkout_price),
                    "discount":     discount_total if is_first_item else 0,
                    "note":         note,
                    "payment":      payment,
                    "ship_date":    ship_date,
                    "is_family":    is_family,
                    "already_taken": already_taken,
                })
                is_first_item = False

            errors.extend(order_errors)
            if order_errors:
                fail_count += len(order_errors)
            if order_items:
                success_count += len(order_items)
                for item in order_items:
                    row_out = [
                        item["order_no"],           # 訂單號碼
                        item["name"],               # 收件人
                        item["address"],            # 完整地址
                        item["phone"],              # 收件人電話
                        item["invoice"],            # 發票號碼
                        item["prod_no"],            # 商品貨號
                        item["prod_name"],          # 商品名稱
                        item["qty"],                # 數量
                        item["unit_price"],         # 商品結帳價
                        0,                          # 商品折扣優惠
                        item["discount"],           # 商品折扣金額（負數加總取絕對值，僅第一列）
                        0,                          # 點數折現分攤
                        item["note"],               # 出貨備註
                        None,                       # 送貨編號
                        item["payment"],            # 付款方式
                    ]
                    if item["is_family"]:
                        rows_cat.append((item, row_out))
                    else:
                        rows_black.append((item, row_out))

                # 黑貓託運單（每訂單一列，非全家）
                if not order_items[0]["is_family"]:
                    rows_slip.append({
                        "no":       len(rows_slip) + 1,
                        "name":     order_items[0]["name"],
                        "phone":    order_items[0]["phone"],
                        "address":  order_items[0]["address"],
                        "ship_date": order_items[0]["ship_date"],
                    })

        # 寫出四份檔案
        output_files = []
        date_obj = _parse_order_date(order_date)

        # 1. 全家取貨
        if rows_cat:
            p = _save_xlsx(
                self.output_dir.parent, self.source_type, date_obj,
                f"全家取貨_{mmdd_str}.xlsx",
                f"全家冷凍取貨訂單 {mmdd_str}",
                [r[1] for r in rows_cat],
            )
            output_files.append(p)

        # 2. 黑貓配送
        if rows_black:
            p = _save_xlsx(
                self.output_dir.parent, self.source_type, date_obj,
                f"黑貓配送_{mmdd_str}.xlsx",
                f"黑貓冷凍配送訂單 {mmdd_str}",
                [r[1] for r in rows_black],
            )
            output_files.append(p)

        # 3. 黑貓託運單
        if rows_slip:
            p = _save_slip_xlsx(
                self.output_dir.parent, self.source_type, date_obj,
                f"黑貓託運單_{mmdd_str}.xlsx", mmdd_str, rows_slip
            )
            output_files.append(p)

        # 4. 黑貓印製托運單 CSV
        if rows_black:
            p = _save_print_csv(
                self.output_dir.parent, self.source_type, date_obj,
                f"黑貓印製托運單_{mmdd_str}.csv", rows_black
            )
            output_files.append(p)

        status = "completed" if fail_count == 0 else ("partial" if success_count > 0 else "failed")
        return ConversionResult(
            source_type   = self.source_type,
            success_count = success_count,
            fail_count    = fail_count,
            manual_count  = manual_count,
            order_date    = order_date[:8] if len(order_date) >= 8 else "",
            output_files  = output_files,
            errors        = errors,
            status        = status,
        )

    def _match_product(self, pos_name: str, amount: float) -> tuple[dict | None, str]:
        """
        從 POS 品名 + 金額 找對應的品號資料。
        回傳 (product_dict, matched_key)

        查找順序：
        1. 後元/水餃等 2-D 系列：金額整除或關鍵字
        2. _extract_key 取代碼 → self._code_map（base 已建，涵蓋所有前綴）
        3. 1P 系列：掃 _product_map 找同 key 的 150g/200g 兩筆
        4. 無代碼：精確品名 → E52 金額比對
        """
        # 1. 後元/水餃系列（2-D）
        if any(k in pos_name for k in ("後元", "水餃", "饅頭", "蘿蔔糕")):
            for p in [119, 129, 149]:
                if amount % p == 0:
                    code = D_PRICE_MAP[p]
                    prod = self._code_map.get(code)
                    if prod:
                        return prod, code
            for kw, code in (("水餃", "2-D01"), ("饅頭", "2-D02"), ("蘿蔔糕", "2-D03")):
                if kw in pos_name:
                    prod = self._code_map.get(code)
                    if prod:
                        return prod, code

        # 2. 提取代碼前綴
        key = _extract_key(pos_name)
        if not key:
            if pos_name in self._product_map:
                return self._product_map[pos_name], pos_name
            for name, prod in self._product_map.items():
                if not str(prod.get("品號", "")).startswith("E52"):
                    continue
                cp = float(prod.get("商品結帳價") or 0)
                if cp > 0 and amount % cp == 0:
                    return prod, name
            return None, ""

        # 3. 1P 系列：150g / 200g 各有一筆，需掃 _product_map
        if key.startswith("1P-"):
            all_1p = [p for n, p in self._product_map.items() if _extract_key(n) == key]
            if not all_1p:
                return None, key
            if "150" in pos_name:
                chosen = [p for p in all_1p if "150g" in p.get("品名", "")]
            elif "200" in pos_name:
                chosen = [p for p in all_1p if "200g" in p.get("品名", "")]
            else:
                chosen = [p for p in all_1p
                          if float(p.get("商品結帳價") or 1) > 0
                          and amount % float(p.get("商品結帳價") or 1) == 0]
            return (chosen[0] if chosen else all_1p[0]), key

        # 4. _code_map 直接取（涵蓋 0-N, 1-NN, 2-N, 2-SN, 2-MN, 3-N）
        prod = self._code_map.get(key)
        if prod:
            return prod, key

        # 5. 精確品名比對（如 2-M寶貝義大利麵 等不在 _code_map 的品項）
        if pos_name in self._product_map:
            return self._product_map[pos_name], pos_name

        return None, key


# ── 工具函式 ──────────────────────────────────────────────────────────────

def _find_pair(xlsx_files: list[Path]) -> tuple[Path | None, Path | None]:
    """找 MMDD.xlsx 與 MMDD-1.xlsx 配對"""
    detail_files = [f for f in xlsx_files
                    if re.search(r"-1\.xlsx$", f.name, re.IGNORECASE)]
    for detail in detail_files:
        main_name = re.sub(r"-1\.xlsx$", ".xlsx", detail.name, flags=re.IGNORECASE)
        for f in xlsx_files:
            if f.name == main_name:
                return f, detail
    return None, None


def _extract_key(name: str) -> str:
    """從品名中提取品號 key"""
    m = KEY_PATTERN.search(name)
    return m.group(1) if m else ""


def _parse_remark(remark: str) -> tuple[str, str]:
    """
    解析訂單標籤與備註，回傳 (ship_date_str, address)。
    格式多樣：'3/30\\n地址', '3/5出貨、地址', '出貨3/5,地址：地址'
    """
    text = remark.replace("\n", "、").replace("，", "、").replace(",", "、")
    text = re.sub(r"出貨", "", text)
    text = re.sub(r"地址：", "", text)
    text = re.sub(r"地址:", "", text)
    parts = re.split(r"[、；;]", text)
    parts = [p.strip() for p in parts if p.strip()]

    date_str = ""
    address  = ""
    for part in parts:
        if re.match(r"^\d{1,2}/\d{1,2}$", part):
            date_str = part
        elif date_str == "" and re.match(r"^\d{1,2}/\d{1,2}", part):
            date_str = re.match(r"(\d{1,2}/\d{1,2})", part).group(1)
        elif re.search(r"[市縣鄉鎮區]", part) and not re.match(r"^\d{1,2}/", part):
            address = part
        elif "全家" in part:
            address = part   # 全家店號
        elif part and not address:
            if not re.match(r"^\d{1,2}/\d{1,2}", part):
                address = part

    return date_str, address


def _build_note(remark: str) -> str:
    """建立出貨備註（取原始備註簡化版）"""
    # 保留日期部分
    text = remark.replace("\n", " ").strip()
    if len(text) > 50:
        text = text[:50] + "..."
    return text


def _clean_payment(payment: str) -> str:
    """簡化付款方式（去括號說明）"""
    m = re.match(r"^([^(（]+)", payment)
    return m.group(1).strip() if m else payment


def _save_xlsx(base_dir: Path, source_type: str, date_obj: datetime,
               filename: str, title: str, rows: list) -> Path:
    """儲存帶標題列的 xlsx"""
    from app.utils import output_path as op
    p = op(base_dir, source_type, date_obj, filename)
    wb = Workbook()
    ws = wb.active
    ws.append([title] + [None] * 14)
    ws.append(["訂單號碼", "收件人", "完整地址", "收件人電話號碼", "發票號碼",
               "商品貨號", "商品名稱", "數量", "商品結帳價", "商品折扣優惠",
               "商品折扣金額", "點數折現分攤", "出貨備註", "送貨編號", "付款方式"])
    for row in rows:
        ws.append(row)
    wb.save(p)
    return p


def _save_slip_xlsx(base_dir: Path, source_type: str, date_obj: datetime,
                    filename: str, mmdd_str: str, slips: list) -> Path:
    from app.utils import output_path as op
    p = op(base_dir, source_type, date_obj, filename)
    wb = Workbook()
    ws = wb.active
    ws.append([f"黑貓宅急便 冷凍宅配託運單 {mmdd_str}"] + [None] * 8)
    ws.append(["編號", "收件人姓名", "收件人電話", "收件人地址",
               "出貨日期", "溫層", "件數", "代收金額", "備註"])
    for i, slip in enumerate(slips, 1):
        ws.append([i, slip["name"], slip["phone"], slip["address"],
                   slip["ship_date"], "冷凍", 1, None, None])
    wb.save(p)
    return p


def _save_print_csv(base_dir: Path, source_type: str, date_obj: datetime,
                    filename: str, rows_black: list) -> Path:
    from app.utils import output_path as op
    p = op(base_dir, source_type, date_obj, filename)

    # 依訂單去重（每訂單一列托運單）
    seen = {}
    for item, _ in rows_black:
        inv = item["invoice"]
        if inv not in seen:
            seen[inv] = item

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
        for item in seen.values():
            # 出貨日期轉成 YYYY/MM/DD
            ship_date_fmt  = _format_ship_date(item["ship_date"], date_obj)
            ship_date_next = _format_ship_date_next(ship_date_fmt)
            writer.writerow([
                item["name"],           # 收件人姓名
                "",                     # 收件人電話（空）
                item["phone"],          # 收件人手機
                item["address"],        # 收件人地址
                "",                     # 代收金額
                1,                      # 件數
                1,                      # 品名
                1,                      # 備註
                item["invoice"],        # 訂單編號
                SENDER["deliver"],      # 希望配達時間
                ship_date_fmt,          # 出貨日期
                ship_date_next,         # 預定配達日期
                SENDER["temp"],         # 溫層
                SENDER["size"],         # 尺寸
                SENDER["name"],         # 寄件人姓名
                SENDER["tel"],          # 寄件人電話
                SENDER["mobile"],       # 寄件人手機
                SENDER["address"],      # 寄件人地址
                "", "", "", "", "", "", "", "", "",
            ])
    return p


def _format_ship_date(ship_date: str, base_date: datetime) -> str:
    """把 M/D 格式轉成 YYYY/MM/DD（依訂單年份）"""
    m = re.match(r"(\d{1,2})/(\d{1,2})", ship_date)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        year = base_date.year
        try:
            return f"{year}/{month:02d}/{day:02d}"
        except ValueError:
            pass
    return ""


def _format_ship_date_next(ship_date_fmt: str) -> str:
    """出貨日期 + 1 天"""
    try:
        d = datetime.strptime(ship_date_fmt, "%Y/%m/%d") + timedelta(days=1)
        return d.strftime("%Y/%m/%d")
    except ValueError:
        return ""


def _parse_order_date(order_date: str) -> datetime:
    try:
        return datetime.strptime(order_date[:8], "%Y%m%d")
    except (ValueError, TypeError):
        return datetime.today()
