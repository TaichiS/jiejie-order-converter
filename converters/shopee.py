"""
converters/shopee.py
蝦皮訂單轉換器。
輸入：Order.toship.YYYYMMDD_YYYYMMDD.xlsx（工作表 orders）
輸出：鼎新-蝦皮{MMDD}.xlsx、運送及發票資料.xlsx

欄位定位：以「欄位標題」動態對應，避免蝦皮匯出格式新增/調整欄位時整體位移而失效。
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl
from openpyxl import Workbook

from converters.base import BaseConverter, ConversionResult, RowError, open_xlsx, SHOPEE_XLSX_PASSWORD
from app.utils import output_path


# 欄位用途 → 可能的標題名稱（依序比對，取第一個存在者）
COLUMN_NAMES: dict[str, list[str]] = {
    "order_id":     ["訂單編號"],
    "product":      ["商品名稱"],          # 完整商品名稱
    "prod_id":      ["商品ID"],            # 商品ID（品號）
    "option":       ["商品選項名稱"],       # 商品選項名稱（短品名/規格）
    "orig_price":   ["商品原價"],
    "act_price":    ["商品活動價格"],
    "qty":          ["數量"],
    "shipping_fee": ["買家支付運費"],
    "coin":         ["蝦幣折抵"],
    "coupon":       ["優惠券", "賣家負擔優惠券"],
    "order_date":   ["訂單成立日期"],
    "ship_date":    ["最晚出貨日期"],
    "address":      ["收件地址"],
    "contact":      ["蝦皮專線和包裹查詢碼 \n(請複製下方完整編號提供給您配合的物流商當做聯絡電話)"],
    "recipient":    ["收件者姓名"],
    "total":        ["買家總支付金額"],
    "buyer_note":   ["買家備註"],
}


def _resolve_columns(header: list) -> dict[str, int | None]:
    """依標題名稱建立欄位用途 → 欄位索引的對照表。找不到者為 None。"""
    norm = [str(h).strip() if h is not None else "" for h in header]
    idx: dict[str, int | None] = {}
    for key, names in COLUMN_NAMES.items():
        found = None
        for n in names:
            if n in norm:
                found = norm.index(n)
                break
        idx[key] = found
    return idx


class ShopeeConverter(BaseConverter):

    @property
    def source_type(self) -> str:
        return "shopee"

    def _validate(self, input_files: list[Path]) -> list[RowError]:
        errors = []
        xlsx = [f for f in input_files if f.suffix.lower() == ".xlsx"]
        if not xlsx:
            errors.append(RowError(0, "file", "", "找不到 .xlsx 訂單檔"))
            return errors

        try:
            wb = open_xlsx(xlsx[0], password=SHOPEE_XLSX_PASSWORD)
            if "orders" not in wb.sheetnames:
                errors.append(RowError(0, "sheet", xlsx[0].name, "工作表 'orders' 不存在",
                                       source_file=xlsx[0].name))
                return errors
            # 檢查關鍵欄位是否存在（標題動態定位）
            ws = wb["orders"]
            header = next(ws.iter_rows(max_row=1, values_only=True), None)
            if header is None:
                errors.append(RowError(0, "sheet", xlsx[0].name, "工作表為空",
                                       source_file=xlsx[0].name))
                return errors
            col = _resolve_columns(list(header))
            missing = [COLUMN_NAMES[k][0] for k in ("order_id", "product", "prod_id", "qty")
                       if col[k] is None]
            if missing:
                errors.append(RowError(0, "header", xlsx[0].name,
                                       f"缺少必要欄位：{', '.join(missing)}",
                                       source_file=xlsx[0].name))
        except Exception as e:
            errors.append(RowError(0, "file", xlsx[0].name, f"無法開啟：{e}",
                                   source_file=xlsx[0].name))
        return errors

    def _process(self, input_files: list[Path]) -> ConversionResult:
        order_file = next(f for f in input_files if f.suffix.lower() == ".xlsx")

        wb_in  = open_xlsx(order_file, password=SHOPEE_XLSX_PASSWORD)
        ws_in  = wb_in["orders"]
        rows   = list(ws_in.iter_rows(values_only=True))
        header = list(rows[0])

        # 依標題動態定位欄位
        col = _resolve_columns(header)
        c_order = col["order_id"]
        c_prod  = col["product"]
        c_pid   = col["prod_id"]
        c_opt   = col["option"]
        c_orig  = col["orig_price"]
        c_act   = col["act_price"]
        c_qty   = col["qty"]
        c_fee   = col["shipping_fee"]
        c_coin  = col["coin"]
        c_coup  = col["coupon"]
        c_order_date = col["order_date"]
        c_ship  = col["ship_date"]
        c_address = col["address"]
        c_contact = col["contact"]
        c_recipient = col["recipient"]
        c_total = col["total"]
        c_buyer_note = col["buyer_note"]

        def cell(row, c, default=None):
            return row[c] if c is not None and c < len(row) else default

        # 檔名日期取本批訂單最後一個「訂單成立日期」。
        # 例如訂單涵蓋 08/07～08/10，輸出為「鼎新-蝦皮0810.xlsx」。
        order_dates = [
            _date_text(cell(row, c_order_date))
            for row in rows[1:]
            if cell(row, c_order_date)
        ]
        order_dates = [d for d in order_dates if d]
        latest_order_date = max(order_dates) if order_dates else ""
        order_date = _batch_date_from_filename(order_file.name, latest_order_date)
        if not order_date:
            order_date = latest_order_date
        if not order_date and len(rows) > 1:
            order_date = _date_text(cell(rows[1], c_ship))

        # MMDD 供輸出檔名
        mmdd = order_date[4:8] if len(order_date) == 8 else "0000"

        # 建立輸出活頁簿
        wb_out = Workbook()
        ws_out = wb_out.active
        ws_out.title = "orders"
        ws_out.append(header)

        # 出貨資料：每張訂單一列，並加入買家備註作為「出貨備註」。
        shipping_headers = [
            "訂單編號",
            "收件地址",
            COLUMN_NAMES["contact"][0],
            "收件者姓名",
            "金額",
            "出貨備註",
        ]
        shipping_rows: list[list] = []
        shipping_seen: set[str] = set()
        for source_row in rows[1:]:
            source_order_id = str(cell(source_row, c_order) or "").strip()
            if not source_order_id or source_order_id in shipping_seen:
                continue
            shipping_seen.add(source_order_id)
            shipping_rows.append([
                source_order_id,
                cell(source_row, c_address, "") or "",
                cell(source_row, c_contact, "") or "",
                cell(source_row, c_recipient, "") or "",
                cell(source_row, c_total, "") or "",
                cell(source_row, c_buyer_note, "") or "",
            ])

        success_count = 0
        fail_count    = 0
        manual_count  = 0
        errors: list[RowError] = []
        seen_orders: set[str] = set()

        for row_idx, raw_row in enumerate(rows[1:], start=2):
            row = list(raw_row)
            # 跳過全空列（read_only 模式可能多讀 Excel dimension 尾端空列）
            if cell(row, c_order) is None and cell(row, c_prod) is None:
                continue
            order_id     = str(cell(row, c_order)) if cell(row, c_order) else ""
            full_name    = str(cell(row, c_prod))  if cell(row, c_prod)  else ""
            qty_order    = cell(row, c_qty) or 0

            # 1. 從完整品名提取短品名（去前綴、去 | 後面、正規化空格）
            short_name = _extract_short_name(full_name)

            # 2. 查品號資料（代碼前綴 → 精確品名 → 模糊比對 + 金額整除）
            act_price = float(cell(row, c_act)) if cell(row, c_act) else 0
            product, candidates = self._lookup_by_name(short_name, amount=act_price)
            if not product:
                errors.append(RowError(
                    row_idx, "商品名稱", full_name,
                    f"找不到「{short_name}」",
                    candidates=candidates,
                    source_file=order_file.name,
                ))
                fail_count += 1
                continue

            # 3. 轉換數值
            pack_count       = int(product.get("包數") or 1)
            checkout_price   = float(product.get("商品結帳價") or 0)
            new_qty          = int(qty_order) * pack_count

            if c_pid  is not None: row[c_pid]  = product["品號"]
            if c_opt  is not None: row[c_opt]  = product["品名"]
            if c_orig is not None: row[c_orig] = act_price        # 份數定價（原始訂單值）
            if c_act  is not None: row[c_act]  = checkout_price   # 每包單價
            if c_qty  is not None: row[c_qty]  = new_qty

            # 4. 同訂單第二列起費用清零（蝦幣、優惠券、運費只保留第一列）
            if order_id in seen_orders:
                if c_coin is not None: row[c_coin] = 0
                if c_coup is not None: row[c_coup] = 0
                if c_fee  is not None: row[c_fee]  = 0
            else:
                seen_orders.add(order_id)

            ws_out.append(row)
            success_count += 1

        # 儲存輸出
        out_file = output_path(
            self.output_dir.parent,
            self.source_type,
            _parse_date(order_date),
            f"鼎新-蝦皮{mmdd}.xlsx",
        )
        wb_out.save(out_file)

        shipping_wb = Workbook()
        shipping_ws = shipping_wb.active
        shipping_ws.title = "orders"
        shipping_ws.append(shipping_headers)
        for shipping_row in shipping_rows:
            shipping_ws.append(shipping_row)
        shipping_file = output_path(
            self.output_dir.parent,
            self.source_type,
            _parse_date(order_date),
            "運送及發票資料.xlsx",
        )
        shipping_wb.save(shipping_file)

        status = "completed" if fail_count == 0 else ("partial" if success_count > 0 else "failed")
        return ConversionResult(
            source_type   = self.source_type,
            success_count = success_count,
            fail_count    = fail_count,
            manual_count  = manual_count,
            order_date    = order_date,
            output_files  = [out_file, shipping_file],
            errors        = errors,
            status        = status,
        )


def _extract_short_name(full_name: str) -> str:
    """
    從完整商品名稱中提取短品名。
    例：「【捷捷寶寶粥】2-1 醬燒什錦豬肉 | 冷凍副食品 ...」→ 「2-1醬燒什錦豬肉」
    例：「寶寶副食品-2-D後元彩虹水餃(14M+)」→ 「2-D後元彩虹水餃」
    """
    # 移除品牌前綴 【xxx】
    name = re.sub(r"^【[^】]*】\s*", "", full_name)
    # 取 | 之前的部分
    name = name.split("|")[0]
    # 移除所有空白
    name = re.sub(r"\s+", "", name)
    # 移除中文類別前綴（如「寶寶副食品-」）
    name = re.sub(r"^[一-鿿＀-￯]+-", "", name)
    # 移除結尾月齡標示（如 (14M+)、(9M+)）
    name = re.sub(r"\(\d+[Mm]\+?\)$", "", name)
    return name.strip()


def _parse_date(order_date: str):
    from datetime import datetime
    try:
        return datetime.strptime(order_date, "%Y%m%d")
    except (ValueError, TypeError):
        return datetime.today()


def _date_text(value) -> str:
    """將 Excel 日期或日期字串正規化為 YYYYMMDD。"""
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y%m%d")
    digits = re.sub(r"[^0-9]", "", str(value)[:10])
    return digits[:8] if len(digits) >= 8 else ""


def _batch_date_from_filename(filename: str, fallback_date: str = "") -> str:
    """從蝦皮原始檔名取批次截止日；支援 YYYYMMDD 與 MMDD。"""
    full_dates = re.findall(r"(?<!\d)(20\d{6})(?!\d)", filename)
    if full_dates:
        return max(full_dates)

    short_dates = re.findall(r"(?<!\d)(\d{4})(?!\d)", Path(filename).stem)
    if short_dates and len(fallback_date) == 8:
        candidate = fallback_date[:4] + short_dates[-1]
        try:
            from datetime import datetime
            datetime.strptime(candidate, "%Y%m%d")
            return candidate
        except ValueError:
            pass
    return ""
