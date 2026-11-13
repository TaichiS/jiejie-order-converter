"""樂齡官網：讀取前台 Sales 格式並輸出鼎新與三種物流檔。"""
from __future__ import annotations

import csv
import math
import re
from collections import OrderedDict
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl import Workbook

from app.utils import output_path
from converters.base import BaseConverter, ConversionResult, RowError

COD_FEE_SKU = "F59900003"
COD_FEE_PRICE = 30
FREIGHT_SKU = "F59900001"
FREIGHT_NAME = "運費"

# 樂齡官網前台活動貨號 → 鼎新品號
SKU_ALIASES = {
    "N1HF00003340999": "E52990008",
}

BLACKCAT_SENDER = {
    "name": "賦寶股份有限公司", "tel": "05-5870993",
    "mobile": "0909-870-993", "address": "雲林縣莿桐鄉大美村溪美22-3號1樓",
    "temp": "3", "size": "2", "deliver": "4",
}
FAMILY_SENDER = {**BLACKCAT_SENDER, "name": "濬詮股份有限公司"}

OUTPUT_HEADER = [
    "訂單號碼", "收件人", "完整地址", "收件人電話號碼", "發票號碼",
    "商品貨號", "商品名稱", "數量", "商品結帳價", "商品折扣優惠",
    "商品折扣金額", "點數折現分攤", "出貨備註", "送貨編號", "付款方式",
    "發票日期",
]

CSV_HEADER = [
    "收件人姓名", "收件人電話", "收件人手機", "收件人地址",
    "代收金額或到付", "件數", "品名(詳參數表)", "備註", "訂單編號",
    "希望配達時間((詳參數表))", "出貨日期(YYYY/MM/DD)", "預定配達日期(YYYY/MM/DD)",
    "溫層((詳參數表))", "尺寸((詳參數表))",
    "寄件人姓名", "寄件人電話", "寄件人手機", "寄件人地址",
    "保值金額(20001~10萬之間)-會產生額外費用", "品名說明", "是否列印",
    "是否捐贈", "統一編號", "手機載具", "愛心碼", "可刷卡(Y/N)", "手機支付(Y/N)",
]


class A1LeageConverter(BaseConverter):
    @property
    def source_type(self) -> str:
        return "a1leage"

    def _validate(self, input_files: list[Path]) -> list[RowError]:
        files = [f for f in input_files if f.suffix.lower() == ".xlsx"]
        if not files:
            return [RowError(0, "file", "", "找不到 .xlsx 訂單檔")]
        try:
            wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
            if "Sales" not in wb.sheetnames:
                return [RowError(
                    0, "sheet", files[0].name,
                    "工作表 'Sales' 不存在，請上傳與前台相同格式的訂單檔",
                    source_file=files[0].name,
                )]
        except Exception as exc:
            return [RowError(0, "file", files[0].name, f"無法開啟：{exc}",
                             source_file=files[0].name)]
        return []

    def _process(self, input_files: list[Path]) -> ConversionResult:
        order_file = next(f for f in input_files if f.suffix.lower() == ".xlsx")
        wb_in = openpyxl.load_workbook(order_file, read_only=True, data_only=True)
        ws_in = wb_in["Sales"]
        all_rows = [r for r in ws_in.iter_rows(values_only=True)
                    if any(v is not None for v in r)]
        cm = {str(v).strip(): i for i, v in enumerate(all_rows[0]) if v is not None}
        sku_map = {
            v.get("品號", "").strip(): v for v in self._product_map.values()
            if v.get("品號", "").strip()
        }

        date_obj = datetime.today()
        order_date = date_obj.strftime("%Y%m%d")
        label = _date_label(order_file.stem, date_obj)
        invoice_date = _invoice_date_from_label(label, date_obj)

        orders: OrderedDict[str, dict] = OrderedDict()
        errors: list[RowError] = []
        success_count = 0
        fail_count = 0

        for row_idx, raw in enumerate(all_rows[1:], start=2):
            def cell(name: str, default=None):
                idx = cm.get(name)
                return raw[idx] if idx is not None and idx < len(raw) else default

            order_id = str(cell("訂單號碼") or "").strip()
            if not order_id:
                continue
            raw_sku = str(cell("商品貨號") or "").strip()
            raw_name = str(cell("商品名稱") or "").strip()
            if not raw_sku:  # 折扣彙總等非商品列
                continue

            lookup_sku = SKU_ALIASES.get(raw_sku, raw_sku)
            product = sku_map.get(lookup_sku)
            if not product:
                errors.append(RowError(
                    row_idx, "商品貨號", raw_sku,
                    f"品號 {raw_sku!r}（{raw_name}）在樂齡官網品號資料中找不到",
                    source_file=order_file.name,
                ))
                fail_count += 1
                continue

            rcv_name = str(cell("收件人") or "").strip()
            rcv_phone = _normalize_phone(str(cell("收件人電話號碼") or ""))
            family_no = str(cell("全家服務編號 / 7-11 店號") or "").strip()
            is_family = bool(family_no)
            shipping = str(cell("送貨方式") or "").strip()
            is_cod = ("貨到付款" in shipping) and not is_family
            address = None if is_family else _parse_front_address(str(cell("完整地址") or ""))
            order_note = str(cell("訂單備註") or "").strip()
            ship_note = str(cell("出貨備註") or "").strip()
            note = "；".join(n for n in [order_note, ship_note] if n) or None
            freight = _num(cell("運費"))
            input_qty = int(_num(cell("數量") or 1) or 1)
            input_price = _num(cell("商品結帳價") or 0)

            pack_count = int(_num(product.get("包數") or 0) or 0)
            if pack_count > 0:
                output_qty = input_qty * pack_count
                ref_price = _num(product.get("商品結帳價") or 0)
                unit_price = 0.0 if input_price == 0 else (
                    ref_price if ref_price else input_price / pack_count
                )
            else:
                output_qty = input_qty
                ref_price = _num(product.get("份數價格") or 0)
                unit_price = 0.0 if input_price == 0 else (
                    ref_price if ref_price else input_price
                )

            # 原單 D 欄「優惠折扣」對應鼎新 J 欄「商品折扣優惠」。
            discount_offer = int(_num(cell("優惠折扣")))
            # 鼎新 K 欄須包含商品折扣、全單折扣與購物金分攤。
            product_discount = int(
                _num(cell("商品折扣金額"))
                + _num(cell("全單折扣金額"))
                + _num(cell("折抵購物金分攤"))
            )
            shopping_credit = int(_num(cell("折抵購物金")))
            shopping_credit_share = int(_num(cell("折抵購物金分攤")))
            points = int(_num(cell("點數折現分攤")))
            invoice_no = str(cell("發票號碼") or "").strip() or None

            if order_id not in orders:
                orders[order_id] = {
                    "rcv_name": rcv_name, "rcv_phone": rcv_phone,
                    "address": address, "family_no": family_no,
                    "is_family": is_family, "is_cod": is_cod,
                    "note": note, "freight": freight,
                    "shopping_credit": 0,
                    "shopping_credit_allocated": 0, "rows": [],
                }
            else:
                if note and not orders[order_id]["note"]:
                    orders[order_id]["note"] = note
                if freight > 0:
                    orders[order_id]["freight"] = freight

            orders[order_id]["rows"].append({
                "order_id": order_id, "rcv_name": rcv_name, "address": address,
                "rcv_phone": rcv_phone, "invoice_no": invoice_no,
                "sku": product.get("品號", raw_sku),
                "name": product.get("品名", raw_name), "qty": output_qty,
                "price": unit_price, "discount_offer": discount_offer,
                "discount_amount": product_discount, "points": points,
                "note": note, "invoice_date": invoice_date,
            })
            orders[order_id]["shopping_credit"] += shopping_credit
            orders[order_id]["shopping_credit_allocated"] += shopping_credit_share
            success_count += 1

        if not orders:
            return ConversionResult(
                source_type=self.source_type, fail_count=fail_count or len(errors),
                order_date=order_date, errors=errors, status="failed",
            )

        xlsx_rows: list[dict] = []
        normal_orders: list[dict] = []
        cod_orders: list[dict] = []
        family_orders: list[dict] = []

        for order_id, order in orders.items():
            # 新版原檔已有每項購物金分攤，僅在未提供或有尾差時補足差額。
            remaining_credit = max(
                order["shopping_credit"] - order["shopping_credit_allocated"], 0
            )
            _allocate_shopping_credit(order["rows"], remaining_credit)
            xlsx_rows.extend(order["rows"])
            first = order["rows"][0]
            if order["freight"] > 0:
                xlsx_rows.append({
                    "order_id": order_id, "rcv_name": order["rcv_name"],
                    "address": order["address"], "rcv_phone": order["rcv_phone"],
                    "invoice_no": first["invoice_no"], "sku": FREIGHT_SKU,
                    "name": FREIGHT_NAME, "qty": 1, "price": order["freight"],
                    "discount_offer": 0, "discount_amount": 0, "points": 0,
                    "note": None, "invoice_date": invoice_date,
                })
            if order["is_cod"]:
                xlsx_rows.append({
                    "order_id": order_id, "rcv_name": order["rcv_name"],
                    "address": order["address"], "rcv_phone": order["rcv_phone"],
                    "invoice_no": first["invoice_no"], "sku": COD_FEE_SKU,
                    "name": "黑貓貨到付款手續費", "qty": 1, "price": COD_FEE_PRICE,
                    "discount_offer": 0, "discount_amount": 0, "points": 0,
                    "note": None, "invoice_date": invoice_date,
                })

            subtotal = (
                sum(r["qty"] * r["price"] for r in order["rows"])
                + order["freight"]
            )
            discounts = sum(r["discount_amount"] + r["points"] for r in order["rows"])
            shipping_order = {
                "order_id": order_id, "rcv_name": order["rcv_name"],
                "rcv_phone": order["rcv_phone"], "address": order["address"],
                "family_no": order["family_no"], "note": order["note"],
                "cod": int(round(subtotal - discounts + COD_FEE_PRICE)) if order["is_cod"] else "",
            }
            if order["is_family"]:
                family_orders.append(shipping_order)
            elif order["is_cod"]:
                cod_orders.append(shipping_order)
            else:
                normal_orders.append(shipping_order)

        output_files = [
            _write_xlsx(self.output_dir.parent, self.source_type, date_obj,
                        f"鼎新-樂齡{label}.xlsx", xlsx_rows),
            _write_shipping_csv(self.output_dir.parent, self.source_type, date_obj,
                                f"黑貓託運-樂齡{label}.csv", normal_orders,
                                BLACKCAT_SENDER, "normal"),
            _write_shipping_csv(self.output_dir.parent, self.source_type, date_obj,
                                f"黑貓到付-樂齡{label}.csv", cod_orders,
                                BLACKCAT_SENDER, "cod"),
            _write_shipping_csv(self.output_dir.parent, self.source_type, date_obj,
                                f"全家託運-樂齡{label}.csv", family_orders,
                                FAMILY_SENDER, "family"),
        ]
        return ConversionResult(
            source_type=self.source_type, success_count=success_count,
            fail_count=fail_count, order_date=order_date,
            output_files=output_files, errors=errors,
            status="completed" if fail_count == 0 else "partial",
        )


def _write_xlsx(base_dir: Path, source_type: str, date_obj: datetime,
                filename: str, rows: list[dict]) -> Path:
    path = output_path(base_dir, source_type, date_obj, filename)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(OUTPUT_HEADER)
    for row in rows:
        ws.append([
            row["order_id"], row["rcv_name"], row["address"], row["rcv_phone"],
            row["invoice_no"], row["sku"], row["name"], row["qty"], row["price"],
            row["discount_offer"], row["discount_amount"], row["points"],
            row["note"], None, None, row["invoice_date"],
        ])
    wb.save(path)
    return path


def _write_shipping_csv(base_dir: Path, source_type: str, date_obj: datetime,
                        filename: str, orders: list[dict], sender: dict,
                        route: str) -> Path:
    path = output_path(base_dir, source_type, date_obj, filename)
    ship_date = date_obj.strftime("%Y/%m/%d")
    next_date = (date_obj + timedelta(days=1)).strftime("%Y/%m/%d")
    with open(path, "w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(CSV_HEADER)
        for order in orders:
            if route == "family":
                address = order["family_no"]
                remark = f"全家店號：{order['family_no']}"
            else:
                address = order["address"]
                remark = order["note"] or 1
            writer.writerow([
                order["rcv_name"], "", order["rcv_phone"], address,
                order["cod"] if route == "cod" else "", 1, 1, remark,
                order["order_id"], sender["deliver"], ship_date, next_date,
                sender["temp"], sender["size"], sender["name"], sender["tel"],
                sender["mobile"], sender["address"],
                "", "", "", "", "", "", "", "", "",
            ])
    return path


def _date_label(stem: str, date_obj: datetime) -> str:
    match = re.search(r"(\d{4}(?:-\d{4})?)", stem)
    return match.group(1) if match else date_obj.strftime("%m%d")


def _invoice_date_from_label(label: str, date_obj: datetime) -> str:
    end = label.split("-")[-1]
    if re.fullmatch(r"\d{4}", end):
        return f"{date_obj.year}{end}"
    return date_obj.strftime("%Y%m%d")


def _parse_front_address(raw: str) -> str | None:
    if not raw:
        return None
    value = re.sub(r"^台灣\s*", "", raw.strip())
    value = re.sub(r"^\d{3,6}\s*", "", value)
    value = value.replace(" ", "")
    return value or None


def _normalize_phone(raw: str) -> str:
    value = re.sub(r"\s+", "", raw.strip())
    if value.startswith("+886"):
        value = "0" + value[4:]
    elif value.startswith("886") and len(value) > 9:
        value = "0" + value[3:]
    return value[:10] if len(value) >= 10 else value


def _num(value) -> float:
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return 0.0


def _allocate_shopping_credit(rows: list[dict], amount: int) -> None:
    """依每項折扣後金額比例分攤訂單購物金，並確保整數合計不變。"""
    if amount <= 0 or not rows:
        return
    weights = [max(0.0, row["qty"] * row["price"] - row["discount_amount"])
               for row in rows]
    total = sum(weights)
    if total <= 0:
        rows[0]["discount_amount"] += amount
        return
    exact = [amount * weight / total for weight in weights]
    allocated = [math.floor(value + 0.5) for value in exact]
    # 官網既有格式是各列四捨五入，尾差統一補到折扣後金額最大的商品。
    allocated[max(range(len(rows)), key=lambda i: weights[i])] += amount - sum(allocated)
    for row, share in zip(rows, allocated):
        row["discount_amount"] += share
