"""
converters/licai.py
麗兒采家嬰童館採購單轉換器。

輸入：個別門市 xlsx（由 split_orders.py 拆分後的檔案）
  Row 0: ('門市：短名', None, '門市資訊：地址 / 電話', ...)
  Row 1: ('單號：UWXXXXXX', None, '採購日期：YYYY/M/D', ...)
  Row 2: 空列
  Row 3+: (序號, 條碼, 品名, 品牌, 市價, 數量)

輸出：MMDD 麗采{門市短名}.xlsx
  欄位：訂單編號, 收件人, 地址, 電話, 產品編號, 產品名稱, 數量, 單價, 備註, 備註.1
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

from converters.base import BaseConverter, ConversionResult, RowError
from app.utils import output_path

HEADERS = ["訂單編號", "收件人", "地址", "電話", "產品編號", "產品名稱", "數量", "單價", "備註", "備註.1"]
OUTPUT_FONT = Font(name="新細明體", size=12)

_STORE_RE = re.compile(r"門市：(\S+)")
_ORDER_RE = re.compile(r"單號：(\S+)")
_DATE_RE  = re.compile(r"採購日期：(\d{4})/(\d{1,2})/(\d{1,2})")
_STORE_SHORT_RE = re.compile(r"麗兒采家-(.+?)店")


class LicaiConverter(BaseConverter):

    def __init__(self, repository, output_dir: Path):
        super().__init__(repository, output_dir)
        self.data_dir = Path(__file__).resolve().parent.parent / "麗兒采家"
        self._barcode_map: dict[str, str]  = {}  # barcode → sku
        self._name_map:    dict[str, str]  = {}  # sku → 品名
        self._store_map:   dict[str, dict] = {}  # 短名 → {store_full, phone, address}

    @property
    def source_type(self) -> str:
        return "licai"

    def _load_reference(self) -> None:
        self._barcode_map = self.repository.load_barcodes("licai")
        from app.models import Product
        self._name_map = {p.sku: p.name for p in Product.query.all()}
        self._store_map = self._load_store_map()

    def _load_store_map(self) -> dict[str, dict]:
        wb = openpyxl.load_workbook(
            self.data_dir / "通路資料.xlsx", read_only=True, data_only=True
        )
        ws = wb.active
        store_map: dict[str, dict] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            store_full, phone, address = (str(v) if v else "" for v in row[:3])
            if "麗兒采家" not in store_full:
                continue
            m = _STORE_SHORT_RE.search(store_full)
            if not m:
                continue
            short = m.group(1).replace("旗艦", "")
            store_map[short] = {
                "store_full": store_full,
                "phone":      phone,
                "address":    address,
            }
        return store_map

    def _validate(self, input_files: list[Path]) -> list[RowError]:
        xlsx = [f for f in input_files if f.suffix.lower() == ".xlsx"]
        if not xlsx:
            return [RowError(0, "file", "", "找不到 .xlsx 採購單")]
        for f in xlsx:
            try:
                wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
                ws = wb.active
                rows = list(ws.iter_rows(max_row=2, values_only=True))
                if not rows or not any("門市：" in str(v) for v in rows[0] if v):
                    return [RowError(0, "sheet", f.name,
                                     "找不到「門市：」標頭列",
                                     source_file=f.name)]
            except Exception as e:
                return [RowError(0, "file", f.name, f"無法開啟：{e}",
                                 source_file=f.name)]
        return []

    def _process(self, input_files: list[Path]) -> ConversionResult:
        xlsx_files = [f for f in input_files if f.suffix.lower() == ".xlsx"]
        all_output:  list[Path]     = []
        all_errors:  list[RowError] = []
        total_success = 0
        total_fail    = 0
        order_date    = ""

        for f in xlsx_files:
            out, success, fail, errors, date = self._process_one(f)
            if out:
                all_output.append(out)
            total_success += success
            total_fail    += fail
            all_errors.extend(errors)
            if date and not order_date:
                order_date = date

        return ConversionResult(
            source_type   = self.source_type,
            success_count = total_success,
            fail_count    = total_fail,
            order_date    = order_date,
            output_files  = all_output,
            errors        = all_errors,
            status        = "completed" if total_success > 0 else "failed",
        )

    def _process_one(self, f: Path) -> tuple[Path | None, int, int, list[RowError], str]:
        wb   = openpyxl.load_workbook(f, read_only=True, data_only=True)
        ws   = wb.active
        rows = list(ws.iter_rows(values_only=True))

        row0_str = " ".join(str(v) for v in rows[0] if v)
        row1_str = " ".join(str(v) for v in rows[1] if v)

        store_m = _STORE_RE.search(row0_str)
        order_m = _ORDER_RE.search(row1_str)
        date_m  = _DATE_RE.search(row1_str)

        store_short = store_m.group(1) if store_m else ""
        order_id    = order_m.group(1) if order_m else ""
        order_date  = ""
        if date_m:
            y, mo, d = date_m.groups()
            order_date = f"{y}{int(mo):02d}{int(d):02d}"

        info        = self._store_map.get(store_short, {})
        store_full  = info.get("store_full", f"麗兒采家-{store_short}店")
        phone       = info.get("phone", "")
        address     = info.get("address", "")

        wb_out = Workbook()
        ws_out = wb_out.active
        ws_out.append(HEADERS)

        success = 0
        errors: list[RowError] = []

        for row_idx, row in enumerate(rows[3:], start=4):
            if row[0] is None:
                continue
            try:
                barcode = str(row[1]).strip() if row[1] is not None else ""
                price   = float(row[4]) if row[4] is not None else 0.0
                qty     = int(row[5])   if row[5] is not None else 0
            except (TypeError, ValueError) as e:
                errors.append(RowError(row_idx, "data", str(row), str(e),
                                       source_file=f.name))
                continue

            sku  = self._barcode_map.get(barcode, "")
            name = self._name_map.get(sku, str(row[2]) if row[2] else "")

            if not sku:
                errors.append(RowError(
                    row_idx, "barcode", barcode,
                    f"條碼 {barcode!r} 查無品號",
                    source_file=f.name,
                ))

            ws_out.append([
                order_id, store_full, address, phone,
                sku, name, qty, price, None, None,
            ])
            success += 1

        for row in ws_out.iter_rows():
            for cell in row:
                cell.font = OUTPUT_FONT

        mmdd = order_date[4:] if len(order_date) == 8 else "0000"
        date_obj = (
            datetime.strptime(order_date, "%Y%m%d") if order_date else datetime.today()
        )
        out_name = f"{mmdd} 麗采{store_short}.xlsx"
        out_path = output_path(self.output_dir.parent, self.source_type, date_obj, out_name)

        wb_out.save(out_path)
        return out_path, success, len(errors), errors, order_date
