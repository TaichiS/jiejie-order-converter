"""
converters/tuanma.py
其他團媽訂單轉換器。
輸入：後台匯出 xlsx（工作表 Sales）
輸出：MMDD其他團媽.xlsx

轉換邏輯：
- 驗證 Sales 工作表存在且含必要欄位
- 直接複製所有列，僅移除「訂單備註」欄位
- 發票開立日期用於產生檔名
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl import Workbook

from converters.base import BaseConverter, ConversionResult, RowError
from app.utils import output_path

OUTPUT_HEADER = [
    "訂單號碼", "收件人", "完整地址", "收件人電話號碼", "發票號碼",
    "商品貨號", "商品名稱", "數量", "商品結帳價", "商品折扣優惠",
    "商品折扣金額", "加購品類型", "點數折現分攤", "全單折扣優惠", "全單折扣金額",
    "折抵購物金分攤", "送貨編號", "出貨備註", "全家服務編號 / 7-11 店號",
    "發票開立日期", "運費", "附加費",
]

_DATE_SLASH = re.compile(r'(\d{4})/(\d{2})/(\d{2})')


class TuanmaConverter(BaseConverter):

    @property
    def source_type(self) -> str:
        return "tuanma"

    def _validate(self, input_files: list[Path]) -> list[RowError]:
        xlsx = [f for f in input_files if f.suffix.lower() == ".xlsx"]
        if not xlsx:
            return [RowError(0, "file", "", "找不到 .xlsx 訂單檔")]
        errors = []
        for f in xlsx:
            try:
                wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
                if "Sales" not in wb.sheetnames:
                    errors.append(RowError(0, "sheet", f.name, "工作表 'Sales' 不存在",
                                           source_file=f.name))
                    continue
                ws = wb["Sales"]
                row = next(ws.iter_rows(max_row=1, values_only=True), None)
                if not row:
                    errors.append(RowError(0, "sheet", f.name, "工作表為空",
                                           source_file=f.name))
                    continue
                hdrs = {str(v).strip() for v in row if v}
                needed = {"訂單號碼", "商品貨號", "全家服務編號 / 7-11 店號", "訂單備註"}
                missing = needed - hdrs
                if missing:
                    errors.append(RowError(0, "header", f.name,
                                           f"缺少欄位：{', '.join(missing)}",
                                           source_file=f.name))
            except Exception as e:
                errors.append(RowError(0, "file", f.name, f"無法開啟：{e}",
                                       source_file=f.name))
        return errors

    def _process(self, input_files: list[Path]) -> ConversionResult:
        xlsx_files = [f for f in input_files if f.suffix.lower() == ".xlsx"]
        total_success = 0
        total_fail = 0
        all_errors: list[RowError] = []
        output_files: list[Path] = []
        order_date = ""

        for order_file in xlsx_files:
            wb = openpyxl.load_workbook(order_file, read_only=True, data_only=True)
            ws = wb["Sales"]
            all_rows = list(ws.iter_rows(values_only=True))
            if not all_rows:
                continue

            hdr = [str(v).strip() if v else "" for v in all_rows[0]]
            idx_map = {name: i for i, name in enumerate(hdr)}

            # 從第一筆有值的發票開立日期提取訂單日期
            date_col = idx_map.get("發票開立日期")
            if date_col is not None and not order_date:
                for row in all_rows[1:]:
                    if date_col < len(row) and row[date_col]:
                        d = row[date_col]
                        m = _DATE_SLASH.search(str(d))
                        if m:
                            order_date = m.group(1) + m.group(2) + m.group(3)
                            break

            # fallback：從檔名前 4 碼取 MMDD
            if not order_date:
                stem_match = re.match(r'^(\d{4})', order_file.stem)
                if stem_match:
                    order_date = datetime.today().strftime("%Y") + stem_match.group(1)

            # 組建輸出列：移除「訂單備註」欄位
            output_rows: list[tuple] = []
            note_idx = idx_map.get("訂單備註")

            for row_idx, row in enumerate(all_rows[1:], start=2):
                if not any(v is not None for v in row):
                    continue

                # 跳過無訂單號碼的列
                order_id_idx = idx_map.get("訂單號碼")
                if order_id_idx is None or order_id_idx >= len(row) or not row[order_id_idx]:
                    continue

                new_row = []
                for col_name in OUTPUT_HEADER:
                    col_idx = idx_map.get(col_name)
                    val = row[col_idx] if col_idx is not None and col_idx < len(row) else None
                    new_row.append(val)

                output_rows.append(tuple(new_row))
                total_success += 1

            if not output_rows:
                continue

            try:
                date_obj = datetime.strptime(order_date, "%Y%m%d")
            except (ValueError, TypeError):
                date_obj = datetime.today()

            mmdd = date_obj.strftime("%m%d")
            out_path = _write_xlsx(self.output_dir.parent, self.source_type,
                                   date_obj, f"{mmdd}其他團媽.xlsx", output_rows)
            output_files.append(out_path)

        if not order_date:
            order_date = datetime.today().strftime("%Y%m%d")

        return ConversionResult(
            source_type=self.source_type,
            success_count=total_success,
            fail_count=total_fail,
            order_date=order_date,
            output_files=output_files,
            errors=all_errors,
            status="completed" if total_success > 0 else "failed",
        )


def _write_xlsx(base_dir: Path, source_type: str, date_obj: datetime,
                filename: str, rows: list[tuple]) -> Path:
    p = output_path(base_dir, source_type, date_obj, filename)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(OUTPUT_HEADER)
    for row in rows:
        ws.append(list(row))
    wb.save(p)
    return p
