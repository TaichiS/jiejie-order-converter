"""
converters/yodee.py
優迪通路（Yodee）訂單轉換器。
輸入：後台匯出 xlsx（含日期命名工作表，標頭列：訂單編號, 收件人, 地址…）
輸出：MMDD.xlsx（新細明體 12pt，固定欄寬）
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

HEADERS = ["訂單編號", "收件人", "地址", "電話", "產品編號", "產品名稱", "數量", "單價", "備註"]

# 欄寬（CLAUDE.md 規格）
COL_WIDTHS = {
    "A": 13.875, "B": 9.75,  "C": 26.125,
    "D": 14.25,  "E": 11.0,  "F": 28.125,
}

OUTPUT_FONT = Font(name="新細明體", size=12)

_DATE_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")


class YodeeConverter(BaseConverter):

    @property
    def source_type(self) -> str:
        return "yodee"

    def _validate(self, input_files: list[Path]) -> list[RowError]:
        xlsx = [f for f in input_files if f.suffix.lower() == ".xlsx"]
        if not xlsx:
            return [RowError(0, "file", "", "找不到 .xlsx 訂單檔")]
        try:
            wb = openpyxl.load_workbook(xlsx[0], read_only=True, data_only=True)
            if not _find_data_sheets(wb):
                return [RowError(0, "sheet", xlsx[0].name,
                                 "找不到含標頭列（訂單編號…）的工作表",
                                 source_file=xlsx[0].name)]
        except Exception as e:
            return [RowError(0, "file", xlsx[0].name, f"無法開啟：{e}",
                             source_file=xlsx[0].name)]
        return []

    def _process(self, input_files: list[Path]) -> ConversionResult:
        order_file = next(f for f in input_files if f.suffix.lower() == ".xlsx")
        wb_in = openpyxl.load_workbook(order_file, read_only=True, data_only=True)

        # 收集所有有效工作表的資料列
        all_rows:  list[tuple] = []
        order_date = ""

        for shname, rows in _find_data_sheets(wb_in).items():
            # 嘗試從工作表名稱提取日期
            m = _DATE_RE.search(shname)
            if m and not order_date:
                order_date = m.group(1) + m.group(2) + m.group(3)
            all_rows.extend(rows)

        # 日期 fallback：從資料第一列電話欄位試取（格式 0912345678-YYYYMMDD），或今日
        if not order_date:
            order_date = datetime.today().strftime("%Y%m%d")

        try:
            date_obj = datetime.strptime(order_date, "%Y%m%d")
        except ValueError:
            date_obj = datetime.today()
            order_date = date_obj.strftime("%Y%m%d")
        mmdd = order_date[4:8]

        success_count = len(all_rows)
        out_path = _write_xlsx(self.output_dir.parent, self.source_type,
                               date_obj, f"{mmdd}.xlsx", all_rows)

        return ConversionResult(
            source_type   = self.source_type,
            success_count = success_count,
            fail_count    = 0,
            order_date    = order_date,
            output_files  = [out_path],
            errors        = [],
            status        = "completed" if success_count > 0 else "failed",
        )


# ── 輸出 ──────────────────────────────────────────────────────────────────

def _write_xlsx(base_dir: Path, source_type: str, date_obj: datetime,
                filename: str, rows: list[tuple]) -> Path:
    p = output_path(base_dir, source_type, date_obj, filename)
    wb = Workbook()
    ws = wb.active
    ws.title = date_obj.strftime("%Y-%m-%d")

    # 欄寬
    for col, width in COL_WIDTHS.items():
        ws.column_dimensions[col].width = width

    # 標頭列
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = OUTPUT_FONT

    # 資料列
    for row in rows:
        ws.append(list(row))
        for cell in ws[ws.max_row]:
            cell.font = OUTPUT_FONT

    wb.save(p)
    return p


# ── 工具 ─────────────────────────────────────────────────────────────────

def _find_data_sheets(wb: openpyxl.Workbook) -> dict[str, list[tuple]]:
    """
    回傳 {工作表名稱: [資料列 tuple, ...]}，
    只收錄第一列為標準標頭的工作表（跳過無標頭的歷史資料表）。
    """
    result: dict[str, list[tuple]] = {}
    for shname in wb.sheetnames:
        ws = wb[shname]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        first = [str(v).strip() if v else "" for v in rows[0]]
        if first[0] == "訂單編號":
            # 收集非空資料列（跳過標頭）
            data = [r for r in rows[1:] if any(v is not None for v in r)]
            if data:
                result[shname] = data
    return result
