"""
converters/chocho.py
CHOCHO 通路訂單轉換器。
輸入：CSV（標準標頭：訂單編號,收件人,地址,電話,產品編號,產品名稱,數量,單價）
輸出：ERP_YYYY-MM-DD_to_YYYY-MM-DD.csv

轉換邏輯：
- 根據產品編號查詢資料庫中的 pack_size 與 pack_price
- 數量 = 原始數量 // pack_size
- 單價 = ceil(pack_price * 1.05)（含稅並無條件進位）
- 其餘欄位直接複製
"""
from __future__ import annotations

import csv
import io
import math
import re
from datetime import datetime
from pathlib import Path

from converters.base import BaseConverter, ConversionResult, RowError
from app.models import UnifiedProduct
from app.utils import output_path

_DATE_RE = re.compile(r"(\d{2})(\d{2})")

HEADERS = ["訂單編號", "收件人", "地址", "電話", "產品編號", "產品名稱", "數量", "單價"]


class ChochoConverter(BaseConverter):

    def __init__(self, repository, output_dir: Path):
        super().__init__(repository, output_dir)
        self._sku_map: dict[str, dict] = {}  # sku → {pack_size, pack_price}

    @property
    def source_type(self) -> str:
        return "chocho"

    def _load_reference(self) -> None:
        """載入 chocho 通路的品號資料。"""
        rows = UnifiedProduct.query.filter_by(channel="chocho通路").all()
        self._sku_map = {
            r.sku: {
                "pack_size": r.pack_size or 1,
                "pack_price": float(r.pack_price) if r.pack_price is not None else 0.0,
            }
            for r in rows
        }

    def _validate(self, input_files: list[Path]) -> list[RowError]:
        csv_files = [f for f in input_files if f.suffix.lower() == ".csv"]
        if not csv_files:
            return [RowError(0, "file", "", "找不到 .csv 訂單檔")]

        errors: list[RowError] = []
        for f in csv_files:
            try:
                text = f.read_text(encoding="utf-8-sig")
                reader = csv.reader(io.StringIO(text))
                hdr = next(reader, None)
                if not hdr:
                    errors.append(RowError(0, "file", f.name, "CSV 為空",
                                           source_file=f.name))
                    continue
                hdr_set = {h.strip() for h in hdr if h}
                needed = set(HEADERS)
                missing = needed - hdr_set
                if missing:
                    errors.append(RowError(0, "header", f.name,
                                           f"缺少欄位：{', '.join(missing)}",
                                           source_file=f.name))
            except Exception as e:
                errors.append(RowError(0, "file", f.name, f"無法開啟：{e}",
                                       source_file=f.name))
        return errors

    def _process(self, input_files: list[Path]) -> ConversionResult:
        csv_files = [f for f in input_files if f.suffix.lower() == ".csv"]
        all_output: list[Path] = []
        all_errors: list[RowError] = []
        total_success = 0
        order_date = ""

        for f in csv_files:
            text = f.read_text(encoding="utf-8-sig")
            reader = csv.DictReader(io.StringIO(text))

            output_rows: list[list] = []
            for row_idx, row in enumerate(reader, start=2):
                sku = (row.get("產品編號") or "").strip()
                qty_str = (row.get("數量") or "").strip()
                name = (row.get("產品名稱") or "").strip()

                if not sku:
                    continue

                info = self._sku_map.get(sku)
                if not info:
                    all_errors.append(RowError(
                        row_idx, "產品編號", sku,
                        "品號資料表中找不到對應品項",
                        source_file=f.name,
                    ))
                    continue

                pack_size = info.get("pack_size", 1) or 1
                try:
                    raw_qty = int(float(qty_str)) if qty_str else 0
                except ValueError:
                    raw_qty = 0

                out_qty = raw_qty // pack_size
                if out_qty <= 0:
                    all_errors.append(RowError(
                        row_idx, "數量", qty_str,
                        f"轉換後數量為 0（原始數量 {raw_qty} / 包數 {pack_size}）",
                        source_file=f.name,
                    ))
                    continue

                pack_price = info.get("pack_price", 0.0) or 0.0
                out_price = math.ceil(pack_price * 1.05)

                output_rows.append([
                    row.get("訂單編號", "").strip(),
                    row.get("收件人", "").strip(),
                    row.get("地址", "").strip(),
                    row.get("電話", "").strip(),
                    sku,
                    name,
                    out_qty,
                    out_price,
                ])
                total_success += 1

            # 提取日期：從檔名找 MMDD
            mmdd_match = _DATE_RE.search(f.stem)
            if mmdd_match:
                mm = mmdd_match.group(1)
                dd = mmdd_match.group(2)
                year = datetime.today().year
                order_date = f"{year}{mm}{dd}"
            else:
                order_date = datetime.today().strftime("%Y%m%d")

            date_obj = datetime.strptime(order_date, "%Y%m%d")
            date_str = date_obj.strftime("%Y-%m-%d")
            out_filename = f"ERP_{date_str}_to_{date_str}.csv"
            out_path = _write_csv(self.output_dir.parent, self.source_type,
                                  date_obj, out_filename, output_rows)
            all_output.append(out_path)

        return ConversionResult(
            source_type=self.source_type,
            success_count=total_success,
            fail_count=len(all_errors),
            order_date=order_date,
            output_files=all_output,
            errors=all_errors,
            status="completed" if total_success > 0 else "failed",
        )


def _write_csv(base_dir: Path, source_type: str, date_obj: datetime,
               filename: str, rows: list[list]) -> Path:
    p = output_path(base_dir, source_type, date_obj, filename)
    with open(p, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        writer.writerows(rows)
    return p
