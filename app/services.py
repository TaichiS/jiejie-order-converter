"""
app/services.py
ConversionService：協調偵測→轉換→歸檔→寫入DB。
也提供資料夾掃描邏輯（判斷已處理/未處理）。
"""
from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path

from app import db
from app.models import ConversionLog, ConversionError
from app.utils import archive_path, output_path, SOURCE_LABELS
from converters import detector

BASE_DIR               = Path(__file__).resolve().parent.parent
OUTPUT_DIR             = BASE_DIR / "output"
ARCHIVE_DIR            = BASE_DIR / "archive"


# ── 資料夾掃描 ────────────────────────────────────────────────────────────

def scan_folder(folder: str) -> list[dict]:
    """
    掃描資料夾，回傳每個檔案的偵測結果列表。
    格式：[{name, path, source_type, source_label, status}]
    status: 'pending' | 'archived' | 'unknown'
    """
    folder_path = Path(folder)
    if not folder_path.is_dir():
        raise ValueError(f"資料夾不存在：{folder}")

    # 收集所有 xlsx / pdf 檔案
    all_files = sorted(
        [f for f in folder_path.iterdir()
         if f.is_file() and f.suffix.lower() in (".xlsx", ".pdf")],
        key=lambda f: f.name
    )

    if not all_files:
        return []

    # 已歸檔檔名集合（比對原始檔名部分）
    archived_names = _archived_names()

    # 逐檔偵測來源
    file_types = detector.detect_each(all_files)

    results = []
    for f in all_files:
        if _is_archived(f.name, archived_names):
            status = "archived"
            src    = None
        else:
            src    = file_types.get(f)
            status = "pending" if src else "unknown"

        results.append({
            "name":         f.name,
            "path":         str(f),
            "source_type":  src,
            "source_label": SOURCE_LABELS.get(src, "未識別") if src else "未識別",
            "status":       status,
        })

    return results


def _archived_names() -> set[str]:
    """回傳 archive/ 下所有已歸檔的原始檔名（去除前綴後部分）"""
    names = set()
    for f in ARCHIVE_DIR.rglob("*"):
        if f.is_file():
            # 格式：{source_type}_{YYYYMMDD}_{原始檔名}
            parts = f.name.split("_", 2)
            if len(parts) == 3:
                names.add(parts[2])
    return names


def _is_archived(filename: str, archived_names: set[str]) -> bool:
    return filename in archived_names


# ── 轉換執行 ──────────────────────────────────────────────────────────────

def run_conversion(folder: str, operator: str,
                   forced_source: str | None = None,
                   selected_files: list[str] | None = None) -> list[ConversionLog]:
    """
    執行完整轉換流程，支援同一資料夾中多種來源類型混合。
    1. 掃描資料夾，收集待處理檔案
    2. 逐檔偵測來源（或使用 forced_source）
    3. 依來源類型分組，各組呼叫對應 Converter
    4. 歸檔原始檔，寫入 DB
    回傳 list[ConversionLog]（每種來源一筆）
    """
    folder_path = Path(folder)
    all_files   = sorted(
        [f for f in folder_path.iterdir()
         if f.is_file() and f.suffix.lower() in (".xlsx", ".pdf")],
        key=lambda f: f.name
    )
    # 排除已歸檔
    archived = _archived_names()
    pending  = [f for f in all_files if not _is_archived(f.name, archived)]

    # 依使用者勾選過濾（selected_files 為相對/絕對路徑字串清單）
    if selected_files:
        selected_names = {Path(p).name for p in selected_files}
        pending = [f for f in pending if f.name in selected_names]

    if not pending:
        raise ValueError("資料夾中沒有待處理的檔案")

    # 依 forced_source 或逐檔偵測建立分組
    if forced_source:
        groups: dict[str, list[Path]] = {forced_source: pending}
    else:
        file_types = detector.detect_each(pending)
        groups = {}
        for f, src in file_types.items():
            if src:
                groups.setdefault(src, []).append(f)
        if not groups:
            raise ValueError("無法自動判斷訂單來源，請手動指定")

    logs = []
    for source_type, files in groups.items():
        log = _convert_group(source_type, files, operator)
        logs.append(log)

    db.session.commit()
    return logs


def _convert_group(source_type: str, files: list[Path], operator: str) -> ConversionLog:
    """對單一來源類型的檔案群組執行轉換、歸檔，並寫入 DB（不 commit）。"""
    from converters.shopee      import ShopeeConverter
    from converters.a1baby      import A1BabyConverter
    from converters.leage       import LeageConverter
    from converters.a1leage     import A1LeageConverter
    from converters.jjofficial  import JJOfficialConverter
    from converters.yodee       import YodeeConverter
    from converters.kadomo      import KadomoConverter

    CONVERTER_MAP = {
        "shopee":     ShopeeConverter,
        "a1baby":     A1BabyConverter,
        "leage":      LeageConverter,
        "a1leage":    A1LeageConverter,
        "jjofficial": JJOfficialConverter,
        "yodee":      YodeeConverter,
        "kadomo":     KadomoConverter,
    }

    from app.repositories.product_repo import ProductRepository
    repo = ProductRepository()
    converter = CONVERTER_MAP[source_type](repo, OUTPUT_DIR)
    result    = converter.convert(files)

    # 解析訂單日期
    order_date = _parse_order_date(result.order_date)

    # 歸檔原始檔
    for f in files:
        dest = archive_path(BASE_DIR, source_type, order_date, f.name)
        shutil.copy2(f, dest)

    # 寫入 DB（由呼叫端統一 commit）
    log = ConversionLog(
        operator      = operator or "未填寫",
        source_type   = source_type,
        input_files   = json.dumps([f.name for f in files], ensure_ascii=False),
        success_count = result.success_count,
        fail_count    = result.fail_count,
        manual_count  = result.manual_count,
        output_files  = json.dumps([str(p) for p in result.output_files], ensure_ascii=False),
        status        = result.status,
    )
    db.session.add(log)
    db.session.flush()  # 取得 log.id

    for err in result.errors:
        db.session.add(ConversionError(
            log_id          = log.id,
            row_number      = err.row_number,
            field_name      = err.field_name,
            original_value  = err.original_value,
            reason          = err.reason,
            source_file     = err.source_file or "",
            candidates_json = json.dumps(
                [{k: v for k, v in c.items() if k in ("品號", "品名", "商品結帳價", "match_type")}
                 for c in (err.candidates or [])],
                ensure_ascii=False
            ) if err.candidates else None,
        ))

    return log


def _parse_order_date(order_date_str: str) -> datetime:
    """解析訂單日期字串 YYYYMMDD，失敗則用今日"""
    try:
        return datetime.strptime(order_date_str, "%Y%m%d")
    except (ValueError, TypeError):
        return datetime.today()
