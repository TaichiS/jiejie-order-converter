import re
from pathlib import Path
from datetime import datetime


SOURCE_LABELS = {
    "shopee":  "蝦皮",
    "a1baby":  "A1婦幼展",
    "leage":   "樂齡網",
    "a1leage": "A1樂齡官網",
}

SOURCE_COLORS = {
    "shopee":  "warning",
    "a1baby":  "info",
    "leage":   "success",
    "a1leage": "danger",
}


def safe_filename(name: str) -> str:
    """移除檔名中不合法的字元（Windows 相容）"""
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def archive_path(base_dir: Path, source_type: str, date: datetime, original_name: str) -> Path:
    """
    建立歸檔路徑：archive/{source_type}/{YYYY-MM}/{source_type}_{YYYYMMDD}_{原始檔名}
    """
    month_dir = base_dir / "archive" / source_type / date.strftime("%Y-%m")
    month_dir.mkdir(parents=True, exist_ok=True)
    prefix = f"{source_type}_{date.strftime('%Y%m%d')}_"
    return month_dir / (prefix + safe_filename(original_name))


def output_path(base_dir: Path, source_type: str, date: datetime, filename: str) -> Path:
    """
    建立輸出歸檔路徑：output/{source_type}/{YYYY-MM}/{filename}
    """
    month_dir = base_dir / "output" / source_type / date.strftime("%Y-%m")
    month_dir.mkdir(parents=True, exist_ok=True)
    return month_dir / safe_filename(filename)
