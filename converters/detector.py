"""
converters/detector.py
來源自動偵測：檔名為初步提示，讀取內部欄位/文字確認（雙重驗證）。
"""
from __future__ import annotations

import re
from pathlib import Path

from converters.base import open_xlsx, SHOPEE_XLSX_PASSWORD


def detect_each(files: list[Path]) -> dict[Path, str | None]:
    """
    逐檔偵測來源類型，回傳 {path: source_type} 對應表。
    a1baby 需要成對檔案，兩個都會標記為 'a1baby'。
    """
    result: dict[Path, str | None] = {}

    xlsx_files = [f for f in files if f.suffix.lower() == ".xlsx"]
    pdf_files  = [f for f in files if f.suffix.lower() == ".pdf"]

    # PDF：各自確認是否為樂齡網
    for f in pdf_files:
        result[f] = "leage" if _confirm_leage(f) else None

    # xlsx：逐檔確認
    unmatched_xlsx = []
    for f in xlsx_files:
        if _confirm_shopee(f):
            result[f] = "shopee"
        elif _confirm_a1leage(f):
            result[f] = "a1leage"
        elif _confirm_jjofficial(f):
            result[f] = "jjofficial"
        elif _confirm_yodee(f):
            result[f] = "yodee"
        elif _confirm_licai(f):
            result[f] = "licai"
        elif _confirm_kadomo(f):
            result[f] = "kadomo"
        else:
            unmatched_xlsx.append(f)

    # 未命中的 xlsx：嘗試 a1baby 配對（支援多組）
    for main_file, detail_file in _find_a1baby_pairs(unmatched_xlsx):
        if _confirm_a1baby(main_file, detail_file):
            result[main_file]   = "a1baby"
            result[detail_file] = "a1baby"
            unmatched_xlsx = [f for f in unmatched_xlsx if f not in (main_file, detail_file)]

    for f in unmatched_xlsx:
        result[f] = None

    return result


def detect(files: list[Path]) -> str | None:
    """
    從一組（同來源）檔案中偵測訂單來源類型，供 run_conversion 使用。
    回傳 'shopee' | 'a1baby' | 'leage' | 'a1leage'，無法識別回傳 None。
    """
    types = set(detect_each(files).values()) - {None}
    if len(types) == 1:
        return types.pop()
    # 多種類型或全部無法識別
    return None


# ── 內部確認函式 ──────────────────────────────────────────────────────────

def _confirm_shopee(path: Path) -> bool:
    """確認 xlsx 為蝦皮訂單（工作表 orders + 含蝦皮專線欄位；支援加密檔）"""
    # 蝦皮固定檔名前綴，加密檔解密後欄位可能無法讀取，以此作為輔助辨識
    name_match = re.match(r"Order\.toship\.", path.name, re.IGNORECASE) is not None
    try:
        wb = open_xlsx(path, password=SHOPEE_XLSX_PASSWORD)
        if "orders" not in wb.sheetnames:
            return False
        ws   = wb["orders"]
        hdrs = [str(c.value) if c.value else "" for c in next(ws.iter_rows(max_row=1))]
        # 優先：欄位精確比對
        if any("蝦皮專線和包裹查詢碼" in h for h in hdrs):
            return True
        # 備用：加密檔解密成功 + orders 工作表存在 + 符合蝦皮檔名格式
        return name_match
    except Exception:
        return False


def _find_a1baby_pairs(xlsx_files: list[Path]) -> list[tuple[Path, Path]]:
    """尋找所有 MMDD.xlsx 與 MMDD-1.xlsx 配對，回傳 [(main, detail), ...]"""
    file_set = set(xlsx_files)
    pairs = []
    for f in xlsx_files:
        if re.search(r"-1\.xlsx$", f.name, re.IGNORECASE):
            main_name = re.sub(r"-1\.xlsx$", ".xlsx", f.name, flags=re.IGNORECASE)
            main_file = f.parent / main_name
            if main_file in file_set:
                pairs.append((main_file, f))
    return pairs


def _confirm_a1baby(main: Path, detail: Path) -> bool:
    """確認婦幼展主檔含 原始單號+訂單標籤與備註，明細含 發票號碼+商品名稱"""
    try:
        import openpyxl
        wb_main = openpyxl.load_workbook(main,   read_only=True, data_only=True)
        wb_det  = openpyxl.load_workbook(detail, read_only=True, data_only=True)

        ws_main = wb_main.active
        ws_det  = wb_det.active

        def hdrs(ws):
            row = next(ws.iter_rows(max_row=1), [])
            return {str(c.value).strip() if c.value else "" for c in row}

        main_ok   = {"原始單號", "訂單標籤與備註"}.issubset(hdrs(ws_main))
        detail_ok = {"發票號碼", "商品名稱"}.issubset(hdrs(ws_det))
        return main_ok and detail_ok
    except Exception:
        return False


def _confirm_a1leage(path: Path) -> bool:
    """確認 xlsx 為 A1樂齡官網訂單（Orders 工作表 + 特有欄位名稱組合）"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if "Orders" not in wb.sheetnames:
            return False
        ws   = wb["Orders"]
        row  = next(ws.iter_rows(max_row=1, values_only=True), None)
        if not row:
            return False
        hdrs = {str(v).strip() for v in row if v}
        return {"貨號", "收件人地址", "購買品項"}.issubset(hdrs)
    except Exception:
        return False


def _confirm_jjofficial(path: Path) -> bool:
    """確認 xlsx 為捷捷寶寶粥官網訂單（Sales 工作表 + 全家服務編號欄位）"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if "Sales" not in wb.sheetnames:
            return False
        ws   = wb["Sales"]
        row  = next(ws.iter_rows(max_row=1, values_only=True), None)
        if not row:
            return False
        hdrs = {str(v).strip() for v in row if v}
        return {"訂單號碼", "全家服務編號 / 7-11 店號", "加購品類型"}.issubset(hdrs)
    except Exception:
        return False


def _confirm_yodee(path: Path) -> bool:
    """確認 xlsx 為優迪通路訂單（含標頭列 訂單編號…產品名稱 的工作表）"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for shname in wb.sheetnames:
            ws = wb[shname]
            row = next(ws.iter_rows(max_row=1, values_only=True), None)
            if not row:
                continue
            hdrs = [str(v).strip() if v else "" for v in row]
            if hdrs[0] == "訂單編號" and "產品名稱" in hdrs:
                return True
        return False
    except Exception:
        return False


def _confirm_kadomo(path: Path) -> bool:
    """確認 xlsx 為卡多摩採購單（含「倉別」與「商品條碼」的標題列，前 15 列內）"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(max_row=15, values_only=True):
            cells = {str(v).strip() for v in row if v is not None}
            if "倉別" in cells and "商品條碼" in cells:
                return True
        return False
    except Exception:
        return False


def _confirm_leage(path: Path) -> bool:
    """確認 PDF 含樂齡網文字特徵"""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            text = pdf.pages[0].extract_text() or ""
            return bool(re.search(r"樂齡生活事業|PO\d{8}", text))
    except Exception:
        return False


def _confirm_licai(path: Path) -> bool:
    """確認 xlsx 為麗兒采家個別門市採購單（row 0 有「門市：」與「門市資訊：」）"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(max_row=2, values_only=True))
        if len(rows) < 2:
            return False
        row0 = [str(v) if v is not None else "" for v in rows[0]]
        row1 = [str(v) if v is not None else "" for v in rows[1]]
        return (
            any("門市：" in v for v in row0) and
            any("門市資訊：" in v for v in row0) and
            any("單號：" in v for v in row1)
        )
    except Exception:
        return False
