"""
converters/base.py
BaseConverter 抽象基底類，以及共用資料結構 ConversionResult、RowError。
"""
from __future__ import annotations

import io
import re
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from difflib import get_close_matches
from pathlib import Path

SHOPEE_XLSX_PASSWORD = "870993"


def open_xlsx(path: Path, password: str | None = None,
              read_only: bool = True, data_only: bool = True):
    """
    開啟 xlsx 檔案，自動處理加密保護。
    若直接開啟失敗且提供了 password，嘗試以 msoffcrypto 解密後再開。
    回傳 openpyxl.Workbook；失敗時拋出原始例外。
    """
    import openpyxl
    try:
        return openpyxl.load_workbook(path, read_only=read_only, data_only=data_only)
    except Exception as first_err:
        if password is None:
            raise
        # 嘗試解密
        try:
            import msoffcrypto
            with open(path, "rb") as f:
                office = msoffcrypto.OfficeFile(f)
                office.load_key(password=password)
                buf = io.BytesIO()
                office.decrypt(buf)
            buf.seek(0)
            # read_only=True 搭配 BytesIO 在 openpyxl 中無法正確偵測工作表範圍，
            # 解密後的記憶體流必須用 read_only=False 開啟。
            return openpyxl.load_workbook(buf, read_only=False, data_only=data_only)
        except Exception:
            raise first_err  # 解密也失敗，回傳原始錯誤

# 品名中的「階段代碼-序號」前綴，如 0-1、1-01、1P-05、2-S11、2-D01、3-3
# \d+[A-Z]? = 前段（數字 + 可選字母，如 1P）
# -[A-Z]?   = 連字號 + 可選系列字母（如 S、D、M）
# \d+        = 後段序號
_CODE_RE      = re.compile(r'^(\d+[A-Z]?-[A-Z]?\d+)')
# 蝦皮特例：缺少「2-」前綴的燉飯/義麵代碼，如 S11、M3
_BARE_CODE_RE = re.compile(r'^([A-Z]\d+)')


@dataclass
class RowError:
    """單一列的轉換錯誤資訊"""
    row_number:     int
    field_name:     str
    original_value: str
    reason:         str
    candidates:     list = field(default_factory=list)  # [{品號,品名,商品結帳價,match_type}]
    source_file:    str  = ""                           # 來源檔名（空字串表示未知）


@dataclass
class ConversionResult:
    """單次轉換的完整結果"""
    source_type:   str
    success_count: int              = 0
    fail_count:    int              = 0
    manual_count:  int              = 0   # 需人工確認（數量異常等）
    order_date:    str              = ""  # 從訂單內部讀取的日期 YYYYMMDD
    output_files:  list[Path]       = field(default_factory=list)
    errors:        list[RowError]   = field(default_factory=list)
    status:        str              = "completed"  # completed | partial | failed


class BaseConverter(ABC):
    """所有轉換器的抽象基底類"""

    def __init__(self, repository, output_dir: Path):
        self.repository    = repository
        self.output_dir    = output_dir
        self._product_map: dict[str, dict] = {}   # 品名 → row dict
        self._code_map:    dict[str, dict] = {}   # 代碼前綴 → row dict（如 "1-08" → {...}）
        self._sku_map:     dict[str, dict] = {}   # 品號 → row dict

    # ── 公開入口 ──────────────────────────────────────────────────────────
    def convert(self, input_files: list[Path]) -> ConversionResult:
        """固定流程：載入品號 → 驗證 → 轉換"""
        self._load_reference()
        errors = self._validate(input_files)
        if errors:
            return ConversionResult(
                source_type=self.source_type,
                fail_count=len(errors),
                errors=errors,
                status="failed",
            )
        return self._process(input_files)

    # ── 共用：載入品號資料 ─────────────────────────────────────────────────
    def _load_reference(self) -> None:
        """從 ProductRepository 載入指定通路的品號資料到 in-memory cache。"""
        self._product_map, self._code_map = self.repository.load_channel(self.source_type)
        self._sku_map = {
            str(p.get("品號", "")).strip(): p
            for p in self._product_map.values()
            if p.get("品號")
        }

    def _lookup_by_name(self, query: str,
                        scope: str | None = None,
                        amount: float = 0) -> tuple[dict | None, list[dict]]:
        """
        多層查找，回傳 (最佳匹配 | None, 候選清單)。
        候選清單在找不到精確結果時提供，供人工確認。

        查找順序：
        1. 從 query 提取代碼前綴 → _code_map（如 "1-08木耳..." → "1-08"）
        2. 蝦皮特例：裸代碼補 "2-" 前綴（如 "S11..." → "2-S11"）
        3. 精確品名比對 _product_map
        4. difflib 模糊比對（名稱相近）+ 金額整除比對 → 候選清單

        scope:  限定品號前綴，如 "F" 只比對樂齡網品項
        amount: 訂單金額，>0 時加入金額整除候選
        """
        q = query.strip()
        pool = (
            {k: v for k, v in self._product_map.items()
             if str(v.get("品號", "")).startswith(scope)}
            if scope else self._product_map
        )

        # 1. 代碼前綴比對（code_map 每個 key 存 list，支援同代碼多規格）
        def _pick_from_code(prods: list[dict]) -> dict | None:
            """從候選清單中依 scope / amount 取最佳匹配。"""
            if scope:
                prods = [p for p in prods if str(p.get("品號", "")).startswith(scope)]
            if not prods:
                return None
            if len(prods) == 1:
                return prods[0]
            # 多規格：用金額整除區分（如 150g/200g）
            if amount > 0:
                for p in prods:
                    cp = float(p.get("商品結帳價") or 0)
                    if cp > 0 and amount % cp == 0:
                        return p
            return prods[0]

        m = _CODE_RE.match(q)
        if m:
            code = m.group(1)
            prods = self._code_map.get(code)
            if prods:
                prod = _pick_from_code(prods)
                if prod:
                    return prod, []

        # 2. 蝦皮裸代碼（S11 → 2-S11）
        m2 = _BARE_CODE_RE.match(q)
        if m2:
            prods = self._code_map.get("2-" + m2.group(1))
            if prods:
                prod = _pick_from_code(prods)
                if prod:
                    return prod, []

        # 3. 精確品名
        prod = pool.get(q)
        if prod:
            return prod, []

        # 4. 候選清單：名稱相似 + 金額整除
        names      = list(pool.keys())
        close_keys = get_close_matches(q, names, n=3, cutoff=0.35)
        seen_nos: set[str] = set()
        candidates: list[dict] = []

        for k in close_keys:
            p = dict(pool[k])
            p["match_type"] = "name"
            no = p.get("品號", "")
            if no not in seen_nos:
                seen_nos.add(no)
                candidates.append(p)

        if amount > 0:
            for name, p in pool.items():
                cp = float(p.get("商品結帳價") or 0)
                if cp > 0 and amount % cp == 0:
                    no = p.get("品號", "")
                    if no not in seen_nos:
                        seen_nos.add(no)
                        c = dict(p)
                        c["match_type"] = "price"
                        candidates.append(c)

        return None, candidates[:5]

    def _lookup_product(self, name: str) -> dict | None:
        """精確比對品名，找不到回傳 None（向下相容舊呼叫）"""
        return self._product_map.get(name.strip())

    def _lookup_by_sku(self, sku: str) -> dict | None:
        """依品號精確查詢，找不到回傳 None。"""
        return self._sku_map.get(sku.strip())

    # ── 抽象方法（子類實作）────────────────────────────────────────────────
    @property
    @abstractmethod
    def source_type(self) -> str:
        """回傳 'shopee' | 'a1baby' | 'leage'"""

    @abstractmethod
    def _validate(self, input_files: list[Path]) -> list[RowError]:
        """驗證輸入檔案格式，回傳錯誤列表（空列表代表通過）"""

    @abstractmethod
    def _process(self, input_files: list[Path]) -> ConversionResult:
        """實際轉換邏輯"""
