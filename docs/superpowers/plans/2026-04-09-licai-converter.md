# 麗兒采家轉換器 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 將麗兒采家（licai）嬰童館採購單整合進現有轉換系統，支援個別門市 xlsx → 標準格式 xlsx 的自動轉換。

**Architecture:** 麗兒采家的原始大檔由使用者透過既有 `split_orders.py` 手動拆分成各門市個別 xlsx；這些拆分後的檔案丟入掃描資料夾後，偵測器辨識（`_confirm_licai`）→ `LicaiConverter` 將每筆條碼查 DB 轉成品號 → 輸出格式與卡多摩相同（訂單編號、收件人、地址、電話、產品編號、產品名稱、數量、單價、備註、備註.1）。通路定價由訂單原檔讀取（市價欄），DB 中 licai channel_prices 以 0.0 作佔位，讓 `load_barcodes("licai")` 正常運作。

**Tech Stack:** Flask-SQLAlchemy, openpyxl, Python 3.12, uv

---

## 檔案結構

| 動作 | 路徑 | 說明 |
|------|------|------|
| **建立** | `converters/licai.py` | LicaiConverter 主體 |
| **修改** | `converters/detector.py` | 新增 `_confirm_licai` + 加入偵測鏈 |
| **修改** | `app/services.py` | 登錄 licai → LicaiConverter |
| **修改** | `app/utils.py` | 新增 licai 標籤與顏色 |
| **修改** | `app/models.py` | ConversionLog.to_dict() 中 source_label 補 licai |
| **修改** | `scripts/migrate_reference_data.py` | 新增 import_licai() 函式 |
| **修改** | `templates/help.html` | 新增麗兒采家說明卡片 |
| **修改** | `CHANGELOG.md` | 記錄本版本變更 |

---

## 關鍵資料格式（實作前必讀）

### 輸入檔案結構（split 後的個別門市 xlsx）
```
Row 0: ('門市：佳里', None, '門市資訊：臺南市佳里區進學路169號 / 06-7232730', None, None, None)
Row 1: ('單號：UW26033008', None, '採購日期：2026/3/30', None, None, None)
Row 2: (None, None, None, None, None, None)   ← 空列
Row 3: (1, '4710586220100', '010青花椰米泥(50克)/4', '捷捷寶寶粥', 40, 1)
Row 4: (2, '4710586220117', '011高麗菜米泥(50克)/4', '捷捷寶寶粥', 40, 1)
...
欄位順序：(序號int, 條碼str, 品名, 品牌, 市價float, 數量int)
```

### 輸出檔案格式（與卡多摩相同）
```
欄位：訂單編號, 收件人, 地址, 電話, 產品編號, 產品名稱, 數量, 單價, 備註, 備註.1
範例：('UW26033008', '麗兒采家-佳里店(中午12點後到貨)', '臺南市佳里區進學路169號(中午12點後到貨)', '06-7232730', 'D50100010', '0-10青花椰菜米泥', 1, 40, None, None)
```
輸出檔名：`MMDD 麗采{門市短名}.xlsx`（如 `0330 麗采佳里.xlsx`）

### 麗兒采家門市短名映射邏輯
`通路資料.xlsx` 的店名欄範例：`麗兒采家-佳里店(中午12點後到貨)`
- regex：`麗兒采家-(.+?)店`  → group(1)
- 再把 `旗艦` 去掉 → 得到短名 `佳里`（`大墩旗艦` → `大墩`）
- 短名須與 row 0 `門市：佳里` 的值一致

### 品號資料_含條碼.csv 格式
```
品號,條碼,品名
D50100011,4710586220117,0-11高麗菜米泥
...
```
（無「商品結帳價」欄位，與現有 reference CSV 格式不同）

---

## Task 1：匯入 麗兒采家 barcodes 與 channel_prices 至 DB

**Files:**
- Modify: `scripts/migrate_reference_data.py`

- [ ] **Step 1：在 `migrate_reference_data.py` 末端（main() 之前）新增 `import_licai()` 函式**

在第 142 行（`import_kadomo_prices` 函式結束後）插入：

```python
LICAI_CSV  = BASE_DIR / "麗兒采家" / "品號資料_含條碼.csv"


def import_licai() -> None:
    """
    匯入麗兒采家品號 / 條碼 / channel_prices(licai)。
    CSV 格式：品號, 條碼, 品名（無商品結帳價欄）。
    channel_prices 以 price=0.0 佔位（實際價格從訂單 xlsx 讀取）。
    """
    with open(LICAI_CSV, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))

    for row in rows:
        sku     = row.get("品號", "").strip()
        barcode = row.get("條碼", "").strip()
        name    = row.get("品名", "").strip()
        if not sku or not barcode or not name:
            continue

        # Upsert Product（若已存在則只更新 name）
        prod = db.session.get(Product, sku)
        if prod is None:
            prod = Product(sku=sku, name=name)
            db.session.add(prod)
        else:
            prod.name = name

        # Upsert ProductBarcode（條碼全域唯一，可能已由 kadomo 匯入）
        exists_bc = ProductBarcode.query.filter_by(barcode=barcode).first()
        if not exists_bc:
            db.session.add(ProductBarcode(sku=sku, barcode=barcode))

        # Upsert ChannelPrice for licai（price=0.0 佔位，讓 load_barcodes 正常）
        _upsert_price(sku, "licai", 0.0)

    db.session.flush()
```

- [ ] **Step 2：在 `main()` 末尾新增呼叫**

在 `main()` 函式的最後 print 與 commit 之前加入：

```python
        print("  匯入 麗兒采家/品號資料_含條碼.csv → 通路：licai")
        import_licai()
```

完整 `main()` 應如下：

```python
def main() -> None:
    app = create_app()
    with app.app_context():
        print("開始匯入品號資料...")

        for src in CSV_SOURCES:
            p = src["path"]
            print(f"  匯入 {p.relative_to(BASE_DIR)} → 通路：{src['channels']}")
            import_csv(p, src["channels"])

        print("  匯入 卡多摩/條碼對照表.json")
        import_kadomo_barcodes()

        print("  匯入 卡多摩/單價對照表.json → 通路：kadomo")
        import_kadomo_prices()

        print("  匯入 麗兒采家/品號資料_含條碼.csv → 通路：licai")
        import_licai()

        db.session.commit()
        print("✓ 全部完成！")


if __name__ == "__main__":
    main()
```

- [ ] **Step 3：執行匯入並確認**

```bash
cd /Users/lung/Documents/企業顧問/02_客戶管理/捷捷寶寶粥/訂單格式轉換系統
uv run python3 scripts/migrate_reference_data.py
```

預期輸出包含：
```
  匯入 麗兒采家/品號資料_含條碼.csv → 通路：licai
✓ 全部完成！
```

- [ ] **Step 4：驗證 DB 資料**

```bash
uv run python3 -c "
from app import create_app
app = create_app()
with app.app_context():
    from app.models import ChannelPrice, ProductBarcode
    licai_count = ChannelPrice.query.filter_by(channel='licai').count()
    bc_count    = ProductBarcode.query.count()
    print(f'licai channel_prices: {licai_count} 筆')
    print(f'product_barcodes 總計: {bc_count} 筆')
    # 抽查
    cp = ChannelPrice.query.filter_by(sku='D50100001', channel='licai').first()
    print(f'D50100001 licai price: {cp.price if cp else \"NOT FOUND\"}')
"
```

預期：licai channel_prices ≥ 90 筆，D50100001 licai price: 0.0

- [ ] **Step 5：Commit**

```bash
git add scripts/migrate_reference_data.py
git commit -m "feat(licai): 新增 import_licai 將品號/條碼/channel_prices 匯入 DB"
```

---

## Task 2：新增 `_confirm_licai` 至 detector.py

**Files:**
- Modify: `converters/detector.py`

**識別特徵：**
- Row 0 的第一個儲存格包含 `門市：`
- Row 0 的第三個儲存格（index 2）包含 `門市資訊：`
- Row 1 的第一個儲存格包含 `單號：`

- [ ] **Step 1：在 `detect_each()` 的 elif 鏈中，於 `_confirm_kadomo` 之前插入 licai 判斷**

找到 detector.py 的這一段（約第 37-40 行）：

```python
        elif _confirm_yodee(f):
            result[f] = "yodee"
        elif _confirm_kadomo(f):
            result[f] = "kadomo"
```

改成：

```python
        elif _confirm_yodee(f):
            result[f] = "yodee"
        elif _confirm_licai(f):
            result[f] = "licai"
        elif _confirm_kadomo(f):
            result[f] = "kadomo"
```

- [ ] **Step 2：在 detector.py 末尾新增 `_confirm_licai` 函式**

在 `_confirm_leage` 函式之後（約第 190-198 行末尾）附加：

```python

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
```

- [ ] **Step 3：手動測試偵測函式**

```bash
cd /Users/lung/Documents/企業顧問/02_客戶管理/捷捷寶寶粥/訂單格式轉換系統
uv run python3 -c "
from pathlib import Path
from converters.detector import _confirm_licai, detect_each

test_file = Path('麗兒采家/split_orders/0330麗兒采家採購單_濬詮_佳里.xlsx')
print('_confirm_licai:', _confirm_licai(test_file))

files = list(Path('麗兒采家/split_orders').glob('*.xlsx'))
result = detect_each(files)
for p, t in result.items():
    print(p.name, '->', t)
"
```

預期：全部 split_orders/*.xlsx 都回傳 `licai`

- [ ] **Step 4：Commit**

```bash
git add converters/detector.py
git commit -m "feat(licai): 新增 _confirm_licai 偵測函式"
```

---

## Task 3：實作 `converters/licai.py`

**Files:**
- Create: `converters/licai.py`

- [ ] **Step 1：建立 `converters/licai.py`**

```python
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
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

from converters.base import BaseConverter, ConversionResult, RowError
from app.utils import output_path

HEADERS = ["訂單編號", "收件人", "地址", "電話", "產品編號", "產品名稱", "數量", "單價", "備註", "備註.1"]

OUTPUT_FONT = Font(name="新細明體", size=12)

# 從「門市：佳里」擷取短名
_STORE_RE = re.compile(r"門市：(\S+)")
# 從「單號：UW26033008」擷取訂單編號
_ORDER_RE = re.compile(r"單號：(\S+)")
# 從「採購日期：2026/3/30」擷取日期 → YYYYMMDD
_DATE_RE  = re.compile(r"採購日期：(\d{4})/(\d{1,2})/(\d{1,2})")
# 從通路資料店名擷取短名：「麗兒采家-大墩旗艦店(...)」→「大墩」
_STORE_SHORT_RE = re.compile(r"麗兒采家-(.+?)店")


class LicaiConverter(BaseConverter):

    def __init__(self, repository, output_dir: Path):
        super().__init__(repository, output_dir)
        self.data_dir = Path(__file__).resolve().parent.parent / "麗兒采家"
        self._barcode_map: dict[str, str]   = {}  # barcode → sku
        self._name_map:    dict[str, str]   = {}  # sku → 品名
        self._store_map:   dict[str, dict]  = {}  # 短名 → {store_full, phone, address}

    @property
    def source_type(self) -> str:
        return "licai"

    # ── 覆寫：載入麗兒采家專屬對照表 ─────────────────────────────────────

    def _load_reference(self) -> None:
        # 條碼 → sku（由 DB licai channel_prices 過濾）
        self._barcode_map = self.repository.load_barcodes("licai")
        # sku → 品名（從 products 表）
        from app.models import Product
        self._name_map = {p.sku: p.name for p in Product.query.all()}
        # 通路資料：讀 xlsx，只取麗兒采家門市
        self._store_map = self._load_store_map()

    def _load_store_map(self) -> dict[str, dict]:
        """從通路資料.xlsx 建立 短名 → {store_full, phone, address}。"""
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

    # ── 驗證 ──────────────────────────────────────────────────────────────

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

    # ── 轉換 ──────────────────────────────────────────────────────────────

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

        # 解析 header（row 0 / row 1）
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

        # 通路資訊
        info        = self._store_map.get(store_short, {})
        store_full  = info.get("store_full", f"麗兒采家-{store_short}店")
        phone       = info.get("phone", "")
        address     = info.get("address", "")

        # 輸出 workbook
        wb_out = Workbook()
        ws_out = wb_out.active
        ws_out.append(HEADERS)

        success = 0
        errors: list[RowError] = []

        # 資料從 row 3 開始（index 2 = empty，index 3+ = data）
        for row_idx, row in enumerate(rows[3:], start=4):
            if row[0] is None:  # 略過空列
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

        # 樣式
        for row in ws_out.iter_rows():
            for cell in row:
                cell.font = OUTPUT_FONT

        # 輸出檔名：MMDD 麗采{短名}.xlsx
        mmdd = order_date[4:] if len(order_date) == 8 else "0000"
        from datetime import datetime
        date_obj = (
            datetime.strptime(order_date, "%Y%m%d") if order_date else datetime.today()
        )
        out_name = f"{mmdd} 麗采{store_short}.xlsx"
        out_path = output_path(self.output_dir.parent, self.source_type, date_obj, out_name)

        wb_out.save(out_path)
        return out_path, success, len(errors), errors, order_date
```

- [ ] **Step 2：手動執行轉換器確認輸出**

```bash
cd /Users/lung/Documents/企業顧問/02_客戶管理/捷捷寶寶粥/訂單格式轉換系統
uv run python3 -c "
from pathlib import Path
from app import create_app
from app.repositories.product_repo import ProductRepository
from converters.licai import LicaiConverter

app = create_app()
with app.app_context():
    repo = ProductRepository()
    conv = LicaiConverter(repo, Path('output'))
    files = list(Path('麗兒采家/split_orders').glob('*.xlsx'))
    result = conv.convert(files)
    print('success:', result.success_count)
    print('fail:   ', result.fail_count)
    print('outputs:', [p.name for p in result.output_files])
    if result.errors:
        for e in result.errors[:5]:
            print('ERROR:', e)
"
```

預期：success_count ≥ 200，fail_count = 0（若有新條碼才會有 error）

- [ ] **Step 3：Commit**

```bash
git add converters/licai.py
git commit -m "feat(licai): 實作 LicaiConverter"
```

---

## Task 4：登錄 licai 至 services.py、utils.py、models.py

**Files:**
- Modify: `app/services.py`
- Modify: `app/utils.py`
- Modify: `app/models.py`

- [ ] **Step 1：在 `app/services.py` 的 `_convert_group` 中登錄 LicaiConverter**

找到 `_convert_group` 的 import 區塊和 CONVERTER_MAP（約第 147-157 行）：

```python
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
```

改成：

```python
    from converters.shopee      import ShopeeConverter
    from converters.a1baby      import A1BabyConverter
    from converters.leage       import LeageConverter
    from converters.a1leage     import A1LeageConverter
    from converters.jjofficial  import JJOfficialConverter
    from converters.yodee       import YodeeConverter
    from converters.kadomo      import KadomoConverter
    from converters.licai       import LicaiConverter

    CONVERTER_MAP = {
        "shopee":     ShopeeConverter,
        "a1baby":     A1BabyConverter,
        "leage":      LeageConverter,
        "a1leage":    A1LeageConverter,
        "jjofficial": JJOfficialConverter,
        "yodee":      YodeeConverter,
        "kadomo":     KadomoConverter,
        "licai":      LicaiConverter,
    }
```

- [ ] **Step 2：在 `app/utils.py` 的 SOURCE_LABELS 和 SOURCE_COLORS 新增 licai**

`SOURCE_LABELS`（約第 6-14 行）加入：

```python
SOURCE_LABELS = {
    "shopee":     "蝦皮",
    "a1baby":     "婦幼展",
    "leage":      "樂齡網",
    "a1leage":    "樂齡官網",
    "jjofficial": "捷捷官網",
    "yodee":      "優迪通路",
    "kadomo":     "卡多摩",
    "licai":      "麗兒采家",
}
```

`SOURCE_COLORS`（約第 16-24 行）加入：

```python
SOURCE_COLORS = {
    "shopee":     "warning",
    "a1baby":     "info",
    "leage":      "success",
    "a1leage":    "danger",
    "jjofficial": "primary",
    "yodee":      "secondary",
    "kadomo":     "dark",
    "licai":      "warning",
}
```

（注意：licai 與蝦皮同用 warning 橘色無妨，或改用 `"light"` 亦可。）

- [ ] **Step 3：更新 `app/models.py` 的 `to_dict()` 中 source_label 查找字典**

找到 `ConversionLog.to_dict()`（約第 37 行）：

```python
            "source_label":  {"shopee": "蝦皮", "a1baby": "婦幼展", "leage": "樂齡網"}.get(self.source_type, self.source_type),
```

改為引用 utils.SOURCE_LABELS，避免重複維護：

```python
            "source_label":  __import__("app.utils", fromlist=["SOURCE_LABELS"]).SOURCE_LABELS.get(self.source_type, self.source_type),
```

（注意：此處為 inline import，避免循環 import；或在 models.py 頂端加 `from app.utils import SOURCE_LABELS`，但需確認無循環依賴。推薦在頂端加 import。）

實際修改：在 `app/models.py` 第 4 行（`from app import db` 之後）加入：

```python
from app.utils import SOURCE_LABELS
```

然後將第 37 行改為：

```python
            "source_label":  SOURCE_LABELS.get(self.source_type, self.source_type),
```

- [ ] **Step 4：確認系統可正常啟動**

```bash
uv run python3 -c "
from app import create_app
app = create_app()
with app.app_context():
    from app.utils import SOURCE_LABELS
    print('licai label:', SOURCE_LABELS.get('licai'))
    from app.models import ConversionLog
    print('models OK')
"
```

預期：`licai label: 麗兒采家`，無錯誤

- [ ] **Step 5：Commit**

```bash
git add app/services.py app/utils.py app/models.py
git commit -m "feat(licai): 登錄 licai 轉換器至 services/utils/models"
```

---

## Task 5：更新 help.html 與 CHANGELOG.md

**Files:**
- Modify: `templates/help.html`
- Modify: `CHANGELOG.md`

- [ ] **Step 1：在 `templates/help.html` 的卡多摩說明區塊之後新增麗兒采家卡片**

找到 `<!-- 通用說明 -->` 的位置（約第 136 行），在其前方插入：

```html
    <!-- 麗兒采家 -->
    <div class="glass-card p-4 mb-3">
      <div class="d-flex align-items-center gap-2 mb-3">
        <span class="badge-source badge-licai px-3 py-1" style="font-size:.85rem">麗兒采家</span>
        <h6 class="fw-bold mb-0">麗兒采家嬰童館採購單</h6>
      </div>
      <table class="table table-sm table-borderless mb-0" style="font-size:.9rem">
        <tbody>
          <tr><td class="text-muted" style="width:7rem">輸入格式</td><td>已拆分的個別門市 xlsx（使用 <code>split_orders.py</code> 拆分後放入資料夾）</td></tr>
          <tr><td class="text-muted">識別特徵</td><td>前兩列含「門市：」、「門市資訊：」、「單號：」等標頭</td></tr>
          <tr><td class="text-muted">對照表</td><td><code>麗兒采家/品號資料_含條碼.csv</code>、<code>通路資料.xlsx</code></td></tr>
          <tr><td class="text-muted">門市資訊</td><td>訂單編號格式：單號（如 <code>UW26033008</code>），收件人為完整店名含備注（如「中午12點後到貨」）</td></tr>
          <tr><td class="text-muted">每檔獨立</td><td>一個門市 xlsx 輸出一個結果檔</td></tr>
          <tr><td class="text-muted">輸出檔案</td><td>{MMDD} 麗采{門市}.xlsx（如 <code>0330 麗采佳里.xlsx</code>）</td></tr>
        </tbody>
      </table>
    </div>

```

同時，將頂部說明文字從「7 種」更新為「8 種」（約第 15 行）：

```html
      <p class="text-muted small mb-0">目前支援以下 8 種訂單來源，系統會自動偵測格式，無需手動選擇。</p>
```

- [ ] **Step 2：新增 badge-licai CSS（若 custom.css 有 badge 樣式的話）**

```bash
grep -n "badge-kadomo\|badge-yodee" templates/base.html static/css/custom.css 2>/dev/null | head
```

若 `badge-kadomo` 定義在 CSS 中，在同處加入：

```css
.badge-licai   { background: #f59e0b; color: #fff; border-radius: 4px; }
```

若 badge 樣式定義在 `base.html` 的 `<style>` 區塊中，也比照加入。

- [ ] **Step 3：在 CHANGELOG.md 頂端新增版本區塊**

在 `## v5 — 2026-04-10` 之前插入：

```markdown
## v6 — 2026-04-09

### 新功能
- **麗兒采家（licai）轉換模組**
  - 輸入：已拆分的個別門市 xlsx（`MMDD麗兒采家採購單_XXX_門市.xlsx`）
  - 輸出：`{MMDD} 麗采{門市}.xlsx`，欄位：訂單編號、收件人、地址、電話、產品編號、產品名稱、數量、單價、備註、備註.1
  - 條碼查品號來自 DB（品號資料_含條碼.csv 匯入 `product_barcodes` + `channel_prices(licai)`）
  - 門市完整名稱、地址、電話查自 `麗兒采家/通路資料.xlsx`
  - 識別特徵：row 0 含「門市：」與「門市資訊：」，row 1 含「單號：」
- **品號資料庫**：新增 licai 通路品號匯入（scripts/migrate_reference_data.py）

---

```

- [ ] **Step 4：Commit**

```bash
git add templates/help.html CHANGELOG.md
# 若有修改 CSS
git add static/css/custom.css
git commit -m "docs(licai): 更新 help.html 與 CHANGELOG"
```

---

## Task 6：整合驗證

**Files:** — (只執行，不修改)

- [ ] **Step 1：啟動伺服器，測試完整掃描→轉換流程**

```bash
uv run python3 run.py &
sleep 2
# 在瀏覽器開啟 http://localhost:5099
```

在 UI 中：
1. 輸入掃描路徑：`麗兒采家/split_orders`
2. 點「掃描」→ 確認全部 8 個 xlsx 都被識別為「麗兒采家」
3. 全選 → 「開始轉換」
4. 確認輸出：`output/licai/2026-03/` 內有 8 個 xlsx 檔

- [ ] **Step 2：確認輸出格式**

```bash
uv run python3 -c "
import openpyxl
from pathlib import Path
outputs = list(Path('output/licai').rglob('*.xlsx'))
for p in outputs[:3]:
    wb = openpyxl.load_workbook(p)
    ws = wb.active
    print(p.name)
    for row in ws.iter_rows(max_row=3, values_only=True):
        print(' ', row)
"
```

預期：header row 為 `('訂單編號', '收件人', '地址', '電話', '產品編號', '產品名稱', '數量', '單價', '備註', '備註.1')`，資料列有正確品號（如 `D50100001`）。

- [ ] **Step 3：確認轉換記錄出現在 DB**

```bash
uv run python3 -c "
from app import create_app
app = create_app()
with app.app_context():
    from app.models import ConversionLog
    logs = ConversionLog.query.filter_by(source_type='licai').all()
    for log in logs:
        print(log.to_dict())
"
```

預期：有 source_label 為 `麗兒采家` 的紀錄。

- [ ] **Step 4：最終 commit + push**

```bash
git add -A
git status  # 確認沒有意外的檔案
git commit -m "feat: 完成麗兒采家轉換器整合驗證"
```

---

## 自查清單（Self-Review）

1. **Spec coverage：** 所有需求均已覆蓋 ✓
   - 條碼查品號 → Task 3（barcode_map）
   - 門市完整名稱/地址/電話 → Task 3（store_map）
   - DB 匯入 → Task 1
   - 偵測器 → Task 2
   - 登錄系統 → Task 4
   - 說明頁面 → Task 5

2. **型別一致性：**
   - `_barcode_map: dict[str, str]` — Task 1 匯入 barcodes，Task 3 使用 `self._barcode_map.get(barcode, "")`
   - `_store_map: dict[str, dict]` — dict value 有 `store_full`, `phone`, `address` 三個 key
   - `output_path()` 第一個參數是 `self.output_dir.parent`（BASE_DIR），因為 `output_path` 本身會在內部組成 `output/{source_type}/...`

3. **Placeholder scan：** 無 TBD / TODO ✓

4. **output_dir 使用說明：** `LicaiConverter.__init__` 的 `output_dir` 參數由 `services.py` 傳入 `OUTPUT_DIR`（= `BASE_DIR / "output"`），而 `output_path(base_dir, ...)` 的第一個參數需要 `BASE_DIR`（非 `output_dir`）。在 `_process_one` 中使用 `self.output_dir.parent` 取得 BASE_DIR。
