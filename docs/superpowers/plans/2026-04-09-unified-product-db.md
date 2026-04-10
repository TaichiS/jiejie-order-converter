# 品號資料統一資料庫 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 將散落在多個 CSV/JSON 檔案中的品號參考資料，整合進 SQLAlchemy 關聯式資料庫，支援各通路獨立定價，並抽象化查詢層以便日後無痛切換至 Supabase PostgreSQL。

**Architecture:** 新增 4 張關聯表（products / product_aliases / product_barcodes / channel_prices），以 `ProductRepository` 封裝所有查詢邏輯。Converter 改為接收 Repository 而非 CSV Path。抽象層即 SQLAlchemy ORM 本身——切換 Supabase 只需改 `DATABASE_URI`，不需另寫 Repository 實作。

**Tech Stack:** Flask-SQLAlchemy（已有）、SQLite（現階段）、PostgreSQL/Supabase（未來）、Python csv/json（匯入腳本）

---

## 檔案結構

| 狀態 | 路徑 | 責任 |
|------|------|------|
| 修改 | `app/models.py` | 新增 Product、ProductAlias、ProductBarcode、ChannelPrice 四個 Model |
| 新增 | `app/repositories/product_repo.py` | ProductRepository：封裝所有品號查詢邏輯 |
| 新增 | `app/repositories/__init__.py` | 空 init |
| 新增 | `scripts/migrate_reference_data.py` | 一次性匯入腳本：CSV + JSON → DB |
| 修改 | `converters/base.py` | `__init__` 改接 ProductRepository；`_load_reference()` 改從 repo 載入 |
| 修改 | `converters/kadomo.py` | `_load_reference()` 改從 repo 取條碼與定價 |
| 修改 | `app/services.py` | 移除 CSV 路徑常數；建立並注入 ProductRepository |
| 修改 | `app/__init__.py` | 無需改動（db 物件已存在） |
| 修改 | `templates/help.html` | 更新「品號資料」說明，改為「集中在系統資料庫」 |

---

## Task 1: 新增 DB Models

**Files:**
- Modify: `app/models.py`

- [ ] **Step 1: 在 `app/models.py` 末端新增四個 Model**

```python
class Product(db.Model):
    __tablename__ = "products"

    sku        = db.Column(db.String(20),  primary_key=True)   # 品號
    name       = db.Column(db.String(100), nullable=False)      # 主品名
    category   = db.Column(db.String(50))                       # 類別
    quantity   = db.Column(db.Integer,  default=1)              # 份數
    pack_size  = db.Column(db.Integer)                          # 包數（nullable）
    unit_price = db.Column(db.Float)                            # 份數價格（nullable）
    erp_source = db.Column(db.String(50))                       # 來源（鼎新等）

    aliases    = db.relationship("ProductAlias",   back_populates="product",
                                 cascade="all, delete-orphan")
    barcodes   = db.relationship("ProductBarcode", back_populates="product",
                                 cascade="all, delete-orphan")
    prices     = db.relationship("ChannelPrice",   back_populates="product",
                                 cascade="all, delete-orphan")


class ProductAlias(db.Model):
    __tablename__ = "product_aliases"

    id      = db.Column(db.Integer, primary_key=True)
    sku     = db.Column(db.String(20), db.ForeignKey("products.sku"), nullable=False)
    alias   = db.Column(db.String(100), nullable=False, unique=True)

    product = db.relationship("Product", back_populates="aliases")


class ProductBarcode(db.Model):
    __tablename__ = "product_barcodes"

    id      = db.Column(db.Integer, primary_key=True)
    sku     = db.Column(db.String(20), db.ForeignKey("products.sku"), nullable=False)
    barcode = db.Column(db.String(50), nullable=False, unique=True)

    product = db.relationship("Product", back_populates="barcodes")


class ChannelPrice(db.Model):
    __tablename__ = "channel_prices"

    id      = db.Column(db.Integer, primary_key=True)
    sku     = db.Column(db.String(20), db.ForeignKey("products.sku"), nullable=False)
    channel = db.Column(db.String(20), nullable=False)   # shopee|a1baby|leage|a1leage|jjofficial|yodee|kadomo
    price   = db.Column(db.Float,      nullable=False)

    product = db.relationship("Product", back_populates="prices")

    __table_args__ = (
        db.UniqueConstraint("sku", "channel", name="uq_sku_channel"),
    )
```

- [ ] **Step 2: 啟動 Flask 讓 `db.create_all()` 建立新表**

```bash
cd /Users/lung/Documents/企業顧問/02_客戶管理/捷捷寶寶粥/訂單格式轉換系統
uv run python3 -c "
from app import create_app
app = create_app()
print('Tables created OK')
"
```

Expected: 無 Exception，印出 `Tables created OK`

---

## Task 2: 建立 ProductRepository

**Files:**
- Create: `app/repositories/__init__.py`
- Create: `app/repositories/product_repo.py`

- [ ] **Step 1: 建立空的 `__init__.py`**

```python
# app/repositories/__init__.py
```

- [ ] **Step 2: 建立 `product_repo.py`**

```python
"""
app/repositories/product_repo.py
ProductRepository：封裝所有品號查詢邏輯。
呼叫端（Converter）只看這個介面，不直接操作 ORM 或 CSV。
切換 Supabase：只改 DATABASE_URI，此類別不需改動。
"""
from __future__ import annotations

from difflib import get_close_matches

from app.models import Product, ProductAlias, ProductBarcode, ChannelPrice


class ProductRepository:
    """
    提供 Converter 所需的品號查詢方法。
    所有方法回傳與原 CSV row dict 相容的 dict（含 品號、品名、商品結帳價 等欄位），
    確保現有 Converter 邏輯無需改動。
    """

    def load_channel(self, channel: str) -> tuple[dict[str, dict], dict[str, dict]]:
        """
        載入指定通路的所有品號資料，回傳 (product_map, code_map)。
        product_map: 品名（含別名）→ row_dict
        code_map:    代碼前綴（如 "1-08"）→ row_dict
        row_dict 格式與舊 CSV DictReader 一致：
          {'品號','品名','類別','份數','包數','份數價格','商品結帳價','來源'}
        """
        import re
        _CODE_RE = re.compile(r'^(\d+[A-Z]?-[A-Z]?\d+)')

        # 取得該通路所有定價
        price_rows = ChannelPrice.query.filter_by(channel=channel).all()
        sku_set    = {cp.sku for cp in price_rows}
        price_map  = {cp.sku: cp.price for cp in price_rows}

        # 批次取商品主資料
        products   = Product.query.filter(Product.sku.in_(sku_set)).all()

        product_map: dict[str, dict] = {}
        code_map:    dict[str, dict] = {}

        for prod in products:
            price   = price_map.get(prod.sku, 0)
            row     = _to_dict(prod, price)
            # 主品名
            product_map[prod.name] = row
            # 別名
            for alias in prod.aliases:
                product_map[alias.alias] = row
            # 代碼前綴索引
            m = _CODE_RE.match(prod.name)
            if m:
                code_map[m.group(1)] = row

        return product_map, code_map

    def load_barcodes(self, channel: str) -> dict[str, str]:
        """
        回傳指定通路的條碼對照表：barcode → sku。
        目前只有 kadomo 通路使用。
        """
        # 取得該通路有定價的 sku 集合
        sku_set = {
            cp.sku
            for cp in ChannelPrice.query.filter_by(channel=channel).all()
        }
        result = {}
        for bc in ProductBarcode.query.all():
            if bc.sku in sku_set:
                result[bc.barcode] = bc.sku
        return result

    def get_price(self, sku: str, channel: str) -> float:
        """回傳指定通路的品號單價，找不到回傳 0.0。"""
        cp = ChannelPrice.query.filter_by(sku=sku, channel=channel).first()
        return float(cp.price) if cp else 0.0

    def fuzzy_lookup(self, query: str, channel: str,
                     scope_prefix: str | None = None,
                     amount: float = 0) -> tuple[dict | None, list[dict]]:
        """
        多層查找，與 BaseConverter._lookup_by_name() 邏輯一致，
        但直接查 DB（不需先 load_channel）。
        回傳 (最佳匹配 | None, 候選清單)。
        注意：Converter 仍使用 in-memory product_map，此方法供未來 API 用。
        """
        product_map, _ = self.load_channel(channel)
        pool = (
            {k: v for k, v in product_map.items()
             if str(v.get("品號", "")).startswith(scope_prefix)}
            if scope_prefix else product_map
        )
        prod = pool.get(query.strip())
        if prod:
            return prod, []
        names      = list(pool.keys())
        close_keys = get_close_matches(query.strip(), names, n=3, cutoff=0.35)
        candidates = [dict(pool[k], match_type="name") for k in close_keys]
        if amount > 0:
            seen = {c["品號"] for c in candidates}
            for p in pool.values():
                cp = float(p.get("商品結帳價") or 0)
                if cp > 0 and amount % cp == 0 and p["品號"] not in seen:
                    seen.add(p["品號"])
                    candidates.append(dict(p, match_type="price"))
        return None, candidates[:5]


def _to_dict(prod: Product, price: float) -> dict:
    """將 Product ORM 物件轉為與 CSV row 相容的 dict。"""
    return {
        "品號":       prod.sku,
        "品名":       prod.name,
        "類別":       prod.category or "",
        "份數":       str(prod.quantity or 1),
        "包數":       str(prod.pack_size) if prod.pack_size else "",
        "份數價格":   str(prod.unit_price) if prod.unit_price else "",
        "商品結帳價": str(price),
        "來源":       prod.erp_source or "",
    }
```

- [ ] **Step 3: 確認可以 import**

```bash
uv run python3 -c "
from app import create_app
app = create_app()
with app.app_context():
    from app.repositories.product_repo import ProductRepository
    repo = ProductRepository()
    pm, cm = repo.load_channel('shopee')
    print('load_channel OK, rows:', len(pm))
"
```

Expected: `load_channel OK, rows: 0`（此時 DB 尚無資料，0 是正常的）

---

## Task 3: 資料匯入腳本

**Files:**
- Create: `scripts/migrate_reference_data.py`

- [ ] **Step 1: 建立 `scripts/migrate_reference_data.py`**

```python
#!/usr/bin/env python3
"""
scripts/migrate_reference_data.py
一次性資料匯入腳本：將所有 CSV/JSON 品號資料寫入 DB。

執行：
    cd /Users/lung/.../訂單格式轉換系統
    uv run python3 scripts/migrate_reference_data.py
"""
from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

# 必須先設定 sys.path 才能 import app
sys.path.insert(0, str(BASE_DIR))

from app import create_app, db
from app.models import Product, ProductAlias, ProductBarcode, ChannelPrice

# ── 各 CSV 對應的通路清單 ─────────────────────────────────────────────────────
# reference/ 的品號同時供應 shopee / a1baby / leage / kadomo（價格另覆蓋）
REFERENCE_CHANNELS = ["shopee", "a1baby", "leage"]

CSV_SOURCES = [
    {
        "path":     BASE_DIR / "reference" / "品號資料.csv",
        "channels": REFERENCE_CHANNELS,
    },
    {
        "path":     BASE_DIR / "A1樂齡官網" / "品號資料.csv",
        "channels": ["a1leage"],
    },
    {
        "path":     BASE_DIR / "捷捷寶寶粥官網" / "品號資料.csv",
        "channels": ["jjofficial"],
    },
    {
        "path":     BASE_DIR / "優迪通路" / "品號資料.csv",
        "channels": ["yodee"],
    },
]

BARCODE_JSON  = BASE_DIR / "卡多摩" / "條碼對照表.json"   # barcode → sku
KADOMO_PRICE_JSON = BASE_DIR / "卡多摩" / "單價對照表.json"  # sku → price


def import_csv(path: Path, channels: list[str]) -> None:
    """匯入單份 CSV。主行建 Product，別名行建 ProductAlias，各通路建 ChannelPrice。"""
    main_by_sku: dict[str, dict] = {}  # sku → main row

    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))

    # 第一遍：收集主行（商品結帳價有值）
    for row in rows:
        sku   = row.get("品號", "").strip()
        price = row.get("商品結帳價", "").strip()
        if sku and price:
            main_by_sku[sku] = row

    # 第二遍：寫入 DB
    for row in rows:
        name  = row.get("品名", "").strip()
        sku   = row.get("品號", "").strip()
        price = row.get("商品結帳價", "").strip()

        if not name:
            continue

        # 別名行（商品結帳價空白）
        if not price and sku in main_by_sku:
            _upsert_alias(sku, name)
            continue

        if not sku:
            continue

        # 主行：upsert Product
        prod = Product.query.get(sku)
        if prod is None:
            prod = Product(sku=sku)
            db.session.add(prod)

        prod.name       = name
        prod.category   = row.get("類別", "").strip() or None
        prod.quantity   = _int(row.get("份數"))
        prod.pack_size  = _int(row.get("包數"))
        prod.unit_price = _float(row.get("份數價格"))
        prod.erp_source = row.get("來源", "").strip() or None

        # ChannelPrice
        for ch in channels:
            _upsert_price(sku, ch, float(price))

    db.session.flush()


def _upsert_alias(sku: str, alias: str) -> None:
    exists = ProductAlias.query.filter_by(sku=sku, alias=alias).first()
    if not exists:
        db.session.add(ProductAlias(sku=sku, alias=alias))


def _upsert_price(sku: str, channel: str, price: float) -> None:
    cp = ChannelPrice.query.filter_by(sku=sku, channel=channel).first()
    if cp:
        cp.price = price
    else:
        db.session.add(ChannelPrice(sku=sku, channel=channel, price=price))


def import_kadomo_barcodes() -> None:
    """將條碼對照表.json 寫入 product_barcodes。"""
    barcode_map: dict[str, str] = json.loads(BARCODE_JSON.read_text(encoding="utf-8"))
    for barcode, sku in barcode_map.items():
        exists = ProductBarcode.query.filter_by(barcode=barcode).first()
        if not exists:
            db.session.add(ProductBarcode(sku=sku, barcode=barcode))
    db.session.flush()


def import_kadomo_prices() -> None:
    """將單價對照表.json 寫入 channel_prices（channel='kadomo'）。"""
    price_map: dict[str, float] = json.loads(KADOMO_PRICE_JSON.read_text(encoding="utf-8"))
    for sku, price in price_map.items():
        _upsert_price(sku, "kadomo", float(price))
    db.session.flush()


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

        db.session.commit()

        # 統計
        print(f"\n匯入完成：")
        print(f"  products       : {Product.query.count()} 筆")
        print(f"  product_aliases: {ProductAlias.query.count()} 筆")
        print(f"  product_barcodes:{ProductBarcode.query.count()} 筆")
        print(f"  channel_prices : {ChannelPrice.query.count()} 筆")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 執行匯入腳本**

```bash
cd /Users/lung/Documents/企業顧問/02_客戶管理/捷捷寶寶粥/訂單格式轉換系統
uv run python3 scripts/migrate_reference_data.py
```

Expected（數字僅供參考）：
```
開始匯入品號資料...
  匯入 reference/品號資料.csv → 通路：['shopee', 'a1baby', 'leage']
  匯入 A1樂齡官網/品號資料.csv → 通路：['a1leage']
  匯入 捷捷寶寶粥官網/品號資料.csv → 通路：['jjofficial']
  匯入 優迪通路/品號資料.csv → 通路：['yodee']
  匯入 卡多摩/條碼對照表.json
  匯入 卡多摩/單價對照表.json → 通路：kadomo

匯入完成：
  products       : XXX 筆
  product_aliases: XXX 筆
  product_barcodes: 96 筆
  channel_prices : XXX 筆
```

- [ ] **Step 3: 確認資料正確性（抽查）**

```bash
uv run python3 -c "
from app import create_app
from app.models import ChannelPrice, ProductBarcode
app = create_app()
with app.app_context():
    # 同品號跨通路定價應不同
    from app.models import ChannelPrice
    d50 = ChannelPrice.query.filter_by(sku='D50100001').all()
    for cp in d50:
        print(cp.channel, cp.price)
    # 條碼查品號
    bc = ProductBarcode.query.filter_by(barcode='4710586220018').first()
    print('條碼 4710586220018 →', bc.sku if bc else '找不到')
"
```

Expected:
- D50100001 在 shopee/a1baby/leage 為 35.0，在 kadomo 為 26.0
- 條碼查品號得到 `D50100001`

---

## Task 4: 更新 BaseConverter 改用 Repository

**Files:**
- Modify: `converters/base.py`

- [ ] **Step 1: 修改 `__init__` 簽名並更新 `_load_reference()`**

將 `base.py` 中的 `__init__` 與 `_load_reference` 改為：

```python
def __init__(self, repository: "ProductRepository", output_dir: Path):
    from app.repositories.product_repo import ProductRepository
    self.repository  = repository
    self.output_dir  = output_dir
    self._product_map: dict[str, dict] = {}
    self._code_map:    dict[str, dict] = {}

def _load_reference(self) -> None:
    """從 Repository 載入指定通路的品號資料到 in-memory cache。"""
    self._product_map, self._code_map = self.repository.load_channel(self.source_type)
```

> **注意**：移除 `self.reference_csv = reference_csv` 那行，以及 `import csv` 與舊的 `_load_reference` 實作（約 40 行）。其餘 `_lookup_by_name`、`_lookup_product`、`convert` 等方法**完全不動**。

- [ ] **Step 2: 移除不再需要的 `import csv`（若其他方法都不用的話）**

確認 `base.py` 中是否還有其他地方使用 `csv` 模組。若無，移除該行。

```bash
grep -n "^import csv\|csv\." converters/base.py
```

若只有原 `_load_reference` 用到，則刪除 `import csv`。

---

## Task 5: 更新 KadomoConverter 改用 Repository

**Files:**
- Modify: `converters/kadomo.py`

- [ ] **Step 1: 修改 `__init__` 與 `_load_reference`**

```python
def __init__(self, repository, output_dir: Path):
    # 注意：不呼叫 super().__init__() 傳 reference_csv，
    # 但仍需初始化 output_dir 與 maps
    from converters.base import BaseConverter
    self.repository  = repository
    self.output_dir  = output_dir
    self._product_map: dict = {}
    self._code_map:    dict = {}
    self._barcode_map: dict[str, str]   = {}
    self._price_map:   dict[str, float] = {}
    self._store_map:   dict[str, str]   = {}
    self._stores:      dict[str, dict]  = {}
    # 通路資料（stores/warehouses）仍從 JSON 讀，不進 DB
    self.data_dir = Path(__file__).resolve().parent.parent / "卡多摩"

def _load_reference(self) -> None:
    # 條碼與定價從 Repository 取
    self._barcode_map = self.repository.load_barcodes("kadomo")
    # 重建 price_map：sku → price
    from app.models import ChannelPrice
    self._price_map = {
        cp.sku: float(cp.price)
        for cp in ChannelPrice.query.filter_by(channel="kadomo").all()
    }
    # 通路資料仍讀 JSON（stores/warehouses 是營運資料，不進 DB）
    import json
    data            = json.loads((self.data_dir / "通路資料.json").read_text(encoding="utf-8"))
    self._store_map = data.get("warehouse_mapping", {})
    self._stores    = {s["店名"]: s for s in data.get("stores", [])}
```

- [ ] **Step 2: 確認 `_process_one` 中的 `self.data_dir` 引用已不存在（或仍正確）**

`_process_one` 只用 `self._barcode_map` 與 `self._price_map`，不直接用 `self.data_dir`，應無問題。

---

## Task 6: 更新 services.py

**Files:**
- Modify: `app/services.py`

- [ ] **Step 1: 移除 CSV 路徑常數，改為建立 Repository**

刪除：
```python
REFERENCE_CSV            = BASE_DIR / "reference" / "品號資料.csv"
A1LEAGE_REFERENCE_CSV    = BASE_DIR / "A1樂齡官網" / "品號資料.csv"
JJOFFICIAL_REFERENCE_CSV = BASE_DIR / "捷捷寶寶粥官網" / "品號資料.csv"
KADOMO_REF               = BASE_DIR / "卡多摩" / "條碼對照表.json"
```

在 `_convert_group` 中，移除 `ref_csv` 的邏輯，改為：

```python
def _convert_group(source_type: str, files: list[Path], operator: str) -> ConversionLog:
    from converters.shopee      import ShopeeConverter
    from converters.a1baby      import A1BabyConverter
    from converters.leage       import LeageConverter
    from converters.a1leage     import A1LeageConverter
    from converters.jjofficial  import JJOfficialConverter
    from converters.yodee       import YodeeConverter
    from converters.kadomo      import KadomoConverter
    from app.repositories.product_repo import ProductRepository

    CONVERTER_MAP = {
        "shopee":     ShopeeConverter,
        "a1baby":     A1BabyConverter,
        "leage":      LeageConverter,
        "a1leage":    A1LeageConverter,
        "jjofficial": JJOfficialConverter,
        "yodee":      YodeeConverter,
        "kadomo":     KadomoConverter,
    }

    repo      = ProductRepository()
    converter = CONVERTER_MAP[source_type](repo, OUTPUT_DIR)
    result    = converter.convert(files)
    # ... 以下歸檔與 DB 寫入邏輯不變 ...
```

- [ ] **Step 2: 確認 `routes.py` 中是否有直接引用 CSV 路徑（需同步移除）**

```bash
grep -n "REFERENCE_CSV\|reference_csv\|品號資料.csv" app/routes.py
```

若有，根據上下文改為使用 Repository 或直接移除。

---

## Task 7: Supabase 遷移準備

**Files:**
- Modify: `app/__init__.py`（新增環境變數讀取）

- [ ] **Step 1: 讓 DATABASE_URI 可由環境變數覆蓋**

```python
import os

# 原本：
# app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{BASE_DIR / 'data' / 'conversion.db'}"

# 改為：
default_uri = f"sqlite:///{BASE_DIR / 'data' / 'conversion.db'}"
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL", default_uri)
```

- [ ] **Step 2: 新增 `.env.example` 說明切換方式**

```bash
# .env.example（新增此檔案）
# SQLite（預設，本機開發）
# DATABASE_URL=sqlite:///data/conversion.db

# Supabase PostgreSQL（上線時改用此行，填入 Supabase 的 Session Mode 連線字串）
# DATABASE_URL=postgresql://postgres.[project-ref]:[password]@aws-0-[region].pooler.supabase.com:5432/postgres
```

> 切換 Supabase 時唯一要做的事：設定環境變數 `DATABASE_URL`，其餘程式碼不需改動。

---

## Task 8: 整合測試

- [ ] **Step 1: 啟動系統，確認現有轉換流程正常**

```bash
cd /Users/lung/Documents/企業顧問/02_客戶管理/捷捷寶寶粥/訂單格式轉換系統
uv run python3 run.py
```

用瀏覽器開啟 http://localhost:5099，進行一次各通路的轉換測試（或用現有的測試檔案）。

- [ ] **Step 2: 確認品號查找結果與舊版一致**

```bash
uv run python3 -c "
from app import create_app
app = create_app()
with app.app_context():
    from app.repositories.product_repo import ProductRepository
    repo = ProductRepository()
    pm, cm = repo.load_channel('shopee')
    # 確認 '0-1純米泥' 可查到
    row = pm.get('0-1純米泥')
    print('shopee 0-1純米泥:', row)
    pm2, _ = repo.load_channel('kadomo')
    # 確認 kadomo 定價是 26，不是 35
    row2 = pm2.get('0-1純米泥') if pm2.get('0-1純米泥') else None
    # kadomo 不一定有品名，改查 price_map
    from app.models import ChannelPrice
    cp = ChannelPrice.query.filter_by(sku='D50100001', channel='kadomo').first()
    print('kadomo D50100001 price:', cp.price if cp else 'N/A')
"
```

Expected:
- shopee `0-1純米泥` 商品結帳價為 `35.0`
- kadomo `D50100001` price 為 `26.0`

- [ ] **Step 3: 更新 `templates/help.html` 中的品號資料說明**

將 a1leage 那行：
```html
<tr><td class="text-muted">品號資料</td><td>使用獨立的 <code>A1樂齡官網/品號資料.csv</code>（與其他來源分開）</td></tr>
```
改為：
```html
<tr><td class="text-muted">品號資料</td><td>集中管理於系統資料庫（各通路獨立定價）</td></tr>
```

---

## Self-Review

**Spec coverage check:**
- ✅ 四張關聯表（products / aliases / barcodes / channel_prices）
- ✅ 各通路獨立定價（channel_prices）
- ✅ 條碼整合（product_barcodes）
- ✅ Repository 封裝查詢邏輯
- ✅ SQLite → Supabase 只需改 URI
- ✅ Converter 不直接接觸 CSV/JSON
- ✅ 卡多摩通路資料（stores/warehouses）保留 JSON，不強行塞入 DB

**風險提醒：**
1. `BaseConverter.__init__` 簽名變更是 breaking change——Task 4~6 必須一起完成，中間狀態會導致系統無法啟動。
2. `KadomoConverter` 的 `super().__init__()` 呼叫需要確認是否傳對參數（Task 5 要注意）。
3. `ChannelPrice.query` 在 `_load_reference` 中被呼叫時，必須在 Flask app context 內執行——Converter 在 `services._convert_group` 中呼叫，此處已有 context，無需額外處理。
