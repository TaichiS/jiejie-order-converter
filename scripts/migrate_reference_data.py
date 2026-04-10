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
sys.path.insert(0, str(BASE_DIR))

from app import create_app, db
from app.models import Product, ProductAlias, ProductBarcode, ChannelPrice

CSV_SOURCES = [
    {
        "path":     BASE_DIR / "reference" / "品號資料.csv",
        "channels": ["shopee", "a1baby", "leage"],
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

BARCODE_JSON      = BASE_DIR / "卡多摩" / "條碼對照表.json"
KADOMO_PRICE_JSON = BASE_DIR / "卡多摩" / "單價對照表.json"
LICAI_CSV         = BASE_DIR / "麗兒采家" / "品號資料_含條碼.csv"


def import_csv(path: Path, channels: list[str]) -> None:
    """匯入單份 CSV。主行建 Product，別名行建 ProductAlias，各通路建 ChannelPrice。"""
    main_by_sku: dict[str, dict] = {}

    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))

    for row in rows:
        sku   = row.get("品號", "").strip()
        price = row.get("商品結帳價", "").strip()
        if sku and price:
            main_by_sku[sku] = row

    for row in rows:
        name  = row.get("品名", "").strip()
        sku   = row.get("品號", "").strip()
        price = row.get("商品結帳價", "").strip()

        if not name:
            continue

        # 有 SKU、price 為空 → 視為別名行
        if not price and sku:
            if sku in main_by_sku:
                _upsert_alias(sku, name)
            else:
                print(f"  [WARN] 別名行 {name!r} 的品號 {sku!r} 在本 CSV 無主行，跳過")
            continue

        if not sku:
            continue

        prod = db.session.get(Product, sku)
        if prod is None:
            prod = Product(sku=sku)
            db.session.add(prod)

        prod.name       = name
        prod.category   = row.get("類別", "").strip() or None
        prod.quantity   = _int(row.get("份數"))
        prod.pack_size  = _int(row.get("包數"))
        prod.unit_price = _float(row.get("份數價格"))
        prod.erp_source = row.get("來源", "").strip() or None

        # 有 price 才建立通路定價；沒有 price 的商品（如組合商品）只建 Product
        if price:
            price_val = _float(price)
            if price_val is not None:
                for ch in channels:
                    _upsert_price(sku, ch, price_val)

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


def _int(val: str | None) -> int | None:
    try:
        return int(val) if val and val.strip() else None
    except (ValueError, TypeError):
        return None


def _float(val: str | None) -> float | None:
    try:
        return float(val) if val and val.strip() else None
    except (ValueError, TypeError):
        return None


def import_kadomo_barcodes() -> None:
    barcode_map: dict[str, str] = json.loads(BARCODE_JSON.read_text(encoding="utf-8"))
    for barcode, sku in barcode_map.items():
        exists = ProductBarcode.query.filter_by(barcode=barcode).first()
        if not exists:
            db.session.add(ProductBarcode(sku=sku, barcode=barcode))
    db.session.flush()


def import_kadomo_prices() -> None:
    price_map: dict[str, float] = json.loads(KADOMO_PRICE_JSON.read_text(encoding="utf-8"))
    for sku, price in price_map.items():
        _upsert_price(sku, "kadomo", float(price))
    db.session.flush()


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

        # Upsert Product（licai 使用條碼查詢，不覆蓋已有品名）
        prod = db.session.get(Product, sku)
        if prod is None:
            prod = Product(sku=sku, name=name)
            db.session.add(prod)

        # Upsert ProductBarcode（條碼全域唯一，可能已由 kadomo 匯入）
        exists_bc = ProductBarcode.query.filter_by(barcode=barcode).first()
        if not exists_bc:
            db.session.add(ProductBarcode(sku=sku, barcode=barcode))

        # Upsert ChannelPrice for licai（price=0.0 佔位）
        _upsert_price(sku, "licai", 0.0)

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

        print("  匯入 麗兒采家/品號資料_含條碼.csv → 通路：licai")
        import_licai()

        db.session.commit()

        print(f"\n匯入完成：")
        print(f"  products        : {Product.query.count()} 筆")
        print(f"  product_aliases : {ProductAlias.query.count()} 筆")
        print(f"  product_barcodes: {ProductBarcode.query.count()} 筆")
        print(f"  channel_prices  : {ChannelPrice.query.count()} 筆")


if __name__ == "__main__":
    main()
