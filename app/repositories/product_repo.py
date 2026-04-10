"""
app/repositories/product_repo.py
ProductRepository：封裝所有品號查詢邏輯。
呼叫端（Converter）只看這個介面，不直接操作 ORM 或 CSV。
切換 Supabase：只改 DATABASE_URI，此類別不需改動。
"""
from __future__ import annotations

import re
from difflib import get_close_matches

from sqlalchemy.orm import joinedload

from app.models import Product, ProductAlias, ProductBarcode, ChannelPrice

_CODE_RE = re.compile(r'^(\d+[A-Z]?-[A-Z]?\d+)')


class ProductRepository:
    """
    提供 Converter 所需的品號查詢方法。
    所有方法回傳與原 CSV row dict 相容的 dict，
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
        price_rows = ChannelPrice.query.filter_by(channel=channel).all()
        sku_set    = {cp.sku for cp in price_rows}
        price_map  = {cp.sku: cp.price for cp in price_rows}

        products = (
            Product.query
            .filter(Product.sku.in_(sku_set))
            .options(joinedload(Product.aliases))
            .all()
        )

        product_map: dict[str, dict] = {}
        code_map:    dict[str, dict] = {}

        for prod in products:
            price = price_map.get(prod.sku, 0)
            row   = _to_dict(prod, price)
            product_map[prod.name] = row
            for alias_obj in prod.aliases:
                product_map[alias_obj.alias] = row
            m = _CODE_RE.match(prod.name)
            if m:
                code_map[m.group(1)] = row

        return product_map, code_map

    def load_barcodes(self, channel: str) -> dict[str, str]:
        """
        回傳指定通路的條碼對照表：barcode → sku。
        目前只有 kadomo 通路使用。
        """
        sku_set = {
            cp.sku
            for cp in ChannelPrice.query.filter_by(channel=channel).all()
        }
        return {
            bc.barcode: bc.sku
            for bc in ProductBarcode.query.filter(
                ProductBarcode.sku.in_(sku_set)
            ).all()
        }

    def get_price(self, sku: str, channel: str) -> float:
        """回傳指定通路的品號單價，找不到回傳 0.0。"""
        cp = ChannelPrice.query.filter_by(sku=sku, channel=channel).first()
        return float(cp.price) if cp else 0.0


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
