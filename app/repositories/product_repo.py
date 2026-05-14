"""
app/repositories/product_repo.py
ProductRepository：封裝所有品號查詢邏輯。
本次改為以 unified_products 為單一事實來源。
"""
from __future__ import annotations

import re

from app.models import UnifiedProduct

_CODE_RE      = re.compile(r'^(\d+[A-Z]?-[A-Z]?\d+)')
_BARE_CODE_RE = re.compile(r'^([A-Z]\d+)')


# CSV 通路名稱 → converter source_type 的對應
# 注意：jjofficial 一般商品只對應「寶寶粥官網」，加購品需額外查詢「寶寶粥官網-加購」
_SOURCE_TYPE_TO_CHANNELS: dict[str, list[str]] = {
    "shopee":     ["蝦皮"],
    "a1baby":     ["婦幼展"],
    "leage":      ["樂齡PDF"],
    "a1leage":    ["樂齡官網"],
    "jjofficial": ["寶寶粥官網"],
    "xuantu":     ["炫兔團"],
    "tuanma":     ["其他團媽"],
    "yodee":      ["吉寶通路"],
    "kadomo":     ["卡多摩"],
    "licai":      ["麗采"],
}


class ProductRepository:
    """
    提供 Converter 所需的品號查詢方法。
    所有方法回傳與原 CSV row dict 相容的 dict。
    """

    def load_channel(self, source_type: str) -> tuple[dict[str, dict], dict[str, dict]]:
        """
        載入指定通路的所有品號資料，回傳 (product_map, code_map)。
        product_map: 品名 → row_dict
        code_map:    代碼前綴 → row_dict
        """
        channels = _SOURCE_TYPE_TO_CHANNELS.get(source_type, [])
        if not channels:
            return {}, {}

        rows = UnifiedProduct.query.filter(
            UnifiedProduct.channel.in_(channels)
        ).all()

        product_map: dict[str, dict] = {}
        code_map:    dict[str, dict] = {}

        # code_map 儲存 list[dict]，支援同代碼多規格（如 150g/200g）
        code_map: dict[str, list[dict]] = {}

        def _add_code(code: str, row: dict) -> None:
            code_map.setdefault(code, []).append(row)

        for up in rows:
            row = _to_dict(up)
            product_map[up.name] = row
            codes: list[str] = []
            m = _CODE_RE.match(up.name)
            if m:
                codes.append(m.group(1))
            else:
                m2 = _BARE_CODE_RE.match(up.name)
                if m2:
                    codes.append(m2.group(1))

            for code in codes:
                _add_code(code, row)
                # 蝦皮：同時註冊帶 2- 與不帶 2- 的兩種 key，因訂單品名可能兩種格式都有
                if source_type == "shopee":
                    if code.startswith("2-"):
                        _add_code(code[2:], row)
                    elif re.match(r"^[A-Z]\d+$", code):
                        _add_code("2-" + code, row)

        return product_map, code_map

    def load_barcodes(self, source_type: str) -> dict[str, str]:
        """
        回傳指定通路的條碼對照表：barcode → sku。
        目前只有 kadomo、licai 通路使用。
        """
        channels = _SOURCE_TYPE_TO_CHANNELS.get(source_type, [])
        if not channels:
            return {}

        rows = (
            UnifiedProduct.query
            .filter(
                UnifiedProduct.channel.in_(channels),
                UnifiedProduct.barcode.isnot(None),
                UnifiedProduct.barcode != "",
            )
            .all()
        )
        return {up.barcode: up.sku for up in rows}

    def get_price(self, sku: str, source_type: str) -> float:
        """回傳指定通路的品號單價，找不到回傳 0.0。"""
        channels = _SOURCE_TYPE_TO_CHANNELS.get(source_type, [])
        if not channels:
            return 0.0

        up = (
            UnifiedProduct.query
            .filter(
                UnifiedProduct.sku == sku,
                UnifiedProduct.channel.in_(channels),
            )
            .first()
        )
        return float(up.checkout_price) if up and up.checkout_price is not None else 0.0

    def lookup_by_sku(self, sku: str, channel: str) -> dict | None:
        """
        直接依 sku + channel 查詢單一品號資料。
        供 jjofficial 查詢「寶寶粥官網-加購」等子通路使用。
        """
        up = (
            UnifiedProduct.query
            .filter_by(sku=sku, channel=channel)
            .first()
        )
        return _to_dict(up) if up else None


def _to_dict(up: UnifiedProduct) -> dict:
    """將 UnifiedProduct ORM 物件轉為與舊 CSV row 相容的 dict。"""
    return {
        "品號":       up.sku,
        "品名":       up.name,
        "類別":       up.category or "",
        "份數":       str(up.quantity or 1),
        "包數":       str(up.pack_size) if up.pack_size is not None else "",
        "份數價格":   str(up.unit_price) if up.unit_price is not None else "",
        "包數價格":   str(up.pack_price) if up.pack_price is not None else "",
        "商品結帳價": str(up.checkout_price) if up.checkout_price is not None else "",
        "折扣":       str(up.discount) if up.discount is not None else "",
        "來源":       up.channel or "",
    }
