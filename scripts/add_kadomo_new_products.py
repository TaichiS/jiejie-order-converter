"""
新增 D50700002 / E51010018 至卡多摩通道
UnifiedProduct channel="卡多摩" + ChannelPrice channel="kadomo"
"""
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))
sys.stdout.reconfigure(encoding='utf-8')

from app import create_app, db
from app.models import UnifiedProduct, ChannelPrice

NEW_PRODUCTS = [
    # (barcode, sku, name, kadomo_price)
    ("4710586224122", "D50700002", "捷捷蘋果洋蔥水",          60.0),
    ("4710586224115", "E51010018", "後元雲朵水餃20顆(14M+)",  102.0),
]

app = create_app()
with app.app_context():
    for barcode, sku, name, price in NEW_PRODUCTS:
        # UnifiedProduct channel="卡多摩"
        up = UnifiedProduct.query.filter_by(channel="卡多摩", sku=sku).first()
        if up:
            up.barcode = barcode
            up.name    = name
            print(f"[UP] 更新 {sku} {name}")
        else:
            db.session.add(UnifiedProduct(
                channel=  "卡多摩",
                sku=      sku,
                name=     name,
                barcode=  barcode,
            ))
            print(f"[UP] 新增 {sku} {name}")

        # ChannelPrice channel="kadomo"
        cp = ChannelPrice.query.filter_by(channel="kadomo", sku=sku).first()
        if cp:
            cp.price = price
            print(f"[CP] 更新 {sku} price={price}")
        else:
            db.session.add(ChannelPrice(channel="kadomo", sku=sku, price=price))
            print(f"[CP] 新增 {sku} price={price}")

    db.session.commit()
    print("完成")
