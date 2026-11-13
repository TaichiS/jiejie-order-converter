import csv
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))
sys.stdout.reconfigure(encoding='utf-8')

kadomo_csv = {}
with open(BASE_DIR / '品號資料統整.csv', newline='', encoding='utf-8-sig') as f:
    for row in csv.DictReader(f):
        if row.get('通路') == '卡多摩':
            kadomo_csv[row['品號']] = row

print(f'=== CSV 卡多摩品項 {len(kadomo_csv)} 筆 ===')
for sku, r in sorted(kadomo_csv.items()):
    cols = list(r.keys())
    price_col = next((c for c in cols if '價' in c), cols[-1] if cols else '')
    barcode = r.get('條碼', r.get('barcode', ''))
    price = r.get(price_col, '')
    print(f'  {sku}  {r["品名"]}  barcode={barcode}  {price_col}={price}')

from app import create_app, db
from app.models import UnifiedProduct, ChannelPrice
app = create_app()
with app.app_context():
    db_skus_up = {u.sku for u in UnifiedProduct.query.filter_by(channel='卡多摩').all()}
    db_skus_cp = {c.sku for c in ChannelPrice.query.filter_by(channel='kadomo').all()}
    csv_skus   = set(kadomo_csv.keys())

    missing_up = csv_skus - db_skus_up
    missing_cp = csv_skus - db_skus_cp

    print(f'\n=== DB UnifiedProduct 卡多摩: {len(db_skus_up)} 筆 ===')
    print(f'=== DB ChannelPrice kadomo: {len(db_skus_cp)} 筆 ===')

    if missing_up:
        print(f'\n⚠ CSV 有但 UnifiedProduct 無 ({len(missing_up)} 筆):')
        for s in sorted(missing_up):
            print(f'  {s}  {kadomo_csv[s]["品名"]}')
    else:
        print('\n✓ UnifiedProduct 卡多摩 與 CSV 一致')

    if missing_cp:
        print(f'\n⚠ CSV 有但 ChannelPrice kadomo 無 ({len(missing_cp)} 筆):')
        for s in sorted(missing_cp):
            print(f'  {s}  {kadomo_csv[s]["品名"]}')
    else:
        print('✓ ChannelPrice kadomo 與 CSV 一致')
