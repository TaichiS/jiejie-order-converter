"""
【炫兔團購】訂單轉換程式 v1.0
功能：將原始訂單格式轉換為「全家v1」欄位配置
作者：Kimi AI
日期：2026-04-12
"""

import pandas as pd
import numpy as np
import re
from datetime import datetime
import os


def build_product_mapping(product_df):
    """
    建立品號對照字典

    從品號資料建立多種形式的對照關係：
    1. 完整品號 (D50500004) → 商品資訊
    2. 商品名稱 (2-1醬燒什錦豬肉) → 商品資訊  
    3. 短代碼 (2-1, 2-01) → 商品資訊
    """
    mapping = {}

    for _, row in product_df.iterrows():
        standard_code = row['品號']
        product_name = row['品名']

        product_info = {
            'standard_code': standard_code,
            'product_name': product_name,
            'package_count': row['包數'],
            'unit_price': row['商品結帳價'],
            'category': row['類別'],
            'batch_price': row['份數價格']
        }

        # 1. 完整品號
        mapping[standard_code] = product_info
        # 2. 商品名稱
        mapping[product_name] = product_info
        # 3. 短代碼
        if '-' in product_name:
            parts = product_name.split('-')
            if len(parts) >= 2:
                short_code = parts[0] + '-' + parts[1].lstrip('0')
                mapping[short_code] = product_info
                original_short = parts[0] + '-' + parts[1]
                mapping[original_short] = product_info

    return mapping


def convert_order_row(row, product_mapping):
    """
    轉換單筆訂單資料
    """
    original_code = row['商品貨號']

    # 排除贈品（貨號為空）
    if pd.isna(original_code) or original_code == '':
        return None

    # 取得商品資訊
    product_info = None
    if original_code in product_mapping:
        product_info = product_mapping[original_code]
    else:
        product_name = row['商品名稱']
        if pd.notna(product_name):
            clean_name = product_name.replace('團購限定B|', '')
            if clean_name in product_mapping:
                product_info = product_mapping[clean_name]
            else:
                match = re.match(r'(\d+)-(\d+)', clean_name)
                if match:
                    short_code = match.group(1) + '-' + match.group(2)
                    if short_code in product_mapping:
                        product_info = product_mapping[short_code]

    # 找不到對照則使用原始資料
    if product_info is None:
        product_info = {
            'standard_code': original_code,
            'product_name': row['商品名稱'],
            'package_count': 1,
            'unit_price': row['商品結帳價']
        }

    # 數量與價格轉換
    original_qty = row['數量']
    original_price = row['商品結帳價']
    package_count = product_info['package_count']

    converted_qty = original_qty * package_count
    converted_price = 0 if original_price == 0 else original_price / package_count

    # 清理商品名稱
    original_name = row['商品名稱']
    converted_name = original_name.replace('團購限定B|', '') if pd.notna(original_name) else product_info['product_name']

    # 加購折扣處理
    add_on_discount = row['加購品類型']
    if pd.isna(add_on_discount):
        add_on_discount = 0

    return {
        '訂單號碼': row['訂單號碼'],
        '收件人': row['收件人'],
        '完整地址': row['完整地址'],
        '收件人電話號碼': row['收件人電話號碼'],
        '發票號碼': row['發票號碼'],
        '商品貨號': product_info['standard_code'],
        '商品名稱': converted_name,
        '數量': converted_qty,
        '商品結帳價': converted_price,
        '商品折扣優惠': row['商品折扣優惠'] if pd.notna(row['商品折扣優惠']) else 0,
        '商品折扣金額': row['商品折扣金額'] if pd.notna(row['商品折扣金額']) else '',
        '加購折扣': add_on_discount,
        '點數折現分攤': row['點數折現分攤'] if pd.notna(row['點數折現分攤']) else '',
        '出貨備註': row['出貨備註'] if pd.notna(row['出貨備註']) else '',
        '送貨編號': row['全家服務編號 / 7-11 店號'] if pd.notna(row['全家服務編號 / 7-11 店號']) else '',
        '付款方式': ''
    }


def convert_orders(original_df, product_df):
    """
    完整的訂單轉換函數
    """
    print("=" * 80)
    print("【炫兔團購】訂單轉換程式")
    print("=" * 80)

    # 步驟 1: 建立品號對照字典
    print("\n步驟 1: 建立品號對照字典...")
    product_mapping = build_product_mapping(product_df)
    print(f"✓ 已建立 {len(product_mapping)} 個對照項目")

    # 步驟 2: 篩選有效訂單
    print("\n步驟 2: 篩選有效訂單資料...")
    valid_mask = original_df['商品貨號'].notna() & (original_df['商品貨號'] != '')
    valid_orders = original_df[valid_mask].copy()
    print(f"✓ 有效訂單: {len(valid_orders)} 筆（排除 {len(original_df) - len(valid_orders)} 筆贈品）")

    # 步驟 3: 執行轉換
    print("\n步驟 3: 執行資料轉換...")
    converted_data = []
    error_count = 0

    for idx, row in valid_orders.iterrows():
        try:
            converted_row = convert_order_row(row, product_mapping)
            if converted_row:
                converted_data.append(converted_row)
        except Exception as e:
            error_count += 1

    print(f"✓ 成功轉換: {len(converted_data)} 筆")
    if error_count > 0:
        print(f"✗ 轉換錯誤: {error_count} 筆")

    # 步驟 4: 建立輸出資料表
    print("\n步驟 4: 建立輸出資料表...")
    result_df = pd.DataFrame(converted_data)

    target_columns = [
        '訂單號碼', '收件人', '完整地址', '收件人電話號碼', '發票號碼',
        '商品貨號', '商品名稱', '數量', '商品結帳價', '商品折扣優惠',
        '商品折扣金額', '加購折扣', '點數折現分攤', '出貨備註', '送貨編號', '付款方式'
    ]

    result_df = result_df[target_columns]
    print(f"✓ 輸出欄位: {len(target_columns)} 個")

    # 步驟 5: 資料驗證
    print("\n步驟 5: 資料驗證...")
    numeric_issues = result_df[result_df['數量'] <= 0]
    if len(numeric_issues) > 0:
        print(f"⚠ 警告: 發現 {len(numeric_issues)} 筆數量異常資料")
    else:
        print("✓ 數量欄位檢查通過")

    print("\n" + "=" * 80)
    print("轉換完成！")
    print("=" * 80)

    return result_df, product_mapping


def main():
    """主程式"""
    input_file = '0402-0404.xlsx'
    product_file = '品號資料_方案B(炫兔團).csv'
    output_dir = 'output'

    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    try:
        print(f"讀取原始訂單: {input_file}")
        original_orders = pd.read_excel(input_file)

        print(f"讀取品號資料: {product_file}")
        product_data = pd.read_csv(product_file, encoding='utf-8-sig')

        print(f"原始訂單筆數: {len(original_orders)}")
        print(f"品號資料筆數: {len(product_data)}")

        result, _ = convert_orders(original_orders, product_data)

        output_excel = f'{output_dir}/炫兔團購_訂單轉換_{timestamp}.xlsx'
        output_csv = f'{output_dir}/炫兔團購_訂單轉換_{timestamp}.csv'

        result.to_excel(output_excel, index=False, engine='openpyxl')
        print(f"\n✓ Excel 格式已保存: {output_excel}")

        result.to_csv(output_csv, index=False, encoding='utf-8-sig')
        print(f"✓ CSV 格式已保存: {output_csv}")

        print(f"\n【轉換結果摘要】")
        print(f"總筆數: {len(result)}")
        print(f"總數量: {result['數量'].sum()}")
        print(f"總金額: {result['商品結帳價'].sum():.2f} 元")

    except Exception as e:
        print(f"\n✗ 執行錯誤: {e}")
        raise


if __name__ == "__main__":
    main()
