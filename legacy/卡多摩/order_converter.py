#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
訂單轉換程式
功能：將卡多摩採購單格式轉換為出貨單格式
作者：AI Assistant
版本：1.1
"""

import pandas as pd
import json
import re
import os
import sys
from pathlib import Path
from datetime import datetime


class OrderConverter:
    """訂單轉換器主類別"""

    REQUIRED_COLUMNS = ['倉別', '商品條碼', '商品名稱', '進貨數量', '採購單號']

    def __init__(self, data_dir='.'):
        self.data_dir = Path(data_dir)
        self.channel_data = None
        self.barcode_map = None
        self.product_map = None
        self.price_table = None
        self.warehouse_mapping = None
        self.load_reference_data()

    def load_reference_data(self):
        try:
            with open(self.data_dir / '通路資料.json', 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.channel_data = {item['店名']: item for item in data['stores']}
                self.keyword_mapping = data.get('keyword_mapping', {})
                self.warehouse_mapping = data.get('warehouse_mapping', {})

            with open(self.data_dir / '條碼對照表.json', 'r', encoding='utf-8') as f:
                self.barcode_map = json.load(f)

            with open(self.data_dir / '品號對照表.json', 'r', encoding='utf-8') as f:
                self.product_map = json.load(f)

            with open(self.data_dir / '單價對照表.json', 'r', encoding='utf-8') as f:
                self.price_table = json.load(f)

            print("對照表資料載入成功")

        except FileNotFoundError as e:
            print(f"無法找到對照表檔案: {e}")
            raise

    def validate_source_file(self, file_path):
        errors = []
        try:
            df = pd.read_excel(file_path, header=None)

            if df.empty:
                errors.append("檔案為空")
                return False, errors, None

            header_row = None
            for idx, row in df.iterrows():
                if '倉別' in row.values:
                    header_row = idx
                    break

            if header_row is None:
                errors.append("無法找到標題列（缺少'倉別'欄位）")
                return False, errors, None

            df = pd.read_excel(file_path, header=header_row)

            missing_columns = [col for col in self.REQUIRED_COLUMNS if col not in df.columns]
            if missing_columns:
                errors.append(f"缺少必填欄位: {', '.join(missing_columns)}")

            data_rows = df.dropna(subset=['商品條碼'])
            if len(data_rows) == 0:
                errors.append("沒有找到任何商品資料")

            warehouses = df['倉別'].dropna().unique()
            if len(warehouses) > 1:
                errors.append(f"發現多個倉別: {warehouses}")

            is_valid = len(errors) == 0
            return is_valid, errors, df if is_valid else None

        except Exception as e:
            errors.append(f"檔案讀取錯誤: {str(e)}")
            return False, errors, None

    def extract_store_info(self, file_name, warehouse_code):
        store_name = None

        if warehouse_code in self.warehouse_mapping:
            store_name = self.warehouse_mapping[warehouse_code]

        if not store_name:
            for keyword, full_name in self.keyword_mapping.items():
                if keyword in file_name:
                    store_name = full_name
                    break

        if store_name and store_name in self.channel_data:
            return self.channel_data[store_name]

        return {
            '店名': f'未知門市({warehouse_code})',
            '電話': '',
            '地址': ''
        }

    def get_product_code(self, barcode, product_name):
        barcode_str = str(barcode).strip()

        if barcode_str in self.barcode_map:
            return self.barcode_map[barcode_str]

        for code, info in self.product_map.items():
            if info['品名'] in product_name or product_name in info['品名']:
                return code

        return None

    def get_unit_price(self, product_code):
        if product_code in self.price_table:
            return self.price_table[product_code]
        return 0.0

    def generate_order_number(self, warehouse_code, file_name):
        date_match = re.search(r'(\d{4})(\d{2})(\d{2})', file_name)
        if date_match:
            month = date_match.group(2)
            day = date_match.group(3)
            return f"{warehouse_code}-{month}{day}"

        today = datetime.now()
        return f"{warehouse_code}-{today.strftime('%m%d')}"

    def convert(self, source_file, output_file=None):
        source_path = Path(source_file)

        print(f"正在驗證檔案: {source_path.name}")
        is_valid, errors, df = self.validate_source_file(source_file)

        if not is_valid:
            print("檔案驗證失敗:")
            for error in errors:
                print(f"  - {error}")
            raise ValueError("原始檔格式不符合要求")

        print("檔案驗證通過")

        warehouse_code = df['倉別'].dropna().iloc[0]
        store_info = self.extract_store_info(source_path.name, warehouse_code)
        order_number = self.generate_order_number(warehouse_code, source_path.name)

        print(f"  倉別: {warehouse_code}")
        print(f"  門市: {store_info['店名']}")
        print(f"  訂單編號: {order_number}")

        output_data = []

        for idx, row in df.iterrows():
            if pd.isna(row['商品條碼']):
                continue

            barcode = str(row['商品條碼']).strip()
            product_name = str(row['商品名稱']).strip() if not pd.isna(row['商品名稱']) else ''
            quantity = int(row['進貨數量']) if not pd.isna(row['進貨數量']) else 0
            po_number = str(row['採購單號']).strip() if not pd.isna(row['採購單號']) else ''

            product_code = self.get_product_code(barcode, product_name)
            unit_price = self.get_unit_price(product_code) if product_code else 0.0

            output_data.append({
                '訂單編號': order_number,
                '收件人': store_info['店名'],
                '地址': store_info['地址'],
                '電話': store_info['電話'],
                '產品編號': product_code if product_code else '',
                '產品名稱': product_name,
                '數量': quantity,
                '單價': unit_price,
                '備註': po_number,
                '備註.1': ''
            })

        output_df = pd.DataFrame(output_data)

        if not output_file:
            date_match = re.search(r'(\d{4})(\d{2})(\d{2})', source_path.name)
            if date_match:
                month = date_match.group(2)
                day = date_match.group(3)
                store_short = store_info['店名'].replace('卡多摩嬰童-', '').replace('卡多摩-', '')
                output_file = f"{month}{day} {store_short}.xlsx"
            else:
                output_file = f"轉換後_{source_path.stem}.xlsx"

        output_path = self.data_dir / output_file
        output_df.to_excel(output_path, index=False, sheet_name='訂單資料')

        print(f"轉換完成！")
        print(f"  輸出檔案: {output_path}")
        print(f"  共 {len(output_data)} 筆資料")

        return str(output_path)


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description='訂單轉換程式 - 將卡多摩採購單轉換為出貨單格式'
    )

    parser.add_argument('input', help='原始檔案路徑')
    parser.add_argument('-o', '--output', help='輸出檔案路徑（選填）')
    parser.add_argument('-d', '--data-dir', default='.', help='對照表資料目錄（預設: 當前目錄）')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"錯誤: 找不到輸入檔案 '{args.input}'")
        sys.exit(1)

    try:
        converter = OrderConverter(args.data_dir)
        converter.convert(args.input, args.output)
    except Exception as e:
        print(f"轉換失敗: {str(e)}")
        sys.exit(1)


if __name__ == '__main__':
    main()
