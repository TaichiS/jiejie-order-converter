#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
訂單拆分工具
將合併的多門市訂單Excel檔案拆分為各門市的獨立檔案

功能：
1. 自動識別不同門市的訂單（透過「門市：」關鍵字）
2. 移除跨頁列印時產生的重複標題行
3. 每個門市的訂單輸出為獨立的Excel檔案

使用方法：
    python split_orders.py <輸入檔案> [輸出目錄]

範例：
    python split_orders.py 0330麗兒采家採購單_濬詮.xlsx
    python split_orders.py 0330麗兒采家採購單_濬詮.xlsx ./output
"""

import pandas as pd
import re
import os
import sys
from pathlib import Path


def split_orders(input_file, output_dir=None):
    """
    將合併的訂單檔案拆分為各門市的獨立檔案
    
    Args:
        input_file: 輸入的Excel檔案路徑
        output_dir: 輸出目錄，預設為輸入檔案所在目錄下的 'split_orders' 資料夾
    
    Returns:
        list: 輸出的檔案路徑列表
    """
    # 檢查輸入檔案是否存在
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"找不到檔案：{input_file}")
    
    # 設定輸出目錄
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(input_file) or '.', 'split_orders')
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"正在處理：{input_file}")
    print(f"輸出目錄：{output_dir}\n")
    
    # 讀取Excel檔案
    df = pd.read_excel(input_file, header=None)
    
    orders = []
    current_order = None
    current_order_rows = []
    
    for idx, row in df.iterrows():
        row_values = row.values
        row_str = ' '.join([str(v) for v in row_values if pd.notna(v)])
        
        # 檢查是否為新訂單開始（包含「門市：」且前面有「麗兒采家採購單」）
        if '門市：' in row_str and '麗兒采家採購單' in str(df.iloc[max(0, idx-1)].values):
            # 儲存前一個訂單
            if current_order and current_order_rows:
                orders.append({
                    'store': current_order,
                    'rows': current_order_rows
                })
            
            # 提取門市名稱
            match = re.search(r'門市：(\S+)', row_str)
            current_order = match.group(1) if match else f'unknown_{idx}'
            current_order_rows = [row]
        
        # 檢查是否為跨頁重複標題行（需要移除）
        # 特徵：只包含「門市：XXX 單號：XXX」
        elif re.match(r'^\s*門市：\S+\s+單號：\S+\s*$', row_str):
            # 這是跨頁標題行，跳過
            continue
        # 檢查是否為跨頁後的表頭行
        elif re.match(r'^\s*序\s+貨號\s+品名\s+品牌\s+市價\s+\S+\s*$', row_str):
            # 檢查前一行是否為跨頁標題行
            if current_order_rows:
                prev_row_str = ' '.join([str(v) for v in current_order_rows[-1].values if pd.notna(v)])
                if re.match(r'^\s*門市：\S+\s+單號：\S+', prev_row_str):
                    # 移除前一行的跨頁標題行
                    current_order_rows.pop()
            # 跳過這個表頭行
            continue
        elif current_order is not None:
            current_order_rows.append(row)
    
    # 儲存最後一個訂單
    if current_order and current_order_rows:
        orders.append({
            'store': current_order,
            'rows': current_order_rows
        })
    
    # 輸出每個訂單為獨立檔案
    output_files = []
    for order in orders:
        store_name = order['store']
        rows = order['rows']
        
        # 建立DataFrame
        order_df = pd.DataFrame(rows)
        
        # 產生輸出檔名
        base_name = Path(input_file).stem
        output_file = os.path.join(output_dir, f'{base_name}_{store_name}.xlsx')
        
        # 寫入Excel
        order_df.to_excel(output_file, index=False, header=False)
        output_files.append(output_file)
        print(f'✓ 已輸出：{store_name}')
    
    print(f'\n共拆分出 {len(orders)} 個訂單')
    print(f'檔案已儲存至：{output_dir}')
    
    return output_files


def main():
    # 檢查命令列參數
    if len(sys.argv) < 2:
        print("使用方法：python split_orders.py <輸入檔案> [輸出目錄]")
        print("範例：python split_orders.py 0330麗兒采家採購單_濬詮.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        split_orders(input_file, output_dir)
    except Exception as e:
        print(f"錯誤：{e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
