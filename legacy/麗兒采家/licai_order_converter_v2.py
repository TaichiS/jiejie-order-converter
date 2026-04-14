#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
麗采通路訂單轉換工具 v2.0
========================

功能：
    1. 自動識別麗兒采家採購單格式
    2. 拆分多門市合併的訂單
    3. 移除跨頁重複標題行
    4. 使用條碼精確對照品號資料
    5. 輸出為標準格式：MMDD 麗采{門市}.xlsx

使用方法：
    python licai_order_converter_v2.py <輸入檔案> [輸出目錄]

範例：
    python licai_order_converter_v2.py 0330麗兒采家採購單_濬詮.xlsx
    python licai_order_converter_v2.py 0330麗兒采家採購單_濬詮.xlsx ./output

相依檔案（需放在同一目錄）：
    - 品號資料_含條碼.csv（品號、條碼、品名三欄位）
    - 通路資料.xlsx

版本：2.0.0
更新：改用條碼精確匹配，移除模糊匹配邏輯
"""

import pandas as pd
import re
import os
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Optional


# =============================================================================
# 資料載入函數
# =============================================================================

def load_reference_data(pinhao_file: str, tonglu_file: str) -> Tuple[Dict, Dict, Dict]:
    """
    載入品號資料（含條碼）和通路資料
    
    Args:
        pinhao_file: 品號資料CSV檔案路徑（需包含品號、條碼、品名三欄位）
        tonglu_file: 通路資料Excel檔案路徑
    
    Returns:
        tuple: (條碼對照字典, 品號品名對照字典, 通路對照字典)
    """
    # 讀取品號資料（含條碼）
    pinhao_df = pd.read_csv(pinhao_file)
    
    # 建立條碼→品號對照字典（精確匹配）
    barcode_dict = dict(zip(pinhao_df['條碼'].astype(str), pinhao_df['品號']))
    
    # 建立品號→品名對照字典
    pinhao_name_dict = dict(zip(pinhao_df['品號'], pinhao_df['品名']))
    
    # 讀取通路資料
    tonglu_df = pd.read_excel(tonglu_file)
    tonglu_dict = {}
    
    for _, row in tonglu_df.iterrows():
        store_full = row['店名']
        # 只處理麗兒采家的門市
        if '麗兒采家' in store_full:
            match = re.search(r'麗兒采家-(.+?)店', store_full)
            if match:
                # 提取簡稱（如「大墩旗艦」→「大墩」）
                store_short = match.group(1).replace('旗艦', '')
                tonglu_dict[store_short] = {
                    'store_full': store_full,
                    'phone': row['電話'],
                    'address': row['地址']
                }
    
    return barcode_dict, pinhao_name_dict, tonglu_dict


# =============================================================================
# 格式驗證函數
# =============================================================================

def is_licai_format(df: pd.DataFrame) -> bool:
    """
    檢查是否為麗采通路格式
    
    特徵：
        - 包含「麗兒采家採購單」字樣
        - 包含「門市：」字樣
    
    Args:
        df: 輸入的DataFrame
    
    Returns:
        bool: 是否為麗采通路格式
    """
    df_str = df.to_string()
    if '麗兒采家採購單' in df_str and '門市：' in df_str:
        return True
    return False


# =============================================================================
# 訂單解析函數
# =============================================================================

def parse_orders(df: pd.DataFrame) -> List[Dict]:
    """
    解析舊格式訂單，拆分成各門市的訂單
    
    Args:
        df: 輸入的DataFrame
    
    Returns:
        list: 訂單列表，每個訂單包含門市名稱、訂單資訊和資料列
    """
    orders = []
    current_order = None
    current_order_rows = []
    current_order_info = {'order_id': '', 'order_date': ''}
    prev_row_str_for_check = ""
    
    for idx, row in df.iterrows():
        row_values = row.values
        row_str = ' '.join([str(v) for v in row_values if pd.notna(v)])
        
        # 檢查是否包含單號和採購日期
        if '單號：' in row_str and '採購日期：' in row_str:
            order_match = re.search(r'單號：(\S+)', row_str)
            date_match = re.search(r'採購日期：(\S+)', row_str)
            current_order_info = {
                'order_id': order_match.group(1) if order_match else '',
                'order_date': date_match.group(1) if date_match else ''
            }
            prev_row_str_for_check = row_str
            continue
        
        # 檢查是否為新訂單開始（門市行）
        if '門市：' in row_str:
            # 檢查前一行是否為「麗兒采家採購單」
            prev_row_str = str(df.iloc[max(0, idx-1)].values)
            if '麗兒采家採購單' in prev_row_str:
                # 儲存前一個訂單
                if current_order and current_order_rows:
                    orders.append({
                        'store': current_order,
                        'info': current_order_info.copy(),
                        'rows': current_order_rows
                    })
                
                # 提取門市名稱
                match = re.search(r'門市：(\S+)', row_str)
                current_order = match.group(1) if match else f'unknown_{idx}'
                current_order_rows = []
                prev_row_str_for_check = row_str
        
        # 跳過跨頁重複標題行
        elif re.match(r'^\s*門市：\S+\s+單號：\S+\s*$', row_str):
            prev_row_str_for_check = row_str
            continue
        elif re.match(r'^\s*序\s+貨號\s+品名\s+品牌\s+市價\s+\S+\s*$', row_str):
            prev_row_str_for_check = row_str
            continue
        # 檢查是否為資料列（有序號、貨號、品名等）
        elif current_order is not None and re.match(r'^\s*\d+\s+\d+', row_str):
            current_order_rows.append(row_values)
            prev_row_str_for_check = row_str
        else:
            prev_row_str_for_check = row_str
    
    # 儲存最後一個訂單
    if current_order and current_order_rows:
        orders.append({
            'store': current_order,
            'info': current_order_info.copy(),
            'rows': current_order_rows
        })
    
    return orders


# =============================================================================
# 格式轉換函數
# =============================================================================

def convert_to_target_format(
    orders: List[Dict], 
    barcode_dict: Dict, 
    pinhao_name_dict: Dict, 
    tonglu_dict: Dict
) -> Tuple[List[Dict], List[Dict]]:
    """
    將訂單轉換為目標格式（使用條碼精確匹配）
    
    目標欄位：
        - 訂單編號
        - 收件人
        - 地址
        - 電話
        - 產品編號
        - 產品名稱
        - 數量
        - 單價
        - 備註
        - 備註.1
    
    Args:
        orders: 解析後的訂單列表
        barcode_dict: 條碼→品號對照字典
        pinhao_name_dict: 品號→品名對照字典
        tonglu_dict: 通路對照字典
    
    Returns:
        tuple: (轉換後的訂單列表, 未匹配的品項列表)
    """
    result_orders = []
    unmatched_items = []
    
    for order in orders:
        store_name = order['store']
        order_info = order['info']
        rows = order['rows']
        
        # 取得通路資訊
        tonglu_info = tonglu_dict.get(store_name, {})
        store_full = tonglu_info.get('store_full', f'麗兒采家-{store_name}')
        phone = tonglu_info.get('phone', '')
        address = tonglu_info.get('address', '')
        
        converted_rows = []
        for row in rows:
            try:
                # row 格式: [序號, 貨號/條碼, 品名, 品牌, 市價, 門市數量]
                barcode = str(row[1]).strip()
                old_product_name = str(row[2])
                price = row[4]
                qty = row[5]
                
                # 使用條碼精確匹配品號
                product_code = barcode_dict.get(barcode)
                
                if product_code:
                    # 取得標準品名
                    product_name = pinhao_name_dict.get(product_code, old_product_name)
                else:
                    # 條碼找不到對應，記錄下來
                    unmatched_items.append({
                        'store': store_name,
                        'barcode': barcode,
                        'old_name': old_product_name
                    })
                    product_code = ''
                    product_name = old_product_name
                
                converted_row = {
                    '訂單編號': order_info['order_id'],
                    '收件人': store_full,
                    '地址': address,
                    '電話': phone,
                    '產品編號': product_code,
                    '產品名稱': product_name,
                    '數量': qty,
                    '單價': price,
                    '備註': '',
                    '備註.1': ''
                }
                converted_rows.append(converted_row)
            except Exception as e:
                print(f"  ⚠ 處理列時發生錯誤: {e}")
                continue
        
        result_orders.append({
            'store': store_name,
            'order_id': order_info['order_id'],
            'order_date': order_info['order_date'],
            'rows': converted_rows
        })
    
    return result_orders, unmatched_items


# =============================================================================
# 主處理函數
# =============================================================================

def process_licai_order(
    input_file: str, 
    pinhao_file: str, 
    tonglu_file: str, 
    output_dir: str
) -> List[str]:
    """
    處理麗采通路訂單
    
    Args:
        input_file: 輸入的Excel檔案路徑
        pinhao_file: 品號資料CSV檔案路徑（需包含品號、條碼、品名三欄位）
        tonglu_file: 通路資料Excel檔案路徑
        output_dir: 輸出目錄
    
    Returns:
        list: 輸出的檔案路徑列表
    """
    # 載入對照資料
    barcode_dict, pinhao_name_dict, tonglu_dict = load_reference_data(
        pinhao_file, tonglu_file
    )
    
    # 讀取輸入檔案
    df = pd.read_excel(input_file, header=None)
    
    print(f"處理檔案: {input_file}")
    print(f"已載入條碼對照: {len(barcode_dict)} 筆")
    
    # 檢查格式
    if not is_licai_format(df):
        print("✗ 不符合麗采通路格式，不予轉換。")
        return []
    
    print("✓ 檔案符合麗采通路格式，開始處理...\n")
    
    # 解析訂單
    orders = parse_orders(df)
    print(f"共解析出 {len(orders)} 個訂單\n")
    
    # 轉換格式
    result_orders, unmatched_items = convert_to_target_format(
        orders, barcode_dict, pinhao_name_dict, tonglu_dict
    )
    
    # 統計結果
    total_items = sum(len(o['rows']) for o in result_orders)
    matched_items = sum(1 for o in result_orders for r in o['rows'] if r['產品編號'])
    match_rate = matched_items / total_items * 100 if total_items > 0 else 0
    
    print(f"轉換結果:")
    print(f"  總品項: {total_items}")
    print(f"  成功匹配: {matched_items}")
    print(f"  匹配率: {match_rate:.1f}%")
    
    if unmatched_items:
        print(f"\n⚠ 未匹配的品項（{len(unmatched_items)} 筆）：")
        for item in unmatched_items[:10]:
            print(f"    {item['barcode']} - {item['old_name']}")
        if len(unmatched_items) > 10:
            print(f"    ... 還有 {len(unmatched_items) - 10} 筆")
    
    # 建立輸出目錄
    os.makedirs(output_dir, exist_ok=True)
    
    output_files = []
    
    print()
    for order in result_orders:
        store_name = order['store']
        order_df = pd.DataFrame(order['rows'])
        
        # 產生檔名: MMDD 麗采{門市}.xlsx
        file_stem = Path(input_file).stem
        date_match = re.search(r'(\d{2})(\d{2})', file_stem)
        if date_match:
            month, day = date_match.groups()
            output_filename = f"{month}{day} 麗采{store_name}.xlsx"
        else:
            output_filename = f"麗采{store_name}.xlsx"
        
        output_path = os.path.join(output_dir, output_filename)
        order_df.to_excel(output_path, index=False)
        output_files.append(output_path)
        
        # 計算該門市的匹配率
        store_total = len(order['rows'])
        store_matched = sum(1 for r in order['rows'] if r['產品編號'])
        store_rate = store_matched / store_total * 100 if store_total > 0 else 0
        
        print(f"✓ {store_name}: {store_total} 品項, 匹配率 {store_rate:.0f}% → {output_filename}")
    
    print(f"\n轉換完成！檔案已儲存至: {output_dir}")
    return output_files


# =============================================================================
# 命令列介面
# =============================================================================

def main():
    """主程式入口"""
    # 檢查命令列參數
    if len(sys.argv) < 2:
        print(__doc__)
        print("\n使用方法：python licai_order_converter_v2.py <輸入檔案> [輸出目錄]")
        print("範例：python licai_order_converter_v2.py 0330麗兒采家採購單_濬詮.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else './converted_orders'
    
    # 檢查輸入檔案是否存在
    if not os.path.exists(input_file):
        print(f"錯誤：找不到輸入檔案: {input_file}")
        sys.exit(1)
    
    # 檢查相依檔案
    script_dir = Path(__file__).parent
    pinhao_file = script_dir / '品號資料_含條碼.csv'
    tonglu_file = script_dir / '通路資料.xlsx'
    
    if not pinhao_file.exists():
        print(f"錯誤：找不到品號資料檔案: {pinhao_file}")
        print("請確保品號資料_含條碼.csv與本程式放在同一目錄")
        print("檔案格式：品號,條碼,品名")
        sys.exit(1)
    
    if not tonglu_file.exists():
        print(f"錯誤：找不到通路資料檔案: {tonglu_file}")
        print("請確保通路資料.xlsx與本程式放在同一目錄")
        sys.exit(1)
    
    try:
        process_licai_order(
            input_file,
            str(pinhao_file),
            str(tonglu_file),
            output_dir
        )
    except Exception as e:
        print(f"錯誤：{e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
