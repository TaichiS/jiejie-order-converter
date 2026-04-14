#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
麗采通路訂單轉換工具
==================

功能：
    1. 自動識別麗兒采家採購單格式
    2. 拆分多門市合併的訂單
    3. 移除跨頁重複標題行
    4. 對照品號資料和通路資料進行格式轉換
    5. 輸出為標準格式：MMDD 麗采{門市}.xlsx

使用方法：
    python licai_order_converter.py <輸入檔案> [輸出目錄]

範例：
    python licai_order_converter.py 0330麗兒采家採購單_濬詮.xlsx
    python licai_order_converter.py 0330麗兒采家採購單_濬詮.xlsx ./output

相依檔案（需放在同一目錄）：
    - 品號資料.csv
    - 通路資料.xlsx

作者：AI Assistant
版本：1.0.0
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

def load_reference_data(pinhao_file: str, tonglu_file: str) -> Tuple[Dict, Dict]:
    """
    載入品號資料和通路資料
    
    Args:
        pinhao_file: 品號資料CSV檔案路徑
        tonglu_file: 通路資料Excel檔案路徑
    
    Returns:
        tuple: (品號對照字典, 通路對照字典)
    """
    # 讀取品號資料
    pinhao_df = pd.read_csv(pinhao_file)
    # 清理重複的標題行
    pinhao_df = pinhao_df[pinhao_df['品號'] != '品號'].copy()
    # 建立品名→品號對照字典
    pinhao_dict = dict(zip(pinhao_df['品名'], pinhao_df['品號']))
    
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
    
    return pinhao_dict, tonglu_dict


# =============================================================================
# 品名轉換函數
# =============================================================================

def convert_old_to_new_product_name(old_name: str) -> str:
    """
    將舊格式品名轉換為新格式品名
    
    轉換規則：
        - 011高麗菜米泥(50克)/4 → 0-11高麗菜米泥
        - 1P02高麗野蔬雞蓉粥150g/2 → 1P-02高麗野蔬雞蓉大寶寶粥150g
        - 2S11田園雞肉燉飯/2 → S11田園雞肉燉飯
        - 後元彩虹水餃15顆(14M+) → 2-D01後元彩虹水餃
    
    Args:
        old_name: 舊格式品名
    
    Returns:
        str: 新格式品名
    """
    # 移除括號內容和後面的/數字
    name = re.sub(r'\([^)]*\)', '', old_name)
    name = re.sub(r'/\d+$', '', name)
    name = name.strip()
    
    # 類型1: 1XX開頭（小寶寶粥）
    # 例如：101玉米菇菇雞茸粥 → 1-01玉米菇菇雞蓉粥
    match = re.match(r'^1(\d{2})(.+)', name)
    if match:
        num, rest = match.groups()
        rest = rest.replace('雞茸', '雞蓉')
        return f"1-{num}{rest}"
    
    # 類型2: 2XX開頭（燴料）
    # 例如：211香甜玉米雞肉 → 2-11香甜玉米雞肉
    match = re.match(r'^2(\d{2})(.+)', name)
    if match:
        num, rest = match.groups()
        return f"2-{num}{rest}"
    
    # 類型3: 純數字開頭（2位數或3位數，米泥）
    # 例如：011高麗菜米泥 → 0-11高麗菜米泥
    # 例如：02南瓜米泥 → 0-2南瓜米泥
    match = re.match(r'^(\d{2,3})(.+)', name)
    if match:
        num, rest = match.groups()
        num_clean = num.lstrip('0') or '0'
        return f"0-{num_clean}{rest}"
    
    # 類型4: 1P開頭（大寶寶粥150g/200g）
    # 例如：1P02高麗野蔬雞蓉粥150g → 1P-02高麗野蔬雞蓉大寶寶粥150g
    match = re.match(r'^1P(\d{2})(.+)', name)
    if match:
        num, rest = match.groups()
        # 將各種粥名統一為「大寶寶粥」
        rest = rest.replace('雞茸粥', '雞蓉大寶寶粥')
        rest = rest.replace('雞蓉粥', '雞蓉大寶寶粥')
        rest = rest.replace('豬豬粥', '豬豬大寶寶粥')
        rest = rest.replace('牛肉粥', '牛肉大寶寶粥')
        rest = rest.replace('鱸魚粥', '鱸魚大寶寶粥')
        rest = rest.replace('蛋黃粥', '蛋黃大寶寶粥')
        rest = rest.replace('野蕈粥', '野蕈大寶寶粥')
        rest = rest.replace('牛奶魚粥', '牛奶魚大寶寶粥')
        rest = rest.replace('針菇粥', '針菇大寶寶粥')
        rest = rest.replace('魚肉粥', '魚肉大寶寶粥')
        rest = rest.replace('嫩雞粥', '嫩雞大寶寶粥')
        rest = rest.replace('鮭魚粥', '鮭魚大寶寶粥')
        rest = rest.replace('豆雞蓉粥', '豆豆雞蓉大寶寶粥')
        return f"1P-{num}{rest}"
    
    # 類型5: 2S開頭（燉飯）
    # 例如：2S11田園雞肉燉飯 → S11田園雞肉燉飯
    # 例如：2S1和風豚肉燉飯 → 2-S1和風豚肉燉飯
    match = re.match(r'^2S(\d{1,2})(.+)', name)
    if match:
        num, rest = match.groups()
        if len(num) == 1:
            return f"2-S{num}{rest}"
        else:
            return f"S{num}{rest}"
    
    # 類型6: 2M開頭（義麵）
    # 例如：2M3青醬海陸義麵 → 2-M3青醬海陸義麵
    match = re.match(r'^2M(\d)(.+)', name)
    if match:
        num, rest = match.groups()
        return f"2-M{num}{rest}"
    
    # 類型7: 後元系列
    if '後元' in name:
        if '彩虹水餃' in name:
            return '2-D01後元彩虹水餃'
        elif '饅頭' in name:
            return '2-D02寶寶後元饅頭'
        elif '蘿蔔糕' in name:
            return '2-D03寶寶後元蘿蔔糕'
    
    # 類型8: 寶寶豬肉鬆
    if '豬肉鬆' in name:
        return '寶寶豬肉鬆'
    
    return name


def find_best_match(old_name: str, pinhao_dict: Dict) -> Tuple[str, Optional[str]]:
    """
    使用模糊匹配找到最匹配的品號資料品名
    
    Args:
        old_name: 舊格式品名
        pinhao_dict: 品號對照字典
    
    Returns:
        tuple: (匹配後的品名, 品號) 若找不到則品號為None
    """
    # 先進行標準轉換
    converted_name = convert_old_to_new_product_name(old_name)
    
    # 如果直接匹配成功，返回結果
    if converted_name in pinhao_dict:
        return converted_name, pinhao_dict[converted_name]
    
    # 提取轉換後品名的關鍵特徵
    # 支持: 1P-03, 2-S2, S11, 0-2, 1-01, 2-11, 2-M3 等格式
    match = re.match(r'^([\dA-Z]+-[\dA-Z]+|[DS]\d+)(.+?)(\d+g)?$', converted_name)
    if not match:
        return converted_name, None
    
    series, product, spec = match.groups()
    
    best_match = None
    best_score = 0
    
    for pinhao_name in pinhao_dict.keys():
        score = 0
        
        # 檢查系列編號是否匹配（最高優先）
        if series in pinhao_name:
            score += 5
        
        # 提取品號資料中的產品名稱部分
        pm_match = re.match(r'^([\dA-Z]+-[\dA-Z]+|[DS]\d+)(.+?)(\d+g)?$', pinhao_name)
        if pm_match:
            pm_series, pm_product, pm_spec = pm_match.groups()
            
            # 檢查產品名稱相似度（計算共同中文字比例）
            old_chars = set(re.findall(r'[\u4e00-\u9fff]', product or ''))
            pm_chars = set(re.findall(r'[\u4e00-\u9fff]', pm_product or ''))
            
            if old_chars and pm_chars:
                common = old_chars & pm_chars
                similarity = len(common) / max(len(old_chars), len(pm_chars))
                score += similarity * 5
            
            # 檢查規格是否匹配
            if spec and pm_spec:
                if spec == pm_spec:
                    score += 2
        
        if score > best_score:
            best_score = score
            best_match = pinhao_name
    
    # 分數門檻為4分（系列匹配5分或高相似度）
    if best_match and best_score >= 4:
        return best_match, pinhao_dict[best_match]
    
    return converted_name, None


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
    pinhao_dict: Dict, 
    tonglu_dict: Dict
) -> List[Dict]:
    """
    將訂單轉換為目標格式
    
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
        pinhao_dict: 品號對照字典
        tonglu_dict: 通路對照字典
    
    Returns:
        list: 轉換後的訂單列表
    """
    result_orders = []
    
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
                # row 格式: [序號, 貨號, 品名, 品牌, 市價, 門市數量]
                barcode = row[1]
                old_product_name = str(row[2])
                price = row[4]
                qty = row[5]
                
                # 使用模糊匹配找到最佳對應
                new_product_name, product_code = find_best_match(old_product_name, pinhao_dict)
                
                converted_row = {
                    '訂單編號': order_info['order_id'],
                    '收件人': store_full,
                    '地址': address,
                    '電話': phone,
                    '產品編號': product_code if product_code else '',
                    '產品名稱': new_product_name,
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
    
    return result_orders


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
        pinhao_file: 品號資料CSV檔案路徑
        tonglu_file: 通路資料Excel檔案路徑
        output_dir: 輸出目錄
    
    Returns:
        list: 輸出的檔案路徑列表
    """
    # 載入對照資料
    pinhao_dict, tonglu_dict = load_reference_data(pinhao_file, tonglu_file)
    
    # 讀取輸入檔案
    df = pd.read_excel(input_file, header=None)
    
    print(f"處理檔案: {input_file}")
    
    # 檢查格式
    if not is_licai_format(df):
        print("✗ 不符合麗采通路格式，不予轉換。")
        return []
    
    print("✓ 檔案符合麗采通路格式，開始處理...\n")
    
    # 解析訂單
    orders = parse_orders(df)
    print(f"共解析出 {len(orders)} 個訂單\n")
    
    # 轉換格式
    result_orders = convert_to_target_format(orders, pinhao_dict, tonglu_dict)
    
    # 建立輸出目錄
    os.makedirs(output_dir, exist_ok=True)
    
    output_files = []
    
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
        
        # 統計匹配率
        matched = sum(1 for r in order['rows'] if r['產品編號'])
        total = len(order['rows'])
        match_rate = matched / total * 100 if total > 0 else 0
        
        print(f"✓ {store_name}: {total} 品項, 匹配率 {match_rate:.0f}% → {output_filename}")
    
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
        print("\n使用方法：python licai_order_converter.py <輸入檔案> [輸出目錄]")
        print("範例：python licai_order_converter.py 0330麗兒采家採購單_濬詮.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else './converted_orders'
    
    # 檢查輸入檔案是否存在
    if not os.path.exists(input_file):
        print(f"錯誤：找不到輸入檔案: {input_file}")
        sys.exit(1)
    
    # 檢查相依檔案
    script_dir = Path(__file__).parent
    pinhao_file = script_dir / '品號資料.csv'
    tonglu_file = script_dir / '通路資料.xlsx'
    
    if not pinhao_file.exists():
        print(f"錯誤：找不到品號資料檔案: {pinhao_file}")
        print("請確保品號資料.csv與本程式放在同一目錄")
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
