#!/bin/bash
# 訂單轉換程式 - 批次處理

echo "======================================"
echo "   訂單轉換程式 - 批次處理"
echo "======================================"
echo ""

if [ $# -eq 0 ]; then
    echo "使用方法: ./batch_convert.sh [檔案名稱或萬用字元]"
    echo ""
    echo "範例:"
    echo "  ./batch_convert.sh 20260330_採購單K009.xlsx"
    echo "  ./batch_convert.sh '採購單*.xlsx'"
    exit 1
fi

count=0
for file in $1; do
    if [ -f "$file" ]; then
        echo "正在轉換: $file"
        if python3 order_converter.py "$file"; then
            ((count++))
            echo "[成功] $file 轉換完成"
        else
            echo "[錯誤] $file 轉換失敗"
        fi
        echo ""
    fi
done

echo "======================================"
echo "批次轉換完成！共轉換 $count 個檔案"
echo "======================================"
