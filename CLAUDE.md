# 捷捷寶寶粥 訂單格式轉換系統

## 版本管理

每次壓縮新版本之前，必須先：
1. 將本次所有修改項目補寫到 `CHANGELOG.md`（新增一個版本區塊）
2. 清除轉換資料（DB + output/ + archive/），再壓縮

壓縮指令（版本號規則：同日期序號 +1）：
```bash
cd /Users/lung/Documents/企業顧問/02_客戶管理/捷捷寶寶粥
7z a "訂單格式轉換系統_YYYYMMDD_vN.7z" "訂單格式轉換系統" \
  -xr!"__pycache__" -xr!"*.pyc" -xr!".git" -xr!"*.db" \
  -xr!"output" -xr!"archive" -xr!"uploads" -xr!".env" \
  -xr!"*.jpg" -xr!"*.png"
```

## 清除轉換資料指令

```bash
cd 訂單格式轉換系統
uv run python3 -c "
import sqlite3
conn = sqlite3.connect('data/conversion.db')
conn.execute('DELETE FROM conversion_errors')
conn.execute('DELETE FROM conversion_logs')
conn.commit()
"
find output -type f -delete
find archive -type f -delete
```

## 系統架構

- **Flask** 網頁應用，port 5099
- **converters/** 各來源轉換器，繼承 `BaseConverter`
- **品號資料**：`reference/品號資料.csv`（蝦皮/婦幼/樂齡網）、`A1樂齡官網/品號資料.csv`（a1leage）
- **輸出**：`output/{source_type}/{YYYY-MM}/`
- **歸檔**：`archive/{source_type}/{YYYY-MM}/`

## 支援的轉換格式

| 代號 | 名稱 | 輸入 | 品號資料 |
|------|------|------|----------|
| shopee | 蝦皮 | Order.toship.*.xlsx | reference/ |
| a1baby | A1婦幼展 | MMDD.xlsx + MMDD-1.xlsx | reference/ |
| leage | 樂齡網 | PDF | reference/ |
| a1leage | A1樂齡官網 | *.xlsx（Orders 工作表） | A1樂齡官網/ |
