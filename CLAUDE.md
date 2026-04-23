# 捷捷寶寶粥 訂單格式轉換系統

## 版本管理

每次壓縮新版本之前，必須先：
1. 將本次所有修改項目補寫到 `CHANGELOG.md`（新增一個版本區塊）
2. 清除轉換資料（DB + output/ + archive/），再壓縮

壓縮指令（版本號規則：同日期序號 +1）：
```bash
cd /Users/lung/Documents/企業顧問/02_客戶管理/捷捷寶寶粥
7z a "訂單格式轉換系統_YYYYMMDD_vN.7z" "訂單格式轉換系統" \
  -xr!"__pycache__" -xr!"*.pyc" -xr!".git" \
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
- **品號資料**：統一由 `UnifiedProduct`（SQLite）管理，透過 `/admin/import-products` 上傳 `品號資料統整.csv` 覆蓋。舊 CSV 已移至 `legacy/`。
- **輸出**：`output/{source_type}/{YYYY-MM}/`
- **歸檔**：`archive/{source_type}/{YYYY-MM}/`

## 支援的轉換格式

| 代號 | 名稱 | 輸入 | 品號資料 |
|------|------|------|----------|
| shopee | 蝦皮 | Order.toship.*.xlsx（可加密） | UnifiedProduct |
| a1baby | A1婦幼展 | MMDD.xlsx + MMDD-1.xlsx | UnifiedProduct |
| leage | 樂齡網 | PDF | UnifiedProduct |
| a1leage | A1樂齡官網 | *.xlsx（Orders 工作表） | UnifiedProduct |
| jjofficial | 捷捷官網 | *.xlsx（Sales 工作表） | UnifiedProduct |
| yodee | 優迪通路 | *.xlsx（含標頭列「訂單編號」） | UnifiedProduct |

## 關鍵架構決策

### 多來源混合轉換
`services.run_conversion()` 回傳 `list[ConversionLog]`（非單筆）。
內部用 `detector.detect_each()` 逐檔偵測，按 source_type 分組，各組各自執行轉換。
**不要改回** `detector.detect()`（要求全部同類型才成功，混合資料夾會失敗）。

### 錯誤資訊追蹤
`RowError` 有 `source_file` 欄位，記錄錯誤來自哪個輸入檔案。
`ConversionError` DB 表也有 `source_file` 欄位（2026-04-07 加入，需確認舊 DB 有無執行 ALTER TABLE）。

### DB Schema 異動注意事項
`db.create_all()` 不會更新已存在的表。新增欄位需手動執行 ALTER TABLE：
```bash
uv run python3 -c "
import sqlite3
conn = sqlite3.connect('data/conversion.db')
cols = [row[1] for row in conn.execute('PRAGMA table_info(conversion_errors)').fetchall()]
if 'source_file' not in cols:
    conn.execute('ALTER TABLE conversion_errors ADD COLUMN source_file TEXT DEFAULT \"\"')
    conn.commit()
conn.close()
"
```

### 偵測邏輯
- `detector.detect_each(files)` → `{Path: source_type}` 逐檔對應
- `detector.detect(files)` → 單一 source_type（僅在全部同類型時成功，混合回傳 None）
- 掃描（`scan_folder`）用 `detect_each`，轉換（`run_conversion`）也改為 `detect_each` 後分組
