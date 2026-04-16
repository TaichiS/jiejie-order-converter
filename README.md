# 捷捷寶寶粥 訂單格式轉換系統

將各通路（蝦皮、婦幼展、樂齡網、官網等）的異構訂單檔案，統一轉換為內部 ERP 可用格式的 Flask 網頁工具。

---

## 系統需求

- Python 3.12+
- [uv](https://docs.astral.sh/uv/)（套件與虛擬環境管理）

---

## 快速開始

### 1. 安裝相依套件

```bash
uv sync
```

### 2. 啟動伺服器

```bash
# macOS / Linux
./start.sh

# Windows
start.bat
```

預設開啟 `http://localhost:5099`

### 3. 初始化品號資料

進入網頁後，前往「管理 → 匯入品號資料」，上傳 `品號資料統整.csv` 即可覆蓋資料庫中的統一品號表。

---

## 專案結構

```
.
├── app/                    # Flask 應用程式
│   ├── models.py           # SQLAlchemy 資料模型
│   ├── routes.py           # 網頁路由與 API
│   ├── services.py         # 轉換協調服務（掃描、執行、歸檔）
│   ├── utils.py            # 通用工具與常數（SOURCE_LABELS 等）
│   └── repositories/
│       └── product_repo.py # 品號資料查詢層
├── converters/             # 各通路轉換器
│   ├── base.py             # BaseConverter 抽象基底類
│   ├── detector.py         # 檔案來源自動偵測
│   ├── shopee.py           # 蝦皮
│   ├── a1baby.py           # A1 婦幼展
│   ├── leage.py            # 樂齡網
│   ├── a1leage.py          # A1 樂齡官網
│   ├── jjofficial.py       # 捷捷官網
│   ├── yodee.py            # 優迪通路
│   ├── kadomo.py           # 卡多摩
│   ├── licai.py            # 麗兒采家
│   ├── xuantu.py           # 炫兔團購
│   ├── tuanma.py           # 其他團媽
│   └── chocho.py           # CHOCHO 通路
├── templates/              # Jinja2 網頁模板
├── static/                 # CSS / JS / 圖示
├── data/                   # SQLite 資料庫（conversion.db）
├── output/                 # 轉換後輸出檔案（不上傳 Git）
├── archive/                # 歸檔原始檔案（不上傳 Git）
├── legacy/                 # 棄用或測試資料（不上傳 Git）
├── scripts/                # 輔助腳本
├── main.py                 # 應用程式進入點
├── pyproject.toml          # uv 套件設定
└── CHANGELOG.md            # 版本修訂日誌
```

---

## 核心架構

### 轉換流程

```
上傳檔案 → detector.detect_each() 逐檔偵測來源
                ↓
        依 source_type 分組
                ↓
        各組呼叫對應 Converter
                ↓
        輸出至 output/{source_type}/{YYYY-MM}/
        原始檔歸檔至 archive/{source_type}/{YYYY-MM}/
                ↓
        寫入 ConversionLog / ConversionError
```

### 多來源混合轉換

`services.run_conversion()` 使用 `detect_each()` 逐檔偵測，同一資料夾可同時存在多種來源檔案，系統會自動分組、各自轉換，不會強制要求全部同類型。

### 品號查詢機制

統一品號資料統一由 `UnifiedProduct`（SQLite）管理：

- **載入**：`ProductRepository.load_channel(source_type)` 依通路過濾並建立記憶體索引
- **查詢順序**：
  1. 代碼前綴比對（如 `1-08`）
  2. 蝦皮 bare code 補 `2-` 前綴（如 `S11` → `2-S11`）
  3. 精確品名比對
  4. difflib 模糊比對 + 金額整除候選

### 錯誤追蹤

每筆 `RowError` 皆記錄：
- `source_file`：錯誤來自哪個輸入檔案
- `row_number`：資料列號
- `field_name` / `original_value` / `reason`：欄位與錯誤說明
- `candidates`：系統建議的候選品項清單

---

## 支援的轉換格式

| 代號 | 名稱 | 輸入格式 | 品號資料來源 |
|------|------|----------|--------------|
| `shopee` | 蝦皮 | `Order.toship.*.xlsx`（可加密） | `UnifiedProduct` |
| `a1baby` | A1 婦幼展 | `MMDD.xlsx` + `MMDD-1.xlsx` | `UnifiedProduct` |
| `leage` | 樂齡網 | PDF | `UnifiedProduct` |
| `a1leage` | A1 樂齡官網 | `*.xlsx`（Orders 工作表） | `UnifiedProduct` |
| `jjofficial` | 捷捷官網 | `*.xlsx`（Sales 工作表） | `UnifiedProduct` |
| `yodee` | 優迪通路 | `*.xlsx`（含標頭列「訂單編號」） | `UnifiedProduct` |
| `kadomo` | 卡多摩 | 專用格式 | `卡多摩/通路資料.json` |
| `licai` | 麗兒采家 | `*.xlsx` | `UnifiedProduct` |
| `xuantu` | 炫兔團購 | `*.xlsx` | `UnifiedProduct` |
| `tuanma` | 其他團媽 | `*.xlsx`（Sales 工作表） | `UnifiedProduct` |
| `chocho` | CHOCHO 通路 | `*.csv`（檔名含 `CHOCHO`） | `UnifiedProduct` |

---

## 常用管理指令

### 清除轉換資料

```bash
uv run python3 -c "
import sqlite3
conn = sqlite3.connect('data/conversion.db')
conn.execute('DELETE FROM conversion_errors')
conn.execute('DELETE FROM conversion_logs')
conn.commit()
conn.close()
"
find output -type f -delete
find archive -type f -delete
```

### 壓縮發布版本

```bash
cd /Users/lung/Documents/企業顧問/02_客戶管理/捷捷寶寶粥
7z a "訂單格式轉換系統_YYYYMMDD_vN.7z" "訂單格式轉換系統" \
  -xr!"__pycache__" -xr!"*.pyc" -xr!".git" -xr!"*.db" \
  -xr!"output" -xr!"archive" -xr!"uploads" -xr!".env" \
  -xr!"*.jpg" -xr!"*.png"
```

---

## 注意事項

- `db.create_all()` **不會**更新已存在的資料表結構。若新增欄位，需手動執行 `ALTER TABLE`。
- `output/`、`archive/`、`legacy/` 與各通路原始資料資料夾已列入 `.gitignore`，不會進入版本控制。
- 部署前務必先清除轉換資料與輸出檔案，再進行壓縮。
