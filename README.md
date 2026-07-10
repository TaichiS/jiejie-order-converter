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

## SHOPLINE 今日訂單下載

`scripts/fetch_shopline_orders_today.py` 用於直接呼叫 SHOPLINE Open API，下載捷捷寶寶粥官網今日訂單，並整理成既有 `jjofficial` 轉換器可讀取的「捷捷官網轉換前」格式。

此腳本的定位是 API 匯入前置工具，不會直接產生 ERP 檔案。它會先輸出一份原始 JSON 與一份 `Sales` 工作表 xlsx，後續仍交給既有捷捷官網轉換流程處理。

### 參考文件

- [SHOPLINE Open API：How to get access_token](https://open-api.docs.shoplineapp.com/docs/getting-started)
- [SHOPLINE Open API：OpenAPI request example](https://open-api.docs.shoplineapp.com/docs/openapi-request-example)
- [SHOPLINE Open API：Get Orders](https://open-api.docs.shoplineapp.com/docs/get-orders)
- [SHOPLINE Open API：Get Order](https://open-api.docs.shoplineapp.com/docs/get-order)
- [SHOPLINE Open API：Pagination](https://open-api.docs.shoplineapp.com/docs/pagination)

### 環境變數

先複製 `.env.example` 為 `.env`，再填入 SHOPLINE token：

```bash
cp .env.example .env
```

`.env` 至少需要：

```env
SHOPLINE_ACCESS_TOKEN=replace_with_real_shopline_open_api_token
SHOPLINE_USER_AGENT=JiejieOrderConverter/0.1
```

- `SHOPLINE_ACCESS_TOKEN`：從 SHOPLINE 後台「設定 > 管理員設定 > API Auth」產生。
- `SHOPLINE_USER_AGENT`：SHOPLINE Open API 必帶的 `User-Agent` header；若 SHOPLINE 有提供 handle code，請改填該值。
- `.env` 已加入 `.gitignore`，不可提交實際 token。

### 基本用法

下載台北時間今天的訂單，預設輸出到相對路徑 `測試資料/`：

```bash
uv run python "scripts/fetch_shopline_orders_today.py"
```

只顯示統計、不顯示訂單摘要：

```bash
uv run python "scripts/fetch_shopline_orders_today.py" --show 0
```

指定台北日期：

```bash
uv run python "scripts/fetch_shopline_orders_today.py" --date 2026-06-22
```

指定輸出資料夾：

```bash
uv run python "scripts/fetch_shopline_orders_today.py" --output-dir "測試資料"
```

### 輸出檔案

未指定輸出檔名時，腳本會產生：

```text
測試資料/shopline_orders_YYYYMMDD_raw.json
測試資料/YYYYMMDD捷捷官網轉換前.xlsx
```

- `shopline_orders_YYYYMMDD_raw.json`：SHOPLINE API 原始回應，包含完整訂單資料，可能含客戶個資，只供除錯與欄位 mapping 使用。
- `YYYYMMDD捷捷官網轉換前.xlsx`：整理後的 `Sales` 工作表，可被 `detector` 辨識為 `jjofficial`，供既有捷捷官網轉換器使用。

自訂輸出檔名：

```bash
uv run python "scripts/fetch_shopline_orders_today.py" \
  --output-json "測試資料/shopline_orders_20260622_raw.json" \
  --output-xlsx "測試資料/20260622捷捷官網轉換前.xlsx"
```

### 訂單篩選與時間處理

- SHOPLINE Open API 的時間參數使用 UTC。
- 腳本接受台北日期，會自動換算成 UTC 區間。
- 預設查詢 `created_after` / `created_before`，也就是「指定台北日期建立的訂單」。
- 預設不把 `status=cancelled` 的取消訂單寫入 xlsx，但原始 JSON 仍會完整保存。
- 如需把取消訂單也寫入 xlsx，可加上：

```bash
uv run python "scripts/fetch_shopline_orders_today.py" --include-cancelled
```

### 參數說明

| 參數 | 預設值 | 說明 |
|------|--------|------|
| `--date` | 今天 | 指定台北日期，格式 `YYYY-MM-DD` |
| `--per-page` | `50` | 每頁訂單數，會限制在 1 到 50 |
| `--max-pages` | `10` | 最多抓取頁數，避免測試時抓取過量資料 |
| `--show` | `10` | 顯示前 N 筆安全摘要；設為 `0` 可隱藏明細 |
| `--output-dir` | `測試資料` | 預設輸出資料夾，使用相對路徑 |
| `--output-json` | 自動產生 | 自訂原始 JSON 輸出路徑 |
| `--output-xlsx` | 自動產生 | 自訂捷捷官網轉換前 xlsx 輸出路徑 |
| `--include-cancelled` | 關閉 | xlsx 是否包含取消訂單 |

### 整理後 xlsx 欄位

腳本會建立 `Sales` 工作表，並輸出捷捷官網轉換器目前需要的欄位：

```text
訂單號碼、訂單狀態、付款狀態、收件人、完整地址、收件人電話號碼、發票號碼、
商品貨號、商品名稱、數量、商品結帳價、商品折扣優惠、商品折扣金額、
點數折現分攤、出貨備註、送貨編號、付款方式、全家服務編號 / 7-11 店號、
加購品類型、訂單備註、發票開立日期、運費
```

其中 `全家服務編號 / 7-11 店號` 會取 SHOPLINE `delivery_data.location_code`。有值時後續轉換會進全家檔，無值時會進黑貓檔。

### 後續轉換

產生 `測試資料/YYYYMMDD捷捷官網轉換前.xlsx` 後，可用既有網頁工具掃描該資料夾並執行轉換，或透過現有服務流程強制指定 `jjofficial`。

若轉換結果出現 `partial`，通常代表品號資料庫缺少 SHOPLINE 商品品號。此時應先補 `品號資料統整.csv` 並重新匯入品號資料，再重新轉換。

### 安全注意事項

- 不要提交 `.env`。
- 不要提交 `測試資料/` 內的 JSON 或 xlsx，這些檔案可能包含客戶姓名、電話、地址與訂單明細。
- 若只要檢查 API 是否可用，建議使用 `--show 0`，避免終端輸出過多訂單資料。

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
