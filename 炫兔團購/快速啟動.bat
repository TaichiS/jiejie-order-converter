@echo off
chcp 65001
echo ==========================================
echo 【炫兔團購】訂單轉換程式
echo ==========================================
echo.
echo 正在檢查必要檔案...

if not exist "0402-0404.xlsx" (
    echo [錯誤] 找不到原始訂單：0402-0404.xlsx
    pause
    exit /b 1
)

if not exist "品號資料_方案B(炫兔團).csv" (
    echo [錯誤] 找不到品號資料：品號資料_方案B(炫兔團).csv
    pause
    exit /b 1
)

echo [成功] 檔案檢查通過
echo.
echo 正在執行轉換...
python convert_orders.py

if errorlevel 1 (
    echo.
    echo [錯誤] 轉換失敗，請查看錯誤訊息
) else (
    echo.
    echo [成功] 轉換完成！
    echo 結果已儲存至 output 目錄
)

pause
