@echo off
chcp 65001 >nul
echo ======================================
echo    訂單轉換程式 - 批次處理
echo ======================================
echo.

if "%~1"=="" (
    echo 使用方法: batch_convert.bat [檔案名稱或萬用字元]
    echo.
    echo 範例:
    echo   batch_convert.bat 20260330_採購單K009.xlsx
    echo   batch_convert.bat 採購單*.xlsx
    pause
    exit /b 1
)

set count=0
for %%f in (%~1) do (
    echo 正在轉換: %%f
    python order_converter.py "%%f"
    if errorlevel 1 (
        echo [錯誤] %%f 轉換失敗
    ) else (
        set /a count+=1
        echo [成功] %%f 轉換完成
    )
    echo.
)

echo ======================================
echo 批次轉換完成！共轉換 %count% 個檔案
echo ======================================
pause
