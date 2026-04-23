@echo off
chcp 65001 >nul
title 捷捷寶寶粥 訂單轉換系統

cd /d "%~dp0"

echo ========================================
echo  捷捷寶寶粥 訂單轉換系統
echo ========================================
echo.

where uv >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 uv，尚未完成安裝！
    echo.
    echo 請在此視窗輸入以下指令安裝 uv，安裝完畢後再重新點擊 start.bat：
    echo.
    echo     winget install astral-sh.uv
    echo.
    echo 若安裝後仍出現此訊息，請截圖傳給管理員。
    echo.
    pause
    exit /b 1
)

echo  正在啟動伺服器，請稍候...
echo  啟動完成後瀏覽器會自動開啟。
echo  請勿關閉此視窗，關閉即停止系統。
echo.

start /min cmd /c "timeout /t 3 /nobreak >nul && start http://127.0.0.1:5099"

uv run python main.py

echo.
echo ========================================
echo  [錯誤] 伺服器已停止，請截圖此畫面
echo  並將完整畫面傳給管理員處理。
echo ========================================
echo.
pause
