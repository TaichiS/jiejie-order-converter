@echo off
chcp 65001 >nul
title 停止 捷捷寶寶粥 訂單轉換系統

set /a found=0
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":5099 "') do (
    taskkill /PID %%a /F >nul 2>&1
    if not errorlevel 1 (
        echo  已停止服務（PID: %%a）
        set /a found+=1
    )
)

if %found%==0 (
    echo  找不到執行中的服務（port 5099）
) else (
    echo  服務已停止。
)

pause
