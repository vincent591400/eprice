@echo off
setlocal

set SCRIPT_DIR=%~dp0
set VBS_PATH=%SCRIPT_DIR%start_web_silent.vbs
set TASK_NAME=EpriceWebServer

echo 正在設定開機自動啟動工作排程...

:: 先刪除舊的（如果存在）
schtasks /delete /tn "%TASK_NAME%" /f >nul 2>&1

:: 建立新工作排程：登入後自動在背景執行
schtasks /create ^
  /tn "%TASK_NAME%" ^
  /tr "wscript.exe \"%VBS_PATH%\"" ^
  /sc ONLOGON ^
  /rl HIGHEST ^
  /f

if %errorlevel% == 0 (
    echo.
    echo [成功] 工作排程已建立：%TASK_NAME%
    echo 下次登入後，網頁服務將自動在背景啟動。
    echo.
    echo 立即啟動服務...
    wscript.exe "%VBS_PATH%"
    echo.
    echo [完成] 現在可以開啟瀏覽器輸入 http://localhost:5000
) else (
    echo.
    echo [失敗] 請以系統管理員身份執行此批次檔。
)

pause
