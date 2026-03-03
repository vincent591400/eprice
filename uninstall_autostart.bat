@echo off
set TASK_NAME=EpriceWebServer

echo 正在移除開機自動啟動工作排程...
schtasks /delete /tn "%TASK_NAME%" /f

:: 關閉 web_app.py 程序
taskkill /f /im pythonw.exe >nul 2>&1

echo [完成] 已移除自動啟動設定。
pause
