@echo off
setlocal

:: 請先到 https://github.com/new 建立新 repo（例如 eprice）
:: 建立後將下方 YOUR_USERNAME 和 YOUR_REPO 改成你的 GitHub 帳號與 repo 名稱

set GITHUB_USER=vincent591400
set GITHUB_REPO=eprice

if "%GITHUB_USER%"=="YOUR_USERNAME" (
    echo 請先編輯此檔案，將 GITHUB_USER 和 GITHUB_REPO 改成你的 GitHub 資訊
    echo 例如: set GITHUB_USER=myaccount
    echo       set GITHUB_REPO=eprice
    pause
    exit /b 1
)

cd /d "%~dp0"

echo 正在加入 remote...
git remote remove origin 2>nul
git remote add origin https://github.com/%GITHUB_USER%/%GITHUB_REPO%.git

echo 正在 push 到 GitHub...
git branch -M main
git push -u origin main

if %errorlevel% == 0 (
    echo.
    echo [完成] 已推送到 https://github.com/%GITHUB_USER%/%GITHUB_REPO%
) else (
    echo.
    echo [失敗] 請確認：
    echo   1. 已在 GitHub 建立 repo
    echo   2. 已設定 Git 憑證（或使用 GitHub Desktop / Personal Access Token）
    echo   3. 網路連線正常
)

pause
