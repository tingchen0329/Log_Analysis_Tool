@echo off
chcp 65001 >nul
rem ==========================================================================
rem  一鍵推送到 GitHub：雙擊執行即可
rem    1. 顯示目前變更
rem    2. 輸入 commit 訊息（直接 Enter 則略過 commit，只推送已存在的 commit / tag）
rem    3. 推送 main 與所有版本 tag（推 v* tag 會觸發 GitHub Actions 自動打包 exe）
rem  第一次執行會跳出 GitHub 登入視窗（Git Credential Manager），之後自動記住
rem ==========================================================================
cd /d "%~dp0"
where git >nul 2>&1 || (echo 找不到 git，請先安裝 Git for Windows：https://git-scm.com/download/win & pause & exit /b 1)

git status --short
echo.
set "MSG="
set /p MSG=輸入 commit 訊息（Enter 略過）：
if not "%MSG%"=="" (
    git add -A
    git commit -m "%MSG%" || goto :error
)
set "TAG="
set /p TAG=要建立版本 tag 嗎？輸入例如 v23（Enter 略過）：
if not "%TAG%"=="" git tag -a %TAG% -m "%TAG%" || goto :error

git push origin main --follow-tags || goto :error
echo.
echo 推送完成：https://github.com/tingchen0329/Log_Analysis_Tool
pause
exit /b 0

:error
echo.
echo 推送失敗，請看上方錯誤訊息（常見原因：尚未登入 GitHub、遠端有新的 commit 需先 git pull）
pause
exit /b 1
