@echo off
chcp 65001 >nul
rem ==========================================================================
rem  打包成單一 exe（使用者電腦不需安裝 Python）：dist\Log_Analysis_Tool.exe
rem  需求：Windows + Python 3.8 以上；執行一次即可
rem ==========================================================================
cd /d "%~dp0"
python -m pip install -r requirements-build.txt || goto :error

set "ICON_ARGS="
if exist icon_ironman.ico set "ICON_ARGS=--icon icon_ironman.ico --add-data icon_ironman.ico;."

python -m PyInstaller --noconfirm --clean --onefile --windowed --name Log_Analysis_Tool %ICON_ARGS% Cisco_Log_Analysis.py || goto :error
echo.
echo 完成：%~dp0dist\Log_Analysis_Tool.exe
pause
exit /b 0

:error
echo 打包失敗，請確認已安裝 Python 並可連網下載套件。
pause
exit /b 1
