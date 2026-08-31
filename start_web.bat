@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动 HubStudio 批量控制台...
python run_app.py
pause
