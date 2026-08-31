@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo  HubStudio 批量控制台 - 打包 exe
echo ========================================
echo.

python -m pip install -r requirements.txt pyinstaller -q
if errorlevel 1 (
  echo 依赖安装失败，请确认已安装 Python 3.10+
  pause
  exit /b 1
)

echo 正在打包（约 2-5 分钟）...
python -m PyInstaller --noconfirm hubstudio_tool.spec
if errorlevel 1 (
  echo 打包失败
  pause
  exit /b 1
)

echo.
echo 完成！发布文件夹：
echo   dist\HubStudio批量控制台\
echo.
echo 请把整个文件夹复制给朋友，运行其中的 HubStudio批量控制台.exe
echo 首次运行会在 exe 旁生成 config.json，在网页里填 HubStudio 路径即可。
echo.
pause
