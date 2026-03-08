@echo off
setlocal

cd /d "%~dp0"

python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --name Word格式识别与套用工具 --windowed app.py

echo.
echo 打包完成，输出目录：dist\Word格式识别与套用工具
pause
