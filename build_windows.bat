@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install -e ".[build]"
pyinstaller --noconfirm --clean --windowed ^
  --name "Отчет по товарному знаку" ^
  --add-data "assets;assets" ^
  --add-data "data;data" ^
  app.py
echo Готовая программа находится в папке dist.
pause
