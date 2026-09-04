@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install -e ".[build,dev]"
python -m ruff check . --exclude build,dist,tmp
if errorlevel 1 exit /b 1
python -m pytest -q
if errorlevel 1 exit /b 1
pyinstaller --noconfirm --clean --windowed ^
  --name "Отчет по товарному знаку" ^
  --add-data "assets;assets" ^
  --add-data "data;data" ^
  app.py
"dist\Отчет по товарному знаку\Отчет по товарному знаку.exe" --self-test
if errorlevel 1 exit /b 1
echo Готовая программа находится в папке dist.
pause
