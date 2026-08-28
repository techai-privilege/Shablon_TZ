@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
.venv\Scripts\python.exe -c "import PySide6, docx, trademark_report" >nul 2>&1
if errorlevel 1 .venv\Scripts\python.exe -m pip install -e .
.venv\Scripts\python.exe app.py
