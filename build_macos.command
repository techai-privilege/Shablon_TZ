#!/bin/zsh
set -e
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi
.venv/bin/python -m pip install -e '.[build,dev]'
.venv/bin/python -m ruff check . --exclude build,dist,tmp
.venv/bin/python -m pytest -q
.venv/bin/pyinstaller --noconfirm --clean --windowed \
  --name "Отчет по товарному знаку" \
  --add-data "assets:assets" \
  --add-data "data:data" \
  app.py
"dist/Отчет по товарному знаку.app/Contents/MacOS/Отчет по товарному знаку" --self-test
echo "Готовое приложение находится в папке dist."
