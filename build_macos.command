#!/bin/zsh
set -e
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi
.venv/bin/python -m pip install -e '.[build]'
.venv/bin/pyinstaller --noconfirm --clean --windowed \
  --name "Отчет по товарному знаку" \
  --add-data "assets:assets" \
  --add-data "data:data" \
  app.py
echo "Готовое приложение находится в папке dist."
