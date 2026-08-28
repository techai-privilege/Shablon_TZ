#!/bin/zsh
set -e
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi
if ! .venv/bin/python -c 'import PySide6, docx, trademark_report' >/dev/null 2>&1; then
  .venv/bin/python -m pip install -e .
fi
exec .venv/bin/python app.py
