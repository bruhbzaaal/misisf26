#!/bin/zsh
cd -- "$(dirname -- "$0")" || exit 1
if command -v python3 >/dev/null 2>&1; then
  python3 -B app.py
else
  echo "Установите Python 3.9 или новее / Install Python 3.9 or newer."
fi
read -r '?Нажмите Enter, чтобы закрыть окно / Press Enter to close'
