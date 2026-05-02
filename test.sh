#!/usr/bin/env bash
# Production-Control-System 直下で実行: Java（JUnit）+ Python（pytest）
# 例: ./test.sh   pytest へ追加: ./test.sh -k smoke
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"

echo "== [1/2] Maven test (code_java) =="
mvn -f code_java/pom.xml test

echo "== [2/2] pytest (code/python/tests) =="
cd code/python
if [[ -n "${VIRTUAL_ENV:-}" ]] && python3 -c "import pytest" 2>/dev/null; then
  python3 -m pytest tests/ -q --tb=short "$@"
elif [[ -x "$ROOT/.venv/bin/python" ]] && "$ROOT/.venv/bin/python" -c "import pytest" 2>/dev/null; then
  "$ROOT/.venv/bin/python" -m pytest tests/ -q --tb=short "$@"
elif python3 -c "import pytest" 2>/dev/null; then
  python3 -m pytest tests/ -q --tb=short "$@"
else
  echo "pytest がありません。例:" >&2
  echo "  sudo apt install python3-venv && python3 -m venv \"$ROOT/.venv\" && \"$ROOT/.venv/bin/pip\" install pytest" >&2
  echo "  または python3 -m pip install pytest --break-system-packages" >&2
  exit 1
fi
