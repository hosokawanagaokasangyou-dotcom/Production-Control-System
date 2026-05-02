#!/usr/bin/env bash
# Production-Control-System ????????s: Java?iJUnit?j+ Python?ipytest?j??????B
# ????m?F: ./test.sh   |   CI ????: bash test.sh
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"

echo "== [1/2] Maven test (code_java) =="
mvn -f code_java/pom.xml test

echo "== [2/2] pytest (code/python/tests) =="
if ! python3 -c "import pytest" 2>/dev/null; then
  echo "pytest が未インストールです。例: python3 -m pip install pytest" >&2
  exit 1
fi
cd code/python
python3 -m pytest tests/ -q --tb=short "$@"
