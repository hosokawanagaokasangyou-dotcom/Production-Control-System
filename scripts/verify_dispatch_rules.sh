#!/usr/bin/env bash
# Verify dispatch special rules package (Phase 1+).
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PHASE="full"
if [[ "${1:-}" == --phase ]]; then
  PHASE="${2:-1}"
elif [[ "${1:-}" == --full ]]; then
  PHASE="full"
fi
cd "$ROOT"
export PYTHONPATH="$ROOT/code/python${PYTHONPATH:+:$PYTHONPATH}"
PY="python3.14"
if ! command -v "$PY" >/dev/null 2>&1; then
  PY="python3"
fi
echo "[verify] phase=$PHASE python=$($PY --version 2>&1)"
"$PY" -m pytest code/python/tests/dispatch_rules/ -q --tb=short
"$PY" code/python/tools/validate_dispatch_rules.py \
  code/json/dispatch_special_rules/dispatch_special_rules.json --conflicts
if [[ -d code_java ]]; then
  (cd code_java && ./mvnw -q compile)
  (cd code_java && ./mvnw -q test -Dtest='jp.co.pm.ai.desktop.dispatch.rules.**.*Test' 2>/dev/null || \
    ./mvnw -q test -Dtest=DispatchRuleMigrationServiceTest)
fi
echo "[verify] ok"
