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
if [[ "$PHASE" == "1" ]]; then
  "$PY" -m pytest code/python/tests/dispatch_rules/test_migrations.py code/python/tests/dispatch_rules/test_migrations_golden.py code/python/tests/dispatch_rules/test_execution_planner.py code/python/tests/dispatch_rules/test_conflict_checker.py code/python/tests/dispatch_rules/test_simulation.py code/python/tests/dispatch_rules/test_trace_recorder.py -q --tb=short
else
  "$PY" -m pytest code/python/tests/dispatch_rules/ -q --tb=short
fi
"$PY" code/python/tools/validate_dispatch_rules.py \
  code/json/dispatch_special_rules/dispatch_special_rules.json --conflicts
if [[ -d code_java ]]; then
  (cd code_java && ./mvnw -q compile)
  (cd code_java && ./mvnw -q test -Dtest='jp.co.pm.ai.desktop.dispatch.rules.**.*Test' 2>/dev/null || \
    ./mvnw -q test -Dtest=DispatchRuleMigrationServiceTest)
fi
echo "[verify] ok"
