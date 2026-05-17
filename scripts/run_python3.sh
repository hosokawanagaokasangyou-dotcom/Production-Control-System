#!/usr/bin/env bash
# Run a Python 3 script using scripts/resolve_python3.sh (for Git hooks).
set -euo pipefail
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PY="$("$ROOT/scripts/resolve_python3.sh")"
case "$PY" in
  'py -3') py -3 "$@" ;;
  *) "$PY" "$@" ;;
esac
