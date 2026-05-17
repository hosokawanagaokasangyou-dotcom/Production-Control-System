#!/usr/bin/env bash
# Print one working Python 3 launcher for Git hooks (stdout). Exit 1 if none.
# Windows Git Bash: python3 in PATH is often the Store stub (exit 49); prefer python or py -3.
set -euo pipefail

try_python() {
  local cmd=("$@")
  "${cmd[@]}" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' 2>/dev/null
}

if command -v python3 >/dev/null 2>&1 && try_python python3; then
  echo python3
  exit 0
fi
if command -v python >/dev/null 2>&1 && try_python python; then
  echo python
  exit 0
fi
if command -v py >/dev/null 2>&1 && try_python py -3; then
  echo 'py -3'
  exit 0
fi
if [[ -x /usr/bin/python3 ]] && try_python /usr/bin/python3; then
  echo /usr/bin/python3
  exit 0
fi

echo 'resolve_python3: no working Python 3 (tried python3, python, py -3, /usr/bin/python3)' >&2
exit 1
