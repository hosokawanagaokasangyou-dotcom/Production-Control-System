#!/usr/bin/env bash
# リポジトリに保存された段階1の計画入力（既定: output/plan_input_tasks.xlsx）で
# Python 段階2（plan_simulation_stage2.py / _generate_plan_impl 正本）を CLI 検証する。
# planning_core は Python 3.14 以上を要求する場合があります（エラー文言を確認）。
# 必要な PM_AI_* は環境で上書き可。例:
#   PM_AI_MASTER_WORKBOOK=/path/to/master.xlsm rtk bash scripts/stage2_python_dispatch_verify_repo.sh
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
export PM_AI_REPO_ROOT="${PM_AI_REPO_ROOT:-$ROOT}"
export PM_AI_CODE_PYTHON_DIR="${PM_AI_CODE_PYTHON_DIR:-$ROOT/code/python}"
export PM_AI_PLAN_INPUT_PATH="${PM_AI_PLAN_INPUT_PATH:-$ROOT/output/plan_input_tasks.xlsx}"
export PM_AI_MASTER_WORKBOOK="${PM_AI_MASTER_WORKBOOK:-$ROOT/master.xlsm}"
export PM_AI_OUTPUT_DIR="${PM_AI_OUTPUT_DIR:-$ROOT/output}"
export PM_AI_SKIP_WORKBOOK_ENV_SHEET="${PM_AI_SKIP_WORKBOOK_ENV_SHEET:-1}"
export PYTHONPATH="${PM_AI_CODE_PYTHON_DIR}${PYTHONPATH:+:${PYTHONPATH:-}}"
cd "$PM_AI_CODE_PYTHON_DIR"
if command -v rtk >/dev/null 2>&1; then
  exec rtk python3 plan_simulation_stage2.py
else
  exec python3 plan_simulation_stage2.py
fi
