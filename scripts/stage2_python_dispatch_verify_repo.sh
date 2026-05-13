#!/usr/bin/env bash
# リポジトリに保存された段階1の計画入力（既定: output/plan_input_tasks.xlsx）で
# Python 段階2（plan_simulation_stage2.py / _generate_plan_impl 正本）を CLI 検証する。
#
# planning_core は Python 3.14 以上必須（bootstrap.py）。WSL Ubuntu 既定の python3 だけでは不足することが多い。
# 解決済みインタプリタを使う: export PM_AI_PYTHON=/usr/bin/python3.14
#
# 必要な PM_AI_* は環境で上書き可。例:
#   PM_AI_MASTER_WORKBOOK=/path/to/master.xlsm bash scripts/stage2_python_dispatch_verify_repo.sh
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
export PM_AI_REPO_ROOT="${PM_AI_REPO_ROOT:-$ROOT}"
export PM_AI_CODE_PYTHON_DIR="${PM_AI_CODE_PYTHON_DIR:-$ROOT/code/python}"
export PM_AI_PLAN_INPUT_PATH="${PM_AI_PLAN_INPUT_PATH:-$ROOT/output/plan_input_tasks.xlsx}"
export PM_AI_MASTER_WORKBOOK="${PM_AI_MASTER_WORKBOOK:-$ROOT/master.xlsm}"
export PM_AI_OUTPUT_DIR="${PM_AI_OUTPUT_DIR:-$ROOT/output}"
export PM_AI_SKIP_WORKBOOK_ENV_SHEET="${PM_AI_SKIP_WORKBOOK_ENV_SHEET:-1}"
export PYTHONPATH="${PM_AI_CODE_PYTHON_DIR}${PYTHONPATH:+:${PYTHONPATH:-}}"

pick_python_ge_314() {
    local bin
    local -a try_bins=()
    if [ -n "${PM_AI_PYTHON:-}" ]; then
        try_bins+=("${PM_AI_PYTHON}")
    fi
    try_bins+=(python3.14 python3.14t)
    if command -v pyenv >/dev/null 2>&1; then
        local pw
        pw="$(pyenv which python3.14 2>/dev/null || true)"
        [ -n "$pw" ] && [ -x "$pw" ] && try_bins+=("$pw")
    fi
    try_bins+=(python3 python)
    for bin in "${try_bins[@]}"; do
        [ -z "$bin" ] && continue
        if ! command -v "$bin" >/dev/null 2>&1 && [ ! -x "$bin" ]; then
            continue
        fi
        if "$bin" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 14) else 1)' 2>/dev/null; then
            printf '%s\n' "$bin"
            return 0
        fi
    done
    return 1
}

print_wsl_ubuntu_hint() {
    cat <<'EOF' >&2
[stage2-verify] Python 3.14 以上が PATH 上に見つかりません（planning_core の要件）。
次のいずれかで 3.14 を入れ、PM_AI_PYTHON にフルパスを指定してください。

■ Ubuntu / WSL（deadsnakes PPA の例）
  sudo add-apt-repository -y ppa:deadsnakes/ppa
  sudo apt update
  sudo apt install -y python3.14 python3.14-venv python3.14-dev
  export PM_AI_PYTHON=/usr/bin/python3.14

■ pyenv の例
  pyenv install 3.14.0
  pyenv local 3.14.0   # または pyenv shell
  export PM_AI_PYTHON="$(pyenv which python)"

■ 正本の依存関係
  code/python/pyproject.toml は requires-python = ">=3.14"。
  パッケージは requirements.txt / setup_environment.py（VBA「環境構築」）が正本です。
EOF
}

PY_BIN="$(pick_python_ge_314 || true)"
if [ -z "$PY_BIN" ]; then
    print_wsl_ubuntu_hint
    exit 2
fi

cd "$PM_AI_CODE_PYTHON_DIR"
echo "[stage2-verify] 使用する Python: $PY_BIN ($("$PY_BIN" -c 'import sys; print(sys.version)' | head -1))" >&2
echo "[stage2-verify] ModuleNotFoundError のときは（この Python に対して）: \"${PY_BIN}\" -m pip install -r \"${PM_AI_REPO_ROOT}/code/python/requirements.txt\"" >&2
exec "$PY_BIN" plan_simulation_stage2.py
