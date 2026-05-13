#!/usr/bin/env bash
# 段階2 Python→Java 同一検証を JavaFX なしで実行（移行プラン: CI／ローカルでの二重実行足場）。
# 使い方（リポジトリルートで）:
#   export PM_AI_PLAN_INPUT_PATH=... PM_AI_CODE_PYTHON_DIR=... PM_AI_PYTHON=...  # 他 PM_AI_* はアプリと同様
#   bash scripts/stage2_headless_parity.sh
# スクリプトディレクトリのフォールバックが必要なら第1引数（例: bash scripts/stage2_headless_parity.sh /path/to/code/python）。
#
# 終了コード: 0=すべて一致、1=比較不一致または exit 非0、2=起動前致命的エラー

set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT/code_java"
if [[ $# -gt 0 ]]; then
  exec rtk mvn -q exec:java \
    -Dexec.classpathScope=runtime \
    -Dexec.mainClass=jp.co.pm.ai.planning.stage2.cli.Stage2HeadlessParityMain \
    -Dexec.args="$1"
else
  exec rtk mvn -q exec:java \
    -Dexec.classpathScope=runtime \
    -Dexec.mainClass=jp.co.pm.ai.planning.stage2.cli.Stage2HeadlessParityMain
fi
