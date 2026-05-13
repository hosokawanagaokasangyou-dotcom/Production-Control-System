#!/usr/bin/env bash
# 段階2 Python→Java 同一検証ランナー（JUnit）を CI 用に実行する。
# 使い方（リポジトリルート）:
#   bash scripts/stage2_headless_parity_ci.sh
# 前提: Ubuntu 等で Python 3.14+ と code/python/requirements.txt の依存が import 可能であること。
#       GitHub Actions では .github/workflows/ci.yml の stage2-headless-parity ジョブが deadsnakes で 3.14 を入れる。

set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
export PM_AI_STAGE2_HEADLESS_CI=1
export PM_AI_PYTHON="${PM_AI_PYTHON:-/usr/bin/python3.14}"
cd "$ROOT/code_java"
exec rtk mvn -q test -Dtest=Stage2HeadlessParityCiTest
