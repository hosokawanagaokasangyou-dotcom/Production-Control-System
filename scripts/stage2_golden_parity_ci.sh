#!/usr/bin/env bash
# 段階2 Java 足場の CI 用ワンショット（Python 二重起動は含まない）。
# 使い方:
#   bash scripts/stage2_golden_parity_ci.sh
# 前提: リポジトリルートで実行。JDK + Maven が PATH にあること。
#
# PM_AI_STAGE2_GOLDEN_CI=1 で JUnit Stage2GoldenParityCiTest を有効化する。
# 将来: 同一フィクスチャで Python 子を起動し Stage2ProductionPlanJsonParity をヘッドレス実行する段をここに追加する。

set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
export PM_AI_STAGE2_GOLDEN_CI=1
cd "$ROOT/code_java"
exec mvn -q test -Dtest=Stage2GoldenParityCiTest
