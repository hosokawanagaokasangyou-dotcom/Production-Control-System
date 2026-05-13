#!/usr/bin/env bash
set -euo pipefail
# git filter-branch --index-filter から呼ぶ。国分 PMD 初期一式に誤コミットされた jlink modules（100MB超）を履歴から除去する。
git rm --cached --ignore-unmatch "国分工場_実行環境/PMD_initial_install/runtime/lib/modules" || true
