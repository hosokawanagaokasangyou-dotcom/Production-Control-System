# -*- coding: utf-8 -*-
"""
VBA / cmd から起動: マクロブック内に「配台試行順_パターン一覧」シートを作成・更新する。

- 現在の「配台計画_タスク入力」を xlwings で読み（未保存分を反映）、
  パターン P1（納期最優先）・P2（機械名グループ+納期）・R*（ランダム）の
  確定配台試行順を一覧化する。
- ランダム件数: DISPATCH_TRIAL_PATTERN_RANDOM_COUNT（既定 3）。OS 環境変数または
  マクロブック「設定_環境変数」シート（A=変数名・B=値）。import 前の workbook_env_bootstrap で反映
- シート名: 環境変数 DISPATCH_TRIAL_PATTERN_LIST_SHEET（既定 配台試行順_パターン一覧）
- 見出し行を Excel の「表」にし、換算数量列（内部は total_qty_m）を含む。表を付けないときは DISPATCH_TRIAL_PATTERN_LIST_NO_EXCEL_TABLE=1

環境変数 TASK_INPUT_WORKBOOK にマクロブックのフルパス（VBA が設定）。
Excel で本ブックを開いたまま実行すること（xlwings が接続）。
"""
import os
import sys

if os.name == "nt" and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import logging

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass

import planning_core as pc  # noqa: E402


def main() -> int:
    logging.info("apply_dispatch_trial_pattern_list_sheet: 開始")
    ok = pc.refresh_dispatch_trial_pattern_list_sheet_only()
    logging.info("apply_dispatch_trial_pattern_list_sheet: 終了 ok=%s", ok)
    return 0 if ok else 1


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
