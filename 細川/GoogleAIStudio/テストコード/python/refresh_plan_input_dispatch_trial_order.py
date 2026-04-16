# -*- coding: utf-8 -*-
"""
VBA / cmd から起動: マクロブック内「配台計画_タスク入力」の「配台試行順番」を
``planning_core.refresh_plan_input_dispatch_trial_order_only()`` 経由で更新する。

- 環境変数 ``PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY`` を 1 / true / yes / on / y 等にしたときは、
  xlwings で **post_load（``_apply_planning_sheet_post_load_mutations``）を行わず**、
  シート上の表だけを対象に ``fill_plan_dispatch_trial_order_column_stage1`` で試行順を再付与し、
  行を試行順の昇順に並べ替える。

- 上記を未設定・0 / 無効のときは、段階2の ``load_planning_tasks_df`` と同趣旨で
  post_load（設定シートの行同期・分割行の自動配台不要等）を実行してから試行順を更新する。

環境変数 ``TASK_INPUT_WORKBOOK`` にマクロブックのフルパスが入っていること（VBA が設定）。
Excel で本ブックを開いたまま実行すること（xlwings が接続）。
"""
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass

import planning_core as pc


def main() -> int:
    ok = pc.refresh_plan_input_dispatch_trial_order_only()
    return 0 if ok else 1


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
