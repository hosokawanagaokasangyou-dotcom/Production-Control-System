# -*- coding: utf-8 -*-
"""
VBA ボタンから起動: マクロブック内「配台計画_タスク入力」の「配台試行順番」を
小数を含むキーとして昇順に行を並べ替え、1..n に振り直す。

マスタ読込・_apply_planning_sheet_post_load_mutations・fill_plan は行わない。

環境変数 TASK_INPUT_WORKBOOK にマクロブックのフルパスが入っていること（VBA が設定）。
Excel で本ブックを開いたまま実行すること（xlwings が接続）。
"""
import logging
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass

import planning_core as pc  # noqa: E402


def main() -> int:
    logging.info("apply_plan_input_dispatch_trial_order_sort_by_float_keys: 開始")
    try:
        import planning_core._core as _pc_core

        logging.info(
            "planning_core._core 実体パス（旧 _core だと InvalidIndexError のまま）: %s",
            getattr(_pc_core, "__file__", "?"),
        )
    except Exception:
        pass
    ok = pc.sort_plan_input_dispatch_trial_order_by_float_keys_only()
    logging.info(
        "apply_plan_input_dispatch_trial_order_sort_by_float_keys: 終了 ok=%s", ok
    )
    return 0 if ok else 1


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
