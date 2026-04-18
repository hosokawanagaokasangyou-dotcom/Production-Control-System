# -*- coding: utf-8 -*-
"""
VBA ボタンから起動: マクロブック内「配台計画_タスク入力」の「配台試行順番」を
段階2の ``load_planning_tasks_df`` と同趣旨（``apply_exclude_rules_from_config=False`` の
``_apply_planning_sheet_post_load_mutations`` 後に ``fill_plan_dispatch_trial_order_column_stage1``）
で再計算し、行を試行順昇順に並べ替える。

環境変数 TASK_INPUT_WORKBOOK にマクロブックのフルパスが入っていること（VBA が設定）。
Excel で本ブックを開いたまま実行すること（xlwings が接続）。
"""
import os
import sys

# sys.getdefaultencoding() は Python 3 で utf-8。Windows コンソールの標準入出力は
# ロケールで cp932 になりがちなので、ログ出力の文字化けを抑えるため先に UTF-8 へ合わせる。
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
    logging.info("apply_plan_input_dispatch_trial_order: 開始")
    ok = pc.refresh_plan_input_dispatch_trial_order_only()
    logging.info("apply_plan_input_dispatch_trial_order: 終了 ok=%s", ok)
    return 0 if ok else 1


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))