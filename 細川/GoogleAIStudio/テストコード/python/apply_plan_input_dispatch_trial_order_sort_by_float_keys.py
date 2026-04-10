# -*- coding: utf-8 -*-
"""
VBA ボタンから起動: マクロブック内「配台計画_タスク入力」の「配台試行順番」を
小数を含むキーとして昇順に行を並べ替え、1..n に振り直す。

マスタ読込・_apply_planning_sheet_post_load_mutations・fill_plan は行わない。

環境変数 TASK_INPUT_WORKBOOK にマクロブックのフルパスが入っていること（VBA が設定）。
Excel で本ブックを開いたまま実行すること（xlwings が接続）。
"""
from __future__ import annotations

import json
import logging
import os
import sys
import time
import traceback

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass

import planning_core as pc  # noqa: E402

# #region agent log
def _agent_dbg(hypothesis_id: str, location: str, message: str, data: dict) -> None:
    log_path = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "debug-3f29a7.log")
    )
    try:
        rec = {
            "sessionId": "3f29a7",
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": int(time.time() * 1000),
        }
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except OSError:
        pass


# #endregion


def main() -> int:
    # #region agent log
    _agent_dbg(
        "A",
        "apply_plan_input_dispatch_trial_order_sort_by_float_keys.py:main:entry",
        "startup",
        {
            "cwd": os.getcwd(),
            "task_input_workbook": os.environ.get("TASK_INPUT_WORKBOOK", ""),
            "argv": sys.argv[:5],
        },
    )
    # #endregion
    logging.info("apply_plan_input_dispatch_trial_order_sort_by_float_keys: 開始")
    ok = pc.sort_plan_input_dispatch_trial_order_by_float_keys_only()
    logging.info(
        "apply_plan_input_dispatch_trial_order_sort_by_float_keys: 終了 ok=%s", ok
    )
    # #region agent log
    _agent_dbg(
        "D",
        "apply_plan_input_dispatch_trial_order_sort_by_float_keys.py:main:exit",
        "sort finished",
        {"ok": bool(ok)},
    )
    # #endregion
    return 0 if ok else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        # #region agent log
        _agent_dbg(
            "C",
            "apply_plan_input_dispatch_trial_order_sort_by_float_keys.py:__main__",
            "uncaught",
            {"exc_type": type(sys.exc_info()[1]).__name__, "tb": traceback.format_exc()},
        )
        # #endregion
        raise
