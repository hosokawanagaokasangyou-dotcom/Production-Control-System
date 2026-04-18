# -*- coding: utf-8 -*-
"""
サマリシート「配台試行順_パターン別段階2」の B3（採用パターンID）と B2（バッチ出力ルート）に基づき、
当該バッチの pattern_jobs_meta.json を読み、選んだパターンの配台試行順を「配台計画_タスク入力」に反映する。

先に apply_dispatch_trial_pattern_stage2_batch.py を実行してサマリとメタ JSON を作成しておくこと。
環境変数 TASK_INPUT_WORKBOOK、Excel でマクロブックを開いたまま実行すること。
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
    logging.info("apply_dispatch_pattern_stage2_selection: 開始")
    ok = pc.refresh_dispatch_pattern_stage2_selection_to_plan_only()
    logging.info("apply_dispatch_pattern_stage2_selection: 終了 ok=%s", ok)
    return 0 if ok else 1


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
