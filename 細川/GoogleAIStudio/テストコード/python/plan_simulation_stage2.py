# -*- coding: utf-8 -*-
"""
段階2: マクロブックの「配台計画_タスク入力」を読み、特別指定・備考AI反映後に計画シミュレーションを実行する。
"""
import os
import sys
import ctypes

os.chdir(os.path.dirname(os.path.abspath(__file__)))

if os.name == "nt":
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 3)

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass

import planning_core as pc

try:
    from planning_core.agent_debug_session_log import agent_ndjson_log
except Exception:  # pragma: no cover

    def agent_ndjson_log(**kwargs):
        return None


def main():
    # #region agent log
    agent_ndjson_log(
        hypothesis_id="E",
        location="plan_simulation_stage2.py:main",
        message="段階2 main 開始",
        data={"cwd": os.getcwd()},
    )
    # #endregion
    try:
        pc.generate_plan()
    except pc.PlanningValidationError as e:
        msg = str(e).strip() or "配台計画の検証で中断しました。"
        # #region agent log
        agent_ndjson_log(
            hypothesis_id="B",
            location="plan_simulation_stage2.py:main",
            message="PlanningValidationError",
            data={"msg": msg[:800]},
        )
        # #endregion
        if not os.path.isfile(pc.stage2_blocking_message_path):
            pc._write_stage2_blocking_message(msg)
        print(msg, file=sys.stderr)
        sys.exit(3)
    except Exception as e:
        # #region agent log
        agent_ndjson_log(
            hypothesis_id="D",
            location="plan_simulation_stage2.py:main",
            message="段階2 未捕捉例外",
            data={
                "exc_type": type(e).__name__,
                "exc_str": str(e)[:800],
            },
        )
        # #endregion
        raise


if __name__ == "__main__":
    main()
