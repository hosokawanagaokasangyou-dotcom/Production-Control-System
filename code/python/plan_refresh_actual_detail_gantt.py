# -*- coding: utf-8 -*-
"""
実績明細に基づく「結果_設備ガント_実績明細」シートのみを output に出力する（段階2は実行しない）。
マクロから TASK_INPUT_WORKBOOK が渡されたブックを前提に planning_core を読む。
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

import logging

import planning_core as pc


def main():
    try:
        out = pc.refresh_equipment_gantt_actual_detail_only()
        print(out)
    except pc.PlanningValidationError as e:
        msg = str(e).strip() or "実績明細ガントの生成を中断しました。"
        if not os.path.isfile(pc.stage2_blocking_message_path):
            pc._write_stage2_blocking_message(msg)
        print(msg, file=sys.stderr)
        sys.exit(3)
    except SystemExit:
        raise
    except Exception:
        logging.exception("plan_refresh_actual_detail_gantt")
        sys.exit(1)


if __name__ == "__main__":
    main()
