# -*- coding: utf-8 -*-
"""
段階3.0: 入力3表（枝番タスク）で配台Aを実行し、枝番統合まで行う。

前提: 段階2.0 → 配台計画手動修正 → 入力3表生成（stage3_input_builder）で
``plan_input_tasks.xlsx`` の第2シート（配台計画_タスク入力3.0）が生成済み。
"""
import os
import sys
import ctypes

_py_here = os.path.dirname(os.path.abspath(__file__))
if _py_here:
    sys.path.insert(0, _py_here)

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


def main():
    try:
        from planning_core.stage2_identical_dispatch_runner import run_stage3_generate_plan

        run_stage3_generate_plan(qty_strict=False)
    except FileNotFoundError as e:
        print(str(e).strip() or "入力3表が見つかりません。", file=sys.stderr)
        sys.exit(2)
    except pc.PlanningValidationError as e:
        msg = str(e).strip() or "配台計画の検証で中断しました。"
        if not os.path.isfile(pc.stage2_blocking_message_path):
            pc._write_stage2_blocking_message(msg)
        print(msg, file=sys.stderr)
        sys.exit(3)


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
