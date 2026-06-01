# -*- coding: utf-8 -*-
"""
段階3.2(数量厳守): 入力3表で配台Aを実行する（同日完走必須・定常外人ブロック無視）。

env ``PM_AI_STAGE3_2_QTY_STRICT=1`` を付与して ``_generate_plan_impl`` を分岐させ、
実行後は段階3.0 と同一の枝番統合を行う。
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

        run_stage3_generate_plan(qty_strict=True)
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
