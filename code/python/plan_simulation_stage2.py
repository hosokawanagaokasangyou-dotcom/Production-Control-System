# -*- coding: utf-8 -*-
"""
段階2: マクロブックの「配台計画_タスク入力」を読み、特別指定・備考AI反映後に計画シミュレーションを実行する。
"""
import os
import sys
import ctypes

# python -P / PYTHONSAFEPATH ではスクリプト所在ディレクトリが sys.path に入らない。
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
        from planning_core.stage2_identical_dispatch_runner import run_stage2_generate_plan

        run_stage2_generate_plan()
    except FileNotFoundError as e:
        print(str(e).strip() or "マスタブックが見つかりません。", file=sys.stderr)
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
