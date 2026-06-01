# -*- coding: utf-8 -*-
"""
段階3.1(時間外): 入力3表（枝番）で残業/休出シミュ付きで配台Aを実行し、枝番統合まで行う。

時間外は段階2.1 と同じ仕組み（PM_AI_STAGE2_1_OVERTIME=1 + overtime_simulation_overrides.json を
env 経由で適用）。成果物は段階3.0 と同様に正本 output へ出力し、枝番統合する。
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
    os.environ.setdefault("PM_AI_STAGE2_1_OVERTIME", "1")
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
