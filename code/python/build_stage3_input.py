# -*- coding: utf-8 -*-
"""
段階3.0 前処理 CLI: 手動修正後の結果_配台表.json から枝番タスクを分解し入力3表を生成する。

usage: build_stage3_input.py <結果_配台表.json> <plan_input_tasks.xlsx>
（省略時は PM_AI_RESULT_DISPATCH_JSON / PM_AI_PLAN_INPUT_PATH を参照）
"""
import os
import sys

_py_here = os.path.dirname(os.path.abspath(__file__))
if _py_here:
    sys.path.insert(0, _py_here)

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def main() -> int:
    from planning_core.stage3_input_builder import main as builder_main

    return builder_main(sys.argv)


if __name__ == "__main__":
    sys.exit(main())
