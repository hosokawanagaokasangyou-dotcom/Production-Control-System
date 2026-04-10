# -*- coding: utf-8 -*-
"""
VBA / cmd ????: ????_????????????????????

planning_core.refresh_plan_input_dispatch_trial_order_via_xlwings ???
??2? generate_plan ??? _apply_planning_sheet_post_load_mutations ?
fill_plan_dispatch_trial_order_column_stage1 ??????????
????? Windows ? os.system("pause") ??cmd ??????????????????????
"""
import os
import sys
import traceback

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def _cmd_pause() -> None:
    """cmd ? pause???????????"""
    if sys.platform == "win32":
        os.system("pause")


def main() -> int:
    import planning_core as pc

    ok = pc.refresh_plan_input_dispatch_trial_order_only()
    return 0 if ok else 1


if __name__ == "__main__":
    rc = 1
    failed = True
    try:
        rc = main()
        failed = rc != 0
    except Exception:
        traceback.print_exc()
        rc = 1
        failed = True
    if failed:
        _cmd_pause()
    sys.exit(rc)
