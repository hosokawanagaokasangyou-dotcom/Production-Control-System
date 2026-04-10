# -*- coding: utf-8 -*-
"""
VBA / cmd ????: ????_????????????????????

planning_core.refresh_plan_input_dispatch_trial_order_via_xlwings ???
??2? generate_plan ??? _apply_planning_sheet_post_load_mutations ?
fill_plan_dispatch_trial_order_column_stage1 ??????????
"""
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass

import planning_core as pc


def main() -> int:
    ok = pc.refresh_plan_input_dispatch_trial_order_only()
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
