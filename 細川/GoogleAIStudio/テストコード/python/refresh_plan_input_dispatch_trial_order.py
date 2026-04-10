# -*- coding: utf-8 -*-
"""Entry for VBA: refresh dispatch trial order on plan input sheet (see planning_core)."""
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
