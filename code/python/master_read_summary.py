# -*- coding: utf-8 -*-
"""CLI: print master workbook read summary as one JSON line (for JavaFX tab)."""
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("-h", "--help"):
        print("usage: master_read_summary.py", file=sys.stderr)
        sys.exit(0)
    from planning_core.master_read_ui import main

    main()
