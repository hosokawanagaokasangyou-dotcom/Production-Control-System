# -*- coding: utf-8 -*-
"""加工実績明細から速度分布を更新する CLI（背景）。"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

os.chdir(_SCRIPT_DIR)

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def main(argv: list[str] | None = None) -> int:
    from planning_core.actual_speed_distribution import update_speed_distribution, write_ml_readiness
    from planning_core.dispatch_workspace import resolve_dispatch_learning_archive_root

    p = argparse.ArgumentParser()
    p.add_argument("--archive-root")
    p.add_argument("--force-full", action="store_true")
    args = p.parse_args(argv)
    root = Path(args.archive_root or resolve_dispatch_learning_archive_root()).resolve()
    summary = update_speed_distribution(root, force_full=args.force_full)
    write_ml_readiness(root)
    print(summary, flush=True)
    return 0


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(lambda: main()))
