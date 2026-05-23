# -*- coding: utf-8 -*-
"""CLI: print attendance preview for stage-3.5 overtime simulation wizard (one JSON line)."""
from __future__ import annotations

import json
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def _emit(payload: dict, exit_code: int = 0) -> None:
    json.dump(payload, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    sys.exit(exit_code)


def main() -> int:
    try:
        from planning_core._core import build_attendance_overtime_preview_dict

        payload = build_attendance_overtime_preview_dict()
        _emit(payload, 0 if payload.get("ok") else 1)
        return 0
    except Exception as e:
        import traceback

        _emit(
            {
                "format_version": 1,
                "ok": False,
                "error": str(e),
                "traceback": traceback.format_exc()[:3000],
                "members": [],
                "dates": [],
                "cells": {},
            },
            1,
        )
        return 1


if __name__ == "__main__":
    try:
        import workbook_env_bootstrap as _wbe_exit

        sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
    except ImportError:
        sys.exit(main())
