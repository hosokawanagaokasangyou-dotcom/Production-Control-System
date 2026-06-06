# -*- coding: utf-8 -*-
"""planning_core 実装ファサード。実体は planning_core.core を exec 連結した共有名前空間。"""
from __future__ import annotations

import sys as _sys
import types as _types
from pathlib import Path as _Path

_CORE_DIR = _Path(__file__).resolve().parent / "core"
_MODULE_ORDER = [
    "state",
    "columns",
    "gemini_auth",
    "gantt_excel",
    "plan_input",
    "task_queue",
    "stage1",
    "time_utils",
    "master_data",
    "roll_pipeline",
    "dispatch_loop",
    "output_refresh",
    "stage2_impl",
]

_ns = _sys.modules[__name__].__dict__
_ns["__name__"] = "planning_core._core"

def _exec_into_ns(filename: str) -> None:
    path = _CORE_DIR / filename
    code = path.read_text(encoding="utf-8")
    exec(compile(code, str(path), "exec"), _ns)  # noqa: S102

_exec_into_ns("_bootstrap.py")
for _part in _MODULE_ORDER:
    _exec_into_ns(f"{_part}.py")

# サブモジュール互換（import planning_core.core.roll_pipeline 等）
for _part in _MODULE_ORDER:
    _proxy = _types.ModuleType(f"planning_core.core.{_part}")
    _proxy.__dict__.update(_ns)
    _sys.modules[f"planning_core.core.{_part}"] = _proxy
