# -*- coding: utf-8 -*-
"""JSON sidecars for stage-2 plan workbook (result task sheet) to limit Excel re-read."""

from __future__ import annotations

import json
import os

import pandas as pd

# 0/false/no/off/none: do not read or write sidecar JSON
ENV_PLAN_RESULT_TASK_JSON = "PM_AI_PLAN_RESULT_TASK_JSON"
ENV_PLAN_RESULT_TASK_JSON_PATH = "PM_AI_PLAN_RESULT_TASK_JSON_PATH"

RESULT_TASK_JSON_SUFFIX = "_\u7d50\u679c_\u30bf\u30b9\u30af\u4e00\u89a7.json"
SIDE_FORMAT_VERSION = 1


def _plan_result_task_json_disabled() -> bool:
    v = (os.environ.get(ENV_PLAN_RESULT_TASK_JSON) or "").strip().lower()
    return v in ("0", "false", "no", "off", "none")


def result_task_sidecar_path(plan_xlsx: str) -> str:
    base, _ = os.path.splitext(plan_xlsx)
    return base + RESULT_TASK_JSON_SUFFIX


def read_result_task_dataframe(plan_xlsx: str) -> pd.DataFrame | None:
    """
    Returns None if sidecar is disabled, missing, or invalid (caller uses read_excel).
    """
    if not plan_xlsx or _plan_result_task_json_disabled():
        return None
    ex = (os.environ.get(ENV_PLAN_RESULT_TASK_JSON_PATH) or "").strip()
    if ex and os.path.isfile(ex):
        p = ex
    else:
        p = result_task_sidecar_path(plan_xlsx)
        if not os.path.isfile(p):
            return None
    try:
        with open(p, encoding="utf-8-sig") as f:
            data = json.load(f)
        if isinstance(data, dict) and "rows" in data:
            rows = data["rows"]
        elif isinstance(data, list):
            rows = data
        else:
            return pd.DataFrame()
        if not rows:
            return pd.DataFrame()
        return pd.DataFrame(rows)
    except Exception:
        return None


def write_result_task_json_sidecar(plan_xlsx: str, df: pd.DataFrame, *, sheet_name: str) -> str | None:
    if _plan_result_task_json_disabled():
        return None
    try:
        if df is None or getattr(df, "empty", True):
            return None
        out = result_task_sidecar_path(plan_xlsx)
        try:
            if os.path.isfile(out):
                os.remove(out)
        except OSError:
            pass
        rows = json.loads(df.to_json(orient="records", date_format="iso", double_precision=15))
        payload = {
            "format_version": SIDE_FORMAT_VERSION,
            "sheet_name": sheet_name,
            "columns": list(df.columns),
            "row_count": int(len(df)),
            "rows": rows,
        }
        with open(out, encoding="utf-8", newline="\n") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            f.write("\n")
        return out
    except Exception:
        return None
