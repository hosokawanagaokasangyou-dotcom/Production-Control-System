# -*- coding: utf-8 -*-
"""JSON sidecars for stage-2 plan workbook (result task sheet) to limit Excel re-read."""

from __future__ import annotations

import json
import os

import pandas as pd

# 0/false/no/off/none: do not read or write sidecar JSON
ENV_PLAN_RESULT_TASK_JSON = "PM_AI_PLAN_RESULT_TASK_JSON"
ENV_PLAN_RESULT_TASK_JSON_PATH = "PM_AI_PLAN_RESULT_TASK_JSON_PATH"
# 0/false/no/off/none: do not write production_plan_multi_day_*.json (full workbook mirror)
ENV_PLAN_WORKBOOK_JSON = "PM_AI_PLAN_WORKBOOK_JSON"
# 0/false/no/off/none: do not write member_schedule_*.json (mirror of member_schedule_*.xlsx)
ENV_MEMBER_SCHEDULE_JSON = "PM_AI_MEMBER_SCHEDULE_JSON"

RESULT_TASK_JSON_SUFFIX = "_\u7d50\u679c_\u30bf\u30b9\u30af\u4e00\u89a7.json"
SIDE_FORMAT_VERSION = 1
WORKBOOK_JSON_FORMAT_VERSION = 2


def _plan_result_task_json_disabled() -> bool:
    v = (os.environ.get(ENV_PLAN_RESULT_TASK_JSON) or "").strip().lower()
    return v in ("0", "false", "no", "off", "none")


def _plan_workbook_json_disabled() -> bool:
    v = (os.environ.get(ENV_PLAN_WORKBOOK_JSON) or "").strip().lower()
    return v in ("0", "false", "no", "off", "none")


def _member_schedule_json_disabled() -> bool:
    v = (os.environ.get(ENV_MEMBER_SCHEDULE_JSON) or "").strip().lower()
    return v in ("0", "false", "no", "off", "none")


def _dump_xlsx_all_sheets_to_json(
    xlsx_path: str, *, payload_extra: dict | None = None
) -> str | None:
    """
    xlsx を pandas で全シート読み、同名ベースの .json に書き出す内部共通処理。
    """
    if not xlsx_path or not os.path.isfile(xlsx_path):
        return None
    out_path = os.path.splitext(xlsx_path)[0] + ".json"
    try:
        try:
            if os.path.isfile(out_path):
                os.remove(out_path)
        except OSError:
            pass
        sheets_in = pd.read_excel(xlsx_path, sheet_name=None, engine="openpyxl")
    except Exception:
        return None
    sheets_out: dict[str, dict] = {}
    for name, df in (sheets_in or {}).items():
        if df is None or getattr(df, "empty", True):
            sheets_out[name] = {"columns": [], "row_count": 0, "rows": []}
            continue
        rows = json.loads(
            df.to_json(orient="records", date_format="iso", double_precision=15)
        )
        sheets_out[name] = {
            "columns": list(df.columns),
            "row_count": int(len(df)),
            "rows": rows,
        }
    payload = {
        "format_version": WORKBOOK_JSON_FORMAT_VERSION,
        "source_xlsx": os.path.basename(xlsx_path),
        "sheets": sheets_out,
    }
    if payload_extra:
        payload.update(payload_extra)
    try:
        with open(out_path, encoding="utf-8", newline="\n") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            f.write("\n")
        return out_path
    except Exception:
        return None


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


def write_production_plan_workbook_json(plan_xlsx: str) -> str | None:
    """
    ``production_plan_multi_day_*.xlsx`` と同名ベースの ``.json`` に、全シートを表形式で書き出す。
    図形・セル書式は含まず、セル値のみ（Excel 再読込と同様）。
    """
    if _plan_workbook_json_disabled():
        return None
    return _dump_xlsx_all_sheets_to_json(plan_xlsx)


def write_member_schedule_workbook_json(member_xlsx: str) -> str | None:
    """
    ``member_schedule_*.xlsx`` と同名ベースの ``.json`` に全シートを書き出す（計画ブック JSON と同形式。
    シート名はメンバー名）。
    """
    if _member_schedule_json_disabled():
        return None
    return _dump_xlsx_all_sheets_to_json(
        member_xlsx, payload_extra={"workbook_kind": "member_schedule"}
    )
