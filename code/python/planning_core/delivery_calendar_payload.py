# -*- coding: utf-8 -*-
"""Delivery-calendar JSON payload for pm_ai_delivery_calendar_view.py / JavaFX (ASCII source + \\u escapes)."""

from __future__ import annotations

import json
import logging
import math
import os
import time
from pathlib import Path
from collections import defaultdict
from collections.abc import Iterable, Mapping
from datetime import date, datetime, timedelta
from typing import Any

import pandas as pd

import planning_core._core as core
from planning_core.dispatch_workspace import (
    DEFAULT_ACTUAL_DETAIL_SOURCE_DIR,
    DEFAULT_TASK_INPUT_SOURCE_DIR,
    ENV_ACTUAL_DETAIL_SOURCE_DIR,
    ENV_ACTUAL_DETAIL_WORKBOOK,
    ENV_PROCESSING_PLAN_PATH,
    ENV_TASK_INPUT_SOURCE_DIR,
    read_tabular_dataframe,
    resolve_actual_detail_workbook_path,
    resolve_processing_plan_path_from_env,
    resolve_result_dispatch_table_output_dir,
)

_LOG = logging.getLogger(__name__)

# region agent log
_AGENT_DEBUG_LOG = "/mnt/c/????AI??????_JAVA/.cursor/debug-ebddd7.log"


def _agent_debug_ndjson(hypothesis_id: str, location: str, message: str, data: dict) -> None:
    """Append one NDJSON line for Cursor debug mode (ignore failures)."""
    try:
        payload = {
            "sessionId": "ebddd7",
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": int(time.time() * 1000),
        }
        with open(_AGENT_DEBUG_LOG, "a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except Exception:
        pass


# endregion

# Result dispatch JSON column names (avoid non-ASCII literals in this file on CP932 mounts)
_DIS_JSON_MACH = "\u6a5f\u68b0\u540d"
_DIS_JSON_TID = "\u4f9d\u983cNO"
_DIS_JSON_DAY = "\u914d\u53f0\u65e5"
_DIS_JSON_QTY = "\u5f53\u65e5\u914d\u53f0\u6570\u91cf"

# Actual-detail (NO-(roll)-betsu raw) column names referenced from the Power Query in ”z‘äƒ‹[ƒ‹:
#   “úŽŸWŒv = Group({ˆË—ŠNO, H’ö–¼, ‰ÁH“ú•t}, max(ŽÀ‰ÁH”)) where ‰ÁH“ú•t = DateTime.Date(‰ÁHŠJŽn“úŽž)
#   ƒtƒBƒ‹ƒ^: »‘¢ðŒ(“à–ó) == "’·‚³" ‚Ì‚Ý
_ACT_COL_TID = "\u4f9d\u983cNO"
_ACT_COL_PROC = "\u5de5\u7a0b\u540d"
_ACT_COL_QTY = "\u5b9f\u52a0\u5de5\u6570"
_ACT_COL_START_DT = "\u52a0\u5de5\u958b\u59cb\u65e5\u6642"
_ACT_COL_DAY = "\u52a0\u5de5\u65e5"
_ACT_COL_PRODUCTION_DETAIL = "\u88fd\u9020\u6761\u4ef6(\u5185\u8a33)"
_ACT_PRODUCTION_DETAIL_LENGTH = "\u9577\u3055"

__all__ = ("build_delivery_calendar_payload",)

# Debug session f73cbb: NDJSON path (workspace; Python subprocess may also resolve via repo)
_DEBUG_F73CBB_LOG = "/mnt/c/\u5de5\u7a0b\u7ba1\u7406AI\u30d7\u30ed\u30b8\u30a7\u30af\u30c8_JAVA/.cursor/debug-f73cbb.log"
_DEBUG_PROBE_ENV = "PM_AI_DELIVERY_CALENDAR_PROBE_TASK"
_ENV_CAL_PAST_DAYS = "PM_AI_DELIVERY_CALENDAR_PAST_DAYS"
_ENV_CAL_FUTURE_DAYS = "PM_AI_DELIVERY_CALENDAR_FUTURE_DAYS"

# Short weekday for calendar column titles (Mon=\u6708 ... Sun=\u65e5)
_JP_WEEKDAY_SHORT = ("\u6708", "\u706b", "\u6c34", "\u6728", "\u91d1", "\u571f", "\u65e5")


# region agent log
def _f73cbb_log_path_candidates() -> list[str]:
    """Workspace path (WSL) plus repo-relative .cursor (Windows Python cannot open /mnt/c/...)."""
    out: list[str] = [_DEBUG_F73CBB_LOG]
    try:
        root = Path(__file__).resolve().parent.parent.parent.parent
        alt = root / ".cursor" / "debug-f73cbb.log"
        s = str(alt)
        if s not in out:
            out.append(s)
    except Exception:
        pass
    return out


def _debug_ndjson_f73cbb(hypothesis_id: str, location: str, message: str, data: dict[str, Any]) -> None:
    """Append one NDJSON line for Cursor debug session f73cbb (ignore failures)."""
    try:
        payload = {
            "sessionId": "f73cbb",
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": int(time.time() * 1000),
        }
        line = json.dumps(payload, ensure_ascii=False) + "\n"
        for path in _f73cbb_log_path_candidates():
            try:
                p = Path(path)
                p.parent.mkdir(parents=True, exist_ok=True)
                with open(path, "a", encoding="utf-8") as f:
                    f.write(line)
                return
            except Exception:
                continue
    except Exception:
        pass


def _resolve_actual_detail_machine_key_for_delivery_calendar(
    row,
    df: pd.DataFrame,
    equip_lookup: dict,
) -> tuple[str, str]:
    """????: ??????????????????????????????????????

    ``_aggregate_daily_actual_qty_aladdin_max`` ????????????????????
    ???????????????????????????????????
    """
    proc = row.get(_ACT_COL_PROC)
    if proc is not None and not (isinstance(proc, float) and pd.isna(proc)):
        proc_key = core._normalize_equipment_match_key(proc)
        canonical = equip_lookup.get(proc_key)
        if canonical:
            mach_raw = str(canonical).strip()
            _, mn_part = core._split_equipment_line_process_machine(mach_raw)
            mach_display = (mn_part or mach_raw).strip()
            mk = core._normalize_equipment_match_key(mach_display)
            if mk:
                return mk, "proc_equip_lookup"
    if core.TASK_COL_MACHINE_NAME in df.columns:
        mv = row.get(core.TASK_COL_MACHINE_NAME)
        if mv is not None and not (isinstance(mv, float) and pd.isna(mv)):
            ms = str(mv).strip()
            if ms:
                mk = core._normalize_equipment_match_key(ms)
                if mk:
                    return mk, "row_machine_column"
    return "", "none"


def _classify_actual_row_for_delivery_calendar(
    row,
    df: pd.DataFrame,
    equip_lookup: dict,
    date_ok: set | None,
) -> tuple[str, dict[str, Any]]:
    """Mirror _aggregate_daily_actual_qty_aladdin_max row acceptance; return (reason_code, detail).

    reason_code: ACCEPT or SKIP_* (hypotheses H_FILTER, H_MACHINE, H_DATE, H_TID, H_QTY).
    """
    detail: dict[str, Any] = {}
    raw_tid = row.get(_ACT_COL_TID)
    detail["raw_tid_repr"] = repr(raw_tid)[:200]
    tid = core.planning_task_id_str_from_scalar(raw_tid)
    detail["norm_tid"] = tid
    if not tid:
        return "SKIP_NO_TID", detail

    has_filter_col = _ACT_COL_PRODUCTION_DETAIL in df.columns
    if has_filter_col:
        cond = row.get(_ACT_COL_PRODUCTION_DETAIL)
        detail["prod_detail_raw"] = repr(cond)[:120]
        if cond is None or (isinstance(cond, float) and pd.isna(cond)):
            return "SKIP_FILTER_NA_COND", detail
        want = core._nfkc_column_aliases(_ACT_PRODUCTION_DETAIL_LENGTH)
        got = core._nfkc_column_aliases(str(cond)).strip()
        detail["prod_detail_nfkc"] = got[:80]
        if got != want:
            return "SKIP_FILTER_NOT_LENGTH", detail

    proc_dbg = row.get(_ACT_COL_PROC)
    detail["proc_raw"] = repr(proc_dbg)[:120]
    detail["equip_canonical_hit"] = False
    if proc_dbg is not None and not (isinstance(proc_dbg, float) and pd.isna(proc_dbg)):
        pk = core._normalize_equipment_match_key(proc_dbg)
        detail["proc_key_norm"] = pk[:80]
        detail["equip_canonical_hit"] = bool(equip_lookup.get(pk))
    detail["machine_name_raw"] = (
        str(row.get(core.TASK_COL_MACHINE_NAME))[:120]
        if core.TASK_COL_MACHINE_NAME in df.columns
        else ""
    )

    mk, mk_src = _resolve_actual_detail_machine_key_for_delivery_calendar(row, df, equip_lookup)
    detail["machine_key_resolution"] = mk_src
    detail["machine_key_norm"] = mk[:120] if mk else ""
    if not mk:
        if proc_dbg is None or (isinstance(proc_dbg, float) and pd.isna(proc_dbg)):
            return "SKIP_NO_MACHINE_NO_PROC", detail
        if not detail["equip_canonical_hit"]:
            return "SKIP_MACHINE_LOOKUP_MISS", detail
        return "SKIP_NO_MACHINE_KEY", detail

    d = _row_actual_day(row)
    detail["day_iso"] = d.isoformat() if d is not None else ""
    detail["start_dt_raw"] = repr(row.get(_ACT_COL_START_DT))[:80]
    detail["proc_day_raw"] = repr(row.get(_ACT_COL_DAY))[:80]
    if d is None:
        return "SKIP_NO_DAY", detail
    if date_ok is not None and d not in date_ok:
        detail["window_min"] = min(date_ok).isoformat() if date_ok else ""
        detail["window_max"] = max(date_ok).isoformat() if date_ok else ""
        return "SKIP_DAY_OUT_OF_RANGE", detail

    try:
        q = core.parse_float_safe(row.get(_ACT_COL_QTY), None)
    except Exception:
        q = None
    detail["qty_parsed"] = q
    if q is None:
        return "SKIP_BAD_QTY_PARSE", detail
    try:
        qf = float(q)
    except (TypeError, ValueError):
        return "SKIP_BAD_QTY_COERCE", detail
    if qf <= 1e-12 or math.isnan(qf):
        return "SKIP_QTY_NON_POSITIVE", detail

    return "ACCEPT", detail


def _probe_task_rows_delivery_calendar(
    df_actual: pd.DataFrame | None,
    equipment_list,
    sorted_dates: list,
    probe_literal: str,
    actual_agg: dict,
    plan_pairs: set[tuple[str, str]],
    eligible_pairs: set[tuple[str, str]],
) -> dict[str, Any]:
    """Log per-row pipeline outcome for rows matching probe task id (default Y4-59).

    Also returns a JSON-serializable summary for meta[\"deliveryCalendarProbe\"] (runtime evidence when NDJSON file path fails from Python subprocess).
    """
    empty_out: dict[str, Any] = {
        "probe_literal": probe_literal,
        "probe_norm": core.planning_task_id_str_from_scalar(probe_literal),
        "matched_row_count": 0,
        "reason_counts": {},
        "note": "df_actual_empty",
        "log_path_candidates": _f73cbb_log_path_candidates(),
    }
    if df_actual is None or len(df_actual) == 0:
        _debug_ndjson_f73cbb(
            "H_SETUP",
            "delivery_calendar_payload.py:_probe_task_rows",
            "df_actual_empty",
            {"probe": probe_literal},
        )
        return empty_out

    probe_norm = core.planning_task_id_str_from_scalar(probe_literal)
    equip_lookup = core._equipment_lookup_normalized_to_canonical(equipment_list)
    date_ok = set(sorted_dates) if sorted_dates else None

    matched_count = 0
    reason_counts: dict[str, int] = {}
    row_samples: list[dict[str, Any]] = []
    for row_idx, row in df_actual.iterrows():
        ntid = core.planning_task_id_str_from_scalar(row.get(_ACT_COL_TID))
        raw_s = str(row.get(_ACT_COL_TID) if row.get(_ACT_COL_TID) is not None else "")
        if ntid != probe_norm and probe_literal not in raw_s and not (
            probe_norm and probe_norm in raw_s
        ):
            continue
        matched_count += 1
        code, detail = _classify_actual_row_for_delivery_calendar(row, df_actual, equip_lookup, date_ok)
        reason_counts[code] = reason_counts.get(code, 0) + 1
        if "FILTER" in code:
            hyp = "H_FILTER"
        elif "MACHINE" in code or code == "SKIP_NO_MACHINE_NO_PROC":
            hyp = "H_MACHINE"
        elif "DAY" in code:
            hyp = "H_DATE"
        elif "TID" in code:
            hyp = "H_TID"
        elif "QTY" in code or "BAD" in code:
            hyp = "H_QTY"
        else:
            hyp = "H_OTHER"
        _debug_ndjson_f73cbb(
            hyp,
            "delivery_calendar_payload.py:_probe_task_rows",
            "row_classified",
            {"code": code, "df_index": row_idx, **detail},
        )
        if len(row_samples) < 12:
            row_samples.append(
                {
                    "code": code,
                    "df_index": str(row_idx),
                    "norm_tid": detail.get("norm_tid"),
                    "machine_key_norm": (detail.get("machine_key_norm") or "")[:80],
                    "day_iso": detail.get("day_iso"),
                    "prod_detail_nfkc": (detail.get("prod_detail_nfkc") or "")[:40],
                }
            )

    _debug_ndjson_f73cbb(
        "H_SETUP",
        "delivery_calendar_payload.py:_probe_task_rows",
        "probe_summary",
        {
            "probe_literal": probe_literal,
            "probe_norm": probe_norm,
            "df_actual_rows": len(df_actual),
            "matched_row_count": matched_count,
            "calendar_day_count": len(sorted_dates),
            "reason_counts_after_scan": reason_counts,
        },
    )

    in_actual_agg = False
    for (_mk, _d), tmap in actual_agg.items():
        if probe_norm and probe_norm in tmap:
            in_actual_agg = True
            break

    pairs_with_probe = [(a, b) for (a, b) in eligible_pairs if b == probe_norm]
    plan_has_probe = any(b == probe_norm for (_a, b) in plan_pairs)
    elig_sample = [[a, b] for a, b in pairs_with_probe[:20]]
    _debug_ndjson_f73cbb(
        "H_RESULT",
        "delivery_calendar_payload.py:_probe_task_rows",
        "aggregation_eligibility",
        {
            "probe_norm": probe_norm,
            "in_actual_agg_buckets": in_actual_agg,
            "plan_pairs_contains_probe": plan_has_probe,
            "eligible_pairs_for_probe": pairs_with_probe[:20],
            "eligible_pair_count_for_probe": len(pairs_with_probe),
            "reason_counts": reason_counts,
        },
    )

    diagnosis = ""
    if matched_count == 0:
        diagnosis = (
            "NO_ROWS_MATCH_PROBE: no row matched this probe on TASK_ID column "
            "(check spelling / PM_AI_DELIVERY_CALENDAR_PROBE_TASK)."
        )
    elif reason_counts.get("ACCEPT", 0) >= 1 and in_actual_agg:
        diagnosis = "ACTUAL_AGG_INCLUDES_PROBE: rows accepted by classifier and present in aggregation buckets."
    elif reason_counts.get("ACCEPT", 0) >= 1 and not in_actual_agg:
        diagnosis = "INCONSISTENT_ACCEPT_BUT_NOT_IN_AGG: classifier ACCEPT but not in actual_agg (investigate aggregation vs classify)."
    elif reason_counts:
        top = max(reason_counts.items(), key=lambda kv: kv[1])[0]
        diagnosis = f"Dominant_skip={top} (see reason_counts and row_samples)."

    out: dict[str, Any] = {
        "probe_literal": probe_literal,
        "probe_norm": probe_norm,
        "matched_row_count": matched_count,
        "reason_counts": dict(sorted(reason_counts.items())),
        "row_samples": row_samples,
        "in_actual_agg_buckets": in_actual_agg,
        "plan_pairs_contains_probe": plan_has_probe,
        "eligible_pair_count_for_probe": len(pairs_with_probe),
        "eligible_pairs_sample": elig_sample,
        "df_actual_rows": len(df_actual),
        "calendar_day_count": len(sorted_dates),
        "diagnosis": diagnosis,
        "log_path_candidates": _f73cbb_log_path_candidates(),
    }
    return out


# endregion


def _format_delivery_calendar_date_header(d: date) -> str:
    """Display label like 2026\u5e744\u67081\u65e5(\u571f) for Spreadsheet column headers."""
    if not isinstance(d, date):
        return str(d)
    w = _JP_WEEKDAY_SHORT[d.weekday()]
    return f"{d.year}\u5e74{d.month}\u6708{d.day}\u65e5({w})"


def _parse_cell_date(val) -> date | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    ts = pd.to_datetime(val, errors="coerce")
    if pd.isna(ts):
        return None
    if isinstance(ts, pd.Timestamp):
        return ts.date()
    if isinstance(ts, datetime):
        return ts.date()
    return None


def _format_cell(v) -> str:
    if v is None or (isinstance(v, float) and (pd.isna(v) or math.isnan(v))):
        return ""
    if isinstance(v, int) and not isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        s = f"{v:.4f}".rstrip("0").rstrip(".")
        return s if s else "0"
    if isinstance(v, (datetime, date, pd.Timestamp)):
        if hasattr(v, "date"):
            try:
                return v.date().isoformat()
            except Exception:
                pass
        return str(v)[:10]
    return str(v).strip()


def _read_plan_tasks_from_processing_plan_env() -> pd.DataFrame | None:
    resolve_processing_plan_path_from_env()
    pp = (os.environ.get(ENV_PROCESSING_PLAN_PATH) or "").strip()
    if not pp or not os.path.isfile(pp):
        _LOG.warning("delivery_calendar: invalid PM_AI_PROCESSING_PLAN_PATH")
        return None
    sheet_name = (os.environ.get(core.ENV_COMPARE_GANTT_PLAN_TASKS_SHEET, "") or "").strip()
    if not sheet_name:
        sheet_name = core.TASKS_SHEET_NAME
    low = pp.lower()
    try:
        if low.endswith((".csv", ".parquet", ".pq")):
            df = read_tabular_dataframe(pp)
        else:
            df = read_tabular_dataframe(pp, sheet_name=sheet_name)
        df.columns = df.columns.astype(str).str.strip()
        df = core._align_dataframe_headers_to_canonical(df, list(core.SOURCE_BASE_COLUMNS))
        _LOG.info(
            "delivery_calendar: loaded plan tasks path=%s sheet=%s rows=%s",
            os.path.basename(pp),
            sheet_name,
            len(df),
        )
        return df
    except Exception as e:
        _LOG.warning("delivery_calendar: plan task load failed (%s)", e)
        return None


def _resolve_dispatch_json_path(processing_plan_path: str) -> str | None:
    out_dir = resolve_result_dispatch_table_output_dir(processing_plan_path or "")
    if not out_dir:
        return None
    p = os.path.join(out_dir, core.RESULT_DISPATCH_TABLE_JSON_FILENAME)
    return os.path.abspath(p) if p else None


def _load_dispatch_json_rows(path: str | None) -> tuple[list[str], list[dict[str, Any]]]:
    if not path or not os.path.isfile(path):
        return [], []
    try:
        raw = json.loads(open(path, encoding="utf-8").read())
    except Exception as e:
        _LOG.warning("delivery_calendar: dispatch json load failed (%s)", e)
        return [], []
    cols = raw.get("columns")
    rows = raw.get("rows")
    if not isinstance(cols, list) or not isinstance(rows, list):
        return [], []
    header = [str(c) for c in cols]
    out: list[dict[str, Any]] = []
    for r in rows:
        if isinstance(r, Mapping):
            out.append(dict(r))
        elif isinstance(r, list):
            m = {}
            for i, k in enumerate(header):
                if i < len(r):
                    m[k] = r[i]
            out.append(m)
    return header, out


def _aggregate_dispatch_quantities(rows: Iterable[dict[str, Any]]):
    agg: dict[tuple[str, date, str], float] = defaultdict(float)
    for row in rows:
        mach = row.get(_DIS_JSON_MACH)
        mk = core._normalize_equipment_match_key(str(mach or ""))
        if not mk:
            continue
        tid = core.planning_task_id_str_from_scalar(row.get(_DIS_JSON_TID))
        if not tid:
            continue
        d = _parse_cell_date(row.get(_DIS_JSON_DAY))
        if d is None:
            continue
        q = core._parse_optional_float_non_nan(row.get(_DIS_JSON_QTY))
        if q is None:
            continue
        agg[(mk, d, tid)] += float(q)
    return agg


def _qty_from_buckets_for_tid(
    buckets: dict[tuple[str, date], list[tuple[str, float]]],
    mk: str,
    d: date,
    tid: str,
) -> float:
    parts = buckets.get((mk, d))
    if not parts:
        return 0.0
    s = 0.0
    for t, q in parts:
        bt = core.planning_task_id_str_from_scalar(t) or str(t).strip()
        if bt == tid:
            s += float(q)
    return s


def _parse_calendar_window_int_env(name: str, default: int, lo: int, hi: int) -> int:
    try:
        raw = (os.environ.get(name) or "").strip()
        v = int(raw) if raw else default
        return max(lo, min(v, hi))
    except (TypeError, ValueError):
        return default


def _date_bounds_from_actual_df(df_actual: pd.DataFrame | None) -> tuple[date | None, date | None]:
    """Min/max calendar dates from ?????? / ??? (same precedence as _row_actual_day)."""
    if df_actual is None or len(df_actual) == 0:
        return None, None
    s1 = pd.Series([pd.NaT] * len(df_actual))
    if _ACT_COL_START_DT in df_actual.columns:
        s1 = pd.to_datetime(df_actual[_ACT_COL_START_DT], errors="coerce")
    s2 = pd.Series([pd.NaT] * len(df_actual))
    if _ACT_COL_DAY in df_actual.columns:
        s2 = pd.to_datetime(df_actual[_ACT_COL_DAY], errors="coerce")
    eff = s1.where(s1.notna(), s2)
    eff = pd.to_datetime(eff, errors="coerce")
    valid = eff.dropna()
    if len(valid) == 0:
        return None, None
    # tz-naive dates for comparison with date.today()-based windows
    mn = valid.min().date()
    mx = valid.max().date()
    return mn, mx


def _date_bounds_from_plan_date_columns(df_plan: pd.DataFrame | None) -> tuple[date | None, date | None]:
    """Min/max of YYYY/MM/DD-style quantity columns (aligned with _build_compare_gantt_aladdin_qty_lookup)."""
    if df_plan is None or len(df_plan) == 0:
        return None, None
    found: list[date] = []
    for col in df_plan.columns:
        col_key = core._nfkc_column_aliases(col)
        m = core._COMPARE_GANTT_ALADDIN_QTY_COL_RE.match(str(col_key).strip())
        if not m:
            continue
        try:
            y, mo, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
            found.append(date(y, mo, dd))
        except ValueError:
            continue
    if not found:
        return None, None
    return min(found), max(found)


def _collect_sorted_dates(
    df_plan: pd.DataFrame | None,
    df_actual: pd.DataFrame | None,
) -> tuple[list[date], dict[str, Any]]:
    """Calendar columns: rolling window merged with actual/plan data date bounds.

    Formerly fixed at today-14 .. today+30, which dropped actual rows older than 14 days from
    aggregation (eligible_pairs / table rows disappeared even though ???? tab showed them).
    """
    today = date.today()
    past_days = _parse_calendar_window_int_env(_ENV_CAL_PAST_DAYS, 45, 1, 800)
    future_days = _parse_calendar_window_int_env(_ENV_CAL_FUTURE_DAYS, 30, 1, 800)
    display_start = today - timedelta(days=past_days)
    display_end = today + timedelta(days=future_days)

    mn_a, mx_a = _date_bounds_from_actual_df(df_actual)
    mn_p, mx_p = _date_bounds_from_plan_date_columns(df_plan)

    merged_start = min([display_start] + ([mn_a] if mn_a else []) + ([mn_p] if mn_p else []))
    merged_end = max([display_end] + ([mx_a] if mx_a else []) + ([mx_p] if mx_p else []))

    abs_past = today - timedelta(days=800)
    abs_future = today + timedelta(days=800)
    merged_start = max(merged_start, abs_past)
    merged_end = min(merged_end, abs_future)
    if merged_end < merged_start:
        merged_end = merged_start

    out: list[date] = []
    d = merged_start
    while d <= merged_end:
        out.append(d)
        d += timedelta(days=1)

    range_meta = {
        "deliveryCalendarPastDaysDefault": past_days,
        "deliveryCalendarFutureDaysDefault": future_days,
        "deliveryCalendarMergedStart": merged_start.isoformat(),
        "deliveryCalendarMergedEnd": merged_end.isoformat(),
        "deliveryCalendarActualBoundsMin": mn_a.isoformat() if mn_a else "",
        "deliveryCalendarActualBoundsMax": mx_a.isoformat() if mx_a else "",
        "deliveryCalendarPlanBoundsMin": mn_p.isoformat() if mn_p else "",
        "deliveryCalendarPlanBoundsMax": mx_p.isoformat() if mx_p else "",
        "deliveryCalendarRollingStart": display_start.isoformat(),
        "deliveryCalendarRollingEnd": display_end.isoformat(),
        "deliveryCalendarColumnDayCount": len(out),
    }

    # region agent log
    _agent_debug_ndjson(
        "H_CAL_RANGE",
        "delivery_calendar_payload.py:_collect_sorted_dates",
        "calendar_column_range",
        {
            "merged_start": merged_start.isoformat(),
            "merged_end": merged_end.isoformat(),
            "n_days": len(out),
            "mn_a": mn_a.isoformat() if mn_a else None,
            "mx_a": mx_a.isoformat() if mx_a else None,
            "past_days": past_days,
        },
    )
    # endregion
    return out, range_meta


def _machine_display_from_plan_row(row) -> str:
    v = row.get(core.TASK_COL_MACHINE_NAME)
    return str(v).strip() if v is not None and not (isinstance(v, float) and pd.isna(v)) else ""


def _row_actual_day(row) -> date | None:
    """Per Power Query: ???? = DateTime.Date(??????)????? ??? ?? fallback?"""
    d = _parse_cell_date(row.get(_ACT_COL_START_DT))
    if d is not None:
        return d
    return _parse_cell_date(row.get(_ACT_COL_DAY))


def _aggregate_daily_actual_qty_aladdin_max(
    df: pd.DataFrame | None,
    equipment_list,
    sorted_dates: list,
) -> dict[tuple[str, date], dict[str, float]]:
    """
    Power Query ????? = Group({??NO, ???, ????}, max(????))??????? Python ????

    - 1 ? = 1 ?????? (??NO, ???, ???) ???????????????????????
      ``max(????)`` ???? 1 ? 1 ???????????????
    - ``????(??)`` ??????? ``"??"`` ??????????????
    - ????? ``equipment_list`` ?????????canonical??????????
    """
    if df is None or len(df) == 0:
        return {}
    equip_lookup = core._equipment_lookup_normalized_to_canonical(equipment_list)
    date_ok = set(sorted_dates) if sorted_dates else None
    has_filter_col = _ACT_COL_PRODUCTION_DETAIL in df.columns

    grouped: dict[tuple[str, date, str], float] = defaultdict(float)
    for _, row in df.iterrows():
        if has_filter_col:
            cond = row.get(_ACT_COL_PRODUCTION_DETAIL)
            if cond is None or (isinstance(cond, float) and pd.isna(cond)):
                continue
            want = core._nfkc_column_aliases(_ACT_PRODUCTION_DETAIL_LENGTH)
            got = core._nfkc_column_aliases(str(cond)).strip()
            if got != want:
                continue
        tid = core.planning_task_id_str_from_scalar(row.get(_ACT_COL_TID))
        if not tid:
            continue
        # Align machine key with ``_aggregate_actual_qty_for_aladdin_compare_from_detail_df``:
        # ??????????????????????????JSON???????
        mk, _src = _resolve_actual_detail_machine_key_for_delivery_calendar(
            row, df, equip_lookup
        )
        if not mk:
            continue
        d = _row_actual_day(row)
        if d is None:
            continue
        if date_ok is not None and d not in date_ok:
            continue
        try:
            q = core.parse_float_safe(row.get(_ACT_COL_QTY), None)
        except Exception:
            q = None
        if q is None:
            continue
        try:
            qf = float(q)
        except (TypeError, ValueError):
            continue
        if qf <= 1e-12 or math.isnan(qf):
            continue
        key = (mk, d, tid)
        if qf > grouped[key]:
            grouped[key] = qf

    out: dict[tuple[str, date], dict[str, float]] = defaultdict(dict)
    for (mk, d, tid), v in grouped.items():
        if v > 1e-12:
            out[(mk, d)][tid] = v
    return out


def _equipment_sort_index(equipment_list: list, mk_normalized: str) -> int:
    for i, eq in enumerate(equipment_list or []):
        _, mn = core._split_equipment_line_process_machine(str(eq))
        disp = (mn or str(eq)).strip()
        if core._normalize_equipment_match_key(disp) == mk_normalized:
            return i
    return 10_000


def build_delivery_calendar_payload() -> dict[str, Any]:
    meta: dict[str, Any] = {}
    try:
        df_plan = _read_plan_tasks_from_processing_plan_env()
        pp = (os.environ.get(ENV_PROCESSING_PLAN_PATH) or "").strip()
        meta["processingPlanPath"] = pp

        dispatch_path = _resolve_dispatch_json_path(pp)
        meta["dispatchJsonPath"] = dispatch_path or ""
        _disp_header, disp_rows = _load_dispatch_json_rows(dispatch_path)
        dispatch_agg = _aggregate_dispatch_quantities(disp_rows)

        df_actual = core.load_machining_actual_detail_df()
        _tiw = core._excel_plan_input_wb()
        _ad_resolved = resolve_actual_detail_workbook_path(_tiw)
        # Always emit strings so JavaFX meta label can show rows (null omits keys / hasNonNull skips).
        # Env empty -> same defaults as Java AppPaths / resolve_actual_detail_workbook_path (actual detail).
        _tdir_env = (os.environ.get(ENV_TASK_INPUT_SOURCE_DIR) or "").strip()
        meta["pmAiTaskInputSourceDir"] = _tdir_env
        meta["pmAiTaskInputSourceDirEffective"] = (
            _tdir_env if _tdir_env else DEFAULT_TASK_INPUT_SOURCE_DIR
        )
        meta["pmAiTaskInputSourceDirUsesDefaultDir"] = not bool(_tdir_env)

        _adir_env = (os.environ.get(ENV_ACTUAL_DETAIL_SOURCE_DIR) or "").strip()
        meta["pmAiActualDetailSourceDir"] = _adir_env
        meta["pmAiActualDetailSourceDirEffective"] = (
            _adir_env if _adir_env else DEFAULT_ACTUAL_DETAIL_SOURCE_DIR
        )
        meta["pmAiActualDetailSourceDirUsesDefaultDir"] = not bool(_adir_env)
        meta["pmAiActualDetailWorkbook"] = (
            os.environ.get(ENV_ACTUAL_DETAIL_WORKBOOK) or ""
        ).strip()
        meta["actualDetailWorkbookPath"] = (_ad_resolved or "").strip()
        meta["actualDetailRowCount"] = (
            int(len(df_actual)) if df_actual is not None else 0
        )

        sorted_dates, cal_range_meta = _collect_sorted_dates(df_plan, df_actual)
        meta.update(cal_range_meta)

        skills_pack = core.load_skills_and_needs()
        equipment_list = skills_pack[2]
        if not equipment_list:
            return {
                "ok": False,
                "error": "equipment_list empty (check master.xlsm)",
                "meta": meta,
            }

        dates_set = set(sorted_dates)
        _, buckets = core._build_compare_gantt_aladdin_qty_lookup(
            df_plan,
            dates_set,
        )

        actual_agg = _aggregate_daily_actual_qty_aladdin_max(
            df_actual if df_actual is not None else pd.DataFrame(),
            equipment_list,
            sorted_dates,
        )

        plan_pairs: set[tuple[str, str]] = set()
        if df_plan is not None and len(df_plan) > 0:
            for _, row in df_plan.iterrows():
                mk = core._normalize_equipment_match_key(row.get(core.TASK_COL_MACHINE_NAME))
                tid = core.planning_task_id_str_from_scalar(row.get(core.TASK_COL_TASK_ID))
                if mk and tid:
                    plan_pairs.add((mk, tid))

        actual_pairs: set[tuple[str, str]] = set()
        for (mk, _d), tmap in actual_agg.items():
            for tid in tmap:
                if mk and tid:
                    actual_pairs.add((mk, tid))

        eligible_pairs = plan_pairs | actual_pairs

        _probe_lit = (os.environ.get(_DEBUG_PROBE_ENV) or "Y4-59").strip()
        if _probe_lit:
            meta["deliveryCalendarProbe"] = _probe_task_rows_delivery_calendar(
                df_actual if df_actual is not None else None,
                equipment_list,
                sorted_dates,
                _probe_lit,
                actual_agg,
                plan_pairs,
                eligible_pairs,
            )

        pair_plan_row: dict[tuple[str, str], Any] = {}
        if df_plan is not None and len(df_plan) > 0:
            for _, row in df_plan.iterrows():
                mk = core._normalize_equipment_match_key(row.get(core.TASK_COL_MACHINE_NAME))
                tid = core.planning_task_id_str_from_scalar(row.get(core.TASK_COL_TASK_ID))
                if not mk or not tid:
                    continue
                if (mk, tid) not in eligible_pairs:
                    continue
                pair_plan_row[(mk, tid)] = row
        for (mk, tid) in eligible_pairs:
            pair_plan_row.setdefault((mk, tid), None)

        mk_to_display: dict[str, str] = {}
        if df_plan is not None and len(df_plan) > 0:
            for _, row in df_plan.iterrows():
                mk = core._normalize_equipment_match_key(row.get(core.TASK_COL_MACHINE_NAME))
                if mk and mk not in mk_to_display:
                    mk_to_display[mk] = _machine_display_from_plan_row(row)

        # ???????????????????????????????_??????????????????
        left_headers = [
            h
            for h in core.RESULT_DISPATCH_TABLE_STATIC_HEADERS
            if h != "??????"
        ]
        # One column per calendar day: JSON cell {"triple": {p,a,d}} stacked in JavaFX (plan / actual / dispatch).
        cal_cols: list[str] = []
        for d in sorted_dates:
            ds = _format_delivery_calendar_date_header(d) if isinstance(d, date) else str(d)
            cal_cols.append(ds)

        main_columns = left_headers + cal_cols

        ordered_pairs = sorted(
            pair_plan_row.keys(),
            key=lambda kv: (
                _equipment_sort_index(equipment_list, kv[0]),
                kv[0],
                kv[1],
            ),
        )

        # region agent log
        _agent_debug_ndjson(
            "H3",
            "delivery_calendar_payload.py:after_ordered_pairs",
            "source_sizes",
            {
                "plan_rows": int(len(df_plan)) if df_plan is not None else 0,
                "actual_detail_rows": int(len(df_actual)) if df_actual is not None else 0,
                "sorted_dates_n": len(sorted_dates),
                "buckets_day_slots": len(buckets),
                "actual_agg_machine_days": len(actual_agg),
                "dispatch_agg_keys": len(dispatch_agg),
            },
        )
        bk_sample = [f"{a}|{b.isoformat()}" for (a, b) in list(buckets.keys())[:10]]
        ak_sample = [f"{a}|{b.isoformat()}" for (a, b) in list(actual_agg.keys())[:10]]
        _agent_debug_ndjson(
            "H1_H2",
            "delivery_calendar_payload.py:key_samples",
            "bucket_vs_actual_agg_day_keys",
            {"bucket_keys_sample": bk_sample, "actual_agg_keys_sample": ak_sample},
        )
        if ordered_pairs:
            mk0, tid0 = ordered_pairs[0]
            diag = []
            for d in sorted_dates[:5]:
                qi = _qty_from_buckets_for_tid(buckets, mk0, d, tid0)
                qa = float(actual_agg.get((mk0, d), {}).get(tid0, 0.0))
                qd = float(dispatch_agg.get((mk0, d, tid0), 0.0))
                diag.append(
                    {
                        "d": d.isoformat() if isinstance(d, date) else str(d),
                        "q_in": qi,
                        "q_act": qa,
                        "q_disp": qd,
                        "has_bucket_day": (mk0, d) in buckets,
                        "has_actual_day": (mk0, d) in actual_agg,
                    }
                )
            _agent_debug_ndjson(
                "H1_H2_H4",
                "delivery_calendar_payload.py:first_pair_qty_probe",
                "first_ordered_pair",
                {"mk": mk0, "tid": tid0, "by_date": diag},
            )
        # endregion

        main_rows_out: list[dict[str, Any]] = []
        current_mk = ""

        def flush_section(mk_norm: str):
            nonlocal current_mk
            if mk_norm == current_mk:
                return
            current_mk = mk_norm
            label = mk_to_display.get(mk_norm, mk_norm)
            sec_cells = [""] * len(left_headers)
            # ???????????????????????????? left_headers ????
            try:
                mi = left_headers.index(core.TASK_COL_MACHINE_NAME)
            except ValueError:
                mi = -1
            if mi >= 0:
                sec_cells[mi] = label
            elif sec_cells:
                sec_cells[0] = label
            row_cells = sec_cells + [""] * len(cal_cols)
            main_rows_out.append({"kind": "section", "machineLabel": label, "cells": row_cells})

        for mk, tid in ordered_pairs:
            plan_row = pair_plan_row.get((mk, tid))
            flush_section(mk)

            left_cells: list[str] = []
            if plan_row is not None:
                for h in left_headers:
                    left_cells.append(_format_cell(plan_row.get(h)))
            else:
                left_cells = [""] * len(left_headers)
                if core.TASK_COL_MACHINE_NAME in left_headers:
                    idx = left_headers.index(core.TASK_COL_MACHINE_NAME)
                    left_cells[idx] = mk_to_display.get(mk, "")
                if core.TASK_COL_TASK_ID in left_headers:
                    left_cells[left_headers.index(core.TASK_COL_TASK_ID)] = tid

            cal_cells: list[dict[str, Any]] = []
            for d in sorted_dates:
                q_in = _qty_from_buckets_for_tid(buckets, mk, d, tid)
                q_act = float(actual_agg.get((mk, d), {}).get(tid, 0.0))
                q_disp = float(dispatch_agg.get((mk, d, tid), 0.0))
                tp = core._format_qty_short(q_in) if abs(q_in) > 1e-12 else ""
                ta = core._format_qty_short(q_act) if abs(q_act) > 1e-12 else ""
                td = core._format_qty_short(q_disp) if abs(q_disp) > 1e-12 else ""
                cal_cells.append({"triple": {"p": tp, "a": ta, "d": td}})

            main_rows_out.append(
                {
                    "kind": "data",
                    "machineKey": mk,
                    "requestId": tid,
                    "cells": left_cells + cal_cells,
                }
            )

        plan_agg: dict[tuple[str, date, str], float] = defaultdict(float)
        for (mk, dk), parts in buckets.items():
            for t, q in parts:
                btid = core.planning_task_id_str_from_scalar(t) or str(t).strip()
                plan_agg[(mk, dk, btid)] += float(q)

        cmp_keys = sorted(set(dispatch_agg.keys()) | set(plan_agg.keys()))
        compare_columns = [
            "machine_key",
            "machine_display",
            "request_id",
            "calendar_date",
            "qty_dispatch_json",
            "qty_task_input_aladdin",
            "delta",
        ]
        compare_rows_out: list[list[str]] = []

        def disp_for_mk(mk: str) -> str:
            return mk_to_display.get(mk, mk)

        for key in cmp_keys:
            mk, dk, tid = key
            dq = float(dispatch_agg.get(key, 0.0))
            pq = float(plan_agg.get(key, 0.0))
            delta = dq - pq
            if abs(dq) < 1e-12 and abs(pq) < 1e-12:
                continue
            compare_rows_out.append(
                [
                    mk,
                    disp_for_mk(mk),
                    tid,
                    _format_delivery_calendar_date_header(dk)
                    if isinstance(dk, date)
                    else str(dk),
                    core._format_qty_short(dq),
                    core._format_qty_short(pq),
                    core._format_qty_short(delta),
                ]
            )

        # region agent log
        trip = {"triple_cells": 0, "p_nonempty": 0, "a_nonempty": 0, "d_nonempty": 0}
        for row in main_rows_out:
            if row.get("kind") != "data":
                continue
            for cell in row.get("cells") or []:
                if isinstance(cell, dict) and "triple" in cell:
                    tr = cell["triple"]
                    trip["triple_cells"] += 1
                    if str(tr.get("p", "")).strip():
                        trip["p_nonempty"] += 1
                    if str(tr.get("a", "")).strip():
                        trip["a_nonempty"] += 1
                    if str(tr.get("d", "")).strip():
                        trip["d_nonempty"] += 1
        _agent_debug_ndjson(
            "H5",
            "delivery_calendar_payload.py:triple_field_counts",
            "non_empty_counts",
            trip,
        )
        # endregion

        return {
            "ok": True,
            "mainCalendar": {"columns": main_columns, "rows": main_rows_out},
            "planCompareTable": {"columns": compare_columns, "rows": compare_rows_out},
            "meta": meta,
        }
    except Exception as e:
        _LOG.exception("delivery_calendar_payload")
        return {"ok": False, "error": str(e), "meta": meta}
