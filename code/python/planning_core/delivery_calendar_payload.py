# -*- coding: utf-8 -*-
"""Delivery-calendar JSON payload for pm_ai_delivery_calendar_view.py / JavaFX (ASCII source + \\u escapes)."""

from __future__ import annotations

import json
import logging
import math
import os
import time
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


def _agent_debug_log_ebddd7(hyp: str, msg: str, data: dict) -> None:
    """Append one NDJSON line for debug session ebddd7. Tries Windows + WSL paths; never raises."""
    try:
        line = json.dumps(
            {
                "sessionId": "ebddd7",
                "hypothesisId": hyp,
                "location": "delivery_calendar_payload.py",
                "message": msg,
                "data": data,
                "timestamp": int(time.time() * 1000),
            },
            ensure_ascii=False,
        )
        candidates = [
            "C:\\\u5de5\u7a0b\u7ba1\u7406AI\u30d7\u30ed\u30b8\u30a7\u30af\u30c8_JAVA\\.cursor\\debug-ebddd7.log",
            "/mnt/c/\u5de5\u7a0b\u7ba1\u7406AI\u30d7\u30ed\u30b8\u30a7\u30af\u30c8_JAVA/.cursor/debug-ebddd7.log",
        ]
        for p in candidates:
            try:
                with open(p, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
                return
            except OSError:
                continue
    except Exception:
        pass

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

# Short weekday for calendar column titles (Mon=\u6708 ... Sun=\u65e5)
_JP_WEEKDAY_SHORT = ("\u6708", "\u706b", "\u6c34", "\u6728", "\u91d1", "\u571f", "\u65e5")


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


def _collect_sorted_dates(
    df_plan: pd.DataFrame | None,
    df_actual: pd.DataFrame | None,
) -> list:
    aladdin_dates: set[date] = set()
    if df_plan is not None and len(df_plan) > 0:
        for col in df_plan.columns:
            m = core._COMPARE_GANTT_ALADDIN_QTY_COL_RE.match(str(col))
            if m:
                try:
                    y, mo, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
                    aladdin_dates.add(date(y, mo, dd))
                except ValueError:
                    continue
    base_sorted = sorted(aladdin_dates)
    merged = core._sorted_dates_union_actual_bounds_df(base_sorted, df_actual)
    d_from = core._parse_env_optional_date(core.ENV_GANTT_ACTUAL_DETAIL_DATE_FROM)
    d_to = core._parse_env_optional_date(core.ENV_GANTT_ACTUAL_DETAIL_DATE_TO)
    merged = core._sorted_dates_filter_inclusive_range(merged, d_from, d_to)
    if not merged and aladdin_dates:
        merged = sorted(aladdin_dates)
    if not merged:
        t = date.today()
        merged = [t + timedelta(days=i) for i in range(14)]
    return merged


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
            if str(cond).strip() != _ACT_PRODUCTION_DETAIL_LENGTH:
                continue
        tid = core.planning_task_id_str_from_scalar(row.get(_ACT_COL_TID))
        if not tid:
            continue
        proc = row.get(_ACT_COL_PROC)
        if proc is None or (isinstance(proc, float) and pd.isna(proc)):
            continue
        proc_key = core._normalize_equipment_match_key(proc)
        canonical = equip_lookup.get(proc_key)
        if not canonical:
            continue
        _, mn = core._split_equipment_line_process_machine(str(canonical))
        mk = core._normalize_equipment_match_key((mn or canonical).strip())
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

        sorted_dates = _collect_sorted_dates(df_plan, df_actual)

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

        # region agent log
        try:
            buckets_keys = len(buckets)
            buckets_nonempty_qty = 0
            buckets_distinct_tids = set()
            for _k, parts in buckets.items():
                for t, q in parts:
                    if abs(float(q)) > 1e-12:
                        buckets_nonempty_qty += 1
                    tid_norm = core.planning_task_id_str_from_scalar(t) or str(t).strip()
                    buckets_distinct_tids.add(tid_norm)
            actual_agg_keys = len(actual_agg)
            actual_agg_nonzero_pairs = 0
            actual_distinct_tids = set()
            for _k, tmap in actual_agg.items():
                for tid, v in tmap.items():
                    if abs(float(v)) > 1e-12:
                        actual_agg_nonzero_pairs += 1
                        actual_distinct_tids.add(tid)
            dispatch_agg_keys = len(dispatch_agg)
            dispatch_agg_nonzero = 0
            dispatch_distinct_tids = set()
            dispatch_distinct_mks = set()
            for (mk, _d, tid), v in dispatch_agg.items():
                if abs(float(v)) > 1e-12:
                    dispatch_agg_nonzero += 1
                    dispatch_distinct_tids.add(tid)
                    dispatch_distinct_mks.add(mk)
            df_plan_rows = int(len(df_plan)) if df_plan is not None else -1
            df_actual_rows = int(len(df_actual)) if df_actual is not None else -1
            sample_dispatch_keys = []
            for k in list(dispatch_agg.keys())[:5]:
                sample_dispatch_keys.append([k[0], str(k[1]), k[2], float(dispatch_agg[k])])
            sample_actual_keys = []
            for k in list(actual_agg.keys())[:5]:
                first_tid = next(iter(actual_agg[k].keys()), None)
                sample_actual_keys.append(
                    [k[0], str(k[1]), first_tid, float(actual_agg[k].get(first_tid, 0.0)) if first_tid else None]
                )
            sample_buckets_keys = []
            for k in list(buckets.keys())[:5]:
                first = buckets[k][0] if buckets[k] else None
                sample_buckets_keys.append([k[0], str(k[1]), str(first[0]) if first else None, float(first[1]) if first else None])
            _agent_debug_log_ebddd7(
                "H-py-agg",
                "buckets / actual_agg / dispatch_agg sizes and non-zero counts",
                {
                    "df_plan_rows": df_plan_rows,
                    "df_actual_rows": df_actual_rows,
                    "sorted_dates_n": len(sorted_dates),
                    "equipment_list_n": len(equipment_list),
                    "buckets_keys_mk_d": buckets_keys,
                    "buckets_nonempty_qty": buckets_nonempty_qty,
                    "buckets_distinct_tids": len(buckets_distinct_tids),
                    "actual_agg_keys_mk_d": actual_agg_keys,
                    "actual_agg_nonzero_pairs": actual_agg_nonzero_pairs,
                    "actual_distinct_tids": len(actual_distinct_tids),
                    "dispatch_agg_keys_mk_d_tid": dispatch_agg_keys,
                    "dispatch_agg_nonzero": dispatch_agg_nonzero,
                    "dispatch_distinct_tids": len(dispatch_distinct_tids),
                    "dispatch_distinct_mks": len(dispatch_distinct_mks),
                    "sampleDispatchKeysFirst5": sample_dispatch_keys,
                    "sampleActualKeysFirst5": sample_actual_keys,
                    "sampleBucketsKeysFirst5": sample_buckets_keys,
                },
            )
        except Exception:
            pass
        # endregion

        pair_plan_row: dict[tuple[str, str], Any] = {}

        if df_plan is not None and len(df_plan) > 0:
            for _, row in df_plan.iterrows():
                mk = core._normalize_equipment_match_key(row.get(core.TASK_COL_MACHINE_NAME))
                tid = core.planning_task_id_str_from_scalar(row.get(core.TASK_COL_TASK_ID))
                if not mk or not tid:
                    continue
                pair_plan_row[(mk, tid)] = row

        for (mk, _d), parts in buckets.items():
            for t, _q in parts:
                tid = core.planning_task_id_str_from_scalar(t) or str(t).strip()
                if mk and tid:
                    pair_plan_row.setdefault((mk, tid), None)

        for (mk, _d), tmap in actual_agg.items():
            for tid in tmap:
                if mk and tid:
                    pair_plan_row.setdefault((mk, tid), None)

        for mk, _d, tid in dispatch_agg.keys():
            pair_plan_row.setdefault((mk, tid), None)

        mk_to_display: dict[str, str] = {}
        if df_plan is not None and len(df_plan) > 0:
            for _, row in df_plan.iterrows():
                mk = core._normalize_equipment_match_key(row.get(core.TASK_COL_MACHINE_NAME))
                if mk and mk not in mk_to_display:
                    mk_to_display[mk] = _machine_display_from_plan_row(row)

        left_headers = list(core.RESULT_DISPATCH_TABLE_STATIC_HEADERS)
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

        main_rows_out: list[dict[str, Any]] = []
        current_mk = ""
        _dbg_triple_total = 0
        _dbg_triple_p = 0
        _dbg_triple_a = 0
        _dbg_triple_d = 0
        _dbg_triple_any = 0
        _dbg_triple_samples_nonempty: list[dict[str, Any]] = []
        _dbg_triple_samples_empty: list[dict[str, Any]] = []
        _dbg_pair_count = len(ordered_pairs)
        _dbg_pair_nonempty = 0

        def flush_section(mk_norm: str):
            nonlocal current_mk
            if mk_norm == current_mk:
                return
            current_mk = mk_norm
            label = mk_to_display.get(mk_norm, mk_norm)
            sec_cells = [""] * len(left_headers)
            if sec_cells:
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
            _row_has_value = False
            for d in sorted_dates:
                q_in = _qty_from_buckets_for_tid(buckets, mk, d, tid)
                q_act = float(actual_agg.get((mk, d), {}).get(tid, 0.0))
                q_disp = float(dispatch_agg.get((mk, d, tid), 0.0))
                tp = core._format_qty_short(q_in) if abs(q_in) > 1e-12 else ""
                ta = core._format_qty_short(q_act) if abs(q_act) > 1e-12 else ""
                td = core._format_qty_short(q_disp) if abs(q_disp) > 1e-12 else ""
                cal_cells.append({"triple": {"p": tp, "a": ta, "d": td}})
                _dbg_triple_total += 1
                if tp:
                    _dbg_triple_p += 1
                if ta:
                    _dbg_triple_a += 1
                if td:
                    _dbg_triple_d += 1
                if tp or ta or td:
                    _dbg_triple_any += 1
                    _row_has_value = True
                    if len(_dbg_triple_samples_nonempty) < 5:
                        _dbg_triple_samples_nonempty.append(
                            {"mk": mk, "tid": tid, "date": str(d), "p": tp, "a": ta, "d": td,
                             "qIn": q_in, "qAct": q_act, "qDisp": q_disp}
                        )
                else:
                    if len(_dbg_triple_samples_empty) < 5:
                        _dbg_triple_samples_empty.append(
                            {"mk": mk, "tid": tid, "date": str(d), "qIn": q_in, "qAct": q_act, "qDisp": q_disp}
                        )
            if _row_has_value:
                _dbg_pair_nonempty += 1

            main_rows_out.append(
                {
                    "kind": "data",
                    "machineKey": mk,
                    "requestId": tid,
                    "cells": left_cells + cal_cells,
                }
            )

        # region agent log
        try:
            _agent_debug_log_ebddd7(
                "H-py-triple",
                "triple-cell fill summary across (mk,tid)x(date)",
                {
                    "ordered_pairs": _dbg_pair_count,
                    "pairs_with_any_value": _dbg_pair_nonempty,
                    "triple_total_cells": _dbg_triple_total,
                    "triple_nonempty_any": _dbg_triple_any,
                    "triple_nonempty_p": _dbg_triple_p,
                    "triple_nonempty_a": _dbg_triple_a,
                    "triple_nonempty_d": _dbg_triple_d,
                    "samplesNonemptyFirst5": _dbg_triple_samples_nonempty,
                    "samplesEmptyFirst5": _dbg_triple_samples_empty,
                },
            )
        except Exception:
            pass
        # endregion

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

        return {
            "ok": True,
            "mainCalendar": {"columns": main_columns, "rows": main_rows_out},
            "planCompareTable": {"columns": compare_columns, "rows": compare_rows_out},
            "meta": meta,
        }
    except Exception as e:
        _LOG.exception("delivery_calendar_payload")
        return {"ok": False, "error": str(e), "meta": meta}
