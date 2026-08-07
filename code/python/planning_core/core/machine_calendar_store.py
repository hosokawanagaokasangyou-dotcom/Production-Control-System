# -*- coding: utf-8 -*-
"""Machine calendar canonical store (JSON). master.xlsm「機械カレンダー」は互換用。"""

from __future__ import annotations

import json
import logging
import os
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

from planning_core.core.columns import SHEET_MACHINE_CALENDAR
from planning_core.core.gemini_auth import MACHINE_CALENDAR_SLOT_MINUTES
from planning_core.core.machine_calendar_paths import machine_calendar_data_json_path
from planning_core.core.stage1 import DEFAULT_END_TIME, DEFAULT_START_TIME

logger = logging.getLogger(__name__)

FORMAT_VERSION = 1


def empty_store() -> dict:
    return {
        "format_version": FORMAT_VERSION,
        "meta": {
            "schema": "pm-ai-machine-calendar-store",
            "updated_at": None,
            "revision": 0,
            "slot_minutes": MACHINE_CALENDAR_SLOT_MINUTES,
            "imported_from_master_at": None,
            "master_source_path": None,
        },
        "columns": [],
        "defined_slots": {},
        "occupancy": {},
    }


def load_machine_calendar_store(path: Path | None = None) -> dict:
    p = path or machine_calendar_data_json_path()
    if not p.is_file():
        return empty_store()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, dict) and data.get("format_version") == FORMAT_VERSION:
            return data
    except (OSError, json.JSONDecodeError) as e:
        logger.warning("machine-calendar-data.json 読込失敗: %s", e)
    return empty_store()


def save_machine_calendar_store(
    store: dict,
    path: Path | None = None,
    *,
    history_kind: str = "auto_save",
    history_label: str = "保存",
) -> Path:
    p = path or machine_calendar_data_json_path()
    meta = store.setdefault("meta", {})
    meta["updated_at"] = datetime.now().isoformat(timespec="seconds")
    p.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(store, ensure_ascii=False, indent=2)
    tmp = p.with_suffix(p.suffix + ".tmp")
    tmp.write_text(payload, encoding="utf-8")
    os.replace(tmp, p)
    try:
        from planning_core.core.machine_calendar_history_store import (
            append_machine_calendar_snapshot,
        )

        append_machine_calendar_snapshot(p, kind=history_kind, label=history_label)
    except Exception as e:
        logger.warning("機械カレンダー JSON 世代退避失敗: %s", e)
    return p


def store_has_machine_calendar_data(store: dict) -> bool:
    cols = store.get("columns")
    occ = store.get("occupancy")
    if isinstance(cols, list) and cols:
        return True
    if isinstance(occ, dict) and occ:
        return True
    defined = store.get("defined_slots")
    if isinstance(defined, dict) and defined:
        return True
    return False


def validate_store_for_dispatch(store: dict) -> bool:
    if not store_has_machine_calendar_data(store):
        return False
    cols = store.get("columns") or []
    defined = store.get("defined_slots") or {}
    slot_rows = sum(len(v) for v in defined.values() if isinstance(v, list))
    if slot_rows == 0:
        occ = store.get("occupancy") or {}
        slot_rows = len(occ)
    header_pairs = len(cols)
    return slot_rows > 0 and header_pairs > 0


def _roll_helpers():
    from planning_core.core.roll_pipeline import (
        _clip_machine_calendar_slot_to_factory_window,
        _equipment_lookup_normalized_to_canonical,
        _equipment_line_key_to_physical_occupancy_key,
        _machine_cal_cell_is_asterisk_occupancy_only,
        _machine_cal_cell_is_occupied,
        _machine_cal_parse_slot_datetime,
        _machine_cal_resolve_column_to_equipment_key,
        _merge_machine_calendar_intervals,
    )

    return {
        "clip": _clip_machine_calendar_slot_to_factory_window,
        "eq_lookup": _equipment_lookup_normalized_to_canonical,
        "phys_key": _equipment_line_key_to_physical_occupancy_key,
        "is_asterisk": _machine_cal_cell_is_asterisk_occupancy_only,
        "is_occupied": _machine_cal_cell_is_occupied,
        "parse_slot": _machine_cal_parse_slot_datetime,
        "resolve_col": _machine_cal_resolve_column_to_equipment_key,
        "merge": _merge_machine_calendar_intervals,
    }


def _parse_columns_from_raw(raw: pd.DataFrame, equipment_list: list[str]) -> list[dict[str, str]]:
    h = _roll_helpers()
    eq_lookup = h["eq_lookup"](equipment_list)
    elist_set = set(str(x).strip() for x in equipment_list if str(x).strip())
    ncols = raw.shape[1]
    non_empty_pm = 0
    for c in range(2, ncols):
        p = raw.iat[0, c]
        m = raw.iat[1, c]
        if pd.isna(p) or pd.isna(m):
            continue
        p_s = str(p).strip()
        m_s = str(m).strip()
        if p_s and m_s and p_s.lower() != "nan" and m_s.lower() != "nan":
            non_empty_pm += 1
    use_two_header = non_empty_pm > 0
    columns: list[dict[str, str]] = []
    seen: set[str] = set()
    for c in range(2, ncols):
        p = raw.iat[0, c]
        m = raw.iat[1, c] if use_two_header else None
        if use_two_header:
            if pd.isna(p) or pd.isna(m):
                continue
            p_s = str(p).strip()
            m_s = str(m).strip()
            if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
                continue
        else:
            if pd.isna(p):
                continue
            p_s = str(p).strip()
            if not p_s or p_s.lower() == "nan":
                continue
            m_s = ""
        canon = h["resolve_col"](p_s, m_s, eq_lookup, elist_set)
        if not canon or canon in seen:
            continue
        seen.add(canon)
        columns.append(
            {"equipment_key": canon, "process": p_s, "machine": m_s}
        )
    return columns


def import_from_master_workbook(
    store: dict,
    master_path: str,
    equipment_list: list[str],
) -> dict:
    """master.xlsm「機械カレンダー」シートを JSON 正本へ取り込む。"""
    from planning_core.core.master_data import _cached_master_pd_excel_file

    h = _roll_helpers()
    path = str(master_path or "").strip()
    if not path or not os.path.isfile(path):
        raise FileNotFoundError(f"マスタブックが見つかりません: {path}")
    xls = _cached_master_pd_excel_file(path)
    if xls is None or SHEET_MACHINE_CALENDAR not in xls.sheet_names:
        raise ValueError(f"シート「{SHEET_MACHINE_CALENDAR}」がありません")
    raw = pd.read_excel(xls, sheet_name=SHEET_MACHINE_CALENDAR, header=None)
    if raw.shape[0] < 3 or raw.shape[1] < 3:
        raise ValueError("機械カレンダーシートが空または未構成です")

    columns = _parse_columns_from_raw(raw, equipment_list)
    col_keys = [c["equipment_key"] for c in columns]
    col_index = {c: i for i, c in enumerate(col_keys)}

    occupancy: dict[str, dict[str, str]] = {}
    defined_slots: dict[str, list[str]] = defaultdict(list)

    for r in range(2, raw.shape[0]):
        slot0 = h["parse_slot"](raw.iat[r, 0])
        if slot0 is None:
            continue
        slot_key = slot0.replace(microsecond=0).isoformat()
        day_key = slot0.date().isoformat()
        slot_end_row = slot0 + timedelta(minutes=MACHINE_CALENDAR_SLOT_MINUTES)
        clipped_row = h["clip"](slot0.date(), slot0, slot_end_row)
        if clipped_row:
            defined_slots[day_key].append(slot_key)
        row_cells: dict[str, str] = {}
        for c in range(2, raw.shape[1]):
            if c >= raw.shape[1]:
                continue
            cell = raw.iat[r, c]
            if not h["is_occupied"](cell):
                continue
            p = raw.iat[0, c]
            m = raw.iat[1, c] if raw.shape[0] > 1 else None
            if pd.isna(p):
                continue
            p_s = str(p).strip()
            m_s = str(m).strip() if m is not None and not pd.isna(m) else ""
            eq_lookup = h["eq_lookup"](equipment_list)
            elist_set = set(str(x).strip() for x in equipment_list if str(x).strip())
            canon = h["resolve_col"](p_s, m_s, eq_lookup, elist_set)
            if not canon or canon not in col_index:
                continue
            row_cells[canon] = str(cell).strip() if isinstance(cell, str) else str(cell)
        if row_cells:
            occupancy[slot_key] = row_cells

    store["columns"] = columns
    store["occupancy"] = occupancy
    store["defined_slots"] = {k: sorted(v) for k, v in defined_slots.items()}
    meta = store.setdefault("meta", {})
    meta["imported_from_master_at"] = datetime.now().isoformat(timespec="seconds")
    meta["master_source_path"] = os.path.abspath(path)
    meta["revision"] = int(meta.get("revision") or 0) + 1
    return {
        "columns": len(columns),
        "occupancy_slots": len(occupancy),
        "defined_days": len(defined_slots),
    }


def occupancy_blocks_from_store(
    store: dict,
    equipment_list: list[str],
    *,
    interactive_only_asterisk_occupancy: bool = False,
) -> tuple[
    dict[date, dict[str, list[tuple[datetime, datetime]]]],
    dict[date, list[tuple[datetime, datetime]]],
]:
    """JSON 正本から配台用占有ブロックを構築する。"""
    h = _roll_helpers()
    columns = store.get("columns") or []
    col_keys = [
        str(c.get("equipment_key") or "").strip()
        for c in columns
        if isinstance(c, dict) and str(c.get("equipment_key") or "").strip()
    ]
    occupancy = store.get("occupancy") or {}
    defined_slots_raw = store.get("defined_slots") or {}

    defined_slot_windows_by_day: dict[date, list[tuple[datetime, datetime]]] = defaultdict(list)
    for day_key, slot_keys in defined_slots_raw.items():
        try:
            day_d = date.fromisoformat(str(day_key))
        except ValueError:
            continue
        if not isinstance(slot_keys, list):
            continue
        for sk in slot_keys:
            slot0 = h["parse_slot"](sk)
            if slot0 is None:
                continue
            slot_end_row = slot0 + timedelta(minutes=MACHINE_CALENDAR_SLOT_MINUTES)
            clipped = h["clip"](day_d, slot0, slot_end_row)
            if clipped:
                defined_slot_windows_by_day[day_d].append(clipped)

    acc: dict[date, dict[str, list[tuple[datetime, datetime]]]] = defaultdict(
        lambda: defaultdict(list)
    )
    for slot_key, per_eq in occupancy.items():
        if not isinstance(per_eq, dict):
            continue
        slot0 = h["parse_slot"](slot_key)
        if slot0 is None:
            continue
        try:
            day_d = slot0.date()
        except Exception:
            continue
        for eq_key, cell_val in per_eq.items():
            eq_s = str(eq_key).strip()
            if not eq_s or eq_s not in col_keys:
                continue
            if interactive_only_asterisk_occupancy:
                if not h["is_asterisk"](cell_val):
                    continue
            elif not h["is_occupied"](cell_val):
                continue
            slot_start = slot0
            slot_end = slot_start + timedelta(minutes=MACHINE_CALENDAR_SLOT_MINUTES)
            clipped_mc = h["clip"](day_d, slot_start, slot_end)
            if clipped_mc is None:
                continue
            slot_start, slot_end = clipped_mc
            acc[day_d][eq_s].append((slot_start, slot_end))

    out: dict[date, dict[str, list[tuple[datetime, datetime]]]] = {}
    phys = h["phys_key"]
    merge = h["merge"]
    for d, eqmap in acc.items():
        merged_eq = {eq: merge(iv) for eq, iv in eqmap.items() if iv}
        phys_accum: dict[str, list] = defaultdict(list)
        for eq, iv in merged_eq.items():
            pk = phys(str(eq).strip())
            if pk:
                phys_accum[pk].extend(iv)
        merged_all = dict(merged_eq)
        for pk, iv in phys_accum.items():
            merged_all[pk] = merge(iv)
        out[d] = merged_all

    interactive_defined: dict[date, list[tuple[datetime, datetime]]] = {}
    if interactive_only_asterisk_occupancy:
        interactive_defined = {
            d: merge(vs) for d, vs in defined_slot_windows_by_day.items()
        }
    return out, interactive_defined


def _factory_slot_keys_for_day(day: date) -> list[str]:
    """工場稼働枠内の 30 分スロットキー（JSON 未作成・当日未定義時の編集グリッド用）。"""
    h = _roll_helpers()
    w0 = datetime.combine(day, DEFAULT_START_TIME)
    w1 = datetime.combine(day, DEFAULT_END_TIME)
    slot_keys: list[str] = []
    t = w0
    while t < w1:
        slot_end = t + timedelta(minutes=MACHINE_CALENDAR_SLOT_MINUTES)
        if h["clip"](day, t, slot_end):
            slot_keys.append(t.replace(microsecond=0).isoformat())
        t = slot_end
    return slot_keys


def build_editor_payload(
    store: dict,
    day: date,
    equipment_list: list[str],
) -> dict[str, Any]:
    """1日分の編集用グリッド（列=設備、行=スロット）。"""
    h = _roll_helpers()
    columns = store.get("columns") or []
    if not columns:
        columns = [{"equipment_key": eq, "process": "", "machine": eq} for eq in equipment_list]
    day_key = day.isoformat()
    defined = store.get("defined_slots") or {}
    slot_keys = list(defined.get(day_key) or [])
    occupancy = store.get("occupancy") or {}
    if not slot_keys:
        for sk in sorted(occupancy.keys()):
            slot0 = h["parse_slot"](sk)
            if slot0 is not None and slot0.date() == day:
                slot_keys.append(sk)
    slot_keys = sorted(set(slot_keys))
    if not slot_keys:
        slot_keys = _factory_slot_keys_for_day(day)
    rows: list[dict[str, Any]] = []
    for sk in slot_keys:
        slot0 = h["parse_slot"](sk)
        if slot0 is None:
            continue
        cells = occupancy.get(sk) or {}
        row_cells: dict[str, str] = {}
        for col in columns:
            if not isinstance(col, dict):
                continue
            ek = str(col.get("equipment_key") or "").strip()
            if not ek:
                continue
            val = cells.get(ek)
            if val is not None and str(val).strip():
                row_cells[ek] = str(val).strip()
        rows.append({"slot": sk, "cells": row_cells})
    return {
        "format_version": 1,
        "ok": True,
        "date": day_key,
        "columns": columns,
        "rows": rows,
        "revision": store.get("meta", {}).get("revision", 0),
        "slot_minutes": int(
            store.get("meta", {}).get("slot_minutes") or MACHINE_CALENDAR_SLOT_MINUTES
        ),
    }


def apply_machine_calendar_patch(store: dict, patch: dict) -> dict:
    """UI からのセル編集をマージ（空文字は占有削除）。"""
    occupancy = store.setdefault("occupancy", {})
    defined_slots = store.setdefault("defined_slots", {})
    applied = 0
    day_key = str(patch.get("date") or "").strip()
    rows = patch.get("rows") or []
    if not isinstance(rows, list):
        rows = []
    day_slot_keys: list[str] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        sk = str(row.get("slot") or "").strip()
        if not sk:
            continue
        day_slot_keys.append(sk)
        cells = row.get("cells") or {}
        if not isinstance(cells, dict):
            continue
        bucket = occupancy.setdefault(sk, {})
        for ek, val in cells.items():
            key = str(ek).strip()
            if not key:
                continue
            s = str(val or "").strip()
            if s:
                bucket[key] = s
                applied += 1
            elif key in bucket:
                del bucket[key]
                applied += 1
        if not bucket:
            occupancy.pop(sk, None)
    if day_key and day_slot_keys:
        defined_slots[day_key] = sorted(set(day_slot_keys))
    meta = store.setdefault("meta", {})
    meta["revision"] = int(meta.get("revision") or 0) + 1
    return {"applied": applied}
