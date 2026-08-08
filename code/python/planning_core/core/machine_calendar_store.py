# -*- coding: utf-8 -*-
"""Machine calendar canonical store (JSON)."""

from __future__ import annotations

import json
import logging
import os
from collections import defaultdict
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any

from planning_core.core.gemini_auth import MACHINE_CALENDAR_SLOT_MINUTES
from planning_core.core.machine_calendar_paths import machine_calendar_data_json_path

logger = logging.getLogger(__name__)

FORMAT_VERSION = 1
_WEEKDAY_JA = ("月", "火", "水", "木", "金", "土", "日")

# 工場マスタ既定（機械カレンダー UI・初期値生成）
MACHINE_CAL_FACTORY_START = time(8, 0)
MACHINE_CAL_FACTORY_END = time(19, 0)
MACHINE_CAL_REGULAR_START = time(8, 25)
MACHINE_CAL_REGULAR_END = time(17, 0)


def empty_store() -> dict:
    return {
        "format_version": FORMAT_VERSION,
        "meta": {
            "schema": "pm-ai-machine-calendar-store",
            "updated_at": None,
            "revision": 0,
            "slot_minutes": MACHINE_CALENDAR_SLOT_MINUTES,
            "factory_start": MACHINE_CAL_FACTORY_START.strftime("%H:%M"),
            "factory_end": MACHINE_CAL_FACTORY_END.strftime("%H:%M"),
            "regular_start": MACHINE_CAL_REGULAR_START.strftime("%H:%M"),
            "regular_end": MACHINE_CAL_REGULAR_END.strftime("%H:%M"),
        },
        "columns": [],
        "defined_slots": {},
        "occupancy": {},
        "cell_comments": {},
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


def require_machine_calendar_json_for_dispatch(
    context_label: str = "配台",
) -> Path:
    """
    配台用 machine-calendar-data.json 正本の存在・整備を検証する。
    不備時は PlanningValidationError（master.xlsm フォールバックは行わない）。
    """
    from planning_core.bootstrap import PlanningValidationError

    ctx = (context_label or "配台").strip()
    jp = machine_calendar_data_json_path()
    if not jp.is_file():
        raise PlanningValidationError(
            f"{ctx}: machine-calendar-data.json が存在しません。"
            f" パス: {jp} 。"
            " 機械カレンダータブで「初期値を作る」または編集後「保存」してください。"
            " master.xlsm の機械カレンダーシートは使用しません。"
        )
    try:
        store = load_machine_calendar_store(jp)
    except Exception as e:
        raise PlanningValidationError(
            f"{ctx}: machine-calendar-data.json の読込に失敗しました ({e})。"
            f" パス: {jp}"
        ) from e
    if not validate_store_for_dispatch(store):
        raise PlanningValidationError(
            f"{ctx}: machine-calendar-data.json が未整備です（列またはスロットが空）。"
            f" パス: {jp} 。"
            " 機械カレンダータブで初期値作成・保存してください。"
        )
    return jp


def _parse_hhmm(value: str | None, default: time) -> time:
    s = str(value or "").strip()
    if not s:
        return default
    parts = s.split(":")
    if len(parts) < 2:
        return default
    try:
        return time(int(parts[0]), int(parts[1]))
    except (TypeError, ValueError):
        return default


def factory_window_times(store: dict) -> tuple[time, time]:
    meta = store.get("meta") or {}
    start = _parse_hhmm(meta.get("factory_start"), MACHINE_CAL_FACTORY_START)
    end = _parse_hhmm(meta.get("factory_end"), MACHINE_CAL_FACTORY_END)
    return start, end


def _roll_helpers():
    from planning_core.core.roll_pipeline import (
        _clip_machine_calendar_slot_to_factory_window,
        _equipment_line_key_to_physical_occupancy_key,
        _machine_cal_cell_is_asterisk_occupancy_only,
        _machine_cal_cell_is_occupied,
        _machine_cal_parse_slot_datetime,
        _merge_machine_calendar_intervals,
    )

    return {
        "clip": _clip_machine_calendar_slot_to_factory_window,
        "phys_key": _equipment_line_key_to_physical_occupancy_key,
        "is_asterisk": _machine_cal_cell_is_asterisk_occupancy_only,
        "is_occupied": _machine_cal_cell_is_occupied,
        "parse_slot": _machine_cal_parse_slot_datetime,
        "merge": _merge_machine_calendar_intervals,
    }


def slot_keys_for_factory_window(
    day: date,
    factory_start: time,
    factory_end: time,
) -> list[str]:
    """工場稼働枠（例 8:00〜19:00）の 30 分スロットキー一覧。"""
    slot_keys: list[str] = []
    t = datetime.combine(day, factory_start)
    end = datetime.combine(day, factory_end)
    while t < end:
        slot_keys.append(t.replace(microsecond=0).isoformat())
        t += timedelta(minutes=MACHINE_CALENDAR_SLOT_MINUTES)
    return slot_keys


def initialize_machine_calendar_defaults(
    store: dict,
    fiscal_year: int,
    need_columns: list[dict[str, str]],
    *,
    start_month: int = 4,
    start_day: int = 1,
) -> dict:
    """
    会計年度の各日を工場稼働枠（時刻範囲）で初期化する。
    土日・祭日にかかわらず占有は設定しない（空＝稼働可能）。
    配台では人の勤怠ブロックが先に効くため、機械カレンダーは会社カレンダーと連動しない。
    """
    from planning_core.core.attendance_store import fiscal_year_date_range

    if not need_columns:
        raise ValueError("need シートに機械列がありません")
    start, end = fiscal_year_date_range(fiscal_year, start_month, start_day)
    factory_start, factory_end = factory_window_times(store)
    eq_keys = [
        str(c.get("equipment_key") or "").strip()
        for c in need_columns
        if isinstance(c, dict) and str(c.get("equipment_key") or "").strip()
    ]
    if not eq_keys:
        raise ValueError("need シートの機械列キーが空です")

    store["columns"] = list(need_columns)
    occupancy = store.setdefault("occupancy", {})
    defined_slots = store.setdefault("defined_slots", {})
    meta = store.setdefault("meta", {})
    meta.setdefault("factory_start", MACHINE_CAL_FACTORY_START.strftime("%H:%M"))
    meta.setdefault("factory_end", MACHINE_CAL_FACTORY_END.strftime("%H:%M"))
    meta.setdefault("regular_start", MACHINE_CAL_REGULAR_START.strftime("%H:%M"))
    meta.setdefault("regular_end", MACHINE_CAL_REGULAR_END.strftime("%H:%M"))

    initialized_days = 0
    d = start
    while d <= end:
        day_key = d.isoformat()
        slot_keys = slot_keys_for_factory_window(d, factory_start, factory_end)
        defined_slots[day_key] = slot_keys
        for sk in slot_keys:
            occupancy.pop(sk, None)
            store.setdefault("cell_comments", {}).pop(sk, None)
        initialized_days += 1
        d += timedelta(days=1)

    meta["revision"] = int(meta.get("revision") or 0) + 1
    meta["initialized_defaults_at"] = datetime.now().isoformat(timespec="seconds")

    valid_slots: set[str] = set()
    d = start
    while d <= end:
        day_key = d.isoformat()
        valid_slots.update(defined_slots.get(day_key) or [])
        d += timedelta(days=1)
    for sk in list(occupancy.keys()):
        if sk not in valid_slots:
            del occupancy[sk]
    cell_comments = store.setdefault("cell_comments", {})
    for sk in list(cell_comments.keys()):
        if sk not in valid_slots:
            del cell_comments[sk]
    for day_key in list(defined_slots.keys()):
        try:
            dd = date.fromisoformat(day_key)
            if dd < start or dd > end:
                del defined_slots[day_key]
        except ValueError:
            del defined_slots[day_key]

    return {
        "fiscal_start": start.isoformat(),
        "fiscal_end": end.isoformat(),
        "columns": len(need_columns),
        "initialized_days": initialized_days,
    }


def initialize_machine_calendar_from_company_calendar(
    store: dict,
    attendance_store: dict,
    fiscal_year: int,
    need_columns: list[dict[str, str]],
    *,
    start_month: int = 4,
    start_day: int = 1,
) -> dict:
    """後方互換名。attendance_store は無視（会社カレンダー連動は廃止）。"""
    _ = attendance_store
    return initialize_machine_calendar_defaults(
        store,
        fiscal_year,
        need_columns,
        start_month=start_month,
        start_day=start_day,
    )


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


def collect_machine_calendar_export_rows(
    store: dict,
    start: date,
    end: date,
) -> tuple[list[dict[str, str]], list[dict[str, Any]]]:
    """会計年度範囲のフラット行（Excel APP_機械カレンダー用）。"""
    columns = list(store.get("columns") or [])
    if not columns:
        try:
            from planning_core.core.master_data import load_need_machine_columns

            columns = load_need_machine_columns()
        except Exception as e:
            logger.warning("need シートから機械カレンダー列の取得に失敗: %s", e)
            columns = []
    if not columns:
        return [], []
    eq_keys = [
        str(c.get("equipment_key") or "").strip()
        for c in columns
        if isinstance(c, dict) and str(c.get("equipment_key") or "").strip()
    ]
    occupancy = store.get("occupancy") or {}
    defined = store.get("defined_slots") or {}
    factory_start, factory_end = factory_window_times(store)
    rows: list[dict[str, Any]] = []
    d = start
    while d <= end:
        day_key = d.isoformat()
        slot_keys = list(defined.get(day_key) or [])
        if not slot_keys:
            slot_keys = slot_keys_for_factory_window(d, factory_start, factory_end)
        wd = _WEEKDAY_JA[d.weekday()]
        for sk in sorted(slot_keys):
            cells_map = occupancy.get(sk) or {}
            row_cells = {ek: str(cells_map.get(ek) or "").strip() for ek in eq_keys}
            rows.append({"slot": sk, "weekday": wd, "cells": row_cells})
        d += timedelta(days=1)
    return columns, rows


def _resolve_editor_columns(
    store: dict,
    need_columns: list[dict[str, str]],
) -> list[dict[str, str]]:
    """UI 列は need シートを正本とし、store 列とのドリフトを避ける。"""
    if need_columns:
        return list(need_columns)
    stored = store.get("columns") or []
    return list(stored) if isinstance(stored, list) else []


def build_editor_payload(
    store: dict,
    day: date,
    need_columns: list[dict[str, str]],
) -> dict[str, Any]:
    """1日分の編集用グリッド（列=need 機械、行=スロット）。"""
    h = _roll_helpers()
    columns = _resolve_editor_columns(store, need_columns)
    day_key = day.isoformat()
    defined = store.get("defined_slots") or {}
    slot_keys = list(defined.get(day_key) or [])
    occupancy = store.get("occupancy") or {}
    cell_comments = store.get("cell_comments") or {}
    if not slot_keys:
        for sk in sorted(occupancy.keys()):
            slot0 = h["parse_slot"](sk)
            if slot0 is not None and slot0.date() == day:
                slot_keys.append(sk)
    slot_keys = sorted(set(slot_keys))
    if not slot_keys:
        factory_start, factory_end = factory_window_times(store)
        slot_keys = slot_keys_for_factory_window(day, factory_start, factory_end)
    rows: list[dict[str, Any]] = []
    for sk in slot_keys:
        slot0 = h["parse_slot"](sk)
        if slot0 is None:
            continue
        cells = occupancy.get(sk) or {}
        comment_cells = cell_comments.get(sk) or {}
        row_cells: dict[str, str] = {}
        row_comments: dict[str, str] = {}
        for col in columns:
            if not isinstance(col, dict):
                continue
            ek = str(col.get("equipment_key") or "").strip()
            if not ek:
                continue
            val = cells.get(ek)
            if val is not None and str(val).strip():
                row_cells[ek] = str(val).strip()
            cmt = comment_cells.get(ek)
            if cmt is not None and str(cmt).strip():
                row_comments[ek] = str(cmt).strip()
        row: dict[str, Any] = {"slot": sk, "cells": row_cells}
        if row_comments:
            row["comments"] = row_comments
        rows.append(row)
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
    """UI からのセル編集をマージ（占有空文字は削除、コメント空文字は削除）。"""
    occupancy = store.setdefault("occupancy", {})
    cell_comments = store.setdefault("cell_comments", {})
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
        if isinstance(cells, dict):
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
        comments = row.get("comments")
        if isinstance(comments, dict):
            comment_bucket = cell_comments.setdefault(sk, {})
            for ek, val in comments.items():
                key = str(ek).strip()
                if not key:
                    continue
                s = str(val or "").strip()
                if s:
                    comment_bucket[key] = s
                    applied += 1
                elif key in comment_bucket:
                    del comment_bucket[key]
                    applied += 1
            if not comment_bucket:
                cell_comments.pop(sk, None)
    if day_key and day_slot_keys:
        defined_slots[day_key] = sorted(set(day_slot_keys))
    patch_columns = patch.get("columns")
    if isinstance(patch_columns, list) and patch_columns:
        store["columns"] = list(patch_columns)
    meta = store.setdefault("meta", {})
    meta["revision"] = int(meta.get("revision") or 0) + 1
    return {"applied": applied}
