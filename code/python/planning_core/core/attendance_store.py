# -*- coding: utf-8 -*-
"""Attendance store: company calendar + member attendance (JSON canonical)."""

from __future__ import annotations

import calendar
import copy
import json
import logging
import os
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any

from planning_core.core.attendance_paths import (
    APP_MASTER_COMPANY_SHEET,
    APP_MASTER_MEMBER_SHEET_PREFIX,
    attendance_data_json_path,
    attendance_view_xlsx_path,
)
from planning_core.core.japanese_holidays import fetch_national_holidays_for_year

logger = logging.getLogger(__name__)

FORMAT_VERSION = 1
DEFAULT_REG_START = time(8, 45)
DEFAULT_REG_END = time(17, 0)
BREAK1_START = time(12, 0)
BREAK1_END = time(12, 50)
BREAK2_START = time(14, 45)
BREAK2_END = time(15, 0)

DAY_KIND_WORKING = "working_day"
DAY_KIND_PUBLIC = "public_holiday"
DAY_KIND_SPECIAL = "special_holiday"

PRESET_WORK = "WORK"
PRESET_OFF_FULL = "OFF_FULL"
PRESET_OFF_AM = "OFF_AM"
PRESET_OFF_PM = "OFF_PM"
PRESET_NO_DISPATCH = "NO_DISPATCH"

HOURLY_AVAILABLE = "available"
HOURLY_BREAK = "break"
HOURLY_LEAVE = "leave"
HOURLY_AWAY = "away"
HOURLY_OFF_SHIFT = "off_shift"


def empty_store(year: int | None = None) -> dict:
    y = year or date.today().year
    return {
        "format_version": FORMAT_VERSION,
        "meta": {
            "schema": "pm-ai-attendance-store",
            "updated_at": None,
            "company_calendar_revision": 0,
            "member_attendance_revision": 0,
            "master_export_at": None,
            "view_excel_generated_at": None,
            "holidays_fetched_at": None,
        },
        "company_calendar": {
            "year": y,
            "factory_hours": {"start": "06:00", "end": "22:00"},
            "regular_hours": {
                "start": DEFAULT_REG_START.strftime("%H:%M"),
                "end": DEFAULT_REG_END.strftime("%H:%M"),
            },
            "days": {},
        },
        "member_attendance": {},
    }


def load_attendance_store(path: Path | None = None) -> dict:
    p = path or attendance_data_json_path()
    if not p.is_file():
        return empty_store()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, dict) and data.get("format_version") == FORMAT_VERSION:
            return data
    except (OSError, json.JSONDecodeError) as e:
        logger.warning("attendance-data.json 読込失敗: %s", e)
    return empty_store()


def save_attendance_store(store: dict, path: Path | None = None) -> Path:
    p = path or attendance_data_json_path()
    meta = store.setdefault("meta", {})
    meta["updated_at"] = datetime.now().isoformat(timespec="seconds")
    p.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(store, ensure_ascii=False, indent=2)
    tmp = p.with_suffix(p.suffix + ".tmp")
    tmp.write_text(payload, encoding="utf-8")
    os.replace(tmp, p)
    return p


def apply_national_holidays_to_company_calendar(
    store: dict,
    year: int,
    overwrite: bool = False,
    include_weekends: bool = False,
    force_online: bool = False,
) -> dict:
    holidays = fetch_national_holidays_for_year(year, force_online=force_online)
    days = store.setdefault("company_calendar", {}).setdefault("days", {})
    applied = 0
    skipped = 0
    for item in holidays:
        d_key = item.get("date")
        name = item.get("name") or "祝日"
        if not d_key:
            continue
        if d_key in days and days[d_key].get("manual_edit") and not overwrite:
            skipped += 1
            continue
        if d_key in days and not overwrite and days[d_key].get("source") != "national_holiday":
            if days[d_key].get("kind") != DAY_KIND_WORKING:
                skipped += 1
                continue
        days[d_key] = {
            "kind": DAY_KIND_PUBLIC,
            "label": name,
            "source": "national_holiday",
        }
        applied += 1
    if include_weekends:
        start = date(year, 1, 1)
        end = date(year, 12, 31)
        d = start
        while d <= end:
            if d.weekday() >= 5:
                key = d.isoformat()
                if key not in days or overwrite:
                    if key in days and not overwrite and days[key].get("manual_edit"):
                        skipped += 1
                    else:
                        days[key] = {
                            "kind": DAY_KIND_PUBLIC,
                            "label": "土曜" if d.weekday() == 5 else "日曜",
                            "source": "weekend_init",
                        }
                        applied += 1
            d += timedelta(days=1)
    store["company_calendar"]["year"] = year
    store["meta"]["holidays_fetched_at"] = datetime.now().isoformat(timespec="seconds")
    store["meta"]["company_calendar_revision"] = int(
        store["meta"].get("company_calendar_revision") or 0
    ) + 1
    return {"applied": applied, "skipped": skipped, "source": "national_holiday"}


def company_day_kind(store: dict, day: date) -> str:
    key = day.isoformat()
    entry = store.get("company_calendar", {}).get("days", {}).get(key)
    if not entry:
        if day.weekday() >= 5:
            return DAY_KIND_PUBLIC
        return DAY_KIND_WORKING
    return entry.get("kind") or DAY_KIND_WORKING


def day_preset_from_company(store: dict, day: date) -> str:
    kind = company_day_kind(store, day)
    if kind == DAY_KIND_WORKING:
        return PRESET_WORK
    return PRESET_OFF_FULL


def preset_to_leave_and_times(
    preset: str,
    day: date,
    reg_start: time = DEFAULT_REG_START,
    reg_end: time = DEFAULT_REG_END,
    company_kind: str | None = None,
) -> dict:
    """Map day preset to attendance row fields (VBA CollectAttendanceFromCalendar 対称)."""
    row: dict[str, Any] = {
        "day_preset": preset,
        "clock_in": None,
        "clock_out": None,
        "leave_type": "",
        "remark": "",
        "breaks": [],
        "overtime_minutes": 0,
        "efficiency": 1.0,
        "eligible_for_assignment": True,
        "hourly": {},
        "manual_edit": False,
    }
    if preset == PRESET_WORK:
        row.update(
            {
                "clock_in": reg_start.strftime("%H:%M"),
                "clock_out": reg_end.strftime("%H:%M"),
                "leave_type": "通常",
                "remark": "通常",
                "breaks": [
                    {"start": BREAK1_START.strftime("%H:%M"), "end": BREAK1_END.strftime("%H:%M")},
                    {"start": BREAK2_START.strftime("%H:%M"), "end": BREAK2_END.strftime("%H:%M")},
                ],
            }
        )
    elif preset == PRESET_OFF_FULL:
        if company_kind in (DAY_KIND_PUBLIC, DAY_KIND_SPECIAL):
            row["leave_type"] = "公休"
            row["remark"] = "公休"
        else:
            row["leave_type"] = "年休"
            row["remark"] = "休"
        row["eligible_for_assignment"] = False
    elif preset == PRESET_OFF_AM:
        row.update(
            {
                "clock_in": BREAK1_END.strftime("%H:%M"),
                "clock_out": reg_end.strftime("%H:%M"),
                "leave_type": "前休",
                "remark": "前休",
                "breaks": [
                    {
                        "start": BREAK2_START.strftime("%H:%M"),
                        "end": BREAK2_END.strftime("%H:%M"),
                    }
                ],
            }
        )
    elif preset == PRESET_OFF_PM:
        row.update(
            {
                "clock_in": reg_start.strftime("%H:%M"),
                "clock_out": BREAK1_START.strftime("%H:%M"),
                "leave_type": "後休",
                "remark": "後休",
            }
        )
    elif preset == PRESET_NO_DISPATCH:
        row["leave_type"] = "-"
        row["remark"] = "-"
        row["eligible_for_assignment"] = False
    return row


def apply_company_calendar_to_members(
    store: dict,
    members: list[str],
    year: int,
    month: int,
    only_unedited: bool = True,
) -> dict:
    applied = 0
    skipped = 0
    ma = store.setdefault("member_attendance", {})
    ym_days = calendar.monthrange(year, month)[1]
    for day_num in range(1, ym_days + 1):
        d = date(year, month, day_num)
        d_key = d.isoformat()
        preset = day_preset_from_company(store, d)
        kind = company_day_kind(store, d)
        day_bucket = ma.setdefault(d_key, {})
        for member in members:
            existing = day_bucket.get(member)
            if only_unedited and existing and existing.get("manual_edit"):
                skipped += 1
                continue
            day_bucket[member] = preset_to_leave_and_times(preset, d, company_kind=kind)
            applied += 1
    store["meta"]["member_attendance_revision"] = int(
        store["meta"].get("member_attendance_revision") or 0
    ) + 1
    return {"applied": applied, "skipped": skipped}


def build_editor_payload(store: dict, members: list[str], year: int, month: int) -> dict:
    ym_days = calendar.monthrange(year, month)[1]
    dates: list[str] = []
    for day_num in range(1, ym_days + 1):
        dates.append(date(year, month, day_num).isoformat())
    cells: dict[str, dict[str, dict]] = {}
    ma = store.get("member_attendance", {})
    cc_days = store.get("company_calendar", {}).get("days", {})
    for d_key in dates:
        cells[d_key] = {}
        d = date.fromisoformat(d_key)
        cc = cc_days.get(d_key, {})
        for m in members:
            entry = ma.get(d_key, {}).get(m) or preset_to_leave_and_times(
                day_preset_from_company(store, d), d
            )
            cells[d_key][m] = {
                "day_preset": entry.get("day_preset", PRESET_WORK),
                "leave_type": entry.get("leave_type", ""),
                "is_working": bool(entry.get("clock_in")),
                "overtime_minutes": int(entry.get("overtime_minutes") or 0),
                "manual_edit": bool(entry.get("manual_edit")),
                "company_kind": cc.get("kind") or company_day_kind(store, d),
            }
    return {
        "format_version": 1,
        "ok": True,
        "year": year,
        "month": month,
        "members": members,
        "dates": dates,
        "cells": cells,
        "company_calendar_revision": store.get("meta", {}).get("company_calendar_revision", 0),
        "member_attendance_revision": store.get("meta", {}).get("member_attendance_revision", 0),
    }


def build_company_calendar_payload(store: dict, year: int) -> dict:
    days = store.get("company_calendar", {}).get("days", {})
    year_days = {k: v for k, v in sorted(days.items()) if k.startswith(f"{year}-")}
    return {
        "format_version": 1,
        "ok": True,
        "year": year,
        "days": year_days,
        "revision": store.get("meta", {}).get("company_calendar_revision", 0),
    }


def member_attendance_to_dataframe_records(store: dict, members: list[str]) -> list[dict]:
    """Convert store to flat records for load_attendance_and_analyze compatibility."""
    records: list[dict] = []
    ma = store.get("member_attendance", {})
    for d_key, per_member in sorted(ma.items()):
        for member, entry in per_member.items():
            if member not in members:
                continue
            rec = {
                "日付": d_key,
                "メンバー": member,
                "出勤時間": entry.get("clock_in") or "",
                "退勤時間": entry.get("clock_out") or "",
                "休憩時間1_開始": "",
                "休憩時間1_終了": "",
                "休憩時間2_開始": "",
                "休憩時間2_終了": "",
                "作業効率": entry.get("efficiency", 1.0),
                "休暇区分": entry.get("leave_type") or "",
                "備考": entry.get("remark") or "",
                "残業(分)": entry.get("overtime_minutes") or 0,
            }
            breaks = entry.get("breaks") or []
            if len(breaks) >= 1:
                rec["休憩時間1_開始"] = breaks[0].get("start") or ""
                rec["休憩時間1_終了"] = breaks[0].get("end") or ""
            if len(breaks) >= 2:
                rec["休憩時間2_開始"] = breaks[1].get("start") or ""
                rec["休憩時間2_終了"] = breaks[1].get("end") or ""
            records.append(rec)
    return records


def export_attendance_to_master_new_sheets(
    store: dict,
    master_path: str | Path,
    months: list[tuple[int, int]] | None = None,
) -> dict:
    from openpyxl import load_workbook

    path = Path(master_path)
    if not path.is_file():
        raise FileNotFoundError(f"master が見つかりません: {path}")
    keep_vba = str(path).lower().endswith(".xlsm")
    try:
        wb = load_workbook(path, keep_vba=keep_vba)
    except Exception:
        if keep_vba:
            wb = load_workbook(path, keep_vba=False)
        else:
            raise
    _write_app_company_sheet(wb, store)
    if months is None:
        months = _months_in_store(store)
    updated: list[str] = []
    for year, month in months:
        sheet_name = f"{APP_MASTER_MEMBER_SHEET_PREFIX}{year}年{month}月"
        _write_app_member_calendar_sheet(wb, store, year, month, sheet_name)
        updated.append(sheet_name)
    try:
        wb.save(path)
    except OSError as e:
        raise RuntimeError(f"master.xlsm への保存に失敗（Excel で開いている可能性）: {e}") from e
    store["meta"]["master_export_at"] = datetime.now().isoformat(timespec="seconds")
    store["meta"]["master_export_sheets"] = [APP_MASTER_COMPANY_SHEET] + updated
    return {
        "ok": True,
        "sheets_updated": store["meta"]["master_export_sheets"],
        "master_export_at": store["meta"]["master_export_at"],
    }


def _months_in_store(store: dict) -> list[tuple[int, int]]:
    months: set[tuple[int, int]] = set()
    for d_key in store.get("member_attendance", {}):
        try:
            d = date.fromisoformat(d_key)
            months.add((d.year, d.month))
        except ValueError:
            continue
    cc_days = store.get("company_calendar", {}).get("days", {})
    for d_key in cc_days:
        try:
            d = date.fromisoformat(d_key)
            months.add((d.year, d.month))
        except ValueError:
            continue
    if not months:
        today = date.today()
        months.add((today.year, today.month))
    return sorted(months)


def _write_app_company_sheet(wb, store: dict) -> None:
    name = APP_MASTER_COMPANY_SHEET
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.cell(1, 1, "対象年")
    ws.cell(1, 2, store.get("company_calendar", {}).get("year"))
    ws.cell(2, 1, "日付")
    ws.cell(2, 2, "区分")
    ws.cell(2, 3, "名称")
    ws.cell(2, 4, "ソース")
    ws.cell(3, 1, "※本シートはアプリ自動出力。編集はアプリで行うこと。")
    row = 4
    days = store.get("company_calendar", {}).get("days", {})
    for d_key in sorted(days.keys()):
        entry = days[d_key]
        ws.cell(row, 1, d_key)
        ws.cell(row, 2, entry.get("kind"))
        ws.cell(row, 3, entry.get("label") or "")
        ws.cell(row, 4, entry.get("source") or "")
        row += 1


def _write_app_member_calendar_sheet(
    wb, store: dict, year: int, month: int, sheet_name: str
) -> None:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.cell(1, 1, f"{year}年{month}月")
    ws.cell(2, 1, "※本シートはアプリ自動出力。編集はアプリで行うこと。")
    ym_days = calendar.monthrange(year, month)[1]
    col = 2
    date_cols: list[date] = []
    for day_num in range(1, ym_days + 1):
        d = date(year, month, day_num)
        date_cols.append(d)
        ws.cell(4, col, d)
        col += 1
    try:
        from planning_core.core.master_data import load_skills_and_needs

        members = load_skills_and_needs()[1]
    except Exception:
        members = sorted(
            {
                m
                for d_key in store.get("member_attendance", {})
                for m in store["member_attendance"].get(d_key, {})
            }
        )
    row = 5
    ma = store.get("member_attendance", {})
    for member in members:
        ws.cell(row, 1, member)
        for idx, d in enumerate(date_cols):
            d_key = d.isoformat()
            entry = ma.get(d_key, {}).get(member)
            if not entry:
                preset = day_preset_from_company(store, d)
                kind = company_day_kind(store, d)
                entry = preset_to_leave_and_times(preset, d, company_kind=kind)
            symbol = _cell_symbol(entry)
            ws.cell(row, 2 + idx, symbol)
        row += 1


def _cell_symbol(entry: dict) -> str:
    lt = entry.get("leave_type") or ""
    if lt == "年休" or lt == "公休" or entry.get("day_preset") == PRESET_OFF_FULL:
        return "*"
    if lt == "前休":
        return "前休"
    if lt == "後休":
        return "後休"
    if lt == "-":
        return "-"
    hourly = entry.get("hourly") or {}
    if any(v == HOURLY_AWAY for v in hourly.values()):
        return "離"
    return ""


def generate_attendance_view_xlsx(store: dict, path: Path | None = None) -> Path:
    from openpyxl import Workbook

    out = path or attendance_view_xlsx_path()
    wb = Workbook()
    ws0 = wb.active
    ws0.title = "会社カレンダー"
    ws0.cell(1, 1, "祝日データ出典: 内閣府『国民の祝日について』/ api.jp-calendar.com")
    ws0.cell(2, 1, "日付")
    ws0.cell(2, 2, "区分")
    ws0.cell(2, 3, "名称")
    r = 3
    for d_key, entry in sorted(store.get("company_calendar", {}).get("days", {}).items()):
        ws0.cell(r, 1, d_key)
        ws0.cell(r, 2, entry.get("kind"))
        ws0.cell(r, 3, entry.get("label") or "")
        r += 1
    ws1 = wb.create_sheet("メンバー勤怠")
    ws1.cell(1, 1, "閲覧専用 — 編集はアプリで行ってください")
    ws1.cell(2, 1, "日付")
    ws1.cell(2, 2, "メンバー")
    ws1.cell(2, 3, "出勤")
    ws1.cell(2, 4, "退勤")
    ws1.cell(2, 5, "休暇区分")
    ws1.cell(2, 6, "備考")
    ws1.cell(2, 7, "残業(分)")
    r = 3
    for rec in member_attendance_to_dataframe_records(
        store, _all_members_in_store(store)
    ):
        ws1.cell(r, 1, rec["日付"])
        ws1.cell(r, 2, rec["メンバー"])
        ws1.cell(r, 3, rec["出勤時間"])
        ws1.cell(r, 4, rec["退勤時間"])
        ws1.cell(r, 5, rec["休暇区分"])
        ws1.cell(r, 6, rec["備考"])
        ws1.cell(r, 7, rec["残業(分)"])
        r += 1
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    store["meta"]["view_excel_generated_at"] = datetime.now().isoformat(timespec="seconds")
    return out


def _all_members_in_store(store: dict) -> list[str]:
    members: set[str] = set()
    for per in store.get("member_attendance", {}).values():
        members.update(per.keys())
    return sorted(members)


def migrate_master_attendance_to_store(store: dict, master_path: str) -> dict:
    """One-time import from legacy member sheets into store."""
    from planning_core.core.master_data import load_attendance_and_analyze, load_skills_and_needs

    members = load_skills_and_needs()[1]
    attendance_data, _ = load_attendance_and_analyze(members)
    applied = 0
    ma = store.setdefault("member_attendance", {})
    for d, per_member in attendance_data.items():
        d_key = d.isoformat()
        bucket = ma.setdefault(d_key, {})
        for member, st in per_member.items():
            bucket[member] = {
                "day_preset": PRESET_WORK if st.get("is_working") else PRESET_OFF_FULL,
                "clock_in": st.get("start_dt").strftime("%H:%M") if st.get("start_dt") else None,
                "clock_out": st.get("end_dt").strftime("%H:%M") if st.get("end_dt") else None,
                "leave_type": st.get("reason") or "",
                "remark": st.get("reason") or "",
                "overtime_minutes": int(st.get("overtime_minutes") or 0),
                "efficiency": float(st.get("efficiency") or 1.0),
                "breaks": [],
                "hourly": {},
                "manual_edit": False,
            }
            applied += 1
    store["meta"]["member_attendance_revision"] = int(
        store["meta"].get("member_attendance_revision") or 0
    ) + 1
    return {"imported_cells": applied}
