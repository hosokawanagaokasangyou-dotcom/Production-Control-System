# -*- coding: utf-8 -*-
"""Attendance canonical store readiness for stage2 / UI badges."""

from __future__ import annotations

import calendar
import os
from datetime import date
from pathlib import Path
from typing import Any

from planning_core.core.attendance_paths import attendance_data_json_path
from planning_core.core.attendance_store import load_attendance_store, fiscal_year_date_range


def _member_cells_for_month(store: dict, members: list[str], year: int, month: int) -> int:
    ym_days = calendar.monthrange(year, month)[1]
    ma = store.get("member_attendance", {})
    count = 0
    for day_num in range(1, ym_days + 1):
        d_key = date(year, month, day_num).isoformat()
        bucket = ma.get(d_key, {})
        for m in members:
            if m in bucket:
                count += 1
    return count


def build_attendance_readiness(
    store: dict | None = None,
    members: list[str] | None = None,
    year: int | None = None,
    month: int | None = None,
) -> dict[str, Any]:
    today = date.today()
    y = year or today.year
    m = month or today.month
    jp = attendance_data_json_path()
    json_exists = jp.is_file()

    master_path: Path | None = None
    master_exists = False
    calendar_xlsx_path: Path | None = None
    calendar_xlsx_exists = False
    try:
        from planning_core.core.master_data import _master_workbook_path_resolved
        from planning_core.core.attendance_paths import attendance_calendar_xlsx_path

        master_path = Path(_master_workbook_path_resolved())
        master_exists = master_path.is_file()
        calendar_xlsx_path = attendance_calendar_xlsx_path()
        calendar_xlsx_exists = calendar_xlsx_path.is_file()
    except Exception:
        alt = (os.environ.get("PM_AI_MASTER_WORKBOOK") or "").strip()
        if alt:
            master_path = Path(alt)
            master_exists = master_path.is_file()
        try:
            from planning_core.core.attendance_paths import attendance_calendar_xlsx_path

            calendar_xlsx_path = attendance_calendar_xlsx_path()
            calendar_xlsx_exists = calendar_xlsx_path.is_file()
        except Exception:
            pass

    if store is None:
        store = load_attendance_store(jp) if json_exists else {}
    if members is None:
        try:
            from planning_core.core.master_data import load_skills_and_needs

            members = list(load_skills_and_needs()[1])
        except Exception:
            members = []

    meta = store.get("meta", {}) if store else {}
    fiscal_start_month = int(meta.get("fiscal_start_month") or 4)
    fiscal_start_day = int(meta.get("fiscal_start_day") or 1)
    fiscal_label = y
    if m < fiscal_start_month or (m == fiscal_start_month and today.day < fiscal_start_day):
        fiscal_label = y - 1
    try:
        fy_start, fy_end = fiscal_year_date_range(fiscal_label, fiscal_start_month, fiscal_start_day)
    except ValueError:
        fy_start, fy_end = fiscal_year_date_range(fiscal_label, 4, 1)
    cc_days = 0
    if store:
        for k in store.get("company_calendar", {}).get("days", {}):
            try:
                d_val = date.fromisoformat(k)
            except ValueError:
                continue
            if fy_start <= d_val <= fy_end:
                cc_days += 1
    member_cells_month = _member_cells_for_month(store, members, y, m) if store and members else 0
    expected_cells = len(members) * calendar.monthrange(y, m)[1] if members else 0

    issues: list[str] = []
    if not json_exists:
        issues.append(
            "attendance-data.json が未作成です。会社カレンダータブでセットアップしてください。"
            "（段階2は実行不可。master.xlsm のレガシー勤怠シートへはフォールバックしません。）"
        )
    if json_exists and members and member_cells_month == 0:
        issues.append(
            f"{y}年{m}月のメンバー勤怠が未登録です。「会社カレンダーに合わせる」またはグリッド編集してください。"
        )
    if json_exists and members and member_cells_month < expected_cells:
        issues.append(
            f"{y}年{m}月のメンバー勤怠が一部のみ（{member_cells_month}/{expected_cells} セル）。"
        )

    stage2_ready = json_exists and member_cells_month > 0 and not issues

    return {
        "ok": True,
        "format_version": 1,
        "year": y,
        "month": m,
        "json_path": str(jp),
        "json_exists": json_exists,
        "master_workbook_path": str(master_path) if master_path else "",
        "master_workbook_exists": master_exists,
        "calendar_xlsx_path": str(calendar_xlsx_path) if calendar_xlsx_path else "",
        "calendar_xlsx_exists": calendar_xlsx_exists,
        "meta": meta,
        "company_calendar_day_count": cc_days,
        "company_calendar_revision": int(meta.get("company_calendar_revision") or 0),
        "member_attendance_revision": int(meta.get("member_attendance_revision") or 0),
        "member_cells_in_month": member_cells_month,
        "member_cells_expected_in_month": expected_cells,
        "skills_member_count": len(members),
        "master_export_at": meta.get("master_export_at"),
        "calendar_xlsx_export_at": meta.get("calendar_xlsx_export_at"),
        "holidays_fetched_at": meta.get("holidays_fetched_at"),
        "stage2_ready": stage2_ready,
        "issues": issues,
        "needs_setup": not json_exists,
        "needs_member_sync": json_exists and member_cells_month == 0,
    }
