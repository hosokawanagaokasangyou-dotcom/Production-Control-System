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


def _machine_calendar_readiness() -> tuple[bool, str, bool]:
    """machine-calendar-data.json の存在・整備状態。"""
    from planning_core.core.machine_calendar_paths import machine_calendar_data_json_path
    from planning_core.core.machine_calendar_store import (
        load_machine_calendar_store,
        validate_store_for_dispatch,
    )

    jp = machine_calendar_data_json_path()
    exists = jp.is_file()
    if not exists:
        return False, str(jp), False
    try:
        store = load_machine_calendar_store(jp)
        ready = validate_store_for_dispatch(store)
    except Exception:
        ready = False
    return ready, str(jp), exists


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


_skills_load_for_readiness = False


def _attendance_json_ready_for_legacy_skip(store: dict, year: int, month: int) -> bool:
    """JSON 正本のみで段階1/2 用メンバー勤怠が足りるか（skills 読込なし）。"""
    if not store:
        return False
    meta = store.get("meta", {})
    if int(meta.get("company_calendar_revision") or 0) <= 0:
        return False
    from planning_core.core.attendance_member_roster import attendance_grid_member_names

    roster = attendance_grid_member_names(store)
    if not roster:
        return False
    expected = len(roster) * calendar.monthrange(year, month)[1]
    cells = _member_cells_for_month(store, roster, year, month)
    if expected <= 0 or cells < expected:
        return False
    mc_ready, _, mc_exists = _machine_calendar_readiness()
    return bool(mc_ready and mc_exists)


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
        skills_members: list[str] = []
        global _skills_load_for_readiness
        if not _skills_load_for_readiness:
            try:
                _skills_load_for_readiness = True
                from planning_core.core.master_data import load_skills_and_needs

                skills_members = list(load_skills_and_needs()[1])
            except Exception:
                skills_members = []
            finally:
                _skills_load_for_readiness = False
        from planning_core.core.attendance_member_roster import members_for_attendance_analysis

        members = members_for_attendance_analysis(skills_members, store if store else {})

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
    company_calendar_revision = int(meta.get("company_calendar_revision") or 0)
    company_calendar_ready = json_exists and company_calendar_revision > 0
    member_attendance_ready = (
        json_exists
        and members
        and expected_cells > 0
        and member_cells_month >= expected_cells
    )
    machine_calendar_ready, machine_calendar_json_path, machine_calendar_json_exists = (
        _machine_calendar_readiness()
    )

    issues: list[str] = []
    if not json_exists:
        issues.append(
            "attendance-data.json が未作成です。会社カレンダータブでセットアップしてください。"
            "（段階1/2 は実行不可。master.xlsm のレガシー勤怠シートへはフォールバックしません。）"
        )
    if json_exists and company_calendar_revision <= 0:
        issues.append(
            "会社カレンダーが未セットアップです。"
            " 会社カレンダータブで年度初期化・祝日取得または編集を行ってください。"
        )
    if not machine_calendar_json_exists:
        issues.append(
            "machine-calendar-data.json が未作成です。"
            " 機械カレンダータブで「初期値を作る」または保存してください。"
        )
    elif not machine_calendar_ready:
        issues.append(
            "machine-calendar-data.json が未整備です（列またはスロットが空）。"
            " 機械カレンダータブで初期値を作成してください。"
        )
    if json_exists and not members:
        issues.append(
            "メンバー名簿が未設定です。メンバー勤怠タブまたは attendance-data.json の member_roster を確認してください。"
        )
    if json_exists and members and member_cells_month == 0:
        issues.append(
            f"{y}年{m}月のメンバー勤怠が未登録です。「会社カレンダーに合わせる」またはグリッド編集してください。"
        )
    if json_exists and members and member_cells_month > 0 and member_cells_month < expected_cells:
        issues.append(
            f"{y}年{m}月のメンバー勤怠が一部のみ（{member_cells_month}/{expected_cells} セル）。"
        )

    planning_stages_ready = (
        json_exists
        and company_calendar_ready
        and member_attendance_ready
        and machine_calendar_ready
        and not issues
    )
    if not planning_stages_ready and not issues:
        issues.append(
            "カレンダー正本 JSON（attendance-data.json / machine-calendar-data.json）が未準備です。"
            " 会社カレンダー・メンバー勤怠・機械カレンダータブでセットアップしてください。"
        )
    stage2_ready = planning_stages_ready
    stage1_ready = planning_stages_ready

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
        "member_attendance_revision": int(meta.get("member_attendance_revision") or 0),
        "member_cells_in_month": member_cells_month,
        "member_cells_expected_in_month": expected_cells,
        "skills_member_count": len(members),
        "company_calendar_ready": company_calendar_ready,
        "company_calendar_revision": company_calendar_revision,
        "member_attendance_ready": member_attendance_ready,
        "machine_calendar_json_path": machine_calendar_json_path,
        "machine_calendar_json_exists": machine_calendar_json_exists,
        "machine_calendar_ready": machine_calendar_ready,
        "master_export_at": meta.get("master_export_at"),
        "calendar_xlsx_export_at": meta.get("calendar_xlsx_export_at"),
        "holidays_fetched_at": meta.get("holidays_fetched_at"),
        "stage1_ready": stage1_ready,
        "planning_stages_ready": planning_stages_ready,
        "stage2_ready": stage2_ready,
        "issues": issues,
        "needs_setup": not json_exists,
        "needs_member_sync": json_exists and member_cells_month == 0,
    }


def legacy_master_attendance_sheets_required() -> bool:
    """
    master.xlsm 上のメンバー別勤怠シート検証がまだ必要か。

    attendance-data.json 正本でメンバー勤怠・機械カレンダーが整備済みなら不要
    （段階1/2 の勤怠読込は JSON のみ。レガシーシートへはフォールバックしない）。

    load_skills_and_needs 内から呼ばれるため build_attendance_readiness は使わない
    （skills 再読込の無限再帰を避ける）。
    """
    jp = attendance_data_json_path()
    if not jp.is_file():
        return True
    try:
        store = load_attendance_store(jp)
    except Exception:
        return True
    today = date.today()
    if _attendance_json_ready_for_legacy_skip(store, today.year, today.month):
        return False
    return True


def require_calendar_json_for_planning_stages(
    context_label: str = "段階1/2",
) -> None:
    """機械カレンダー・会社カレンダー・メンバー勤怠の JSON 正本が揃うまで段階1/2 を止める。"""
    from planning_core.bootstrap import PlanningValidationError

    readiness = build_attendance_readiness()
    if readiness.get("planning_stages_ready"):
        return
    issues = readiness.get("issues") or []
    if not issues:
        issues = [
            "カレンダー正本 JSON（attendance-data.json / machine-calendar-data.json）が未準備です。"
        ]
    ctx = (context_label or "段階1/2").strip()
    raise PlanningValidationError(ctx + ": " + "\n".join(issues))
