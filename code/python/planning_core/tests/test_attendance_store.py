# -*- coding: utf-8 -*-
"""Tests for attendance_store."""

from __future__ import annotations

from datetime import date

from planning_core.core.attendance_store import (
    DAY_KIND_PUBLIC,
    DAY_KIND_WORKING,
    PRESET_OFF_AM,
    PRESET_OFF_FULL,
    PRESET_WORK,
    apply_company_calendar_to_members,
    apply_member_attendance_patch,
    company_day_kind,
    day_preset_from_company,
    empty_store,
    preset_to_leave_and_times,
)


def test_preset_work_matches_vba_breaks():
    row = preset_to_leave_and_times(PRESET_WORK, date(2026, 8, 6))
    assert row["clock_in"] == "08:45"
    assert row["clock_out"] == "17:00"
    assert len(row["breaks"]) == 2


def test_company_public_holiday_preset():
    store = empty_store(2026)
    store["company_calendar"]["days"]["2026-08-06"] = {
        "kind": DAY_KIND_PUBLIC,
        "label": "休",
    }
    assert day_preset_from_company(store, date(2026, 8, 6)) == PRESET_OFF_FULL
    assert company_day_kind(store, date(2026, 8, 7)) == DAY_KIND_WORKING


def test_apply_member_attendance_patch():
    store = empty_store(2026)
    patch = {
        "year": 2026,
        "month": 8,
        "cells": {
            "2026-08-10": {"A": {"day_preset": PRESET_OFF_AM}},
        },
    }
    result = apply_member_attendance_patch(store, patch)
    assert result["applied"] == 1
    entry = store["member_attendance"]["2026-08-10"]["A"]
    assert entry["day_preset"] == PRESET_OFF_AM
    assert entry["manual_edit"] is True


def test_apply_company_calendar_to_members():
    store = empty_store(2026)
    store["company_calendar"]["days"]["2026-08-06"] = {"kind": DAY_KIND_PUBLIC}
    members = ["A", "B"]
    result = apply_company_calendar_to_members(store, members, 2026, 8)
    assert result["applied"] == 2 * 31
    entry = store["member_attendance"]["2026-08-06"]["A"]
    assert entry["day_preset"] == PRESET_OFF_FULL
    assert entry["leave_type"] == "公休"


def test_export_master_preserves_legacy_calendar_sheets(tmp_path):
    from openpyxl import load_workbook, Workbook

    from planning_core.core.attendance_paths import APP_MASTER_COMPANY_SHEET
    from planning_core.core.attendance_store import export_attendance_to_master_new_sheets

    master = tmp_path / "master.xlsm"
    wb = Workbook()
    wb.active.title = "勤怠カレンダー_2026年5月"
    wb.create_sheet("会社カレンダー")
    wb.create_sheet("skills")
    wb.save(master)

    store = empty_store(2026)
    store["company_calendar"]["days"]["2026-05-01"] = {
        "kind": DAY_KIND_PUBLIC,
        "label": "テスト",
    }
    store["member_attendance"]["2026-05-01"] = {
        "A": {
            "day_preset": PRESET_OFF_FULL,
            "leave_type": "公休",
            "manual_edit": True,
        }
    }

    result = export_attendance_to_master_new_sheets(store, master)
    assert result["ok"]
    assert APP_MASTER_COMPANY_SHEET in result["sheets_updated"]

    wb2 = load_workbook(master)
    assert "勤怠カレンダー_2026年5月" in wb2.sheetnames
    assert "会社カレンダー" in wb2.sheetnames
    assert APP_MASTER_COMPANY_SHEET in wb2.sheetnames
    assert any(n.startswith("APP_勤怠カレンダー_") for n in wb2.sheetnames)


def test_build_dense_company_calendar_days_covers_fiscal_year():
    from planning_core.core.attendance_store import (
        build_dense_company_calendar_days_for_export,
        fiscal_year_date_range,
    )

    store = empty_store(2026)
    store["meta"]["fiscal_start_month"] = 4
    store["meta"]["fiscal_start_day"] = 1
    store["company_calendar"]["days"]["2026-05-06"] = {
        "kind": DAY_KIND_PUBLIC,
        "label": "振替",
        "manual_edit": True,
    }
    dense = build_dense_company_calendar_days_for_export(store)
    start, end = fiscal_year_date_range(2026, 4, 1)
    expected_days = (end - start).days + 1
    assert len(dense) == expected_days
    assert dense["2026-05-06"]["label"] == "振替"
    assert dense["2026-05-03"]["kind"] == DAY_KIND_PUBLIC  # 憲法記念日 enrich


def test_write_app_company_sheet_grid_layout(tmp_path):
    from openpyxl import load_workbook, Workbook

    from planning_core.core.attendance_paths import APP_MASTER_COMPANY_SHEET
    from planning_core.core.attendance_store import export_attendance_to_master_new_sheets

    master = tmp_path / "master.xlsm"
    wb = Workbook()
    wb.save(master)

    store = empty_store(2026)
    store["meta"]["fiscal_start_month"] = 4
    store["meta"]["fiscal_start_day"] = 1
    store["company_calendar"]["days"]["2026-05-06"] = {
        "kind": DAY_KIND_PUBLIC,
        "label": "公休",
    }
    export_attendance_to_master_new_sheets(store, master)

    wb2 = load_workbook(master)
    ws = wb2[APP_MASTER_COMPANY_SHEET]
    assert "2026年度" in str(ws.cell(1, 1).value)
    # 5月グリッドに平日出勤セル（数字のみ）と祝日セル（例: 3祝）がある
    found_day = False
    found_national = False
    for row in ws.iter_rows(min_row=5, max_row=35, min_col=1, max_col=31):
        for cell in row:
            v = cell.value
            if v is None:
                continue
            s = str(v)
            if s.isdigit():
                found_day = True
            if s.endswith("祝"):
                found_national = True
    assert found_day
    assert found_national
