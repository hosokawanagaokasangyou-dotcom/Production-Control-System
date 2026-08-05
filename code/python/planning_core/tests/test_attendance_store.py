# -*- coding: utf-8 -*-
"""Tests for attendance_store."""

from __future__ import annotations

from datetime import date

from planning_core.core.attendance_store import (
    DAY_KIND_PUBLIC,
    DAY_KIND_WORKING,
    PRESET_OFF_FULL,
    PRESET_WORK,
    apply_company_calendar_to_members,
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


def test_apply_company_calendar_to_members():
    store = empty_store(2026)
    store["company_calendar"]["days"]["2026-08-06"] = {"kind": DAY_KIND_PUBLIC}
    members = ["A", "B"]
    result = apply_company_calendar_to_members(store, members, 2026, 8)
    assert result["applied"] == 2 * 31
    entry = store["member_attendance"]["2026-08-06"]["A"]
    assert entry["day_preset"] == PRESET_OFF_FULL
    assert entry["leave_type"] == "公休"
