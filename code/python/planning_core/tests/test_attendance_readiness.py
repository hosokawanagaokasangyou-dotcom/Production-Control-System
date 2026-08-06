# -*- coding: utf-8 -*-
"""Tests for attendance_readiness."""

from __future__ import annotations

from planning_core.core.attendance_readiness import build_attendance_readiness
from planning_core.core.attendance_store import empty_store


def test_readiness_not_ready_without_json():
    result = build_attendance_readiness(store={}, members=["A"], year=2026, month=8)
    assert result["ok"] is True
    assert result["stage2_ready"] is False
    assert result["issues"]


def test_readiness_ready_with_members_synced():
    store = empty_store(2026)
    store["company_calendar"]["days"]["2026-08-06"] = {"kind": "public", "label": "休"}
    members = ["A"]
    from planning_core.core.attendance_store import apply_company_calendar_to_members

    apply_company_calendar_to_members(store, members, 2026, 8)
    result = build_attendance_readiness(store=store, members=members, year=2026, month=8)
    assert result["member_cells_in_month"] == 31
    assert result["member_cells_expected_in_month"] == 31
    assert not any("未登録" in issue for issue in result["issues"])
