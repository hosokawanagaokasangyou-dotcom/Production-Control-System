# -*- coding: utf-8 -*-
"""Tests for attendance_readiness."""

from __future__ import annotations

from planning_core.core.attendance_paths import ENV_ATTENDANCE_JSON
from planning_core.core.attendance_readiness import build_attendance_readiness
from planning_core.core.attendance_store import empty_store
from planning_core.core.machine_calendar_paths import ENV_MACHINE_CALENDAR_JSON
from planning_core.core.machine_calendar_store import apply_machine_calendar_patch

NEED_COLS = [{"equipment_key": "EC+EC機", "process": "EC", "machine": "EC機"}]


def test_readiness_not_ready_without_json():
    result = build_attendance_readiness(store={}, members=["A"], year=2026, month=8)
    assert result["ok"] is True
    assert result["stage2_ready"] is False
    assert result["issues"]


def test_readiness_ready_with_members_synced(tmp_path, monkeypatch):
    store = empty_store(2026)
    store["meta"]["company_calendar_revision"] = 1
    store["company_calendar"]["days"]["2026-08-06"] = {"kind": "public", "label": "休"}
    members = ["A"]
    from planning_core.core.attendance_store import apply_company_calendar_to_members

    apply_company_calendar_to_members(store, members, 2026, 8)
    att = tmp_path / "attendance-data.json"
    monkeypatch.setenv(ENV_ATTENDANCE_JSON, str(att))
    from planning_core.core.attendance_store import save_attendance_store

    save_attendance_store(store, att)
    mc = tmp_path / "machine-calendar-data.json"
    from planning_core.core.machine_calendar_store import initialize_machine_calendar_defaults, save_machine_calendar_store
    from planning_core.core.machine_calendar_store import empty_store as mc_empty

    mc_store = mc_empty()
    initialize_machine_calendar_defaults(mc_store, 2026, NEED_COLS, start_month=4, start_day=1)
    apply_machine_calendar_patch(
        mc_store,
        {"date": "2026-08-01", "rows": [{"slot": "2026-08-01T09:00:00", "cells": {"EC+EC機": "*"}}]},
    )
    save_machine_calendar_store(mc_store, mc)
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(mc))
    result = build_attendance_readiness(store=store, members=members, year=2026, month=8)
    assert result["member_cells_in_month"] == 31
    assert result["member_cells_expected_in_month"] == 31
    assert result["planning_stages_ready"]
    assert not any("未登録" in issue for issue in result["issues"])


def test_readiness_not_empty_after_company_calendar_initialize(tmp_path, monkeypatch):
    """明示日次エントリが無くても JSON 正本があれば既定の平日／週末解釈が有効。"""
    from planning_core.core.attendance_paths import ENV_ATTENDANCE_JSON
    from planning_core.core.attendance_store import initialize_company_calendar, save_attendance_store

    json_path = tmp_path / "attendance-data.json"
    monkeypatch.setenv(ENV_ATTENDANCE_JSON, str(json_path))

    store = empty_store(2026)
    store["company_calendar"]["days"]["2026-08-06"] = {
        "kind": "public",
        "label": "手動",
        "manual_edit": True,
    }
    initialize_company_calendar(store, 2026, start_month=4, start_day=1)
    save_attendance_store(store, history_kind="test", history_label="init")

    result = build_attendance_readiness(members=["A"], year=2026, month=8)
    assert result["json_exists"]
    assert result["company_calendar_day_count"] == 0
    assert result["needs_setup"] is False
    assert not any("会社カレンダーが空" in issue for issue in result["issues"])
