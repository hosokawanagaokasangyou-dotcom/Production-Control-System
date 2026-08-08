# -*- coding: utf-8 -*-
"""段階1/2: カレンダー正本 JSON 3種が揃うまで実行不可。"""

from __future__ import annotations

import pytest

from planning_core.bootstrap import PlanningValidationError
from planning_core.core.attendance_paths import ENV_ATTENDANCE_JSON
from planning_core.core.attendance_readiness import (
    build_attendance_readiness,
    legacy_master_attendance_sheets_required,
    require_calendar_json_for_planning_stages,
)
from planning_core.core.attendance_store import (
    apply_company_calendar_to_members,
    empty_store,
    initialize_company_calendar,
    save_attendance_store,
)
from planning_core.core.machine_calendar_paths import ENV_MACHINE_CALENDAR_JSON
from planning_core.core.machine_calendar_store import (
    apply_machine_calendar_patch,
    empty_store as mc_empty,
    initialize_machine_calendar_defaults,
    save_machine_calendar_store,
)

NEED_COLS = [{"equipment_key": "EC+EC機", "process": "EC", "machine": "EC機"}]


def _write_ready_attendance(path) -> dict:
    store = empty_store(2026)
    initialize_company_calendar(store, 2026, start_month=4, start_day=1)
    apply_company_calendar_to_members(store, ["A"], 2026, 8)
    save_attendance_store(store, path)
    return store


def _write_ready_machine_calendar(path) -> None:
    store = mc_empty()
    initialize_machine_calendar_defaults(store, 2026, NEED_COLS, start_month=4, start_day=1)
    apply_machine_calendar_patch(
        store,
        {
            "date": "2026-08-01",
            "rows": [{"slot": "2026-08-01T09:00:00", "cells": {"EC+EC機": "*"}}],
        },
    )
    save_machine_calendar_store(store, path)


def test_require_calendar_json_raises_when_incomplete(tmp_path, monkeypatch):
    att = tmp_path / "attendance-data.json"
    mc = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_ATTENDANCE_JSON, str(att))
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(mc))

    with pytest.raises(PlanningValidationError, match="attendance-data.json"):
        require_calendar_json_for_planning_stages("テスト")


def test_planning_stages_ready_when_all_json_present(tmp_path, monkeypatch):
    att = tmp_path / "attendance-data.json"
    mc = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_ATTENDANCE_JSON, str(att))
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(mc))
    store = _write_ready_attendance(att)
    _write_ready_machine_calendar(mc)

    result = build_attendance_readiness(store=store, members=["A"], year=2026, month=8)
    assert result["company_calendar_ready"]
    assert result["member_attendance_ready"]
    assert result["machine_calendar_ready"]
    assert result["planning_stages_ready"]
    assert result["stage1_ready"]

    import planning_core.core.attendance_readiness as readiness_mod

    monkeypatch.setattr(
        readiness_mod,
        "build_attendance_readiness",
        lambda *args, **kwargs: result,
    )
    require_calendar_json_for_planning_stages("テスト")


def test_require_calendar_json_uses_roster_when_skills_unavailable(tmp_path, monkeypatch):
    """段階1: skills 読込失敗時も attendance-data.json 名簿で readiness 判定する。"""
    att = tmp_path / "attendance-data.json"
    mc = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_ATTENDANCE_JSON, str(att))
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(mc))

    store = empty_store(2026)
    store["member_roster"] = [{"name": "A", "primary_role": "後加工"}]
    initialize_company_calendar(store, 2026, start_month=4, start_day=1)
    apply_company_calendar_to_members(store, ["A"], 2026, 8)
    save_attendance_store(store, att)
    _write_ready_machine_calendar(mc)

    def _fail_skills():
        raise RuntimeError("skills unavailable")

    import planning_core.core.master_data as master_data_mod

    monkeypatch.setattr(master_data_mod, "load_skills_and_needs", _fail_skills)

    result = build_attendance_readiness()
    assert result["planning_stages_ready"]
    require_calendar_json_for_planning_stages("テスト")


def test_legacy_master_attendance_sheets_not_required_when_json_ready(
    tmp_path, monkeypatch
):
    att = tmp_path / "attendance-data.json"
    mc = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_ATTENDANCE_JSON, str(att))
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(mc))

    store = empty_store(2026)
    store["member_roster"] = [{"name": "A", "primary_role": "後加工"}]
    initialize_company_calendar(store, 2026, start_month=4, start_day=1)
    apply_company_calendar_to_members(store, ["A"], 2026, 8)
    save_attendance_store(store, att)
    _write_ready_machine_calendar(mc)

    assert not legacy_master_attendance_sheets_required()


def test_legacy_master_attendance_check_does_not_load_skills(tmp_path, monkeypatch):
    """legacy 判定が load_skills_and_needs を呼ばない（無限再帰防止）。"""
    att = tmp_path / "attendance-data.json"
    mc = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_ATTENDANCE_JSON, str(att))
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(mc))

    store = empty_store(2026)
    store["member_roster"] = [{"name": "A", "primary_role": "後加工"}]
    initialize_company_calendar(store, 2026, start_month=4, start_day=1)
    apply_company_calendar_to_members(store, ["A"], 2026, 8)
    save_attendance_store(store, att)
    _write_ready_machine_calendar(mc)

    import planning_core.core.master_data as master_data_mod

    def _boom():
        raise AssertionError("load_skills_and_needs must not run during legacy check")

    monkeypatch.setattr(master_data_mod, "load_skills_and_needs", _boom)
    assert not legacy_master_attendance_sheets_required()
