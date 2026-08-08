# -*- coding: utf-8 -*-
"""段階1/2: カレンダー正本 JSON 3種が揃うまで実行不可。"""

from __future__ import annotations

import pytest

from planning_core.bootstrap import PlanningValidationError
from planning_core.core.attendance_paths import ENV_ATTENDANCE_JSON
from planning_core.core.attendance_readiness import (
    build_attendance_readiness,
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
