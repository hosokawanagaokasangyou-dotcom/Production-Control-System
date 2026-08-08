# -*- coding: utf-8 -*-
"""配台: machine-calendar-data.json 必須（master フォールバックなし）。"""

from __future__ import annotations

import pytest

from planning_core.bootstrap import PlanningValidationError
from planning_core.core.machine_calendar_paths import ENV_MACHINE_CALENDAR_JSON
from planning_core.core.machine_calendar_store import (
    apply_machine_calendar_patch,
    empty_store,
    require_machine_calendar_json_for_dispatch,
    save_machine_calendar_store,
)

NEED_COLS = [{"equipment_key": "EC+EC機", "process": "EC", "machine": "EC機"}]


def _write_valid_store(path) -> None:
    store = empty_store()
    store["columns"] = NEED_COLS
    apply_machine_calendar_patch(
        store,
        {
            "date": "2026-08-01",
            "rows": [{"slot": "2026-08-01T09:00:00", "cells": {"EC+EC機": "*"}}],
        },
    )
    save_machine_calendar_store(store, path)


def test_require_machine_calendar_json_missing(tmp_path, monkeypatch):
    jp = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(jp))
    assert not jp.is_file()

    with pytest.raises(PlanningValidationError, match="存在しません"):
        require_machine_calendar_json_for_dispatch("テスト配台")


def test_require_machine_calendar_json_empty_store(tmp_path, monkeypatch):
    jp = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(jp))
    save_machine_calendar_store(empty_store(), jp)

    with pytest.raises(PlanningValidationError, match="未整備"):
        require_machine_calendar_json_for_dispatch("テスト配台")


def test_load_machine_calendar_occupancy_blocks_from_json_only(tmp_path, monkeypatch):
    from planning_core.core.roll_pipeline import load_machine_calendar_occupancy_blocks

    jp = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(jp))
    _write_valid_store(jp)

    blocks = load_machine_calendar_occupancy_blocks(
        str(tmp_path / "master.xlsm"),
        ["EC+EC機"],
        context_label="テスト配台",
    )
    assert blocks


def test_load_machine_calendar_raises_without_json(tmp_path, monkeypatch):
    from planning_core.core.roll_pipeline import load_machine_calendar_occupancy_blocks

    jp = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(jp))

    with pytest.raises(PlanningValidationError, match="存在しません"):
        load_machine_calendar_occupancy_blocks(
            str(tmp_path / "master.xlsm"),
            ["EC+EC機"],
            context_label="テスト配台",
        )
