# -*- coding: utf-8 -*-
"""Tests for machine_calendar_store."""

from __future__ import annotations

from datetime import date

from planning_core.core.attendance_store import DAY_KIND_PUBLIC, empty_store as empty_attendance_store
from planning_core.core.machine_calendar_store import (
    apply_machine_calendar_patch,
    build_editor_payload,
    collect_machine_calendar_export_rows,
    empty_store,
    initialize_machine_calendar_defaults,
    initialize_machine_calendar_from_company_calendar,
    occupancy_blocks_from_store,
    slot_keys_for_factory_window,
    store_has_machine_calendar_data,
    validate_store_for_dispatch,
    MACHINE_CAL_FACTORY_END,
    MACHINE_CAL_FACTORY_START,
)


NEED_COLS = [
    {"equipment_key": "EC+EC機", "process": "EC", "machine": "EC機"},
]


def test_empty_store_not_ready():
    store = empty_store()
    assert not store_has_machine_calendar_data(store)
    assert not validate_store_for_dispatch(store)


def test_apply_patch_and_editor_payload():
    store = empty_store()
    store["columns"] = NEED_COLS
    slot = "2026-08-01T09:00:00"
    apply_machine_calendar_patch(
        store,
        {
            "date": "2026-08-01",
            "rows": [{"slot": slot, "cells": {"EC+EC機": "*"}}],
        },
    )
    assert store_has_machine_calendar_data(store)
    assert validate_store_for_dispatch(store)
    payload = build_editor_payload(store, date(2026, 8, 1), NEED_COLS)
    assert payload["ok"]
    assert len(payload["rows"]) == 1
    blocks, _ = occupancy_blocks_from_store(store, ["EC+EC機"])
    assert blocks


def test_build_editor_payload_uses_need_columns_and_factory_slots():
    store = empty_store()
    payload = build_editor_payload(store, date(2026, 8, 8), NEED_COLS)
    assert payload["ok"]
    assert payload["columns"][0]["equipment_key"] == "EC+EC機"
    assert len(payload["rows"]) > 0
    first_slot = payload["rows"][0]["slot"]
    assert first_slot.startswith("2026-08-08T08:00")


def test_comment_patch_roundtrip():
    store = empty_store()
    store["columns"] = NEED_COLS
    slot = "2026-08-01T09:00:00"
    apply_machine_calendar_patch(
        store,
        {
            "date": "2026-08-01",
            "rows": [
                {
                    "slot": slot,
                    "cells": {"EC+EC機": "*"},
                    "comments": {"EC+EC機": "保全作業"},
                }
            ],
        },
    )
    payload = build_editor_payload(store, date(2026, 8, 1), NEED_COLS)
    assert payload["rows"][0]["comments"]["EC+EC機"] == "保全作業"
    apply_machine_calendar_patch(
        store,
        {
            "date": "2026-08-01",
            "rows": [{"slot": slot, "comments": {"EC+EC機": ""}}],
        },
    )
    payload2 = build_editor_payload(store, date(2026, 8, 1), NEED_COLS)
    assert "comments" not in payload2["rows"][0]


def test_initialize_defaults_all_days_available():
    mc = empty_store()
    att = empty_attendance_store()
    att["company_calendar"]["days"]["2026-08-06"] = {"kind": DAY_KIND_PUBLIC}
    result = initialize_machine_calendar_defaults(
        mc, 2026, NEED_COLS, start_month=4, start_day=1
    )
    assert result["columns"] == 1
    assert result["initialized_days"] > 0
    # Saturday 2026-08-08 — 土日も占有なし（稼働可）
    sat_slots = slot_keys_for_factory_window(
        date(2026, 8, 8), MACHINE_CAL_FACTORY_START, MACHINE_CAL_FACTORY_END
    )
    for sk in sat_slots:
        assert sk not in mc["occupancy"]
    # Monday 2026-08-10
    mon_slots = slot_keys_for_factory_window(
        date(2026, 8, 10), MACHINE_CAL_FACTORY_START, MACHINE_CAL_FACTORY_END
    )
    for sk in mon_slots:
        assert sk not in mc["occupancy"]
    # 後方互換 API も同様（会社カレンダーは無視）
    mc2 = empty_store()
    initialize_machine_calendar_from_company_calendar(
        mc2, att, 2026, NEED_COLS, start_month=4, start_day=1
    )
    for sk in sat_slots:
        assert sk not in mc2["occupancy"]


def test_collect_export_rows_for_fiscal_range():
    mc = empty_store()
    mc["columns"] = NEED_COLS
    initialize_machine_calendar_defaults(
        mc, 2026, NEED_COLS, start_month=4, start_day=1
    )
    cols, rows = collect_machine_calendar_export_rows(
        mc, date(2026, 8, 1), date(2026, 8, 10)
    )
    assert len(cols) == 1
    assert len(rows) >= 22
    sat_rows = [r for r in rows if r["slot"].startswith("2026-08-08")]
    assert sat_rows
    assert sat_rows[0]["cells"]["EC+EC機"] == ""
