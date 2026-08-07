# -*- coding: utf-8 -*-
"""Tests for machine_calendar_store."""

from __future__ import annotations

from datetime import date

from planning_core.core.machine_calendar_store import (
    apply_machine_calendar_patch,
    build_editor_payload,
    empty_store,
    occupancy_blocks_from_store,
    store_has_machine_calendar_data,
    validate_store_for_dispatch,
)


def test_empty_store_not_ready():
    store = empty_store()
    assert not store_has_machine_calendar_data(store)
    assert not validate_store_for_dispatch(store)


def test_apply_patch_and_editor_payload():
    store = empty_store()
    store["columns"] = [
        {"equipment_key": "後加工+EC機", "process": "後加工", "machine": "EC機"}
    ]
    slot = "2026-08-01T09:00:00"
    apply_machine_calendar_patch(
        store,
        {
            "date": "2026-08-01",
            "rows": [{"slot": slot, "cells": {"後加工+EC機": "*"}}],
        },
    )
    assert store_has_machine_calendar_data(store)
    assert validate_store_for_dispatch(store)
    payload = build_editor_payload(store, date(2026, 8, 1), ["後加工+EC機"])
    assert payload["ok"]
    assert len(payload["rows"]) == 1
    blocks, _ = occupancy_blocks_from_store(store, ["後加工+EC機"])
    assert blocks


def test_build_editor_payload_generates_default_factory_slots():
    store = empty_store()
    store["columns"] = [
        {"equipment_key": "後加工+EC機", "process": "後加工", "machine": "EC機"}
    ]
    payload = build_editor_payload(store, date(2026, 8, 8), ["後加工+EC機"])
    assert payload["ok"]
    assert len(payload["rows"]) > 0
    assert payload["rows"][0]["slot"]
