# -*- coding: utf-8 -*-
"""Tests for attendance_history_store."""

from __future__ import annotations

import json

from planning_core.core.attendance_history_store import (
    append_attendance_snapshot,
    list_attendance_history,
    max_entries,
    restore_attendance_snapshot,
)
from planning_core.core.attendance_store import empty_store, save_attendance_store


def test_attendance_history_max_is_twenty():
    assert max_entries() == 20


def test_append_and_restore_roundtrip(tmp_path, monkeypatch):
    jp = tmp_path / "attendance-data.json"
    hist = tmp_path / "attendance-json-history"
    monkeypatch.setenv("PM_AI_ATTENDANCE_JSON", str(jp))
    monkeypatch.setenv("PM_AI_ATTENDANCE_JSON_HISTORY_DIR", str(hist))

    store = empty_store(2026)
    store["meta"]["company_calendar_revision"] = 1
    save_attendance_store(store, jp, history_label="初回")
    save_attendance_store(store, jp, history_label="2回目")
    store["meta"]["company_calendar_revision"] = 2
    save_attendance_store(store, jp, history_label="3回目")

    listed = list_attendance_history(jp)
    assert listed["max_entries"] == 20
    entries = listed["entries"]
    assert len(entries) == 3
    assert entries[0]["label"] == "3回目"

    restore_attendance_snapshot(entries[1]["id"], jp)
    restored = json.loads(jp.read_text(encoding="utf-8"))
    assert restored["meta"]["company_calendar_revision"] == 1


def test_history_trim_to_twenty(tmp_path, monkeypatch):
    jp = tmp_path / "attendance-data.json"
    hist = tmp_path / "attendance-json-history"
    monkeypatch.setenv("PM_AI_ATTENDANCE_JSON", str(jp))
    monkeypatch.setenv("PM_AI_ATTENDANCE_JSON_HISTORY_DIR", str(hist))

    store = empty_store(2026)
    for i in range(22):
        store["meta"]["company_calendar_revision"] = i
        save_attendance_store(store, jp, history_label=f"save-{i}")

    listed = list_attendance_history(jp)
    assert len(listed["entries"]) == 20
