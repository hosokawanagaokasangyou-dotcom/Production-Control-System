# -*- coding: utf-8 -*-

from __future__ import annotations

import json

from planning_core.core.machine_calendar_history_store import (
    append_machine_calendar_snapshot,
    list_machine_calendar_history,
    max_entries,
    restore_machine_calendar_snapshot,
)
from planning_core.core.machine_calendar_paths import ENV_MACHINE_CALENDAR_JSON
from planning_core.core.machine_calendar_store import empty_store, save_machine_calendar_store


def test_machine_calendar_history_max_is_twenty():
    assert max_entries() == 20


def test_append_and_restore_roundtrip(tmp_path, monkeypatch):
    jp = tmp_path / "machine-calendar-data.json"
    hist = tmp_path / "machine-calendar-json-history"
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(jp))
    monkeypatch.setenv("PM_AI_MACHINE_CALENDAR_HISTORY_DIR", str(hist))

    store = empty_store()
    store["columns"] = [{"equipment_key": "EC+EC機", "process": "EC", "machine": "EC機"}]
    store["meta"]["revision"] = 1
    save_machine_calendar_store(store, jp, history_label="初回")
    store["meta"]["revision"] = 2
    save_machine_calendar_store(store, jp, history_label="2回目")
    store["meta"]["revision"] = 3
    save_machine_calendar_store(store, jp, history_label="3回目")

    listed = list_machine_calendar_history(jp)
    assert listed["max_entries"] == 20
    entries = listed["entries"]
    assert len(entries) == 3
    assert entries[0]["label"] == "3回目"
    assert entries[0]["revision"] == 3

    restore_machine_calendar_snapshot(entries[1]["id"], jp)
    restored = json.loads(jp.read_text(encoding="utf-8"))
    assert restored["meta"]["revision"] == 2


def test_history_trim_to_twenty(tmp_path, monkeypatch):
    jp = tmp_path / "machine-calendar-data.json"
    hist = tmp_path / "machine-calendar-json-history"
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(jp))
    monkeypatch.setenv("PM_AI_MACHINE_CALENDAR_HISTORY_DIR", str(hist))

    store = empty_store()
    for i in range(22):
        store["meta"]["revision"] = i
        save_machine_calendar_store(store, jp, history_label=f"save-{i}")

    listed = list_machine_calendar_history(jp)
    assert len(listed["entries"]) == 20


def test_append_snapshot_direct(tmp_path, monkeypatch):
    jp = tmp_path / "machine-calendar-data.json"
    monkeypatch.setenv(ENV_MACHINE_CALENDAR_JSON, str(jp))
    store = empty_store()
    jp.write_text(json.dumps(store, ensure_ascii=False), encoding="utf-8")
    entry = append_machine_calendar_snapshot(jp, kind="test", label="直接退避")
    assert entry is not None
    assert entry["label"] == "直接退避"
