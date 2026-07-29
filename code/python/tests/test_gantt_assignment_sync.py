"""gantt_assignment_sync の検証・トークン・二重割当テスト。"""

from __future__ import annotations

import json
from datetime import date, datetime

import pytest

pytest.importorskip("pandas")


def _sample_contract(events: list[dict]) -> dict:
    from planning_core.gantt_render_contract import make_gantt_render_contract

    attendance = {
        date(2026, 5, 14): {
            "山田 太郎": {"is_working": True},
            "佐藤 花子": {"is_working": False},
        }
    }
    return make_gantt_render_contract(
        timeline_events=events,
        equipment_list=["EC機　湖南"],
        sorted_dates=[date(2026, 5, 14)],
        attendance_data=attendance,
        data_extract_dt_str="2026/05/14 00:00:00",
        base_now_dt=datetime(2026, 5, 14, 8, 0, 0),
        regular_shift_times=None,
        plan_rows=True,
        kind="equipment_gantt",
    )


def _machining_event(
    *,
    task_id: str,
    start_h: int,
    end_h: int,
    op: str = "",
    sub: str = "",
) -> dict:
    return {
        "date": date(2026, 5, 14),
        "machine": "EC機　湖南",
        "task_id": task_id,
        "event_kind": "machining",
        "start_dt": datetime(2026, 5, 14, start_h, 0, 0),
        "end_dt": datetime(2026, 5, 14, end_h, 0, 0),
        "unit_m": 100.0,
        "units_done": 1.0,
        "op": op,
        "sub": sub,
    }


def test_apply_op_sub_updates_roundtrip():
    from planning_core.gantt_assignment_sync import (
        apply_op_sub_updates,
        parse_updates_map,
        timeline_events_from_contract,
    )

    contract = _sample_contract(
        [_machining_event(task_id="Y1-1", start_h=8, end_h=9, op="A", sub="")]
    )
    updates = parse_updates_map({"0": {"op": "B", "sub": "C,D"}})
    updated = apply_op_sub_updates(contract, updates)
    events = timeline_events_from_contract(updated)
    assert events[0]["op"] == "B"
    assert events[0]["sub"] == "C,D"


def test_double_assignment_rejected():
    from planning_core.gantt_assignment_sync import (
        OpSubUpdate,
        find_double_assignments,
        validate_assignment_changes,
    )

    events = [
        _machining_event(task_id="Y1-1", start_h=8, end_h=10, op="山田 太郎"),
        _machining_event(task_id="Y1-2", start_h=9, end_h=11, op="山田 太郎"),
    ]
    contract = _sample_contract(events)
    issues = find_double_assignments(events)
    assert issues
    assert issues[0].code == "double_assignment"

    result = validate_assignment_changes(
        contract,
        [OpSubUpdate(1, "山田 太郎", "")],
    )
    assert not result.ok
    assert any(e.code == "double_assignment" for e in result.errors)


def test_absent_warning_and_confirm_token():
    from planning_core.gantt_assignment_sync import (
        OpSubUpdate,
        make_confirm_token,
        validate_assignment_changes,
        validate_confirm_token,
    )

    events = [
        _machining_event(task_id="Y1-1", start_h=8, end_h=9, op="山田 太郎"),
    ]
    contract = _sample_contract(events)
    result = validate_assignment_changes(
        contract,
        [OpSubUpdate(0, "佐藤 花子", "")],
    )
    assert not result.ok
    assert result.status == "warnings"
    assert result.warnings
    assert result.warnings[0].code == "absent"
    token = result.confirm_token
    assert validate_confirm_token(result.timeline_hash, result.warnings, token)

    forced = validate_assignment_changes(
        contract,
        [OpSubUpdate(0, "佐藤 花子", "")],
        confirm_token=token,
        force_warnings=True,
    )
    assert forced.ok


def test_hash_mismatch_detected():
    from planning_core.gantt_assignment_sync import (
        OpSubUpdate,
        validate_assignment_changes,
    )

    contract = _sample_contract(
        [_machining_event(task_id="Y1-1", start_h=8, end_h=9, op="A")]
    )
    result = validate_assignment_changes(
        contract,
        [OpSubUpdate(0, "B", "")],
        expected_timeline_hash="deadbeef",
    )
    assert not result.ok
    assert any(e.code == "hash_mismatch" for e in result.errors)


def test_parse_updates_nested_format():
    from planning_core.gantt_assignment_sync import parse_updates_map

    raw = {"updates": {"1": {"op": "X", "sub": ""}}}
    updates = parse_updates_map(raw)
    assert len(updates) == 1
    assert updates[0].event_index == 1
    assert updates[0].op == "X"


def test_timeline_hash_changes_when_op_changes():
    from planning_core.gantt_assignment_sync import (
        apply_op_sub_updates,
        OpSubUpdate,
        timeline_assignment_hash,
        timeline_events_from_contract,
    )

    contract = _sample_contract(
        [_machining_event(task_id="Y1-1", start_h=8, end_h=9, op="A")]
    )
    h0 = timeline_assignment_hash(timeline_events_from_contract(contract))
    updated = apply_op_sub_updates(contract, [OpSubUpdate(0, "B", "")])
    h1 = timeline_assignment_hash(timeline_events_from_contract(updated))
    assert h0 != h1
