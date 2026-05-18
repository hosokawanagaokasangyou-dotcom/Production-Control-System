"""依頼切替・休憩再開の準備時間（配台）の単体テスト。"""

from __future__ import annotations

from datetime import date, datetime, timedelta

import pytest

pytest.importorskip("pandas")

import planning_core._core as core


def _reset_stage2_prep_globals() -> None:
    core._STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE = {("スライス", "スライス機1"): 15}
    core._STAGE2_REQUEST_SWITCH_PREP_BY_MACHINE = {}
    core._STAGE2_BREAK_RESUME_PREP_BY_PROC_MACHINE = {("スライス", "スライス機1"): 5}
    core._STAGE2_BREAK_RESUME_PREP_BY_MACHINE = {}


def test_lookup_request_switch_prep_by_proc_machine():
    _reset_stage2_prep_globals()
    assert core._lookup_request_switch_prep_minutes("スライス", "スライス機1") == 15
    assert core._lookup_request_switch_prep_minutes("スライス", "存在しない") == 0


def test_generate_plan_impl_must_assign_module_globals_not_locals():
    """_generate_plan_impl と同型: 関数内ローカル代入では lookup が 0 のまま。"""
    sp = {("スライス", "スライス機3"): 15}
    core._STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE = {}
    core._STAGE2_REQUEST_SWITCH_PREP_BY_MACHINE = {}

    def _assign_local_only() -> int:
        _STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE = sp  # noqa: F841 — ローカル
        return core._lookup_request_switch_prep_minutes("スライス", "スライス機3")

    def _assign_on_module() -> int:
        core._STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE = sp
        return core._lookup_request_switch_prep_minutes("スライス", "スライス機3")

    assert _assign_local_only() == 0
    assert _assign_on_module() == 15


def test_lookup_break_resume_prep():
    _reset_stage2_prep_globals()
    assert core._lookup_break_resume_prep_minutes("スライス", "スライス機1") == 5


def test_prep_segment_extends_machining_start():
    t0 = datetime(2026, 5, 18, 10, 0)
    t1, segs = core._prep_segments_immediately_before_machining(
        machining_start=t0,
        prep_minutes=15,
        event_kind=core.TIMELINE_EVENT_REQUEST_SWITCH_PREP,
        eq_line="スライス+スライス機1",
        machine_occ_key="スライス機1",
    )
    assert t1 == t0 + timedelta(minutes=15)
    assert len(segs) == 1
    assert segs[0]["start_dt"] == t0
    assert segs[0]["end_dt"] == t1


def test_needs_request_switch_prep():
    d = date(2026, 5, 18)
    mh = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
    }
    assert core._needs_request_switch_prep(mh, "occ1", d, "B002")
    assert not core._needs_request_switch_prep(mh, "occ1", d, "A001")


def test_roll_prep_switch_only_not_resume_after_break():
    _reset_stage2_prep_globals()
    d = date(2026, 5, 18)
    break_end = datetime(2026, 5, 18, 12, 50)
    team_breaks = [(datetime(2026, 5, 18, 12, 0), break_end)]
    mh_switch = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
    }
    ts, segs = core._roll_prep_segments_for_assign(
        team_start=break_end,
        team_breaks=team_breaks,
        machine_handoff=mh_switch,
        machine_occ_key="occ1",
        current_date=d,
        task_id="B002",
        machine_proc="スライス",
        machine_name="スライス機1",
        eq_line="スライス+スライス機1",
        abolish_limits=False,
    )
    assert len(segs) == 1
    assert segs[0]["event_kind"] == core.TIMELINE_EVENT_REQUEST_SWITCH_PREP
    assert ts == break_end + timedelta(minutes=15)

    mh_same = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
    }
    ts2, segs2 = core._roll_prep_segments_for_assign(
        team_start=break_end,
        team_breaks=team_breaks,
        machine_handoff=mh_same,
        machine_occ_key="occ1",
        current_date=d,
        task_id="A001",
        machine_proc="スライス",
        machine_name="スライス機1",
        eq_line="スライス+スライス機1",
        abolish_limits=False,
    )
    assert len(segs2) == 1
    assert segs2[0]["event_kind"] == core.TIMELINE_EVENT_BREAK_RESUME_PREP
    assert ts2 == break_end + timedelta(minutes=5)


def test_team_start_is_post_break_resume():
    be = datetime(2026, 5, 18, 12, 50)
    assert core._team_start_is_immediate_post_break_resume(
        be, [(datetime(2026, 5, 18, 12, 0), be)]
    )
    assert not core._team_start_is_immediate_post_break_resume(
        be + timedelta(minutes=1),
        [(datetime(2026, 5, 18, 12, 0), be)],
    )
