"""依頼切替・休憩再開の準備時間（配台）の単体テスト。"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path

import pytest

pytest.importorskip("pandas")

import planning_core._core as core


def _reset_stage2_prep_globals() -> None:
    core._STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE = {("スライス", "スライス機1"): 15}
    core._STAGE2_REQUEST_SWITCH_PREP_BY_MACHINE = {}
    core._STAGE2_BREAK_RESUME_PREP_BY_PROC_MACHINE = {("スライス", "スライス機1"): 5}
    core._STAGE2_BREAK_RESUME_PREP_BY_MACHINE = {}
    core._STAGE2_POST_MACHINING_CLEANUP_BY_PROC_MACHINE = {}
    core._STAGE2_POST_MACHINING_CLEANUP_BY_MACHINE = {}
    core._STAGE2_REQUEST_INTERVAL_BUFFER_BY_PROC_MACHINE = {}
    core._STAGE2_REQUEST_INTERVAL_BUFFER_BY_MACHINE = {}


def _reset_stage2_prep_globals_with_cleanup_buffer() -> None:
    _reset_stage2_prep_globals()
    core._STAGE2_POST_MACHINING_CLEANUP_BY_PROC_MACHINE = {
        ("スライス", "スライス機1"): 10
    }
    core._STAGE2_REQUEST_INTERVAL_BUFFER_BY_PROC_MACHINE = {
        ("スライス", "スライス機1"): 5
    }


def test_load_request_switch_prep_settings_new_column_names(tmp_path: Path):
    pd = pytest.importorskip("pandas")

    master = tmp_path / "master.xlsm"
    df = pd.DataFrame(
        [
            {
                "工程名": "スライス",
                "機械名": "スライス機1",
                "依頼切替準備時間": 15,
                "休憩後再開準備時間": 5,
                "後始末時間": 10,
                "加工依頼間の余裕時間": 5,
            }
        ]
    )
    with pd.ExcelWriter(master, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=core.SHEET_REQUEST_SWITCH_PREP, index=False)

    (
        sp,
        sm,
        rp,
        rm,
        cp,
        cm,
        bp,
        bm,
    ) = core.load_request_switch_prep_settings(str(master))
    assert sp[("スライス", "スライス機1")] == 15
    assert rp[("スライス", "スライス機1")] == 5
    assert cp[("スライス", "スライス機1")] == 10
    assert bp[("スライス", "スライス機1")] == 5
    assert not sm
    assert not rm
    assert not cm
    assert not bm


def test_lookup_request_switch_prep_by_proc_machine():
    _reset_stage2_prep_globals()
    assert core._lookup_request_switch_prep_minutes("スライス", "スライス機1") == 15
    assert core._lookup_request_switch_prep_minutes("スライス", "存在しない") == 0


def test_lookup_post_machining_cleanup_and_interval_buffer():
    _reset_stage2_prep_globals_with_cleanup_buffer()
    assert core._lookup_post_machining_cleanup_minutes("スライス", "スライス機1") == 0
    assert core._lookup_request_interval_buffer_minutes("スライス", "スライス機1") == 5


def test_lookup_post_machining_cleanup_when_enabled(monkeypatch):
    _reset_stage2_prep_globals_with_cleanup_buffer()
    monkeypatch.setattr(core, "_stage2_post_machining_cleanup_enabled", lambda: True)
    assert core._lookup_post_machining_cleanup_minutes("スライス", "スライス機1") == 10


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


def test_prep_segments_from_anchor():
    t0 = datetime(2026, 5, 18, 10, 0)
    t1, segs = core._prep_segments_from_anchor(
        anchor=t0,
        prep_minutes=10,
        event_kind=core.TIMELINE_EVENT_POST_MACHINING_CLEANUP,
        eq_line="スライス+スライス機1",
        machine_occ_key="スライス機1",
    )
    assert t1 == t0 + timedelta(minutes=10)
    assert len(segs) == 1
    assert segs[0]["event_kind"] == core.TIMELINE_EVENT_POST_MACHINING_CLEANUP


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


def test_resolve_prev_machining_end_falls_back_to_machine_avail_dt():
    d = date(2026, 5, 18)
    prev_end = datetime(2026, 5, 18, 10, 0)
    mh = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
        "last_machining_dt": {},
    }
    got = core._resolve_prev_machining_end_for_request_switch(
        machine_handoff=mh,
        machine_occ_key="occ1",
        explicit=None,
        machine_avail_dt={"occ1": prev_end},
        machine_day_floor=datetime(2026, 5, 18, 8, 0),
    )
    assert got == prev_end


def test_roll_prep_cleanup_anchors_prev_end_when_team_start_deferred():
    """後始末無効時: 依頼間余裕は prev_end から、準備は team_start 直前。"""
    _reset_stage2_prep_globals_with_cleanup_buffer()
    d = date(2026, 5, 18)
    prev_end = datetime(2026, 5, 18, 10, 0)
    team_start = datetime(2026, 5, 18, 11, 0)
    mh = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
        "last_eq": {"occ1": "スライス+スライス機1"},
        "last_machining_dt": {},
    }
    ts, segs = core._roll_prep_segments_for_assign(
        team_start=team_start,
        team_breaks=[],
        machine_handoff=mh,
        machine_occ_key="occ1",
        current_date=d,
        task_id="B002",
        machine_proc="スライス",
        machine_name="スライス機1",
        eq_line="スライス+スライス機1",
        abolish_limits=False,
        prev_machining_end=None,
        prev_eq_line="スライス+スライス機1",
        machine_avail_dt={"occ1": prev_end},
        machine_day_floor=datetime(2026, 5, 18, 8, 0),
    )
    assert [s["event_kind"] for s in segs] == [
        core.TIMELINE_EVENT_REQUEST_INTERVAL_BUFFER,
        core.TIMELINE_EVENT_REQUEST_SWITCH_PREP,
    ]
    assert segs[0]["start_dt"] == prev_end
    assert segs[0]["end_dt"] == prev_end + timedelta(minutes=5)
    assert segs[1]["start_dt"] == team_start
    assert segs[1]["end_dt"] == team_start + timedelta(minutes=15)
    assert ts == team_start + timedelta(minutes=15)


def test_roll_prep_switch_includes_cleanup_buffer_and_prep_when_enabled(monkeypatch):
    _reset_stage2_prep_globals_with_cleanup_buffer()
    monkeypatch.setattr(core, "_stage2_post_machining_cleanup_enabled", lambda: True)
    d = date(2026, 5, 18)
    prev_end = datetime(2026, 5, 18, 10, 0)
    team_start = datetime(2026, 5, 18, 10, 0)
    mh = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
        "last_eq": {"occ1": "スライス+スライス機1"},
        "last_machining_dt": {"occ1": prev_end},
    }
    ts, segs = core._roll_prep_segments_for_assign(
        team_start=team_start,
        team_breaks=[],
        machine_handoff=mh,
        machine_occ_key="occ1",
        current_date=d,
        task_id="B002",
        machine_proc="スライス",
        machine_name="スライス機1",
        eq_line="スライス+スライス機1",
        abolish_limits=False,
        prev_machining_end=prev_end,
        prev_eq_line="スライス+スライス機1",
    )
    assert [s["event_kind"] for s in segs] == [
        core.TIMELINE_EVENT_POST_MACHINING_CLEANUP,
        core.TIMELINE_EVENT_REQUEST_INTERVAL_BUFFER,
        core.TIMELINE_EVENT_REQUEST_SWITCH_PREP,
    ]
    assert segs[0]["start_dt"] == prev_end
    assert segs[0]["end_dt"] == prev_end + timedelta(minutes=10)
    assert segs[1]["start_dt"] == segs[0]["end_dt"]
    assert segs[1]["end_dt"] == prev_end + timedelta(minutes=15)
    assert segs[2]["start_dt"] == segs[1]["end_dt"]
    assert segs[2]["end_dt"] == prev_end + timedelta(minutes=30)
    assert ts == prev_end + timedelta(minutes=30)


def test_roll_prep_switch_includes_buffer_and_prep_without_cleanup():
    _reset_stage2_prep_globals_with_cleanup_buffer()
    d = date(2026, 5, 18)
    prev_end = datetime(2026, 5, 18, 10, 0)
    team_start = datetime(2026, 5, 18, 10, 0)
    mh = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
        "last_eq": {"occ1": "スライス+スライス機1"},
        "last_machining_dt": {"occ1": prev_end},
    }
    ts, segs = core._roll_prep_segments_for_assign(
        team_start=team_start,
        team_breaks=[],
        machine_handoff=mh,
        machine_occ_key="occ1",
        current_date=d,
        task_id="B002",
        machine_proc="スライス",
        machine_name="スライス機1",
        eq_line="スライス+スライス機1",
        abolish_limits=False,
        prev_machining_end=prev_end,
        prev_eq_line="スライス+スライス機1",
    )
    assert [s["event_kind"] for s in segs] == [
        core.TIMELINE_EVENT_REQUEST_INTERVAL_BUFFER,
        core.TIMELINE_EVENT_REQUEST_SWITCH_PREP,
    ]
    assert segs[0]["start_dt"] == prev_end
    assert segs[0]["end_dt"] == prev_end + timedelta(minutes=5)
    assert segs[1]["start_dt"] == prev_end + timedelta(minutes=5)
    assert segs[1]["end_dt"] == prev_end + timedelta(minutes=20)
    assert ts == prev_end + timedelta(minutes=20)


def test_roll_prep_skips_buffer_when_prev_end_unknown():
    """直前加工終了が不明なとき、余裕を team_start に付けず（ギャップ中途表示を防ぐ）準備のみ。"""
    _reset_stage2_prep_globals_with_cleanup_buffer()
    d = date(2026, 5, 18)
    team_start = datetime(2026, 5, 18, 11, 0)
    mh = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
        "last_eq": {"occ1": "スライス+スライス機1"},
        "last_machining_dt": {},
    }
    ts, segs = core._roll_prep_segments_for_assign(
        team_start=team_start,
        team_breaks=[],
        machine_handoff=mh,
        machine_occ_key="occ1",
        current_date=d,
        task_id="B002",
        machine_proc="スライス",
        machine_name="スライス機1",
        eq_line="スライス+スライス機1",
        abolish_limits=False,
        prev_machining_end=None,
        prev_eq_line="スライス+スライス機1",
        machine_avail_dt={},
        machine_day_floor=datetime(2026, 5, 18, 8, 0),
    )
    assert [s["event_kind"] for s in segs] == [
        core.TIMELINE_EVENT_REQUEST_SWITCH_PREP,
    ]
    assert segs[0]["start_dt"] == team_start
    assert ts == team_start + timedelta(minutes=15)


def test_resolve_prev_machining_end_for_roll_prep_uses_raw_avail():
    d = date(2026, 5, 18)
    prev_end = datetime(2026, 5, 18, 10, 0)
    mh = {
        "last_machining_dt": {},
        "machining_today_occ": {"occ1"},
    }
    got = core._resolve_prev_machining_end_for_roll_prep(
        mh,
        "occ1",
        None,
        {"occ1": datetime(2026, 5, 18, 11, 30)},
        datetime(2026, 5, 18, 8, 0),
        prev_end,
    )
    assert got == prev_end


def test_team_start_is_post_break_resume():
    be = datetime(2026, 5, 18, 12, 50)
    assert core._team_start_is_immediate_post_break_resume(
        be, [(datetime(2026, 5, 18, 12, 0), be)]
    )
    assert not core._team_start_is_immediate_post_break_resume(
        be + timedelta(minutes=1),
        [(datetime(2026, 5, 18, 12, 0), be)],
    )
