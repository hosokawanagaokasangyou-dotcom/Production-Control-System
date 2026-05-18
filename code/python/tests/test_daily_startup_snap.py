"""日次始業準備のタイムライン寄せ（先頭加工直前）の後処理テスト。"""

from __future__ import annotations

from datetime import date, datetime, timedelta

import planning_core._core as core


def _startup_event(start: datetime, end: datetime, occ: str = "スライス機1") -> dict:
    d = start.date()
    return {
        "date": d,
        "start_dt": start,
        "end_dt": end,
        "machine_occupancy_key": occ,
        "event_kind": core.TIMELINE_EVENT_MACHINE_DAILY_STARTUP,
        "machine": f"スライス+{occ}",
        "op": "",
        "sub": "",
        "breaks": [],
    }


def _machining_event(start: datetime, end: datetime, task_id: str, occ: str = "スライス機1") -> dict:
    d = start.date()
    return {
        "date": d,
        "start_dt": start,
        "end_dt": end,
        "machine_occupancy_key": occ,
        "event_kind": "machining",
        "task_id": task_id,
        "machine": f"スライス+{occ}",
        "op": "京　直道",
        "sub": "",
        "breaks": [
            (
                datetime.combine(d, datetime.strptime("12:00", "%H:%M").time()),
                datetime.combine(d, datetime.strptime("12:50", "%H:%M").time()),
            )
        ],
    }


def test_repair_snaps_daily_startup_before_break_when_first_machining_is_after_break():
    d = date(2026, 5, 19)
    occ = "スライス機1"
    morning_start = datetime.combine(d, datetime.strptime("08:25", "%H:%M").time())
    morning_end = datetime.combine(d, datetime.strptime("08:55", "%H:%M").time())
    break_start = datetime.combine(d, datetime.strptime("12:00", "%H:%M").time())
    mach_start = datetime.combine(d, datetime.strptime("13:00", "%H:%M").time())
    mach_end = datetime.combine(d, datetime.strptime("13:10", "%H:%M").time())

    su = _startup_event(morning_start, morning_end, occ)
    mach = _machining_event(mach_start, mach_end, "Y5-25", occ)
    timeline = [su, mach]
    task_queue = [
        {
            "task_id": "Y5-25",
            "machine": "スライス",
            "machine_name": occ,
        }
    ]

    attendance = {d: {"京　直道": {"breaks_dt": mach["breaks"]}}}
    n = core._repair_timeline_daily_startup_snapped_to_first_machining(
        timeline, task_queue, attendance, {}, {}
    )
    assert n == 1
    assert su["end_dt"] == break_start
    assert su["start_dt"] == break_start - (morning_end - morning_start)


def test_repair_still_snaps_when_first_machining_is_not_after_break():
    d = date(2026, 5, 19)
    occ = "スライス機1"
    morning_start = datetime.combine(d, datetime.strptime("08:25", "%H:%M").time())
    morning_end = datetime.combine(d, datetime.strptime("08:55", "%H:%M").time())
    mach_start = datetime.combine(d, datetime.strptime("09:30", "%H:%M").time())
    mach_end = datetime.combine(d, datetime.strptime("09:40", "%H:%M").time())

    su = _startup_event(morning_start, morning_end, occ)
    mach = _machining_event(mach_start, mach_end, "Y5-25", occ)
    mach["breaks"] = []
    timeline = [su, mach]
    task_queue = [
        {
            "task_id": "Y5-25",
            "machine": "スライス",
            "machine_name": occ,
        }
    ]

    attendance = {d: {"京　直道": {"breaks_dt": []}}}
    n = core._repair_timeline_daily_startup_snapped_to_first_machining(
        timeline, task_queue, attendance, {}, {}
    )
    assert n == 1
    assert su["end_dt"] == mach_start
    assert su["start_dt"] == mach_start - (morning_end - morning_start)


def test_repair_snaps_before_break_when_break_resume_prep_abuts_machining():
    d = date(2026, 5, 19)
    occ = "スライス機1"
    morning_start = datetime.combine(d, datetime.strptime("08:25", "%H:%M").time())
    morning_end = datetime.combine(d, datetime.strptime("08:55", "%H:%M").time())
    break_start = datetime.combine(d, datetime.strptime("12:00", "%H:%M").time())
    prep_start = datetime.combine(d, datetime.strptime("12:50", "%H:%M").time())
    mach_start = datetime.combine(d, datetime.strptime("12:55", "%H:%M").time())
    mach_end = datetime.combine(d, datetime.strptime("13:10", "%H:%M").time())

    su = _startup_event(morning_start, morning_end, occ)
    prep = {
        "date": d,
        "start_dt": prep_start,
        "end_dt": mach_start,
        "machine_occupancy_key": occ,
        "event_kind": core.TIMELINE_EVENT_BREAK_RESUME_PREP,
        "machine": f"スライス+{occ}",
        "op": "",
        "sub": "",
        "breaks": [],
    }
    mach = _machining_event(mach_start, mach_end, "Y5-25", occ)
    timeline = [su, prep, mach]
    task_queue = [
        {
            "task_id": "Y5-25",
            "machine": "スライス",
            "machine_name": occ,
        }
    ]

    attendance = {d: {"京　直道": {"breaks_dt": mach["breaks"]}}}
    n = core._repair_timeline_daily_startup_snapped_to_first_machining(
        timeline, task_queue, attendance, {}, {}
    )
    assert n == 1
    assert su["end_dt"] == break_start
    assert prep["start_dt"] == prep_start
    assert prep["end_dt"] == mach_start
