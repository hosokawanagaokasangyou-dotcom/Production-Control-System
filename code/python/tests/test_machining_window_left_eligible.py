# -*- coding: utf-8 -*-
"""窓不足フィルタ: 1ロールが工場窓に収まるのに avg_eff=0.5 固定で除外しないこと。"""

from __future__ import annotations

from datetime import date, datetime, time

from planning_core.core.roll_pipeline import (
    DEFAULT_END_TIME,
    DEFAULT_START_TIME,
    _task_no_machining_window_left_from_avail_floor,
)


def _day_status(day: date, *, efficiency: float = 1.0) -> dict:
    return {
        "OP1": {
            "eligible_for_assignment": True,
            "is_working": True,
            "start_dt": datetime.combine(day, DEFAULT_START_TIME),
            "end_dt": datetime.combine(day, time(19, 0)),
            "efficiency": efficiency,
            "breaks": [],
        }
    }


def test_window_left_allows_300min_roll_in_factory_day_at_eff_1():
    """
    W9-1 相当: btp=300・工場窓 DEFAULT_START〜DEFAULT_END（約495分）。
    効率1.0なら 1ロール分は入る。旧実装の /0.5 固定だと 600分扱いになり永久除外される。
    """
    day = date(2026, 9, 7)
    task = {
        "task_id": "W9-1",
        "machine": "検査",
        "machine_name": "熱融着機　湖南",
        "equipment_line_key": "検査+熱融着機　湖南",
        "base_time_per_unit": 300.0,
        "task_eff_factor": 1.0,
    }
    members = ["OP1"]
    daily_status = _day_status(day, efficiency=1.0)
    machine_day_start = datetime.combine(day, DEFAULT_START_TIME)
    blocked = _task_no_machining_window_left_from_avail_floor(
        task,
        day,
        daily_status,
        members,
        {},
        machine_day_start,
    )
    assert blocked is False
    # 窓終端は工場マスタ終業（17:00）とメンバー終業の小さい方
    assert DEFAULT_END_TIME == time(17, 0)


def test_window_left_blocks_when_member_efficiency_too_low():
    """実効効率が低く 1ロールが窓に入らないときは除外する。"""
    day = date(2026, 9, 7)
    task = {
        "task_id": "W9-1",
        "machine": "検査",
        "machine_name": "熱融着機　湖南",
        "equipment_line_key": "検査+熱融着機　湖南",
        "base_time_per_unit": 300.0,
        "task_eff_factor": 1.0,
    }
    members = ["OP1"]
    daily_status = _day_status(day, efficiency=0.4)
    machine_day_start = datetime.combine(day, DEFAULT_START_TIME)
    blocked = _task_no_machining_window_left_from_avail_floor(
        task,
        day,
        daily_status,
        members,
        {},
        machine_day_start,
    )
    assert blocked is True
