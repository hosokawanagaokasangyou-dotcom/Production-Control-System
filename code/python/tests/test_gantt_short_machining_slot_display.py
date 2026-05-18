# -*- coding: utf-8 -*-
"""10分未満の加工をガント1スロット枠に拡張して表示する挙動。"""

from __future__ import annotations

from datetime import datetime, timedelta

from planning_core._core import (
    GANTT_TIMELINE_SLOT_MINUTES,
    _eq_grid_slot_overlaps_event,
    _gantt_machining_display_range_for_slot_overlap,
)


def _machining_ev(start: datetime, end: datetime) -> dict:
    return {
        "date": start.date(),
        "machine": "EC機　湖南",
        "task_id": "Y1-1",
        "event_kind": "machining",
        "start_dt": start,
        "end_dt": end,
    }


def test_display_range_expands_sub_ten_minute_machining_to_one_slot():
    st = datetime(2026, 5, 14, 8, 5)
    ed = datetime(2026, 5, 14, 8, 7)
    disp_st, disp_ed = _gantt_machining_display_range_for_slot_overlap(
        _machining_ev(st, ed)
    )
    assert disp_st == datetime(2026, 5, 14, 8, 0)
    assert disp_ed == datetime(2026, 5, 14, 8, 10)


def test_slot_overlap_covers_only_floored_slot():
    ev = _machining_ev(
        datetime(2026, 5, 14, 8, 5), datetime(2026, 5, 14, 8, 7)
    )
    sm = GANTT_TIMELINE_SLOT_MINUTES
    slot_800 = datetime(2026, 5, 14, 8, 0)
    slot_810 = datetime(2026, 5, 14, 8, 10)
    slot_820 = datetime(2026, 5, 14, 8, 20)
    assert _eq_grid_slot_overlaps_event(
        slot_800, slot_800 + timedelta(minutes=sm), ev
    )
    assert not _eq_grid_slot_overlaps_event(
        slot_810, slot_810 + timedelta(minutes=sm), ev
    )
    assert not _eq_grid_slot_overlaps_event(
        slot_820, slot_820 + timedelta(minutes=sm), ev
    )
