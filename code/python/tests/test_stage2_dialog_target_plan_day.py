# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import date

from planning_core.core.roll_pipeline import (
    _first_working_day_strictly_after,
    _stage2_dialog_target_plan_day,
)


def test_dialog_target_plan_day_skip_today_uses_run_date():
    run = date(2026, 7, 11)
    working = [date(2026, 7, 10), date(2026, 7, 11), date(2026, 7, 14)]
    assert _stage2_dialog_target_plan_day(run, working, skip_today=True) == run


def test_dialog_target_plan_day_without_skip_uses_next_working_day():
    run = date(2026, 7, 11)
    working = [date(2026, 7, 10), date(2026, 7, 11), date(2026, 7, 14)]
    assert _stage2_dialog_target_plan_day(run, working, skip_today=False) == date(2026, 7, 14)
    assert _first_working_day_strictly_after(run, working) == date(2026, 7, 14)
