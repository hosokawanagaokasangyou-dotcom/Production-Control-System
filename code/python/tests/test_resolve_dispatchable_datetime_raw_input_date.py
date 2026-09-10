"""resolve_dispatchable_datetime_from_plan_row: 原反投入日が暦日の正。"""

from __future__ import annotations

from datetime import date, datetime, time

from planning_core.core.columns import (
    PLAN_COL_DISPATCHABLE_DATETIME,
    TASK_COL_RAW_INPUT_DATE,
    TASK_COL_STOCK_LOCATION,
)
from planning_core.core.gantt_excel import resolve_dispatchable_datetime_from_plan_row


def test_resolve_prefers_raw_input_date_when_dispatchable_column_is_stale_later():
    """W9-4 相当: 原反投入日を 9/7 にしても列が 9/9 12:45 のままなら 9/7 起算。"""
    row = {
        TASK_COL_RAW_INPUT_DATE: "2026/9/7",
        PLAN_COL_DISPATCHABLE_DATETIME: "2026/9/9 12:45",
        TASK_COL_STOCK_LOCATION: "K",
    }
    got = resolve_dispatchable_datetime_from_plan_row(row, run_date=date(2026, 9, 1))
    assert got == datetime(2026, 9, 7, 8, 45)


def test_resolve_stock_k_uses_845_even_when_column_still_1245(monkeypatch):
    """湖南工場・在庫場所K: 列が 12:45 でも配台可能は当日 8:45。"""
    monkeypatch.setenv("PM_AI_FACTORY_SITE", "KONAN")
    row = {
        TASK_COL_RAW_INPUT_DATE: "2026/9/7",
        PLAN_COL_DISPATCHABLE_DATETIME: "2026/9/7 12:45",
        TASK_COL_STOCK_LOCATION: "K",
    }
    got = resolve_dispatchable_datetime_from_plan_row(row, run_date=date(2026, 9, 1))
    assert got == datetime(2026, 9, 7, 8, 45)


def test_resolve_computed_time_for_konan_local_stock(monkeypatch):
    monkeypatch.setenv("PM_AI_FACTORY_SITE", "KONAN")
    row = {
        TASK_COL_RAW_INPUT_DATE: "2026/9/7",
        PLAN_COL_DISPATCHABLE_DATETIME: "2026/9/7 9:30",
        TASK_COL_STOCK_LOCATION: "湖南",
    }
    got = resolve_dispatchable_datetime_from_plan_row(row, run_date=date(2026, 9, 1))
    assert got == datetime(2026, 9, 7, 8, 45)


def test_resolve_column_only_when_no_raw_input_date():
    row = {
        TASK_COL_RAW_INPUT_DATE: "",
        PLAN_COL_DISPATCHABLE_DATETIME: "2026/9/9 12:45",
    }
    got = resolve_dispatchable_datetime_from_plan_row(row, run_date=date(2026, 9, 1))
    assert got == datetime(2026, 9, 9, 12, 45)


def test_resolve_enforces_raw_floor_when_column_earlier_than_raw():
    row = {
        TASK_COL_RAW_INPUT_DATE: "2026/9/9",
        PLAN_COL_DISPATCHABLE_DATETIME: "2026/9/7 12:45",
    }
    got = resolve_dispatchable_datetime_from_plan_row(row, run_date=date(2026, 9, 1))
    assert got == datetime(2026, 9, 9, 12, 45)
