# -*- coding: utf-8 -*-
"""湖南工場・在庫場所「湖南」タスクの投入日早期配台開始（9:30）の最小テスト。"""

from __future__ import annotations

from datetime import date, time

import pytest

from planning_core.core.columns import (
    DISPATCHABLE_FROM_TIME,
    DISPATCHABLE_FROM_TIME_KONAN_STOCK,
    dispatchable_from_time_for,
)
from planning_core.core.gantt_excel import compute_dispatchable_datetime


def test_default_konan_stock_time_is_9_30():
    assert DISPATCHABLE_FROM_TIME_KONAN_STOCK == time(9, 30)


@pytest.mark.parametrize(
    "stock_location",
    ["湖南", "湖南/中央", "  湖南  "],
)
def test_konan_factory_konan_stock_uses_early_time(monkeypatch, stock_location):
    monkeypatch.setenv("PM_AI_FACTORY_SITE", "KONAN")
    assert dispatchable_from_time_for(stock_location) == DISPATCHABLE_FROM_TIME_KONAN_STOCK


def test_konan_factory_other_stock_uses_default_time(monkeypatch):
    monkeypatch.setenv("PM_AI_FACTORY_SITE", "KONAN")
    assert dispatchable_from_time_for("滋賀") == DISPATCHABLE_FROM_TIME
    assert dispatchable_from_time_for(None) == DISPATCHABLE_FROM_TIME
    assert dispatchable_from_time_for("") == DISPATCHABLE_FROM_TIME


def test_kokubu_factory_konan_stock_still_uses_default_time(monkeypatch):
    monkeypatch.setenv("PM_AI_FACTORY_SITE", "KOKUBU")
    assert dispatchable_from_time_for("湖南") == DISPATCHABLE_FROM_TIME


def test_unset_factory_site_defaults_to_konan(monkeypatch):
    monkeypatch.delenv("PM_AI_FACTORY_SITE", raising=False)
    assert dispatchable_from_time_for("湖南") == DISPATCHABLE_FROM_TIME_KONAN_STOCK


def test_compute_dispatchable_datetime_applies_konan_stock_override(monkeypatch):
    monkeypatch.setenv("PM_AI_FACTORY_SITE", "KONAN")
    raw_input_date = date(2026, 7, 9)

    dt_konan = compute_dispatchable_datetime(raw_input_date, stock_location="湖南")
    assert dt_konan is not None
    assert dt_konan.date() == raw_input_date
    assert dt_konan.time() == time(9, 30)

    dt_other = compute_dispatchable_datetime(raw_input_date, stock_location="滋賀")
    assert dt_other.time() == time(12, 45)


def test_compute_dispatchable_datetime_none_when_no_raw_input_date(monkeypatch):
    monkeypatch.setenv("PM_AI_FACTORY_SITE", "KONAN")
    assert compute_dispatchable_datetime(None, stock_location="湖南") is None
