# -*- coding: utf-8 -*-
"""段階3.2 数量厳守モード（PM_AI_STAGE3_2_QTY_STRICT）の挙動テスト。

- 終業直前デファー・小残デファーが無効化される
- チーム終業上限が当日 23:59 まで拡張される
既定 off では従来挙動（デファー有効・終業拡張なし）であること。
"""
from datetime import datetime, time

import pytest

from planning_core import _core as pc


@pytest.fixture
def _clear_strict_env(monkeypatch):
    monkeypatch.delenv("PM_AI_STAGE3_2_QTY_STRICT", raising=False)
    yield


def _defer_call():
    team_end = datetime(2026, 6, 10, 17, 0)
    team_start = datetime(2026, 6, 10, 16, 30)  # 残 30 分 <= 既定 45 分
    task = {"task_id": "Y3-24-01", "machine": "SL", "remaining_units": 2}
    return pc._defer_team_start_past_prebreak_and_end_of_day(
        task,
        ("\u30c1\u30fc\u30e0A",),
        team_start,
        team_end,
        [],
        lambda x: x,
        min_contiguous_work_mins=None,
    )


def test_eod_defer_active_by_default(_clear_strict_env):
    # 既定: 終業直前・小残のため当日開始不可（None）
    assert _defer_call() is None


def test_eod_defer_disabled_under_qty_strict(monkeypatch):
    monkeypatch.setenv("PM_AI_STAGE3_2_QTY_STRICT", "1")
    # 数量厳守: デファーせず開始時刻を返す
    assert _defer_call() == datetime(2026, 6, 10, 16, 30)


def test_team_end_relax_only_under_qty_strict(monkeypatch):
    d = datetime(2026, 6, 10, 17, 0).date()
    base_end = datetime(2026, 6, 10, 17, 0)

    monkeypatch.delenv("PM_AI_STAGE3_2_QTY_STRICT", raising=False)
    assert pc._interactive_trial_relax_team_end_limit_to_eod(base_end, d) == base_end

    monkeypatch.setenv("PM_AI_STAGE3_2_QTY_STRICT", "1")
    relaxed = pc._interactive_trial_relax_team_end_limit_to_eod(base_end, d)
    assert relaxed == datetime.combine(d, time(23, 59))
