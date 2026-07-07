# -*- coding: utf-8 -*-
"""段階2: アラジン当日・翌日除外 JSON と dispatch_loop 消費の最小テスト。"""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pytest

from planning_core import _core as pc
from planning_core.core.columns import ENV_STAGE2_ALADDIN_TODAY_EXCLUDE_NEXT_DAY_JSON


@pytest.fixture
def _clear_apply_date():
    prev = pc._STAGE2_ALADDIN_EXCLUDE_APPLY_DATE
    pc._STAGE2_ALADDIN_EXCLUDE_APPLY_DATE = None
    yield
    pc._STAGE2_ALADDIN_EXCLUDE_APPLY_DATE = prev


def test_load_stage2_aladdin_today_exclude_next_day_overrides(tmp_path: Path, monkeypatch):
    path = tmp_path / "aladdin-exclude.json"
    path.write_text(
        json.dumps(
            {
                "version": 1,
                "entries": [
                    {
                        "task_id": "T1",
                        "process": "スリット",
                        "machine_name": "M1",
                        "exclude_next_day_m": 3045,
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv(ENV_STAGE2_ALADDIN_TODAY_EXCLUDE_NEXT_DAY_JSON, str(path))

    out = pc._load_stage2_aladdin_today_exclude_next_day_overrides()
    key = pc._stage2_in_progress_next_day_dispatch_key("T1", "スリット", "M1")
    assert key in out
    assert out[key] == pytest.approx(3045.0)


def test_aladdin_next_day_exclude_consumes_roll_on_apply_date_only(_clear_apply_date):
    pc._STAGE2_ALADDIN_EXCLUDE_APPLY_DATE = date(2026, 6, 10)
    task = {
        "task_id": "T1",
        "aladdin_today_exclude_next_day_dialog": True,
        "aladdin_next_day_exclude_remaining_m": 6090.0,
        "unit_m": 3045.0,
    }

    assert not pc._stage2_aladdin_next_day_exclude_consumes_roll(task, date(2026, 6, 9))
    assert task["aladdin_next_day_exclude_remaining_m"] == pytest.approx(6090.0)

    assert pc._stage2_aladdin_next_day_exclude_consumes_roll(task, date(2026, 6, 10))
    assert task["aladdin_next_day_exclude_remaining_m"] == pytest.approx(3045.0)

    assert pc._stage2_aladdin_next_day_exclude_consumes_roll(task, date(2026, 6, 10))
    assert task["aladdin_next_day_exclude_remaining_m"] == pytest.approx(0.0)

    assert not pc._stage2_aladdin_next_day_exclude_consumes_roll(task, date(2026, 6, 10))
