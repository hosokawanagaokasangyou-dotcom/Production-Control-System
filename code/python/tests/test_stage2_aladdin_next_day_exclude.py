# -*- coding: utf-8 -*-
"""段階2: アラジン当日・翌日除外 JSON と dispatch_loop 消費の最小テスト。"""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from planning_core import _core as pc
from planning_core.core.columns import ENV_STAGE2_ALADDIN_TODAY_EXCLUDE_NEXT_DAY_JSON
from planning_core.core.roll_pipeline import (
    append_plan_input_rows_missing_from_dispatch_table,
)


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
        "remaining_units": 4.0,
    }

    assert not pc._stage2_aladdin_next_day_exclude_consumes_roll(task, date(2026, 6, 9))
    assert task["aladdin_next_day_exclude_remaining_m"] == pytest.approx(6090.0)
    assert task["remaining_units"] == pytest.approx(4.0)

    assert pc._stage2_aladdin_next_day_exclude_consumes_roll(task, date(2026, 6, 10))
    assert task["aladdin_next_day_exclude_remaining_m"] == pytest.approx(3045.0)
    assert task["remaining_units"] == pytest.approx(3.0)

    assert pc._stage2_aladdin_next_day_exclude_consumes_roll(task, date(2026, 6, 10))
    assert task["aladdin_next_day_exclude_remaining_m"] == pytest.approx(0.0)
    assert task["remaining_units"] == pytest.approx(2.0)

    assert not pc._stage2_aladdin_next_day_exclude_consumes_roll(task, date(2026, 6, 10))
    assert task["remaining_units"] == pytest.approx(2.0)


def test_append_plan_missing_skips_stub_when_aladdin_exclude_covers_remaining(
    tmp_path: Path, monkeypatch
):
    """翌日配台0ロール（exclude=残量全量）のとき結果_配台表へ数量0スタブを載せない。"""
    path = tmp_path / "aladdin-exclude.json"
    path.write_text(
        json.dumps(
            {
                "version": 1,
                "entries": [
                    {
                        "task_id": "V9-1",
                        "process": "接続",
                        "machine_name": "熱融着機　湖南",
                        "exclude_next_day_m": 1000.0,
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv(ENV_STAGE2_ALADDIN_TODAY_EXCLUDE_NEXT_DAY_JSON, str(path))

    tasks_df = pd.DataFrame(
        [
            {
                "依頼NO": "V9-1",
                "工程名": "接続",
                "機械名": "熱融着機　湖南",
                "換算数量": 1000,
                "実加工数": 0,
                "未加工": 1000,
                "配台使用残数量": 1000,
            }
        ]
    )
    out = append_plan_input_rows_missing_from_dispatch_table(
        pd.DataFrame(), tasks_df, None
    )
    assert out is None or getattr(out, "empty", True)


def test_append_plan_missing_keeps_stub_when_aladdin_exclude_partial(
    tmp_path: Path, monkeypatch
):
    path = tmp_path / "aladdin-exclude-partial.json"
    path.write_text(
        json.dumps(
            {
                "version": 1,
                "entries": [
                    {
                        "task_id": "V9-1",
                        "process": "接続",
                        "machine_name": "熱融着機　湖南",
                        "exclude_next_day_m": 200.0,
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv(ENV_STAGE2_ALADDIN_TODAY_EXCLUDE_NEXT_DAY_JSON, str(path))

    tasks_df = pd.DataFrame(
        [
            {
                "依頼NO": "V9-1",
                "工程名": "接続",
                "機械名": "熱融着機　湖南",
                "換算数量": 1000,
                "実加工数": 0,
                "未加工": 1000,
                "配台使用残数量": 1000,
            }
        ]
    )
    out = append_plan_input_rows_missing_from_dispatch_table(
        pd.DataFrame(), tasks_df, None
    )
    assert out is not None and not out.empty
    assert len(out) == 1
    assert float(out.iloc[0]["当日配台数量"]) == pytest.approx(0.0)
