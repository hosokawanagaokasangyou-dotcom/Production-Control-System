# -*- coding: utf-8 -*-
"""段階2: 計画シート末尾の空白行で未加工検証が落ちないこと。"""

from __future__ import annotations

from datetime import date

import pandas as pd

from planning_core import _core as pc


def _valid_row(**overrides):
    base = {
        pc.TASK_COL_TASK_ID: "C9-8",
        pc.TASK_COL_MACHINE: "SEC",
        pc.TASK_COL_MACHINE_NAME: "SEC機　湖南",
        pc.TASK_COL_QTY: 10000,
        pc.TASK_COL_UNPROCESSED: 10000,
        pc.TASK_COL_ACTUAL_DONE: 0,
        pc.TASK_COL_SPEED: 27,
        pc.PLAN_COL_RAW_ROLL_UNIT_LENGTH: 200,
        pc.TASK_COL_ANSWER_DUE: "2026-09-20",
    }
    base.update(overrides)
    return base


def test_build_task_queue_skips_trailing_blank_rows_without_unprocessed_error():
    """Excel 余白行（依頼NO・工程空・未加工 NaN）があっても exit=3 にしない。"""
    blank = {
        pc.TASK_COL_TASK_ID: None,
        pc.TASK_COL_MACHINE: None,
        pc.TASK_COL_MACHINE_NAME: None,
        pc.TASK_COL_QTY: None,
        pc.TASK_COL_UNPROCESSED: None,
        pc.TASK_COL_ACTUAL_DONE: None,
        pc.TASK_COL_SPEED: None,
    }
    tasks = pd.DataFrame([_valid_row(), blank, blank])
    queue = pc.build_task_queue_from_planning_df(
        tasks,
        date(2026, 9, 7),
        {},
        ai_by_tid={},
        equipment_list=[],
    )
    assert len(queue) == 1
    assert queue[0]["task_id"] == "C9-8"


def test_build_task_queue_still_requires_unprocessed_on_real_dispatch_rows():
    """実データ行で未加工が空なら従来どおり PlanningValidationError。"""
    import pytest

    tasks = pd.DataFrame(
        [
            _valid_row(
                **{
                    pc.TASK_COL_UNPROCESSED: None,
                }
            )
        ]
    )
    with pytest.raises(pc.PlanningValidationError) as ei:
        pc.build_task_queue_from_planning_df(
            tasks,
            date(2026, 9, 7),
            {},
            ai_by_tid={},
            equipment_list=[],
        )
    assert "未加工" in str(ei.value)
