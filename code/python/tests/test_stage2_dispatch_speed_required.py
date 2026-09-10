# -*- coding: utf-8 -*-
"""段階2: 配台対象行の加工速度が空・非正なら PlanningValidationError。"""

from __future__ import annotations

from datetime import date

import pandas as pd
import pytest

from planning_core import _core as pc


def _row(**overrides):
    base = {
        pc.TASK_COL_TASK_ID: "C9-8",
        pc.TASK_COL_MACHINE: "SEC",
        pc.TASK_COL_MACHINE_NAME: "SEC機　湖南",
        pc.TASK_COL_QTY: 10000,
        pc.TASK_COL_UNPROCESSED: 10000,
        pc.TASK_COL_SPEED: 27,
        pc.PLAN_COL_RAW_ROLL_UNIT_LENGTH: 200,
        pc.TASK_COL_ANSWER_DUE: "2026-09-20",
    }
    base.update(overrides)
    return base


def test_positive_processing_speed_helper_rejects_empty_and_nonpositive():
    assert pc._positive_processing_speed_m_per_min(27) == 27.0
    assert pc._positive_processing_speed_m_per_min("27") == 27.0
    assert pc._positive_processing_speed_m_per_min("２７") == 27.0
    assert pc._positive_processing_speed_m_per_min(None) is None
    assert pc._positive_processing_speed_m_per_min("") is None
    assert pc._positive_processing_speed_m_per_min(0) is None
    assert pc._positive_processing_speed_m_per_min(-1) is None
    assert pc._positive_processing_speed_m_per_min("nan") is None


def test_build_task_queue_raises_when_dispatch_speed_empty():
    tasks = pd.DataFrame([_row(**{pc.TASK_COL_SPEED: ""})])
    with pytest.raises(pc.PlanningValidationError) as ei:
        pc.build_task_queue_from_planning_df(
            tasks,
            date(2026, 9, 7),
            {},
            ai_by_tid={},
            equipment_list=[],
        )
    msg = str(ei.value)
    assert "加工速度" in msg
    assert "C9-8" in msg
    assert "Excel行" in msg


def test_build_task_queue_raises_when_dispatch_speed_nonpositive():
    tasks = pd.DataFrame([_row(**{pc.TASK_COL_SPEED: 0})])
    with pytest.raises(pc.PlanningValidationError) as ei:
        pc.build_task_queue_from_planning_df(
            tasks,
            date(2026, 9, 7),
            {},
            ai_by_tid={},
            equipment_list=[],
        )
    assert "C9-8" in str(ei.value)


def test_build_task_queue_lists_multiple_invalid_speed_rows():
    tasks = pd.DataFrame(
        [
            _row(**{pc.TASK_COL_TASK_ID: "A-1", pc.TASK_COL_SPEED: None}),
            _row(**{pc.TASK_COL_TASK_ID: "A-2", pc.TASK_COL_SPEED: -5}),
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
    msg = str(ei.value)
    assert "A-1" in msg
    assert "A-2" in msg


def test_build_task_queue_excludes_配台不要_even_if_speed_empty():
    tasks = pd.DataFrame(
        [
            _row(
                **{
                    pc.TASK_COL_TASK_ID: "SKIP-1",
                    pc.TASK_COL_SPEED: "",
                    pc.PLAN_COL_EXCLUDE_FROM_ASSIGNMENT: "yes",
                }
            ),
            _row(**{pc.TASK_COL_TASK_ID: "OK-1", pc.TASK_COL_SPEED: 27}),
        ]
    )
    queue = pc.build_task_queue_from_planning_df(
        tasks,
        date(2026, 9, 7),
        {},
        ai_by_tid={},
        equipment_list=[],
    )
    assert len(queue) == 1
    assert queue[0]["task_id"] == "OK-1"


def test_build_task_queue_accepts_positive_speed():
    tasks = pd.DataFrame([_row()])
    queue = pc.build_task_queue_from_planning_df(
        tasks,
        date(2026, 9, 7),
        {},
        ai_by_tid={},
        equipment_list=[],
    )
    assert len(queue) == 1
    assert queue[0][pc.TASK_COL_SPEED] == 27.0


def test_build_task_queue_speed_override_rescues_empty_base_speed():
    tasks = pd.DataFrame(
        [
            _row(
                **{
                    pc.TASK_COL_SPEED: "",
                    pc.PLAN_COL_SPEED_OVERRIDE: 30,
                }
            )
        ]
    )
    queue = pc.build_task_queue_from_planning_df(
        tasks,
        date(2026, 9, 7),
        {},
        ai_by_tid={},
        equipment_list=[],
    )
    assert len(queue) == 1
    assert queue[0][pc.TASK_COL_SPEED] == 30.0
