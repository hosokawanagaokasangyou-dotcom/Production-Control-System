# -*- coding: utf-8 -*-
"""両面EC の配台時間倍化（task_queue）。"""

from __future__ import annotations

from datetime import date

import pandas as pd

from planning_core import _core as pc
from planning_core.core.ec_side_classification import EC_SIDE_CLASS_DOUBLE, EC_SIDE_CLASS_SINGLE


def _minimal_ec_plan_row(*, ec_side: str) -> dict:
    return {
        pc.TASK_COL_TASK_ID: "W1-1",
        pc.TASK_COL_MACHINE: "EC",
        pc.TASK_COL_MACHINE_NAME: "EC機　湖南",
        pc.TASK_COL_QTY: 1000,
        pc.TASK_COL_UNPROCESSED: 1000,
        pc.TASK_COL_SPEED: 10,
        pc.PLAN_COL_RAW_ROLL_UNIT_LENGTH: 100,
        pc.PLAN_COL_EC_SIDE_CLASS: ec_side,
        pc.TASK_COL_ANSWER_DUE: "2026-08-20",
    }


def test_build_task_queue_double_sided_ec_doubles_base_time_per_unit():
    tasks = pd.DataFrame([_minimal_ec_plan_row(ec_side=EC_SIDE_CLASS_DOUBLE)])
    queue = pc.build_task_queue_from_planning_df(
        tasks,
        date(2026, 8, 17),
        {},
        ai_by_tid={},
        equipment_list=[],
    )
    assert len(queue) == 1
    task = queue[0]
    assert task["ec_dispatch_pass_count"] == 2
    assert task["base_time_per_unit"] == 20.0


def test_build_task_queue_single_sided_ec_keeps_base_time_per_unit():
    tasks = pd.DataFrame([_minimal_ec_plan_row(ec_side=EC_SIDE_CLASS_SINGLE)])
    queue = pc.build_task_queue_from_planning_df(
        tasks,
        date(2026, 8, 17),
        {},
        ai_by_tid={},
        equipment_list=[],
    )
    assert len(queue) == 1
    task = queue[0]
    assert task["ec_dispatch_pass_count"] == 1
    assert task["base_time_per_unit"] == 10.0


def test_build_task_queue_double_sided_sec_keeps_base_time_per_unit():
    row = _minimal_ec_plan_row(ec_side=EC_SIDE_CLASS_DOUBLE)
    row[pc.TASK_COL_MACHINE] = "SEC"
    row[pc.TASK_COL_MACHINE_NAME] = "SEC機　湖南"
    tasks = pd.DataFrame([row])
    queue = pc.build_task_queue_from_planning_df(
        tasks,
        date(2026, 8, 17),
        {},
        ai_by_tid={},
        equipment_list=[],
    )
    assert len(queue) == 1
    task = queue[0]
    assert task["ec_dispatch_pass_count"] == 1
    assert task["base_time_per_unit"] == 10.0
