# -*- coding: utf-8 -*-
"""実績ガント: 工程名がマスタ不一致でスキップした依頼NOをログに出す。"""
from __future__ import annotations

import logging
from datetime import date

import pandas as pd

from planning_core import _core as pc


def test_equipment_mismatch_warning_includes_task_ids_and_process(caplog):
    df = pd.DataFrame(
        [
            {pc.ACT_COL_TASK_ID: "W7-99", pc.ACT_COL_PROCESS: "幽霊工程A"},
            {pc.ACT_COL_TASK_ID: "A4-3", pc.ACT_COL_PROCESS: "幽霊工程B"},
            {pc.ACT_COL_TASK_ID: "W8-1", pc.ACT_COL_PROCESS: "熱融着機 湖南"},
        ]
    )
    with caplog.at_level(logging.WARNING):
        events = pc.build_actual_timeline_events(
            df,
            ["熱融着機 湖南"],
            [date(2026, 8, 20)],
            log_sheet_name="先頭シート",
        )

    joined = "\n".join(r.getMessage() for r in caplog.records)
    assert "W7-99" in joined
    assert "A4-3" in joined
    assert "幽霊工程A" in joined
    assert "幽霊工程B" in joined
    assert "ガント非表示" in joined
    assert "W8-1" not in joined
    assert events == []


def test_equipment_mismatch_warning_includes_roll_when_detail(caplog):
    df = pd.DataFrame(
        [
            {
                pc.ACT_COL_TASK_ID: "W7-10",
                pc.ACT_COL_PROCESS: "不明工程",
                pc.ACT_DETAIL_COL_ROLL: "R3",
            },
        ]
    )
    with caplog.at_level(logging.WARNING):
        pc.build_actual_timeline_events(
            df,
            ["熱融着機 湖南"],
            [date(2026, 8, 20)],
            log_sheet_name="先頭シート",
            roll_detail=True,
        )

    joined = "\n".join(r.getMessage() for r in caplog.records)
    assert "W7-10/R3" in joined
    assert "不明工程" in joined
