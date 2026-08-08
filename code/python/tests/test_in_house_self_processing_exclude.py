# -*- coding: utf-8 -*-
"""自社加工品（依頼NO先頭が2）の配台不要自動設定。"""
import pandas as pd

from planning_core import _core as pc


def _sample_row(task_id: str, exclude: str = "") -> dict:
    order = pc.plan_input_sheet_column_order()
    rec = {c: "" for c in order}
    rec[pc.TASK_COL_TASK_ID] = task_id
    rec[pc.TASK_COL_MACHINE] = "SL"
    rec[pc.TASK_COL_MACHINE_NAME] = "スライス機1　湖南"
    rec[pc.PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = exclude
    return rec


def test_is_in_house_self_processing_task_id():
    assert pc._is_in_house_self_processing_task_id("2125-02-16")
    assert pc._is_in_house_self_processing_task_id(" 2125-03-27 ")
    assert not pc._is_in_house_self_processing_task_id("C8-11")
    assert not pc._is_in_house_self_processing_task_id("")
    assert not pc._is_in_house_self_processing_task_id(None)


def test_apply_auto_exclude_in_house_self_processing_sets_yes():
    order = pc.plan_input_sheet_column_order()
    df = pd.DataFrame(
        [
            _sample_row("2125-02-16"),
            _sample_row("C8-11"),
        ]
    ).reindex(columns=order).fillna("")
    pc._apply_auto_exclude_in_house_self_processing(df, log_prefix="test")
    assert df.at[0, pc.PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] == "yes"
    assert df.at[1, pc.PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] == ""


def test_apply_auto_exclude_in_house_self_processing_does_not_overwrite_manual():
    order = pc.plan_input_sheet_column_order()
    df = pd.DataFrame([_sample_row("2125-02-16", exclude="no")]).reindex(columns=order).fillna("")
    pc._apply_auto_exclude_in_house_self_processing(df, log_prefix="test")
    assert df.at[0, pc.PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] == "no"
