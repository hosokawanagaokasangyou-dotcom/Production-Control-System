"""結果_配台表: 回答納期は加工計画DATA のみ。無ければ空欄（配台計画入力へフォールバックしない）。"""

import pandas as pd

from planning_core.core.roll_pipeline import _dispatch_table_cell_from_sources


def test_answer_due_blank_when_missing_from_processing_plan_data():
    plan_row = pd.Series(
        {"依頼NO": "C7-4", "工程名": "SEC", "回答納期": "2026/07/18"}
    )
    src_row = pd.Series({"依頼NO": "C7-4", "工程名": "SEC", "回答納期": ""})
    cell = _dispatch_table_cell_from_sources(
        src_row=src_row,
        plan_row=plan_row,
        task_dict={"answer_due_date": "2026/07/20"},
        col_name="回答納期",
    )
    assert cell == ""


def test_answer_due_from_processing_plan_data_when_present():
    plan_row = pd.Series(
        {"依頼NO": "C7-4", "工程名": "SEC", "回答納期": "2026/07/18"}
    )
    src_row = pd.Series(
        {"依頼NO": "C7-4", "工程名": "SEC", "回答納期": "2026/07/22"}
    )
    cell = _dispatch_table_cell_from_sources(
        src_row=src_row,
        plan_row=plan_row,
        task_dict=None,
        col_name="回答納期",
    )
    assert str(cell).replace("-", "/") in ("2026/07/22", "2026/7/22")
