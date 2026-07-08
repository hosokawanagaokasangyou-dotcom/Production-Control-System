"""結果_配台表: 原反投入日は配台計画入力（タスク入力表）を加工計画DATAより優先する。"""

import pandas as pd

from planning_core.core.roll_pipeline import _dispatch_table_cell_from_sources


def test_raw_input_date_prefers_plan_input_over_processing_plan_data():
    plan_row = pd.Series(
        {"依頼NO": "C7-4", "工程名": "SEC", "原反投入日": "2025/7/8"}
    )
    src_row = pd.Series(
        {"依頼NO": "C7-4", "工程名": "SEC", "原反投入日": "2025/7/10"}
    )
    cell = _dispatch_table_cell_from_sources(
        src_row=src_row,
        plan_row=plan_row,
        task_dict=None,
        col_name="原反投入日",
    )
    assert str(cell).replace("-", "/") in ("2025/7/8", "2025/07/08")
