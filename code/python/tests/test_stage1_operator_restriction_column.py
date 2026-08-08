# -*- coding: utf-8 -*-
"""段階1「担当OP_限定」列契約のテスト。"""

from __future__ import annotations

import pandas as pd

from planning_core import _core as pc


def test_stage1_operator_restriction_column_contract():
    assert pc.PLAN_COL_LIMITED_OP == "担当OP_限定"

    stage1_order = pc.plan_input_sheet_column_order()
    assert "担当OP_指定" not in stage1_order
    assert pc.PLAN_COL_LIMITED_OP in stage1_order
    assert pc.PLAN_COL_LIMITED_OP not in pc.PLAN_STAGE1_MERGE_COLUMNS


def test_stage1_operator_restriction_is_initialized_blank():
    record = {pc.PLAN_COL_LIMITED_OP: "既存値"}

    pc._initialize_stage1_manual_input_columns(record)

    output = pd.DataFrame([record]).reindex(
        columns=pc.plan_input_sheet_column_order()
    ).fillna("")
    assert pc.PLAN_COL_LIMITED_OP in output.columns
    assert output.at[0, pc.PLAN_COL_LIMITED_OP] == ""


def test_stage1_rerun_does_not_inherit_operator_restriction(tmp_path, monkeypatch):
    old_path = tmp_path / "existing-plan.xlsx"
    old_df = pd.DataFrame(
        [
            {
                pc.TASK_COL_TASK_ID: "A-1",
                pc.TASK_COL_MACHINE: "SL",
                "担当OP_指定": "継承してはいけない旧OP",
                pc.PLAN_COL_LIMITED_OP: "継承してはいけないOP",
            }
        ]
    )
    old_df.to_excel(old_path, sheet_name=pc.PLAN_INPUT_SHEET_NAME, index=False)
    monkeypatch.setattr(pc, "_excel_plan_input_wb", lambda: str(old_path))

    current = pd.DataFrame(
        [
            {
                pc.TASK_COL_TASK_ID: "A-1",
                pc.TASK_COL_MACHINE: "SL",
                pc.PLAN_COL_LIMITED_OP: "",
            }
        ]
    )

    merged = pc._merge_plan_sheet_user_overrides(current)

    assert "担当OP_指定" not in merged.columns
    assert merged.at[0, pc.PLAN_COL_LIMITED_OP] == ""
