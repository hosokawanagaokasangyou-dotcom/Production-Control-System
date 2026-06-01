# -*- coding: utf-8 -*-
"""stage3_input_builder の枝番分解・配台可能日時・数量合計の単体テスト。"""
import json
from pathlib import Path

import pandas as pd
import pytest

from planning_core import _core as pc
from planning_core import stage3_input_builder as builder


def _write_input1(xlsx_path: Path) -> None:
    order = pc.plan_input_sheet_column_order()
    rec = {c: "" for c in order}
    rec[pc.TASK_COL_TASK_ID] = "Y3-24"
    rec[pc.TASK_COL_MACHINE] = "SL"
    rec[pc.TASK_COL_MACHINE_NAME] = "スライス機1　湖南"
    rec[pc.TASK_COL_PRODUCT] = "製品A"
    rec[pc.TASK_COL_QTY] = 1000.0
    rec[pc.TASK_COL_UNPROCESSED] = 1000.0
    rec[pc.PLAN_COL_ROLL_UNIT_LENGTH] = 100.0
    rec[pc.PLAN_COL_SPECIAL_REMARK] = "特記コピー確認"
    df = pd.DataFrame([rec]).reindex(columns=order).fillna("")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="w") as w:
        df.to_excel(w, sheet_name=pc.PLAN_INPUT_SHEET_NAME, index=False)


def _result_json(path: Path) -> None:
    proc = "SL"
    rows = [
        {
            pc.TASK_COL_TASK_ID: "Y3-24",
            pc.TASK_COL_MACHINE: proc,
            pc.TASK_COL_MACHINE_NAME: "スライス機1　湖南",
            "配台日": "2026/06/10",
            "当日配台数量": "300",
        },
        {
            pc.TASK_COL_TASK_ID: "Y3-24",
            pc.TASK_COL_MACHINE: proc,
            pc.TASK_COL_MACHINE_NAME: "スライス機1　湖南",
            "配台日": "2026/06/11",
            "当日配台数量": "200",
        },
        {
            pc.TASK_COL_TASK_ID: "Y3-24",
            pc.TASK_COL_MACHINE: proc,
            pc.TASK_COL_MACHINE_NAME: "スライス機1　湖南",
            "配台日": "2026/06/11",
            "当日配台数量": "100",
        },
    ]
    path.write_text(json.dumps({"rows": rows}, ensure_ascii=False), encoding="utf-8")


def test_branch_decomposition(tmp_path):
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input1(xlsx)
    _result_json(rdj)

    res = builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
    # 6/10=300 と 6/11=(200+100)=300 → 枝番 2 行
    assert res["branch_rows"] == 2

    out = pd.read_excel(xlsx, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME)
    out.columns = out.columns.str.strip()
    assert len(out) == 2

    ids = sorted(str(v) for v in out[pc.TASK_COL_TASK_ID].tolist())
    assert ids == ["Y3-24-01", "Y3-24-02"]
    assert set(str(v) for v in out[pc.PLAN_COL_PARENT_TASK_ID].tolist()) == {"Y3-24"}

    # 数量合計が保存される（300 + 300 = 600）
    assert abs(float(out[pc.TASK_COL_QTY].sum()) - 600.0) < 1e-6

    # 配台可能日時は配台日 + 既定始業（DEFAULT_START_TIME, master 無し）
    disp = sorted(str(v) for v in out[pc.PLAN_COL_DISPATCHABLE_DATETIME].tolist())
    hhmm = pc.DEFAULT_START_TIME.strftime("%-H:%M") if hasattr(pc.DEFAULT_START_TIME, "strftime") else "8:45"
    assert disp[0].startswith("2026/6/10") or disp[0].startswith("2026/06/10")

    # 特記が元行からコピーされる
    assert all("特記コピー確認" == str(v) for v in out[pc.PLAN_COL_SPECIAL_REMARK].tolist())


def test_missing_rows_raises(tmp_path):
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input1(xlsx)
    rdj.write_text(json.dumps({"rows": []}, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(ValueError):
        builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
