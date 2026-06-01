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

    # 配台可能日時は配台日 + 定常開始時刻（master 無し時 DEFAULT_START_TIME）。上書き列は出力しない。
    assert pc.PLAN_COL_DISPATCHABLE_DATETIME_OVERRIDE not in out.columns
    assert pc.plan_reference_column_name(
        pc.PLAN_COL_DISPATCHABLE_DATETIME_OVERRIDE
    ) not in out.columns
    disp_vals = [str(v) for v in out[pc.PLAN_COL_DISPATCHABLE_DATETIME].tolist()]
    assert len(disp_vals) == 2
    start_hhmm = pc.DEFAULT_START_TIME.strftime("%H:%M")
    start_hhmm_alt = start_hhmm.lstrip("0")
    for v in disp_vals:
        assert start_hhmm in v or start_hhmm_alt in v
        assert "12:45" not in v
    assert any(v.startswith("2026/6/10") or v.startswith("2026/06/10") for v in disp_vals)
    assert any(v.startswith("2026/6/11") or v.startswith("2026/06/11") for v in disp_vals)

    # 特記が元行からコピーされる
    assert all("特記コピー確認" == str(v) for v in out[pc.PLAN_COL_SPECIAL_REMARK].tolist())

    # 配台試行順番は出力行順に 1..n
    orders = [int(v) for v in out[pc.RESULT_TASK_COL_DISPATCH_TRIAL_ORDER].tolist()]
    assert orders == [1, 2]


def test_build_twice_replaces_stage3_sheet(tmp_path):
    """2 回連続生成（初回成功→再生成）で openpyxl append が落ちないこと。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input1(xlsx)
    _result_json(rdj)

    res1 = builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
    res2 = builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
    assert res1["branch_rows"] == 2
    assert res2["branch_rows"] == 2

    out = pd.read_excel(xlsx, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME)
    out.columns = out.columns.str.strip()
    assert len(out) == 2


def _strip_shared_strings_xml(xlsx_path: Path) -> None:
    """段階1出力等で xl/sharedStrings.xml が欠ける OOXML を模倣する。"""
    import os
    import shutil
    import tempfile
    import zipfile

    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            if item.filename == "xl/sharedStrings.xml":
                continue
            zout.writestr(item, zin.read(item.filename))
    shutil.move(tmp, xlsx_path)


def test_build_twice_after_shared_strings_stripped(tmp_path):
    """初回書込後に sharedStrings 欠落でも 2 回目生成が成功すること。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input1(xlsx)
    _result_json(rdj)

    builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
    _strip_shared_strings_xml(xlsx)
    assert pc._ooxml_workbook_missing_shared_strings(str(xlsx))

    res2 = builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
    assert res2["branch_rows"] == 2
    # 2 回目後も openpyxl で読める（sharedStrings 欠落 KeyError の再発防止）
    pd.read_excel(
        xlsx, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME, engine="openpyxl"
    )


def test_stage3_column_order_excludes_dispatchable_override():
    order = pc.plan_input_stage3_sheet_column_order()
    assert pc.PLAN_COL_DISPATCHABLE_DATETIME in order
    assert pc.PLAN_COL_DISPATCHABLE_DATETIME_OVERRIDE not in order
    assert (
        pc.plan_reference_column_name(pc.PLAN_COL_DISPATCHABLE_DATETIME_OVERRIDE) not in order
    )


def test_missing_rows_raises(tmp_path):
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input1(xlsx)
    rdj.write_text(json.dumps({"rows": []}, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(ValueError):
        builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
