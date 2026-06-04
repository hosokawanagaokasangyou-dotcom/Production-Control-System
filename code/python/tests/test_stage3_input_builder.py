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
    rec[pc.TASK_COL_RAW_INPUT_DATE] = "2026/06/02"
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

    # 配台可能日時 = 各枝番の配台日 + 定常開始（入力1の原反日・12:45 は使わない）。
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

    # 原反投入日は入力1表の値（配台日で上書きしない）
    raw_dates = [str(v).strip() for v in out[pc.TASK_COL_RAW_INPUT_DATE].tolist()]
    assert all(d.startswith("2026/6/2") or d.startswith("2026/06/02") for d in raw_dates)
    assert not any(d.startswith("2026/6/10") or d.startswith("2026/06/10") for d in raw_dates)
    assert not any(d.startswith("2026/6/11") or d.startswith("2026/06/11") for d in raw_dates)

    # 配台試行順番は出力行順に 1..n
    orders = [int(v) for v in out[pc.RESULT_TASK_COL_DISPATCH_TRIAL_ORDER].tolist()]
    assert orders == [1, 2]


def test_order_stage3_records_by_parent_branch():
    recs = [
        {"元依頼NO": "V5-5", "配台枝番": "03", "依頼NO": "V5-5-03", "配台試行順番": 3},
        {"元依頼NO": "V5-5", "配台枝番": "01", "依頼NO": "V5-5-01", "配台試行順番": 1},
        {"元依頼NO": "Y3-24", "配台枝番": "02", "依頼NO": "Y3-24-02", "配台試行順番": 5},
        {"元依頼NO": "Y3-24", "配台枝番": "01", "依頼NO": "Y3-24-01", "配台試行順番": 4},
    ]
    ordered = builder._order_stage3_records_by_parent_branch(recs)
    assert [r["依頼NO"] for r in ordered] == [
        "V5-5-01",
        "V5-5-03",
        "Y3-24-01",
        "Y3-24-02",
    ]
    assert [int(r["配台試行順番"]) for r in ordered] == [1, 2, 3, 4]


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
    assert not any(pc._plan_column_is_original_reference(c) for c in order)


def test_plan_input_sheet_column_order_excludes_original_reference_columns():
    order = pc.plan_input_sheet_column_order()
    assert pc.PLAN_COL_PREFERRED_OP in order
    assert pc.PLAN_COL_SPECIAL_REMARK in order
    assert not any(pc._plan_column_is_original_reference(c) for c in order)


def _write_input1_v62(xlsx_path: Path) -> None:
    order = pc.plan_input_sheet_column_order()
    rec = {c: "" for c in order}
    rec[pc.TASK_COL_TASK_ID] = "V6-2"
    rec[pc.TASK_COL_MACHINE] = "分割"
    rec[pc.TASK_COL_MACHINE_NAME] = "スリット機1　湖南"
    rec[pc.TASK_COL_QTY] = 10000.0
    rec[pc.TASK_COL_UNPROCESSED] = 10000.0
    df = pd.DataFrame([rec]).reindex(columns=order).fillna("")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="w") as w:
        df.to_excel(w, sheet_name=pc.PLAN_INPUT_SHEET_NAME, index=False)


def test_baseline_sidecar_ignored_after_stage3_pipeline(tmp_path):
    """段階3.0 実行済み sidecar では baseline(段階2) を入力3目標に使わない。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input1_v62(xlsx)
    proc = "分割"
    rows = [
        {
            pc.TASK_COL_TASK_ID: "V6-2",
            pc.TASK_COL_MACHINE: proc,
            pc.TASK_COL_MACHINE_NAME: "スリット機1　湖南",
            "配台日": "2026-06-12",
            "当日配台数量": 7200,
            "実配台数量": 7200,
        },
        {
            pc.TASK_COL_TASK_ID: "V6-2",
            pc.TASK_COL_MACHINE: proc,
            pc.TASK_COL_MACHINE_NAME: "スリット機1　湖南",
            "配台日": "2026-06-15",
            "当日配台数量": 2800,
            "実配台数量": 2800,
        },
    ]
    rdj.write_text(json.dumps({"rows": rows}, ensure_ascii=False), encoding="utf-8")
    rdj.with_name(rdj.name + ".stage3_planning_meta.json").write_text(
        json.dumps(
            {
                "variant": "3.0",
                "baselineEntries": {
                    "V6-2\u0001スリット機1　湖南\u00012026-06-11": 4000.0,
                    "V6-2\u0001スリット機1　湖南\u00012026-06-12": 6000.0,
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    res = builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
    assert res["branch_rows"] == 2
    out = pd.read_excel(xlsx, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME)
    out.columns = out.columns.str.strip()
    assert len(out) == 2
    assert abs(float(out[pc.TASK_COL_QTY].sum()) - 10000.0) < 1e-6


def test_aladdin_collapsed_json_produces_single_branch_row(tmp_path):
    """アラジン整列後（JSON 1暦日・10000m）は入力3も枝番1行。sidecar の段階2複数日は使わない。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input1(xlsx)
    proc = "SL"
    # 1日化後の JSON（1暦日・10000m）
    rows = [
        {
            pc.TASK_COL_TASK_ID: "Y3-24",
            pc.TASK_COL_MACHINE: proc,
            pc.TASK_COL_MACHINE_NAME: "スライス機1　湖南",
            "配台日": "2026/06/10",
            "当日配台数量": "10000",
        },
    ]
    rdj.write_text(json.dumps({"rows": rows}, ensure_ascii=False), encoding="utf-8")
    sidecar = {
        "baselineEntries": {
            "Y3-24\u0001スライス機1　湖南\u00012026-06-10": 5000.0,
            "Y3-24\u0001スライス機1　湖南\u00012026-06-11": 5000.0,
        }
    }
    rdj.with_name(rdj.name + ".stage3_planning_meta.json").write_text(
        json.dumps(sidecar, ensure_ascii=False), encoding="utf-8"
    )

    res = builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
    assert res["branch_rows"] == 1

    out = pd.read_excel(xlsx, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME)
    out.columns = out.columns.str.strip()
    assert len(out) == 1
    assert float(out[pc.TASK_COL_QTY].iloc[0]) == 10000.0
    assert str(out[pc.TASK_COL_TASK_ID].iloc[0]) == "Y3-24-01"
    disp = str(out[pc.PLAN_COL_DISPATCHABLE_DATETIME].iloc[0])
    assert "2026/06/10" in disp or "2026/6/10" in disp
    assert "12:45" not in disp


def test_input3_dispatchable_is_json_dispatch_date_plus_shift_start(tmp_path):
    """アラジンで 6/12 に 10000m のとき配台可能日時は 6/12+定常開始（入力1の 6/11 12:45 を使わない）。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    order = pc.plan_input_sheet_column_order()
    rec = {c: "" for c in order}
    rec[pc.TASK_COL_TASK_ID] = "V6-2"
    rec[pc.TASK_COL_MACHINE] = "分割"
    rec[pc.TASK_COL_MACHINE_NAME] = "スリット機1　湖南"
    rec[pc.TASK_COL_RAW_INPUT_DATE] = "2026/06/11"
    rec[pc.PLAN_COL_DISPATCHABLE_DATETIME] = "2026/06/11 12:45"
    pd.DataFrame([rec]).reindex(columns=order).fillna("").to_excel(
        xlsx, sheet_name=pc.PLAN_INPUT_SHEET_NAME, index=False
    )
    rows = [
        {
            pc.TASK_COL_TASK_ID: "V6-2",
            pc.TASK_COL_MACHINE: "分割",
            pc.TASK_COL_MACHINE_NAME: "スリット機1　湖南",
            "配台日": "2026/06/12",
            "当日配台数量": "10000",
        },
    ]
    rdj.write_text(json.dumps({"rows": rows}, ensure_ascii=False), encoding="utf-8")
    res = builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
    assert res["branch_rows"] == 1
    out = pd.read_excel(xlsx, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME)
    out.columns = out.columns.str.strip()
    disp = str(out[pc.PLAN_COL_DISPATCHABLE_DATETIME].iloc[0])
    assert "2026/06/12" in disp or "2026/6/12" in disp
    assert "12:45" not in disp
    assert pc.DEFAULT_START_TIME.strftime("%H:%M") in disp


def test_collapse_stage3_two_branch_rows_to_one(tmp_path):
    """段階2由来の2枝番（暦日別数量）を段階3配台前に1行・総量へ統合する。"""
    order = pc.plan_input_stage3_sheet_column_order()
    rows = []
    for seq, (qty, disp) in enumerate(
        [(4000.0, "2026/06/11 08:25"), (6000.0, "2026/06/12 08:25")], start=1
    ):
        rec = {c: "" for c in order}
        rec[pc.PLAN_COL_PARENT_TASK_ID] = "V6-2"
        rec[pc.PLAN_COL_BRANCH_SEQ] = f"{seq:02d}"
        rec[pc.TASK_COL_TASK_ID] = f"V6-2-{seq:02d}"
        rec[pc.TASK_COL_MACHINE] = "分割"
        rec[pc.TASK_COL_MACHINE_NAME] = "スリット機1　湖南"
        rec[pc.TASK_COL_QTY] = qty
        rec[pc.TASK_COL_UNPROCESSED] = qty
        rec[pc.PLAN_COL_DISPATCH_REMAINING_QTY] = qty
        rec[pc.PLAN_COL_DISPATCHABLE_DATETIME] = disp
        rows.append(rec)
    df = pd.DataFrame(rows).reindex(columns=order).fillna("")
    collapsed = builder.collapse_stage3_plan_df_by_parent(pc, df)
    assert len(collapsed) == 1
    assert str(collapsed[pc.TASK_COL_TASK_ID].iloc[0]) == "V6-2-01"
    assert float(collapsed[pc.TASK_COL_QTY].iloc[0]) == 10000.0
    targets = builder.build_stage3_dispatch_targets_from_plan_df(pc, df)
    assert len(targets) == 1
    only_key = next(iter(targets))
    assert only_key[0] == "V6-2-01"
    assert targets[only_key] == 10000.0
    assert str(only_key[3]) == "2026-06-12"
    disp = str(collapsed[pc.PLAN_COL_DISPATCHABLE_DATETIME].iloc[0])
    assert "2026/06/12" in disp or "2026/6/12" in disp


def test_intentional_multi_day_branches_not_collapsed(tmp_path):
    """意図的な複数日枝番（各50%）は統合せずキャップも暦日別のまま。"""
    order = pc.plan_input_stage3_sheet_column_order()
    rows = []
    for seq, (qty, disp) in enumerate(
        [(5000.0, "2026/06/12 08:25"), (5000.0, "2026/06/15 08:25")], start=1
    ):
        rec = {c: "" for c in order}
        rec[pc.PLAN_COL_PARENT_TASK_ID] = "Y3-24"
        rec[pc.PLAN_COL_BRANCH_SEQ] = f"{seq:02d}"
        rec[pc.TASK_COL_TASK_ID] = f"Y3-24-{seq:02d}"
        rec[pc.TASK_COL_MACHINE] = "SL"
        rec[pc.TASK_COL_MACHINE_NAME] = "スライス機1　湖南"
        rec[pc.TASK_COL_QTY] = qty
        rec[pc.TASK_COL_UNPROCESSED] = qty
        rec[pc.PLAN_COL_DISPATCHABLE_DATETIME] = disp
        rows.append(rec)
    df = pd.DataFrame(rows).reindex(columns=order).fillna("")
    collapsed = builder.collapse_stage3_plan_df_by_parent(pc, df)
    assert len(collapsed) == 2
    targets = builder.build_stage3_dispatch_targets_from_plan_df(pc, df)
    assert len(targets) == 2
    assert sum(targets.values()) == 10000.0


def test_build_stage3_dispatch_targets_from_plan_df(tmp_path):
    from datetime import date, datetime

    xlsx = tmp_path / "plan_input_tasks.xlsx"
    order = pc.plan_input_stage3_sheet_column_order()
    recs = []
    for seq, iso, qty in (("01", date(2026, 6, 12), 5000), ("02", date(2026, 6, 15), 5000)):
        rec = {c: "" for c in order}
        rec[pc.TASK_COL_TASK_ID] = f"Y3-24-{seq}"
        rec[pc.PLAN_COL_PARENT_TASK_ID] = "Y3-24"
        rec[pc.PLAN_COL_BRANCH_SEQ] = seq
        rec[pc.TASK_COL_MACHINE] = "SL"
        rec[pc.TASK_COL_MACHINE_NAME] = "スライス機1　湖南"
        rec[pc.TASK_COL_QTY] = float(qty)
        rec[pc.TASK_COL_UNPROCESSED] = float(qty)
        rec[pc.PLAN_COL_DISPATCHABLE_DATETIME] = pc.format_dispatchable_datetime_cell(
            datetime.combine(iso, pc.DEFAULT_START_TIME)
        )
        recs.append(rec)
    df = pd.DataFrame(recs).reindex(columns=order).fillna("")
    with pd.ExcelWriter(xlsx, engine="openpyxl", mode="w") as w:
        df.to_excel(w, sheet_name=pc.PLAN_INPUT_SHEET_NAME, index=False)
        df.to_excel(w, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME, index=False)

    targets = builder.build_stage3_dispatch_targets_from_plan_df(pc, df)
    assert len(targets) == 2
    assert sum(targets.values()) == 10000.0


def test_write_baseline_sidecar_if_absent(tmp_path):
    rdj = tmp_path / "結果_配台表.json"
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
            "当日配台数量": "700",
        },
    ]
    rdj.write_text(json.dumps({"rows": rows}, ensure_ascii=False), encoding="utf-8")
    assert builder.write_stage3_baseline_sidecar_if_absent(rdj, rows) is True
    sidecar = rdj.with_name(rdj.name + ".stage3_planning_meta.json")
    assert sidecar.is_file()
    data = json.loads(sidecar.read_text(encoding="utf-8"))
    assert len(data.get("baselineEntries") or {}) == 2
    assert builder.write_stage3_baseline_sidecar_if_absent(rdj, rows) is False


def test_baseline_env_does_not_override_json_task(tmp_path, monkeypatch):
    """JSON に暦日目標があるタスクは env baseline（段階2複数日）より JSON を優先する。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input1(xlsx)
    proc = "SL"
    rows = [
        {
            pc.TASK_COL_TASK_ID: "Y3-24",
            pc.TASK_COL_MACHINE: proc,
            pc.TASK_COL_MACHINE_NAME: "スライス機1　湖南",
            "配台日": "2026/06/10",
            "当日配台数量": "10000",
        },
    ]
    rdj.write_text(json.dumps({"rows": rows}, ensure_ascii=False), encoding="utf-8")
    entries = {
        "Y3-24\u0001スライス機1　湖南\u00012026-06-10": 4000.0,
        "Y3-24\u0001スライス機1　湖南\u00012026-06-11": 6000.0,
    }
    monkeypatch.setenv(
        "PM_AI_STAGE3_BASELINE_ENTRIES_JSON", json.dumps(entries, ensure_ascii=False)
    )

    res = builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
    assert res["branch_rows"] == 1
    out = pd.read_excel(xlsx, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME)
    out.columns = out.columns.str.strip()
    assert float(out[pc.TASK_COL_QTY].iloc[0]) == 10000.0


def test_missing_rows_raises(tmp_path):
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input1(xlsx)
    rdj.write_text(json.dumps({"rows": []}, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(ValueError):
        builder.build_stage3_input_sheet(rdj, xlsx, master_path="")
