# -*- coding: utf-8 -*-
"""stage3_branch_merge の枝番→元依頼NO統合（min/max時刻・数量合算・メンバー連結）の単体テスト。"""
import json
from pathlib import Path

import pandas as pd

from planning_core import _core as pc
from planning_core import stage3_branch_merge as merge


def _write_input3(xlsx_path: Path) -> None:
    """枝番 2 行（Y3-24-01, Y3-24-02）→ 元依頼NO=Y3-24 の入力3表。"""
    order = pc.plan_input_stage3_sheet_column_order()
    recs = []
    for seq in ("01", "02"):
        rec = {c: "" for c in order}
        rec[pc.TASK_COL_TASK_ID] = f"Y3-24-{seq}"
        rec[pc.PLAN_COL_PARENT_TASK_ID] = "Y3-24"
        rec[pc.PLAN_COL_BRANCH_SEQ] = seq
        rec[pc.TASK_COL_MACHINE] = "SL"
        rec[pc.TASK_COL_MACHINE_NAME] = "スライス機1　湖南"
        recs.append(rec)
    df = pd.DataFrame(recs).reindex(columns=order).fillna("")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="w") as w:
        df.to_excel(w, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME, index=False)


def _result_json_branches(path: Path) -> None:
    columns = [
        "依頼NO",
        "工程名",
        "機械名",
        "加工開始日時",
        "加工終了日時",
        "メンバー名",
        "当日配台数量",
        "2026/06/10",
        "2026/06/11",
    ]
    rows = [
        {
            "依頼NO": "Y3-24-01",
            "工程名": "SL",
            "機械名": "スライス機1　湖南",
            "加工開始日時": "2026/06/10 09:00",
            "加工終了日時": "2026/06/10 12:00",
            "メンバー名": "山田　佐藤",
            "当日配台数量": "300",
            "2026/06/10": "300",
            "2026/06/11": "",
        },
        {
            "依頼NO": "Y3-24-02",
            "工程名": "SL",
            "機械名": "スライス機1　湖南",
            "加工開始日時": "2026/06/11 08:45",
            "加工終了日時": "2026/06/11 15:30",
            "メンバー名": "佐藤　鈴木",
            "当日配台数量": "300",
            "2026/06/10": "",
            "2026/06/11": "300",
        },
    ]
    path.write_text(
        json.dumps({"columns": columns, "rows": rows}, ensure_ascii=False), encoding="utf-8"
    )


def test_branch_merge(tmp_path):
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input3(xlsx)
    _result_json_branches(rdj)

    res = merge.merge_branch_result_dispatch(rdj, xlsx)
    assert res["source_rows"] == 2
    assert res["merged_rows"] == 1

    out = json.loads(rdj.read_text(encoding="utf-8"))
    assert out.get("branch_merged") is True
    row = out["rows"][0]

    # 依頼NO は元依頼NO に戻る
    assert row["依頼NO"] == "Y3-24"
    # 加工開始=min, 加工終了=max
    assert row["加工開始日時"] == "2026/06/10 09:00"
    assert row["加工終了日時"] == "2026/06/11 15:30"
    # 数量は合算
    assert float(row["当日配台数量"]) == 600.0
    assert float(row["2026/06/10"]) == 300.0
    assert float(row["2026/06/11"]) == 300.0
    # メンバーは重複排除して連結（山田・佐藤・鈴木）
    members = set(row["メンバー名"].split("\u3000"))
    assert members == {"山田", "佐藤", "鈴木"}


def test_branch_merge_does_not_sum_converted_qty(tmp_path):
    """枝番2行が親の換算数量(10000)を載せているとき、統合後も10000のまま。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input3(xlsx)
    columns = [
        "依頼NO",
        "工程名",
        "機械名",
        "換算数量",
        "当日配台数量",
        "加工開始日時",
        "加工終了日時",
    ]
    rows = [
        {
            "依頼NO": "Y3-24-01",
            "工程名": "SL",
            "機械名": "スライス機1　湖南",
            "換算数量": 10000,
            "当日配台数量": 5000,
            "加工開始日時": "2026/06/10 09:00",
            "加工終了日時": "2026/06/10 12:00",
        },
        {
            "依頼NO": "Y3-24-02",
            "工程名": "SL",
            "機械名": "スライス機1　湖南",
            "換算数量": 10000,
            "当日配台数量": 5000,
            "加工開始日時": "2026/06/11 08:45",
            "加工終了日時": "2026/06/11 15:30",
        },
    ]
    rdj.write_text(json.dumps({"columns": columns, "rows": rows}, ensure_ascii=False), encoding="utf-8")

    merge.merge_branch_result_dispatch(rdj, xlsx)
    out_rows = json.loads(rdj.read_text(encoding="utf-8"))["rows"]
    assert len(out_rows) == 2
    assert all(float(r["換算数量"]) == 10000.0 for r in out_rows)
    assert sum(float(r["当日配台数量"]) for r in out_rows) == 10000.0


def test_branch_merge_long_format_one_row_per_dispatch_date(tmp_path):
    """枝番2行・配台日列あり: 親依頼NOで暦日別2行、換算数量は10000のまま。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input3(xlsx)
    columns = [
        "依頼NO",
        "工程名",
        "機械名",
        "換算数量",
        "当日配台数量",
        "配台日",
        "加工開始日時",
        "加工終了日時",
    ]
    rows = [
        {
            "依頼NO": "Y3-24-01",
            "工程名": "SL",
            "機械名": "スライス機1　湖南",
            "換算数量": 10000,
            "当日配台数量": 4000,
            "配台日": "2026-06-10",
            "加工開始日時": "2026/06/10 09:00",
            "加工終了日時": "2026/06/10 12:00",
        },
        {
            "依頼NO": "Y3-24-02",
            "工程名": "SL",
            "機械名": "スライス機1　湖南",
            "換算数量": 10000,
            "当日配台数量": 6000,
            "配台日": "2026-06-11",
            "加工開始日時": "2026/06/11 08:45",
            "加工終了日時": "2026/06/11 15:30",
        },
    ]
    rdj.write_text(json.dumps({"columns": columns, "rows": rows}, ensure_ascii=False), encoding="utf-8")

    res = merge.merge_branch_result_dispatch(rdj, xlsx)
    assert res["merged_rows"] == 2
    out = json.loads(rdj.read_text(encoding="utf-8"))["rows"]
    assert all(r["依頼NO"] == "Y3-24" for r in out)
    assert all(float(r["換算数量"]) == 10000.0 for r in out)
    assert sum(float(r["当日配台数量"]) for r in out) == 10000.0


def test_parent_multi_calendar_rows_not_collapsed(tmp_path):
    """元依頼NO行（枝番でない暦日複数行）は統合せず行数を維持する。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input3(xlsx)
    columns = ["依頼NO", "工程名", "機械名", "換算数量", "当日配台数量", "配台日"]
    rows = [
        {
            "依頼NO": "V6-2",
            "工程名": "分割",
            "機械名": "スリット機1　湖南",
            "換算数量": 10000,
            "当日配台数量": 1000,
            "配台日": "2026-06-11",
        },
        {
            "依頼NO": "V6-2",
            "工程名": "分割",
            "機械名": "スリット機1　湖南",
            "換算数量": 10000,
            "当日配台数量": 9000,
            "配台日": "2026-06-16",
        },
    ]
    rdj.write_text(json.dumps({"columns": columns, "rows": rows}, ensure_ascii=False), encoding="utf-8")

    res = merge.merge_branch_result_dispatch(rdj, xlsx)
    assert res["merged_rows"] == 2
    out = json.loads(rdj.read_text(encoding="utf-8"))["rows"]
    assert sum(float(r["当日配台数量"]) for r in out) == 10000.0
    assert all(float(r["換算数量"]) == 10000.0 for r in out)


def test_stale_parent_row_skipped_when_branches_exist(tmp_path):
    """アラジン整列後の親行＋枝番行が共存しても、親行は二重計上しない。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    order = pc.plan_input_stage3_sheet_column_order()
    rec = {c: "" for c in order}
    rec[pc.TASK_COL_TASK_ID] = "V6-2-01"
    rec[pc.PLAN_COL_PARENT_TASK_ID] = "V6-2"
    rec[pc.PLAN_COL_BRANCH_SEQ] = "01"
    rec[pc.TASK_COL_MACHINE] = "分割"
    rec[pc.TASK_COL_MACHINE_NAME] = "スリット機1　湖南"
    pd.DataFrame([rec]).reindex(columns=order).fillna("").to_excel(
        xlsx, sheet_name=pc.PLAN_INPUT_STAGE3_SHEET_NAME, index=False
    )
    columns = ["依頼NO", "工程名", "機械名", "配台日", "当日配台数量", "換算数量"]
    rows = [
        {
            "依頼NO": "V6-2",
            "工程名": "分割",
            "機械名": "スリット機1　湖南",
            "配台日": "2026-06-12",
            "当日配台数量": 10000,
            "換算数量": 10000,
        },
        {
            "依頼NO": "V6-2-01",
            "工程名": "分割",
            "機械名": "スリット機1　湖南",
            "配台日": "2026-06-17",
            "当日配台数量": 4600,
            "換算数量": 10000,
        },
        {
            "依頼NO": "V6-2-01",
            "工程名": "分割",
            "機械名": "スリット機1　湖南",
            "配台日": "2026-06-18",
            "当日配台数量": 5400,
            "換算数量": 10000,
        },
    ]
    rdj.write_text(json.dumps({"columns": columns, "rows": rows}, ensure_ascii=False), encoding="utf-8")

    res = merge.merge_branch_result_dispatch(rdj, xlsx)
    out = json.loads(rdj.read_text(encoding="utf-8"))["rows"]
    v62 = [r for r in out if str(r.get("依頼NO", "")).startswith("V6-2")]
    assert all(r["依頼NO"] == "V6-2" for r in v62)
    assert len(v62) == 2
    assert sum(float(r["当日配台数量"]) for r in v62) == 10000.0
    assert res["merged_rows"] == 2


def test_non_branch_passthrough(tmp_path):
    """マップに無い依頼NO（非枝番）も素通しで1行のまま。"""
    xlsx = tmp_path / "plan_input_tasks.xlsx"
    rdj = tmp_path / "結果_配台表.json"
    _write_input3(xlsx)
    columns = ["依頼NO", "工程名", "機械名", "当日配台数量"]
    rows = [{"依頼NO": "Z9-99", "工程名": "EC", "機械名": "EC1", "当日配台数量": "50"}]
    rdj.write_text(json.dumps({"columns": columns, "rows": rows}, ensure_ascii=False), encoding="utf-8")

    res = merge.merge_branch_result_dispatch(rdj, xlsx)
    assert res["merged_rows"] == 1
    out = json.loads(rdj.read_text(encoding="utf-8"))
    assert out["rows"][0]["依頼NO"] == "Z9-99"
