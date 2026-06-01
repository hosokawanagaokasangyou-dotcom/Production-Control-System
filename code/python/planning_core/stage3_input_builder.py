# -*- coding: utf-8 -*-
"""段階3.0 前処理: 結果_配台表.json（手動修正後）から枝番タスクを分解し入力3表を生成する。

- 入力: ``結果_配台表.json`` の各行（依頼NO×工程×機械名×配台日 と「当日配台数量」）
- 出力: ``plan_input_tasks.xlsx`` の第2シート（``PLAN_INPUT_STAGE3_SHEET_NAME``）に枝番行のみを書き出す

枝番行は元行（入力1表）を複製し、依頼NO=``{元依頼NO}-{枝番2桁}``、換算数量=セル数量、
原反投入日=配台日、配台可能日時=配台日+DISPATCHABLE_FROM_TIME（12:45）、
配台可能日時_上書き=移動先セル暦日+定常開始時刻（master 定常始業）とする。
元タスク行は入力3表に含めない。
特別ルール scope は列「元依頼NO」で親に紐付ける（配台 task_id は枝番依頼NOのまま）。
"""
from __future__ import annotations

import json
import math
import sys
from datetime import datetime
from pathlib import Path


def _norm(v) -> str:
    return str(v).strip() if v is not None else ""


def _aggregate_targets(pc, rows):
    """結果_配台表 rows を (依頼NO, 工程key, 機械名, 配台日) -> 当日配台数量(m) に集約。"""
    targets: dict[tuple, float] = {}
    for r in rows:
        if not isinstance(r, dict):
            continue
        tid = pc._interactive_norm_cell(r.get(pc.TASK_COL_TASK_ID)) or pc._interactive_norm_cell(
            r.get("タスクID")
        )
        proc = pc._interactive_dispatch_target_process_key(r.get(pc.TASK_COL_MACHINE))
        mach = pc._interactive_norm_cell(r.get(pc.TASK_COL_MACHINE_NAME))
        dd = pc._interactive_parse_dispatch_date_cell(r.get("配台日"))
        qcell = r.get("当日配台数量")
        try:
            qty = (
                float(str(qcell).replace(",", "").strip())
                if qcell not in (None, "")
                else 0.0
            )
        except (TypeError, ValueError):
            qty = 0.0
        if tid and mach and dd is not None and qty > 1e-9:
            key = (tid, proc, mach, dd)
            targets[key] = targets.get(key, 0.0) + qty
    return targets


def _write_sheet_preserving_others(pc, xlsx_path: Path, sheet: str, df) -> None:
    """指定シートだけ差し替え、他シートは維持する。

    ``pd.ExcelWriter(..., mode=\"a\")`` は xl/sharedStrings.xml 欠落ブックや
    2 回目の ``if_sheet_exists=\"replace\"`` で openpyxl が読めず落ちることがある。
    既存ブックは calamine で全シート読込 → openpyxl で全書き戻し（正規 OOXML）とする。
    """
    import os
    import shutil
    import tempfile

    import pandas as pd

    xlsx_path = Path(xlsx_path)
    target = str(sheet)[:31]

    if not xlsx_path.exists():
        with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="w") as w:
            df.to_excel(w, sheet_name=target, index=False)
        return

    pc.normalize_ooxml_shared_strings_if_missing(str(xlsx_path))

    xf = pd.ExcelFile(xlsx_path, engine="calamine")
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        replaced = False
        with pd.ExcelWriter(tmp, engine="openpyxl", mode="w") as w:
            for name in xf.sheet_names:
                safe = str(name)[:31] if name else "Sheet1"
                if safe == target or str(name).strip() == str(sheet).strip():
                    df.to_excel(w, sheet_name=target, index=False)
                    replaced = True
                else:
                    other = pd.read_excel(xf, sheet_name=name, header=0)
                    other.to_excel(w, sheet_name=safe, index=False)
            if not replaced:
                df.to_excel(w, sheet_name=target, index=False)
        shutil.move(tmp, str(xlsx_path))
    except Exception:
        if os.path.isfile(tmp):
            try:
                os.unlink(tmp)
            except OSError:
                pass
        raise


def build_stage3_input_sheet(
    result_dispatch_json_path,
    plan_input_xlsx_path,
    *,
    input1_sheet: str | None = None,
    stage3_sheet: str | None = None,
    master_path: str | None = None,
) -> dict:
    """結果_配台表.json から入力3表（枝番）を生成し ``plan_input_xlsx_path`` 第2シートへ書き出す。

    Returns: ``{"branch_rows": int, "output_path": str, "sheet": str}``
    """
    from planning_core import _core as pc
    import pandas as pd

    json_path = Path(result_dispatch_json_path).resolve()
    xlsx_path = Path(plan_input_xlsx_path).resolve()
    input1_sheet = input1_sheet or pc.PLAN_INPUT_SHEET_NAME
    stage3_sheet = stage3_sheet or pc.PLAN_INPUT_STAGE3_SHEET_NAME

    raw = json.loads(json_path.read_text(encoding="utf-8"))
    rows = raw.get("rows") if isinstance(raw, dict) else None
    if not rows:
        raise ValueError("結果_配台表.json に rows がありません。")

    df1 = pc.read_tabular_dataframe(str(xlsx_path), sheet_name=input1_sheet)
    df1.columns = df1.columns.str.strip()
    order1 = pc.plan_input_sheet_column_order()
    df1 = pc._align_dataframe_headers_to_canonical(df1, order1)
    for c in order1:
        if c not in df1.columns:
            df1[c] = ""

    # (依頼NO, 工程key, 機械名) と (依頼NO, 工程key) の2系統で元行を引く。
    plan_lookup: dict[tuple, object] = {}
    plan_lookup2: dict[tuple, object] = {}
    for _, prow in df1.iterrows():
        tid = _norm(pc.planning_task_id_str_from_plan_row(prow))
        proc = pc._interactive_dispatch_target_process_key(
            pc._planning_df_cell_scalar(prow, pc.TASK_COL_MACHINE)
        )
        mach = _norm(pc._planning_df_cell_scalar(prow, pc.TASK_COL_MACHINE_NAME))
        if tid and proc:
            plan_lookup.setdefault((tid, proc, mach), prow)
            plan_lookup2.setdefault((tid, proc), prow)

    targets = _aggregate_targets(pc, rows)

    shift_start = None
    try:
        mp = master_path or pc._master_workbook_path_resolved()
        st, _et = pc._read_master_main_regular_shift_times(mp)
        shift_start = st
    except Exception:
        shift_start = None
    if shift_start is None:
        shift_start = pc.DEFAULT_START_TIME

    out_order = pc.plan_input_stage3_sheet_column_order()
    records: list[dict] = []
    seq_by_parent: dict[str, int] = {}
    n_unmatched = 0

    for key in sorted(targets.keys(), key=lambda k: (k[0], k[3], k[1], k[2])):
        tid, proc, mach, dd = key
        qty = targets[key]
        prow = plan_lookup.get((tid, proc, mach))
        if prow is None:
            prow = plan_lookup2.get((tid, proc))
        if prow is None:
            n_unmatched += 1
            continue

        rec: dict = {c: "" for c in out_order}
        for c in order1:
            try:
                v = pc._planning_df_cell_scalar(prow, c)
            except Exception:
                v = None
            rec[c] = (
                "" if v is None or (isinstance(v, float) and pd.isna(v)) else v
            )

        seq = seq_by_parent.get(tid, 0) + 1
        seq_by_parent[tid] = seq
        branch_id = f"{tid}-{seq:02d}"
        rec[pc.PLAN_COL_PARENT_TASK_ID] = tid
        rec[pc.PLAN_COL_BRANCH_SEQ] = f"{seq:02d}"
        rec[pc.TASK_COL_TASK_ID] = branch_id

        # 全数未加工として、この枝番分（配台日セル数量）を配台する。
        rec[pc.TASK_COL_QTY] = float(qty)
        rec[pc.TASK_COL_UNPROCESSED] = float(qty)
        rec[pc.TASK_COL_ACTUAL_DONE] = 0
        unit = pc._roll_unit_m_estimate_from_plan_row(prow, qty)
        rec[pc.PLAN_COL_DISPATCH_REMAINING_QTY] = float(qty)
        rec[pc.PLAN_COL_DISPATCH_ROLL_COUNT] = (
            int(math.ceil(qty / unit)) if unit > 1e-9 else ""
        )

        rec[pc.TASK_COL_RAW_INPUT_DATE] = dd.strftime("%Y/%m/%d")
        rec[pc.PLAN_COL_RAW_INPUT_DATE_OVERRIDE] = ""
        rec[pc.PLAN_COL_DISPATCHABLE_DATETIME] = pc.format_dispatchable_datetime_cell(
            pc.compute_dispatchable_datetime(dd)
        )
        # 手動修正で移した配台日セル: 上書き列はその暦日 + 定常開始時刻（段階2 解決時は上書き優先）
        rec[pc.PLAN_COL_DISPATCHABLE_DATETIME_OVERRIDE] = pc.format_dispatchable_datetime_cell(
            datetime.combine(dd, shift_start)
        )
        rec[pc.RESULT_TASK_COL_DISPATCH_TRIAL_ORDER] = ""
        records.append(rec)

    out_df = pd.DataFrame(records)
    if out_df.empty:
        out_df = pd.DataFrame(columns=out_order)
    else:
        out_df = out_df.reindex(columns=out_order).fillna("")

    _write_sheet_preserving_others(pc, xlsx_path, stage3_sheet, out_df)

    if n_unmatched:
        print(
            f"[stage3-input] 入力1表に対応行が無く除外した目標: {n_unmatched} 件",
            flush=True,
        )
    print(
        f"[stage3-input] 入力3表を書き出し: {xlsx_path} シート「{stage3_sheet}」 枝番 {len(records)} 行",
        flush=True,
    )
    return {
        "branch_rows": int(len(records)),
        "output_path": str(xlsx_path),
        "sheet": stage3_sheet,
    }


def main(argv: list[str] | None = None) -> int:
    """CLI: argv[1]=結果_配台表.json（省略時 PM_AI_RESULT_DISPATCH_JSON）、argv[2]=入力xlsx（省略時 PM_AI_PLAN_INPUT_PATH）。"""
    import os

    argv = list(sys.argv if argv is None else argv)
    json_path = (
        argv[1]
        if len(argv) > 1
        else (os.environ.get("PM_AI_RESULT_DISPATCH_JSON") or "").strip()
    )
    xlsx_path = (
        argv[2]
        if len(argv) > 2
        else (os.environ.get("PM_AI_PLAN_INPUT_PATH") or "").strip()
    )
    if not json_path or not xlsx_path:
        print(
            "usage: stage3_input_builder <結果_配台表.json> <plan_input_tasks.xlsx>",
            flush=True,
        )
        return 2
    try:
        res = build_stage3_input_sheet(json_path, xlsx_path)
    except Exception as e:
        print(f"[stage3-input] 失敗: {e}", flush=True)
        import traceback

        traceback.print_exc()
        return 1
    print(json.dumps(res, ensure_ascii=False), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
