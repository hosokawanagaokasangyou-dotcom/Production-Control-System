# -*- coding: utf-8 -*-
"""段階3 枝番統合: 配台後の結果_配台表（枝番依頼NO）を元依頼NO単位へ集約する。

枝番タスクは配台・タイムライン上は枝番依頼NO（``{元依頼NO}-{枝番2桁}``）のまま割付くが、
成果物表示は元依頼NO単位へまとめる（プラン「特別ルール適用方針」L235: 統合は表示・成果物の集約のみ）。

集約キー: ``(元依頼NO, 工程名, 機械名)``
- 依頼NO        → 元依頼NO
- 加工開始日時   → 枝番行の最小
- 加工終了日時   → 枝番行の最大
- 配台日別数量 / 当日配台数量 / 換算数量 → 合算
- メンバー名     → 重複排除して全角空白連結
- その他静的列   → 先頭枝番行の値を踏襲（親属性は枝番間で同一）

枝番→親の対応は **入力3表シート**（列「依頼NO」「元依頼NO」）を正本とする。
依頼NO 自体が ``-数字`` で終わりうる（例 ``Y3-24``）ため、接尾辞の文字列剥がしは行わない。
"""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

_DATE_COL_RE = re.compile(r"^\d{4}/\d{1,2}/\d{1,2}$")
_DATETIME_FORMATS = (
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
)
# 配台日別の数量列に加えて常に合算する数量系の固定列。
_ALWAYS_SUM_COLUMNS = ("当日配台数量", "換算数量", "実加工数", "計画合計", "実出来高")


def _norm(v) -> str:
    return str(v).strip() if v is not None else ""


def _parse_dt(v):
    s = _norm(v)
    if not s:
        return None
    for fmt in _DATETIME_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _to_float(v):
    s = _norm(v).replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _fmt_qty(total: float):
    """合算結果を表示用にフォーマット（整数なら整数表記）。"""
    if abs(total - round(total)) < 1e-9:
        return int(round(total))
    return total


def load_branch_parent_map(plan_input_xlsx_path, *, stage3_sheet: str | None = None) -> dict:
    """入力3表シートから ``{枝番依頼NO: 元依頼NO}`` を読み出す。"""
    from planning_core import _core as pc

    sheet = stage3_sheet or pc.PLAN_INPUT_STAGE3_SHEET_NAME
    df = pc.read_tabular_dataframe(str(Path(plan_input_xlsx_path).resolve()), sheet_name=sheet)
    df.columns = df.columns.str.strip()
    out: dict[str, str] = {}
    for _, row in df.iterrows():
        branch = _norm(pc._planning_df_cell_scalar(row, pc.TASK_COL_TASK_ID))
        parent = _norm(pc._planning_df_cell_scalar(row, pc.PLAN_COL_PARENT_TASK_ID))
        if branch:
            out[branch] = parent or branch
    return out


def _sum_columns_for(columns, rows) -> set:
    """合算対象列（配台日列 + 固定数量列のうち存在するもの）。"""
    cols = set()
    for c in columns:
        cs = _norm(c)
        if _DATE_COL_RE.match(cs) or cs in _ALWAYS_SUM_COLUMNS:
            cols.add(c)
    return cols


def merge_branch_rows(rows, columns, branch_parent_map: dict) -> list:
    """結果_配台表 rows を元依頼NO単位へ集約した rows を返す。"""
    sum_cols = _sum_columns_for(columns, rows)
    groups: dict[tuple, dict] = {}
    order: list[tuple] = []

    for r in rows:
        if not isinstance(r, dict):
            continue
        branch_id = _norm(r.get("依頼NO"))
        parent_id = branch_parent_map.get(branch_id, branch_id)
        # 枝番でない（マップに無く親=自身）行も素通しで1グループ化される。
        key = (parent_id, _norm(r.get("工程名")), _norm(r.get("機械名")))
        if key not in groups:
            base = dict(r)
            base["依頼NO"] = parent_id
            groups[key] = {"base": base, "starts": [], "ends": [], "members": [], "sums": {}}
            order.append(key)
        g = groups[key]
        st = _parse_dt(r.get("加工開始日時"))
        if st is not None:
            g["starts"].append((st, _norm(r.get("加工開始日時"))))
        en = _parse_dt(r.get("加工終了日時"))
        if en is not None:
            g["ends"].append((en, _norm(r.get("加工終了日時"))))
        mem = _norm(r.get("メンバー名"))
        if mem:
            for part in re.split(r"[\u3000\s]+", mem):
                p = part.strip()
                if p and p not in g["members"]:
                    g["members"].append(p)
        for c in sum_cols:
            fv = _to_float(r.get(c))
            if fv is not None:
                g["sums"][c] = g["sums"].get(c, 0.0) + fv

    merged: list[dict] = []
    for key in order:
        g = groups[key]
        out = dict(g["base"])
        if g["starts"]:
            out["加工開始日時"] = min(g["starts"], key=lambda x: x[0])[1]
        if g["ends"]:
            out["加工終了日時"] = max(g["ends"], key=lambda x: x[0])[1]
        if g["members"]:
            out["メンバー名"] = "\u3000".join(g["members"])
        for c, total in g["sums"].items():
            out[c] = _fmt_qty(total)
        merged.append(out)
    return merged


def merge_branch_result_dispatch(
    result_dispatch_json_path,
    plan_input_xlsx_path,
    *,
    stage3_sheet: str | None = None,
    output_json_path=None,
) -> dict:
    """結果_配台表.json（枝番）を元依頼NO単位へ統合し JSON へ書き出す。

    Returns: ``{"merged_rows": int, "source_rows": int, "output_path": str}``
    """
    json_path = Path(result_dispatch_json_path).resolve()
    raw = json.loads(json_path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("結果_配台表.json の形式が不正です（dict ではありません）。")
    rows = raw.get("rows") or []
    columns = raw.get("columns") or (list(rows[0].keys()) if rows else [])

    branch_parent_map = load_branch_parent_map(plan_input_xlsx_path, stage3_sheet=stage3_sheet)
    merged_rows = merge_branch_rows(rows, columns, branch_parent_map)

    out = dict(raw)
    out["rows"] = merged_rows
    out["branch_merged"] = True

    out_path = Path(output_json_path).resolve() if output_json_path else json_path
    out_path.write_text(
        json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return {
        "merged_rows": int(len(merged_rows)),
        "source_rows": int(len(rows)),
        "output_path": str(out_path),
    }


def main(argv: list[str] | None = None) -> int:
    """CLI: argv[1]=結果_配台表.json、argv[2]=plan_input_tasks.xlsx、[argv[3]=出力json]。"""
    import os
    import sys

    argv = list(sys.argv if argv is None else argv)
    json_path = argv[1] if len(argv) > 1 else (os.environ.get("PM_AI_RESULT_DISPATCH_JSON") or "").strip()
    xlsx_path = argv[2] if len(argv) > 2 else (os.environ.get("PM_AI_PLAN_INPUT_PATH") or "").strip()
    out_path = argv[3] if len(argv) > 3 else None
    if not json_path or not xlsx_path:
        print("usage: stage3_branch_merge <結果_配台表.json> <plan_input_tasks.xlsx> [出力json]", flush=True)
        return 2
    try:
        res = merge_branch_result_dispatch(json_path, xlsx_path, output_json_path=out_path)
    except Exception as e:
        print(f"[stage3-merge] 失敗: {e}", flush=True)
        import traceback

        traceback.print_exc()
        return 1
    print(json.dumps(res, ensure_ascii=False), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
