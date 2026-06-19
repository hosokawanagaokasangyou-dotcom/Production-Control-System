# -*- coding: utf-8 -*-
"""段階3 枝番統合: 配台後の結果_配台表（枝番依頼NO）を元依頼NO単位へ集約する。

枝番タスクは配台・タイムライン上は枝番依頼NO（``{元依頼NO}-{枝番2桁}``）のまま割付くが、
成果物表示は元依頼NO単位へまとめる（プラン「特別ルール適用方針」L235: 統合は表示・成果物の集約のみ）。

集約キー: ``(元依頼NO, 工程名, 機械名)``
- 依頼NO        → 元依頼NO
- **長形式（``配台日`` 列あり）**: 暦日ごとに1行（``当日配台数量`` はその日のみ合算）
- **ワイド形式（``配台日`` 無し・``yyyy/M/d`` 列あり）**: 1行に暦日列を合算
- 換算数量 → 枝番（同一 工程×機械 グループ）から推定。分割形式は合算、旧「親全量載せ」形式は単一値（_infer_parent_converted_qty）
- メンバー名     → 重複排除して全角空白連結（グループ全体）
- 枝番でない行（依頼NO＝元依頼NO）→ **統合せず**そのまま出力

枝番→親の対応は **入力3表シート**（列「依頼NO」「元依頼NO」）を正本とする。
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

_DATE_COL_RE = re.compile(r"^\d{4}/\d{1,2}/\d{1,2}$")
_DATETIME_FORMATS = (
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
)
# 配台 m のみ合算。換算数量等は各行が親タスク全量を載せるため合算すると二重計上になる。
_ALWAYS_SUM_COLUMNS = ("当日配台数量", "実配台数量")


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


def _normalize_dispatch_date_cell(v) -> str:
    """配台日セルを ISO 日付キー (YYYY-MM-DD) に正規化。空なら空文字。"""
    s = _norm(v)
    if not s:
        return ""
    if "T" in s:
        s = s.split("T", 1)[0]
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        return s[:10]
    for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(s[:10], fmt).date().isoformat()
        except ValueError:
            continue
    return s


def _dispatch_date_key(row: dict, sum_cols: set, *, has_dispatch_date_col: bool) -> str:
    """行の暦日キー。長形式は配台日、無いときは非ゼロ暦日列→加工開始日。"""
    if has_dispatch_date_col:
        dk = _normalize_dispatch_date_cell(row.get("配台日"))
        if dk:
            return dk
    for c in sorted(sum_cols):
        if not _DATE_COL_RE.match(_norm(c)):
            continue
        fv = _to_float(row.get(c))
        if fv is not None and fv > 1e-9:
            return _normalize_dispatch_date_cell(c.replace("/", "-"))
    st = _parse_dt(row.get("加工開始日時"))
    if st is not None:
        return st.date().isoformat()
    return ""


def _row_dispatch_m(row: dict) -> float:
    for col in _ALWAYS_SUM_COLUMNS:
        fv = _to_float(row.get(col))
        if fv is not None and fv > 1e-9:
            return fv
    return 0.0


def _branch_converted_qty_by_id(branch_rows: list[dict]) -> dict[str, float]:
    out: dict[str, float] = {}
    for r in branch_rows:
        bid = _norm(r.get("依頼NO"))
        if not bid:
            continue
        qty = _to_float(r.get("換算数量"))
        if qty is not None and qty > 1e-9:
            out[bid] = qty
    return out


def _branch_dispatch_totals(branch_rows: list[dict]) -> dict[str, float]:
    out: dict[str, float] = {}
    for r in branch_rows:
        bid = _norm(r.get("依頼NO"))
        if not bid:
            continue
        out[bid] = out.get(bid, 0.0) + _row_dispatch_m(r)
    return out


def _infer_parent_converted_qty(branch_rows: list[dict]) -> float | None:
    """枝番行の換算数量から元依頼NO単位の換算数量を推定。

    枝番（同一 工程×機械 グループ）の換算数量は次の2形式があり得る:

    - **分割形式（段階3 input3 builder）**: 各枝番が「残量（換算−実加工数）の暦日分割」を載せる。
      例 C6-6: 4000 + 6000。元依頼NO の換算数量＝**合算**。
    - **親全量形式（旧）**: 各枝番が親の総換算数量を載せる。例 10000 + 10000。
      合算すると二重計上になるため**単一値**。

    判定: 全枝番が同値かつ合算が配台 m 合計を超えるときのみ旧形式（単一値）。
    それ以外は合算（分割形式・単一枝番ともに正しい値になる）。
    """
    by_branch = _branch_converted_qty_by_id(branch_rows)
    if not by_branch:
        return None
    values = list(by_branch.values())
    if len(values) == 1:
        return values[0]
    unique_sum = sum(values)
    if len(set(values)) == 1:
        total_dispatch = sum(_branch_dispatch_totals(branch_rows).values())
        single = values[0]
        if unique_sum > total_dispatch + 1e-3:
            return single
        return unique_sum
    return unique_sum


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


def _columns_have_dispatch_date(columns) -> bool:
    return any(_norm(c) == "配台日" for c in columns)


def _columns_have_calendar_date_columns(columns) -> bool:
    return any(_DATE_COL_RE.match(_norm(c)) for c in columns)


def _build_merged_row_for_day(
    parent_id: str,
    day_rows: list[dict],
    sum_cols: set,
    *,
    has_dispatch_date_col: bool,
    group_members: list[str],
) -> dict:
    base = dict(day_rows[0])
    base["依頼NO"] = parent_id
    if has_dispatch_date_col and day_rows:
        dk = _dispatch_date_key(day_rows[0], sum_cols, has_dispatch_date_col=True)
        if dk:
            base["配台日"] = dk

    starts: list[tuple] = []
    ends: list[tuple] = []
    sums: dict[str, float] = {}
    for r in day_rows:
        st = _parse_dt(r.get("加工開始日時"))
        if st is not None:
            starts.append((st, _norm(r.get("加工開始日時"))))
        en = _parse_dt(r.get("加工終了日時"))
        if en is not None:
            ends.append((en, _norm(r.get("加工終了日時"))))
        for c in sum_cols:
            fv = _to_float(r.get(c))
            if fv is not None:
                sums[c] = sums.get(c, 0.0) + fv

    if starts:
        base["加工開始日時"] = min(starts, key=lambda x: x[0])[1]
    if ends:
        base["加工終了日時"] = max(ends, key=lambda x: x[0])[1]
    if group_members:
        base["メンバー名"] = "\u3000".join(group_members)
    for c, total in sums.items():
        base[c] = _fmt_qty(total)
    return base


def _apply_parent_converted_qty(rows: list[dict], parent_qty: float | None) -> list[dict]:
    if parent_qty is None or parent_qty <= 1e-9:
        return rows
    for row in rows:
        row["換算数量"] = _fmt_qty(parent_qty)
    return rows


def _flush_branch_group(
    groups: dict,
    key: tuple,
    sum_cols: set,
    *,
    has_dispatch_date_col: bool,
    wide_date_columns: bool,
) -> list[dict]:
    g = groups.pop(key, None)
    if not g:
        return []
    day_rows: list[dict] = g.get("rows") or []
    if not day_rows:
        return []

    parent_id = key[0]
    parent_qty = _infer_parent_converted_qty(day_rows)

    per_day_long = has_dispatch_date_col or not wide_date_columns
    if not per_day_long:
        return _apply_parent_converted_qty(
            [
                _build_merged_row_for_day(
                    parent_id,
                    day_rows,
                    sum_cols,
                    has_dispatch_date_col=has_dispatch_date_col,
                    group_members=g.get("members") or [],
                )
            ],
            parent_qty,
        )

    by_day: dict[str, list[dict]] = defaultdict(list)
    for r in day_rows:
        dk = _dispatch_date_key(r, sum_cols, has_dispatch_date_col=has_dispatch_date_col)
        by_day[dk or "__unknown__"].append(r)

    out: list[dict] = []
    for dk in sorted(by_day.keys()):
        rows_for_day = by_day[dk]
        out.append(
            _build_merged_row_for_day(
                parent_id,
                rows_for_day,
                sum_cols,
                has_dispatch_date_col=has_dispatch_date_col,
                group_members=g.get("members") or [],
            )
        )
    return _apply_parent_converted_qty(out, parent_qty)


def merge_branch_rows(rows, columns, branch_parent_map: dict) -> list:
    """結果_配台表 rows を元依頼NO単位へ集約した rows を返す。"""
    sum_cols = _sum_columns_for(columns, rows)
    has_dispatch_date_col = _columns_have_dispatch_date(columns)
    wide_date_columns = _columns_have_calendar_date_columns(columns)
    groups: dict[tuple, dict] = {}
    open_keys: list[tuple] = []
    merged: list[dict] = []

    branch_group_keys: set[tuple] = set()
    for r in rows:
        if not isinstance(r, dict):
            continue
        branch_id = _norm(r.get("依頼NO"))
        parent_id = branch_parent_map.get(branch_id, branch_id)
        if branch_id != parent_id:
            branch_group_keys.add(
                (parent_id, _norm(r.get("工程名")), _norm(r.get("機械名")))
            )

    for r in rows:
        if not isinstance(r, dict):
            continue
        branch_id = _norm(r.get("依頼NO"))
        parent_id = branch_parent_map.get(branch_id, branch_id)
        if branch_id == parent_id:
            parent_key = (parent_id, _norm(r.get("工程名")), _norm(r.get("機械名")))
            if parent_key in branch_group_keys:
                continue
            for key in list(open_keys):
                merged.extend(
                    _flush_branch_group(
                        groups,
                        key,
                        sum_cols,
                        has_dispatch_date_col=has_dispatch_date_col,
                        wide_date_columns=wide_date_columns,
                    )
                )
            open_keys.clear()
            merged.append(dict(r))
            continue

        key = (parent_id, _norm(r.get("工程名")), _norm(r.get("機械名")))
        if key not in groups:
            groups[key] = {"rows": [], "members": []}
            if key not in open_keys:
                open_keys.append(key)
        g = groups[key]
        g["rows"].append(dict(r))
        mem = _norm(r.get("メンバー名"))
        if mem:
            for part in re.split(r"[\u3000\s]+", mem):
                p = part.strip()
                if p and p not in g["members"]:
                    g["members"].append(p)

    for key in list(open_keys):
        merged.extend(
            _flush_branch_group(
                groups,
                key,
                sum_cols,
                has_dispatch_date_col=has_dispatch_date_col,
                wide_date_columns=wide_date_columns,
            )
        )
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
