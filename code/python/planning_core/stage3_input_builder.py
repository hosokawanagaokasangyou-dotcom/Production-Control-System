# -*- coding: utf-8 -*-
"""段階3.0 前処理: 結果_配台表.json（手動修正後）から枝番タスクを分解し入力3表を生成する。

- 入力: ``結果_配台表.json`` の各行（依頼NO×工程×機械名×配台日 と「当日配台数量」）
- 出力: ``plan_input_tasks.xlsx`` の第2シート（``PLAN_INPUT_STAGE3_SHEET_NAME``）に枝番行のみを書き出す

枝番行は元行（入力1表）を複製し、依頼NO=``{元依頼NO}-{枝番2桁}``、換算数量=セル数量、
原反投入日=入力1表の当該行（配台計画_タスク入力と同一。配台日で上書きしない）。
配台可能日時=結果JSONの配台日（暦日目標）+ master 定常開始（A15）。入力1の配台可能日時・原反+12:45 は使わない。
段階3.0〜3.2 では ``配台可能日時_上書き`` 列は使わない。
元タスク行は入力3表に含めない。
特別ルール scope は列「元依頼NO」で親に紐付ける（配台 task_id は枝番依頼NOのまま）。
"""
from __future__ import annotations

import json
import math
import os
import sys
from collections import defaultdict
from datetime import date, datetime, time
from pathlib import Path

_BASELINE_KEY_SEP = "\u0001"


def _norm(v) -> str:
    return str(v).strip() if v is not None else ""


def _aggregate_targets(pc, rows):
    """結果_配台表 rows を (依頼NO, 工程key, 機械名, 配台日) -> 配台 m に集約。"""
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
        qty = _dispatch_row_qty_m(pc, r)
        if tid and mach and dd is not None and qty > 1e-9:
            key = (tid, proc, mach, dd)
            targets[key] = targets.get(key, 0.0) + qty
    return targets


def _stage3_planning_meta_sidecar_path(json_path: Path) -> Path:
    name = json_path.name
    return json_path.resolve().parent / f"{name}.stage3_planning_meta.json"


def _read_stage3_planning_sidecar_variant(json_path: Path) -> str | None:
    sidecar = _stage3_planning_meta_sidecar_path(json_path)
    if not sidecar.is_file():
        return None
    try:
        raw = json.loads(sidecar.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    variant = raw.get("variant") if isinstance(raw, dict) else None
    return str(variant).strip() if variant else None


def _stage3_pipeline_already_applied(json_path: Path) -> bool:
    """段階3.0/3.1/3.2 実行済みなら baseline は段階2比較用のみ（入力3の目標源にしない）。"""
    return _read_stage3_planning_sidecar_variant(json_path) in ("3.0", "3.1", "3.2")


def _dispatch_row_qty_m(pc, row) -> float:
    """結果_配台表行の配台 m。実配台数量列があればそれを正とする。"""
    if not isinstance(row, dict):
        return 0.0
    for col in ("実配台数量", "当日配台数量"):
        qcell = row.get(col)
        if qcell in (None, ""):
            continue
        try:
            qty = float(str(qcell).replace(",", "").strip())
        except (TypeError, ValueError):
            qty = 0.0
        if qty > 1e-9:
            return qty
    return 0.0


def _tid_mach_process_lookup(pc, rows) -> dict[tuple[str, str], str]:
    """結果_配台表 rows から (依頼NO, 機械名) → 工程key を引く。"""
    out: dict[tuple[str, str], str] = {}
    for r in rows:
        if not isinstance(r, dict):
            continue
        tid = pc._interactive_norm_cell(r.get(pc.TASK_COL_TASK_ID)) or pc._interactive_norm_cell(
            r.get("タスクID")
        )
        mach = pc._interactive_norm_cell(r.get(pc.TASK_COL_MACHINE_NAME))
        if not tid or not mach:
            continue
        proc = pc._interactive_dispatch_target_process_key(r.get(pc.TASK_COL_MACHINE))
        out[(tid, mach)] = proc
    return out


def _aggregate_targets_from_baseline_sidecar(
    pc, json_path: Path, rows
) -> dict[tuple, float]:
    """段階2直後に保存した baselineEntries を枝番目標へ復元（アラジン1日化後も暦日別）。"""
    if _stage3_pipeline_already_applied(json_path):
        return {}
    sidecar = _stage3_planning_meta_sidecar_path(json_path)
    if not sidecar.is_file():
        return {}
    try:
        raw = json.loads(sidecar.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    entries = raw.get("baselineEntries") if isinstance(raw, dict) else None
    if not isinstance(entries, dict) or not entries:
        return {}
    return _aggregate_targets_from_baseline_entries(pc, rows, entries)


def _aggregate_targets_from_baseline_entries(
    pc, rows, entries: dict
) -> dict[tuple, float]:
    """Java wideShortfallKey 形式の baselineEntries を枝番目標へ変換。"""
    if not entries:
        return {}
    proc_by_tid_mach = _tid_mach_process_lookup(pc, rows)
    targets: dict[tuple, float] = {}
    for key, val in entries.items():
        if not isinstance(key, str):
            continue
        parts = key.split(_BASELINE_KEY_SEP, 2)
        if len(parts) < 3:
            continue
        tid = pc._interactive_norm_cell(parts[0]) or parts[0].strip()
        mach = pc._interactive_norm_cell(parts[1]) or parts[1].strip()
        iso_date = parts[2].strip()
        if not tid or not mach or not iso_date:
            continue
        try:
            qty = float(val)
        except (TypeError, ValueError):
            qty = 0.0
        if qty <= 1e-9:
            continue
        try:
            dd = date.fromisoformat(iso_date)
        except ValueError:
            dd = pc._interactive_parse_dispatch_date_cell(iso_date)
        if dd is None:
            continue
        proc = proc_by_tid_mach.get((tid, mach), "")
        tkey = (tid, proc, mach, dd)
        targets[tkey] = targets.get(tkey, 0.0) + qty
    return targets


def _baseline_entries_from_env() -> dict:
    raw = (os.environ.get("PM_AI_STAGE3_BASELINE_ENTRIES_JSON") or "").strip()
    if not raw:
        return {}
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _merge_json_targets_with_baseline(
    json_targets: dict[tuple, float], baseline_targets: dict[tuple, float]
) -> dict[tuple, float]:
    """現行 ``結果_配台表.json``（アラジン整列後を含む）を正とし、JSON に無いタスクのみ baseline で補う。

    旧挙動（baseline が JSON の1日化を上書き）は、アラジン後に入力3表が段階2の暦日2行のままになる不具合の原因だった。
    """
    if not baseline_targets:
        return dict(json_targets)

    merged = dict(baseline_targets)
    tasks_in_json: set[tuple[str, str]] = set()
    for key, qty in json_targets.items():
        if qty <= 1e-9:
            continue
        tid, _proc, mach, _dd = key
        if tid and mach:
            tasks_in_json.add((tid, mach))
    for tid, mach in tasks_in_json:
        drop = [k for k in merged if k[0] == tid and k[2] == mach]
        for k in drop:
            del merged[k]
    for key, qty in json_targets.items():
        if qty > 1e-9:
            merged[key] = float(qty)
    return merged


def refresh_post_stage3_dispatch_entries_sidecar(json_path: Path, *, pc=None) -> bool:
    """段階3枝番統合後: 正本の暦日別 m を sidecar ``postStage3Entries`` へ記録（baseline は維持）。"""
    if pc is None:
        from planning_core import _core as pc

    path = Path(json_path).resolve()
    if not path.is_file():
        return False
    try:
        raw_json = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    rows = raw_json.get("rows") if isinstance(raw_json, dict) else None
    if not isinstance(rows, list) or not rows:
        return False
    entries = baseline_entries_from_dispatch_rows(pc, rows)
    sidecar = _stage3_planning_meta_sidecar_path(path)
    sidecar_raw: dict = {}
    if sidecar.is_file():
        try:
            parsed = json.loads(sidecar.read_text(encoding="utf-8"))
            if isinstance(parsed, dict):
                sidecar_raw = parsed
        except (OSError, json.JSONDecodeError):
            pass
    sidecar_raw["postStage3Entries"] = entries
    sidecar.write_text(
        json.dumps(sidecar_raw, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return True


def baseline_entries_from_dispatch_rows(pc, rows) -> dict[str, float]:
    """段階2直後の JSON rows から Java ``wideShortfallKey`` 形式の baseline を構築。"""
    entries: dict[str, float] = {}
    for key, qty in _aggregate_targets(pc, rows).items():
        if qty <= 1e-9:
            continue
        tid, _proc, mach, dd = key
        if not tid or not mach or dd is None:
            continue
        iso = dd.isoformat() if hasattr(dd, "isoformat") else str(dd)
        k = f"{tid}{_BASELINE_KEY_SEP}{mach}{_BASELINE_KEY_SEP}{iso}"
        entries[k] = entries.get(k, 0.0) + float(qty)
    return entries


def write_stage3_baseline_sidecar_if_absent(json_path: Path, rows, *, pc=None) -> bool:
    """``結果_配台表.json`` 直後に baseline sidecar を1回だけ保存（段階2凍結用）。"""
    import os

    if (os.environ.get("PM_AI_PLAN_INPUT_STAGE3") or "").strip() == "1":
        return False
    if (os.environ.get("PM_AI_INTERACTIVE_DISPATCH_TRIAL") or "").strip() == "1":
        return False
    if pc is None:
        from planning_core import _core as pc

    path = Path(json_path).resolve()
    sidecar = _stage3_planning_meta_sidecar_path(path)
    if sidecar.is_file():
        return False
    entries = baseline_entries_from_dispatch_rows(pc, rows or [])
    if not entries:
        return False
    payload = {"baselineEntries": entries}
    sidecar.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return True


def resolve_stage3_dispatchable_datetime_for_input3_write(
    pc, prow, *, dispatch_date: date | None, shift_start: time
) -> datetime | None:
    """入力3枝番行の配台可能日時 = 当該枝番の配台日（JSON暦日）+ 定常開始。

    入力1の「配台可能日時」列や原反投入日+12:45 は参照しない（アラジン6/12・10000m なら 6/12+定常開始）。
    """
    if dispatch_date is None:
        return None
    return datetime.combine(dispatch_date, shift_start)


def _stage3_row_dispatch_qty_m(pc, row) -> float:
    for col in (
        pc.TASK_COL_UNPROCESSED,
        pc.TASK_COL_QTY,
        pc.PLAN_COL_DISPATCH_REMAINING_QTY,
    ):
        try:
            raw = pc._planning_df_cell_scalar(row, col)
            if raw in (None, ""):
                continue
            qty = float(str(raw).replace(",", "").strip())
        except (TypeError, ValueError):
            qty = 0.0
        if qty > 1e-9:
            return qty
    return 0.0


def _stage3_branch_dispatch_datetimes(pc, rows) -> list[datetime]:
    """枝番行の配台可能日時（段階3読込と同じ定常開始フォールバック）。"""
    shift = pc.stage3_regular_shift_start_time()
    out: list[datetime] = []
    for row in rows:
        col = pc.parse_optional_datetime(
            pc._planning_df_cell_scalar(row, pc.PLAN_COL_DISPATCHABLE_DATETIME)
        )
        if col is not None:
            out.append(col)
            continue
        dt = pc.resolve_dispatchable_datetime_from_plan_row(
            row, regular_shift_start=shift
        )
        if dt is not None:
            out.append(dt)
    return out


def _should_collapse_stage2_calendar_split_branches(pc, rows) -> bool:
    """段階2の暦日別枝番（部分量の複数行）のみ統合。意図的な複数日計画は統合しない。"""
    if len(rows) <= 1:
        return False
    qtys = [_stage3_row_dispatch_qty_m(pc, row) for row in rows]
    total = sum(qtys)
    if total <= 1e-9:
        return False
    dts = _stage3_branch_dispatch_datetimes(pc, rows)
    if len({d.date() for d in dts if hasattr(d, "date")}) <= 1:
        return True
    # 各枝番が総量の過半数未満 → 段階2の暦日分割（例 4000+6000）とみなす
    return min(qtys) < 0.45 * total


def collapse_stage3_plan_df_by_parent(pc, tasks_df):
    """同一元依頼NO×工程×機械の段階2暦日分割枝番を1行へまとめる（複数日の本計画は維持）。"""
    import pandas as pd

    if tasks_df is None or getattr(tasks_df, "empty", True):
        return tasks_df
    if pc.PLAN_COL_PARENT_TASK_ID not in tasks_df.columns:
        return tasks_df

    groups: dict[tuple[str, str, str], list] = defaultdict(list)
    for _, row in tasks_df.iterrows():
        tid = _norm(pc.planning_task_id_str_from_plan_row(row))
        parent = _norm(pc._planning_df_cell_scalar(row, pc.PLAN_COL_PARENT_TASK_ID)) or tid
        proc = pc._interactive_dispatch_target_process_key(
            pc._planning_df_cell_scalar(row, pc.TASK_COL_MACHINE)
        )
        mach = _norm(pc._planning_df_cell_scalar(row, pc.TASK_COL_MACHINE_NAME))
        if not parent or not mach:
            continue
        groups[(parent, proc, mach)].append(row)

    if all(len(v) <= 1 for v in groups.values()):
        return tasks_df

    order = list(tasks_df.columns)
    out_records: list[dict] = []
    collapsed: list[str] = []
    shift_start = pc.stage3_regular_shift_start_time()
    for (parent, proc, mach), rows in sorted(groups.items()):
        if len(rows) <= 1 or not _should_collapse_stage2_calendar_split_branches(pc, rows):
            for row in rows:
                out_records.append(row.to_dict())
            continue
        collapsed.append(parent)
        best_row = rows[0]
        best_dt: datetime | None = None
        min_dt: datetime | None = None
        total_qty = 0.0
        for row in rows:
            total_qty += _stage3_row_dispatch_qty_m(pc, row)
            dt = pc.parse_optional_datetime(
                pc._planning_df_cell_scalar(row, pc.PLAN_COL_DISPATCHABLE_DATETIME)
            )
            if dt is None:
                dt = pc.resolve_dispatchable_datetime_from_plan_row(
                    row, regular_shift_start=shift_start
                )
            if dt is None:
                continue
            if min_dt is None or dt < min_dt:
                min_dt = dt
            if best_dt is None or dt > best_dt:
                best_dt = dt
                best_row = row
        rec = {c: best_row.get(c, "") for c in order}
        branch_id = f"{parent}-01"
        rec[pc.PLAN_COL_PARENT_TASK_ID] = parent
        rec[pc.PLAN_COL_BRANCH_SEQ] = "01"
        rec[pc.TASK_COL_TASK_ID] = branch_id
        rec[pc.TASK_COL_QTY] = float(total_qty)
        rec[pc.TASK_COL_UNPROCESSED] = float(total_qty)
        rec[pc.PLAN_COL_DISPATCH_REMAINING_QTY] = float(total_qty)
        unit = pc._roll_unit_m_estimate_from_plan_row(best_row, total_qty)
        rec[pc.PLAN_COL_DISPATCH_ROLL_COUNT] = (
            int(math.ceil(total_qty / unit)) if unit > 1e-9 else ""
        )
        if best_dt is not None:
            rec[pc.PLAN_COL_DISPATCHABLE_DATETIME] = pc.format_dispatchable_datetime_cell(
                best_dt
            )
        # #region agent log
        if parent == "V6-2":
            try:
                from planning_core import agent_debug_ndjson as _adn

                _adn.append_structured(
                    "H-S3-collapse-dd",
                    "stage3_input_builder.collapse_stage3_plan_df_by_parent",
                    "collapse chose latest dispatchable",
                    {
                        "runId": "stage3-collapse-max-dd-v1",
                        "parent": parent,
                        "minDispatchable": min_dt.isoformat() if min_dt else "",
                        "maxDispatchable": best_dt.isoformat() if best_dt else "",
                        "written": rec.get(pc.PLAN_COL_DISPATCHABLE_DATETIME, ""),
                        "totalQty": total_qty,
                        "branchCount": len(rows),
                    },
                )
            except Exception:
                pass
        # #endregion
        out_records.append(rec)

    if not collapsed:
        return tasks_df

    out_df = pd.DataFrame(out_records).reindex(columns=order).fillna("")
    print(
        f"[stage3-input] 段階2暦日枝番を統合: {collapsed} → 各1行（計 {len(collapsed)} 親）",
        flush=True,
    )
    # #region agent log
    try:
        from planning_core import agent_debug_ndjson as _adn

        _adn.append_structured(
            "H-S3-collapse-parent",
            "stage3_input_builder.collapse_stage3_plan_df_by_parent",
            "collapsed multi-branch input3 rows",
            {
                "runId": "stage3-parent-collapse-v1",
                "parents": collapsed,
                "rowsBefore": int(len(tasks_df)),
                "rowsAfter": int(len(out_df)),
            },
        )
    except Exception:
        pass
    # #endregion
    return out_df


def build_stage3_dispatch_targets_from_plan_df(pc, tasks_df) -> dict[tuple, float]:
    """入力3表（枝番行）を段階3配台キャップへ。段階2暦日分割統合後は最遅配台可能日×総量の1キー。"""
    tasks_df = collapse_stage3_plan_df_by_parent(pc, tasks_df)
    targets: dict[tuple, float] = {}
    if tasks_df is None or getattr(tasks_df, "empty", True):
        return targets

    by_parent: dict[tuple[str, str, str], dict] = {}
    for _, row in tasks_df.iterrows():
        tid = pc._interactive_norm_cell(pc.planning_task_id_str_from_plan_row(row))
        parent = (
            pc._interactive_norm_cell(
                pc._planning_df_cell_scalar(row, pc.PLAN_COL_PARENT_TASK_ID)
            )
            or tid
        )
        if not tid or not parent:
            continue
        proc = pc._interactive_dispatch_target_process_key(
            pc._planning_df_cell_scalar(row, pc.TASK_COL_MACHINE)
        )
        mach = pc._interactive_norm_cell(
            pc._planning_df_cell_scalar(row, pc.TASK_COL_MACHINE_NAME)
        )
        dt = pc.parse_optional_datetime(
            pc._planning_df_cell_scalar(row, pc.PLAN_COL_DISPATCHABLE_DATETIME)
        )
        if dt is None:
            dt = pc.resolve_dispatchable_datetime_from_plan_row(
                row, regular_shift_start=pc.stage3_regular_shift_start_time()
            )
        if dt is None:
            continue
        dd = dt.date() if hasattr(dt, "date") else None
        if dd is None:
            continue
        qty = _stage3_row_dispatch_qty_m(pc, row)
        if qty <= 1e-9:
            continue
        gkey = (parent, proc, mach)
        slot = by_parent.get(gkey)
        if slot is None:
            by_parent[gkey] = {
                "tid": tid,
                "dd": dd,
                "qty": float(qty),
            }
        else:
            slot["qty"] += float(qty)
            if dd > slot["dd"]:
                slot["dd"] = dd
                slot["tid"] = tid

    for (_parent, proc, mach), slot in by_parent.items():
        targets[(slot["tid"], proc, mach, slot["dd"])] = float(slot["qty"])

    # #region agent log
    try:
        from planning_core import agent_debug_ndjson as _adn

        v62 = {str(k): targets[k] for k in targets if k[0].startswith("V6-2")}
        _adn.append_structured(
            "H-S3-cap-targets",
            "stage3_input_builder.build_stage3_dispatch_targets_from_plan_df",
            "stage3 cap targets built",
            {
                "runId": "stage3-parent-collapse-v1",
                "targetCount": len(targets),
                "v62": v62,
            },
        )
    except Exception:
        pass
    # #endregion
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

    targets_json = _aggregate_targets(pc, rows)
    targets_sidecar = _aggregate_targets_from_baseline_sidecar(pc, json_path, rows)
    targets_env = _aggregate_targets_from_baseline_entries(
        pc, rows, _baseline_entries_from_env()
    )
    targets_baseline = dict(targets_sidecar)
    for k, v in targets_env.items():
        if v > 1e-9:
            targets_baseline[k] = v
    targets = _merge_json_targets_with_baseline(targets_json, targets_baseline)

    v62_keys = [k for k in targets if k[0] == "V6-2"]
    # #region agent log
    try:
        from planning_core import agent_debug_ndjson as _adn

        _adn.append_structured(
            "H-input3-baseline",
            "stage3_input_builder.py:build_stage3_input_sheet",
            "stage3 input targets merged",
            {
                "runId": "aladdin-input3-prefer-json-v1",
                "jsonTargetDays": len(targets_json),
                "sidecarTargetDays": len(targets_sidecar),
                "envTargetDays": len(targets_env),
                "mergedTargetDays": len(targets),
                "v62BranchDays": len(v62_keys),
                "v62QtyByDay": {str(k[3]): targets[k] for k in v62_keys},
                "baselineSidecar": str(_stage3_planning_meta_sidecar_path(json_path)),
                "sidecarExists": _stage3_planning_meta_sidecar_path(json_path).is_file(),
            },
        )
    except Exception:
        pass
    # #endregion

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
    trial_order = 0

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

        # 原反投入日は入力1表から複製済み（配台日で上書きしない）。
        # 配台可能日時 = この枝番の配台日（JSON）+ 定常開始（入力1列は上書きで使わない）。
        dt_disp = resolve_stage3_dispatchable_datetime_for_input3_write(
            pc, prow, dispatch_date=dd, shift_start=shift_start
        )
        rec[pc.PLAN_COL_DISPATCHABLE_DATETIME] = pc.format_dispatchable_datetime_cell(
            dt_disp
        )
        if tid == "V6-2":
            # #region agent log
            try:
                from planning_core import agent_debug_ndjson as _adn

                _adn.append_structured(
                    "H-input3-dispatchable-dt",
                    "stage3_input_builder.build_stage3_input_sheet",
                    "input3 dispatchable datetime written",
                    {
                        "runId": "input3-dispatch-date-shift-v1",
                        "branchId": branch_id,
                        "dispatchDate": str(dd),
                        "shiftStart": shift_start.strftime("%H:%M"),
                        "written": rec[pc.PLAN_COL_DISPATCHABLE_DATETIME],
                    },
                )
            except Exception:
                pass
            # #endregion
        trial_order += 1
        rec[pc.RESULT_TASK_COL_DISPATCH_TRIAL_ORDER] = trial_order
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
