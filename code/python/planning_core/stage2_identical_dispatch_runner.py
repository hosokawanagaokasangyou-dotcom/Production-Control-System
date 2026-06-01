# -*- coding: utf-8 -*-
"""
段階2／段階3（インタラクティブ配台試行）の共通オーケストレーション。

- 段階2の正本: ``master`` の工場時間コンテキストのうえで ``_generate_plan_impl()`` を
  オーバーライド無しで実行する（従来 ``generate_plan()`` が担っていた処理）。
  配台失敗ポリシーは段階3（段階2同一パリティ）と同一（前提未満足のみ致命・他は後ろ倒し）。
- 段階3の正本: 結果_配台表 JSON の ``rows`` / ``columns`` を読み、タスク DataFrame に
  マージしたうえで ``PM_AI_INTERACTIVE_TRIAL_STAGE2_PARITY=1`` を付与し、
  配台エンジン条件を段階2と同一にした ``_generate_plan_impl`` を実行する。
  不足 JSON のペイロード組み立て・検証失敗時の書き出しもここに集約する。

巨大な配台ループ本体は ``planning_core._core`` の ``_generate_plan_impl`` に残す。
"""
from __future__ import annotations

import json
import os
import traceback
from pathlib import Path

# Runner が設定する環境変数（_core 側で参照）
ENV_INTERACTIVE_TRIAL_STAGE2_PARITY = "PM_AI_INTERACTIVE_TRIAL_STAGE2_PARITY"


def _flush_dispatch_rule_trace_sidecar() -> None:
    try:
        from planning_core.dispatch_rules import trace_recorder

        path = trace_recorder.write_sidecar()
        if path is not None:
            print(f"[dispatch-rules] trace sidecar: {path}", flush=True)
    except Exception as ex:
        print(f"[dispatch-rules] trace sidecar failed: {ex}", flush=True)


def run_stage2_generate_plan() -> None:
    """
    段階2標準: マスタの工場時間オーバーライドのうえで ``_generate_plan_impl()`` を 1 回実行する。
    """
    from planning_core import _core as pc
    from planning_core.dispatch_rules import trace_recorder

    trace_recorder.reset_trace()
    master_abs = pc._master_workbook_path_resolved()
    try:
        with pc._override_default_factory_hours_from_master(master_abs):
            pc._generate_plan_impl()
    finally:
        _flush_dispatch_rule_trace_sidecar()


def _load_stage3_input_tasks_df(pc):
    """入力3表（第2シート ``PLAN_INPUT_STAGE3_SHEET_NAME``）を tasks_df として読み込む。

    列構成は入力1表 + ``元依頼NO`` / ``配台枝番``。読込後の正規化・速度適用は
    ``load_planning_tasks_df`` と同等を最小限で行う（特別ルール scope は build_task_queue が
    ``元依頼NO`` から rule_task_id を解決する）。
    """
    import os

    plan_path = (os.environ.get(pc.ENV_PLAN_INPUT_PATH) or "").strip()
    if not plan_path or not os.path.isfile(plan_path):
        raise FileNotFoundError(
            f"段階3: {pc.ENV_PLAN_INPUT_PATH} が実在しません: {plan_path!r}。"
        )
    sheet = pc.PLAN_INPUT_STAGE3_SHEET_NAME
    df = pc.read_tabular_dataframe(plan_path, sheet_name=sheet)
    df.columns = df.columns.str.strip()
    df = pc._align_dataframe_headers_to_canonical(df, pc.plan_input_stage3_sheet_column_order())
    for c in pc.plan_input_stage3_sheet_column_order():
        if c not in df.columns:
            df[c] = ""
    df = pc._coalesce_plan_plain_remark_into_special(df)
    pc._apply_master_speed_sheet_to_plan_df(df, log_prefix="入力3表読込")
    try:
        from planning_core.actual_speed_apply import apply_learned_speed_to_plan_df

        apply_learned_speed_to_plan_df(df, log_prefix="入力3表読込")
    except Exception:
        pass
    pc._fill_plan_dispatch_remaining_qty_column(df)
    return df


def run_stage3_generate_plan(*, qty_strict: bool = False) -> dict:
    """段階3.0/3.2: 入力3表（枝番）で配台Aを実行し、枝番統合まで行う。

    Args:
        qty_strict: True で段階3.2（同日完走必須・定常外人ブロック無視）。

    Returns:
        枝番統合の結果 dict（``merge_branch_result_dispatch`` の戻り）または ``{}``。
    """
    import os

    from planning_core import _core as pc
    from planning_core.dispatch_rules import trace_recorder

    trace_recorder.reset_trace()
    os.environ["PM_AI_PLAN_INPUT_STAGE3"] = "1"
    if qty_strict:
        os.environ["PM_AI_STAGE3_2_QTY_STRICT"] = "1"

    master_abs = pc._master_workbook_path_resolved()
    try:
        tasks_df = _load_stage3_input_tasks_df(pc)
        with pc._override_default_factory_hours_from_master(master_abs):
            pc._generate_plan_impl(tasks_df_override=tasks_df)
        return _merge_stage3_branches(pc)
    finally:
        if qty_strict:
            os.environ.pop("PM_AI_STAGE3_2_QTY_STRICT", None)
        os.environ.pop("PM_AI_PLAN_INPUT_STAGE3", None)
        _flush_dispatch_rule_trace_sidecar()


def _merge_stage3_branches(pc) -> dict:
    """配台出力（結果_配台表.json）を元依頼NO単位へ統合し正本へ上書きする。"""
    import os

    try:
        from planning_core.dispatch_workspace import resolve_result_dispatch_table_output_dir
        from planning_core import stage3_branch_merge

        plan_path = (os.environ.get(pc.ENV_PLAN_INPUT_PATH) or "").strip()
        out_dir = resolve_result_dispatch_table_output_dir(plan_path)
        if not out_dir:
            return {}
        from pathlib import Path

        json_path = Path(out_dir) / pc.RESULT_DISPATCH_TABLE_JSON_FILENAME
        if not json_path.is_file():
            return {}
        res = stage3_branch_merge.merge_branch_result_dispatch(json_path, plan_path)
        print(
            f"[stage3-merge] 枝番統合: {res.get('source_rows')} 行 → {res.get('merged_rows')} 行",
            flush=True,
        )
        return res
    except Exception as ex:
        print(f"[stage3-merge] 枝番統合スキップ: {ex}", flush=True)
        return {}


def run_interactive_dispatch_trial_from_result_dispatch_json(
    path: Path,
) -> tuple[int, Path | None]:
    """
    段階3: ``結果_配台表.json`` を入力とし、段階2同一条件で配台を実行する。

    Returns:
        (exit_code, shortage_json_path_or_none)
        exit_code: 0 成功、1 一般失敗、2 usage、3 PlanningValidationError
    """
    path = path.resolve()
    if not path.is_file():
        return 1, None

    from planning_core import _core as pc

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"json read failed: {e}", flush=True)
        return 1, None

    rows = raw.get("rows") if isinstance(raw, dict) else None
    if rows is None:
        print("missing rows array", flush=True)
        return 1, None
    json_columns = raw.get("columns") if isinstance(raw, dict) else None

    os.environ["PM_AI_INTERACTIVE_DISPATCH_TRIAL"] = "1"
    os.environ[ENV_INTERACTIVE_TRIAL_STAGE2_PARITY] = "1"

    from planning_core.dispatch_rules import trace_recorder

    trace_recorder.reset_trace()
    shortage_path = path.with_name("dispatch_trial_shortages.json")

    try:
        print("[dispatch trial] 計画タスクを読み込み、表データをマージ中…", flush=True)
        tasks_df = pc.load_planning_tasks_df()
        try:
            df_src = pc.load_tasks_df()
        except Exception as _e_src:
            print(
                f"[dispatch trial] 加工計画DATA 読込に失敗（静的列の補完のみスキップ）: {_e_src}",
                flush=True,
            )
            df_src = None
        n_fill = pc.fill_interactive_result_dispatch_json_rows_from_planning_sources(
            rows, tasks_df, df_src
        )
        if n_fill:
            print(f"[dispatch trial] 計画入力・加工計画DATA から {n_fill} セルを補完しました。", flush=True)
        merged_df, targets = pc.merge_interactive_result_dispatch_json_into_tasks_df(
            tasks_df, rows
        )
        print("[dispatch trial] 段階2同一条件で配台を実行中…（時間がかかる場合があります）", flush=True)
        master_abs = pc._master_workbook_path_resolved()
        with pc._override_default_factory_hours_from_master(master_abs):
            paths = pc._generate_plan_impl(
                tasks_df_override=merged_df,
                return_output_paths=True,
                interactive_relax_intraday=False,
                interactive_dispatch_targets=targets if targets else None,
                interactive_result_dispatch_json_rows=rows,
                interactive_result_dispatch_json_columns=json_columns
                if isinstance(json_columns, list)
                else None,
            )
        snap = pc.interactive_trial_shortages_snapshot()
        md_snap = pc.interactive_trial_meters_done_snapshot()
        plan_targets = pc.interactive_trial_plan_targets_snapshot()
        shortfall_targets = plan_targets if plan_targets else (targets if targets else None)
        dispatch_qty_shortfall = pc.compute_interactive_trial_dispatch_qty_shortfall(
            shortfall_targets,
            md_snap if md_snap else None,
        )
        _meta_miss = pc.interactive_trial_meta_miss_shortfall_snapshot()
        if _meta_miss:
            _seen_sf = {
                (
                    str(r.get("task_id") or ""),
                    str(r.get("machine_name") or ""),
                    str(r.get("dispatch_date") or ""),
                )
                for r in dispatch_qty_shortfall
            }
            for _row in _meta_miss:
                _k = (
                    str(_row.get("task_id") or ""),
                    str(_row.get("machine_name") or ""),
                    str(_row.get("dispatch_date") or ""),
                )
                if _k not in _seen_sf:
                    dispatch_qty_shortfall.append(_row)
                    _seen_sf.add(_k)
        s3_meta = {}
        try:
            s3_meta = pc.interactive_stage3_last_run_meta_snapshot()
        except Exception:
            s3_meta = {}
        _remaining_at_end = []
        try:
            _remaining_at_end = pc.interactive_trial_remaining_tasks_at_calendar_end_snapshot()
        except Exception:
            _remaining_at_end = []
        _ot_sim = (os.environ.get(pc.ENV_OVERTIME_SIMULATION_JSON) or "").strip()
        _note_base = (
            "段階3配台試行（段階2同一条件）。"
            "配台ループのブロック条件は段階2と同一。"
            "JSON 暦日×数量（interactive_dispatch_targets）がある手動修正試行ではループ内キャップも有効。"
            "致命: 機械カレンダー未作成・勤怠未作成。"
            "勤怠最終日まで未割当のタスクは remaining_tasks_at_calendar_end に載せて試行は続行。"
            "計画暦日未達は dispatch_qty_shortfall に載せる。"
            "勤怠日付の自動拡張は行わない。"
        )
        if _ot_sim:
            _note_base += f" 残業シミュレーション適用: {_ot_sim}"
        shortage_payload: dict = {
            "format_version": 3,
            "source_json": str(path),
            "note": _note_base,
            "op_shortage": snap["op_shortage"],
            "as_shortage": snap["as_shortage"],
            "dispatch_qty_shortfall": dispatch_qty_shortfall,
            "remaining_tasks_at_calendar_end": _remaining_at_end,
            "stage3": s3_meta,
        }
        if isinstance(paths, dict):
            shortage_payload["production_plan"] = str(paths.get("production_plan") or "")
            shortage_payload["member_schedule"] = str(paths.get("member_schedule") or "")
        if _remaining_at_end:
            print(
                "[dispatch trial] 警告: 勤怠最終日まで未割当 "
                f"{len(_remaining_at_end)} 件（remaining_tasks_at_calendar_end を参照）",
                flush=True,
            )
        shortage_path.write_text(
            json.dumps(shortage_payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        print("[dispatch trial] 不足情報JSONを書き出しました。", flush=True)
        try:
            from planning_core.dispatch_workspace import resolve_result_dispatch_table_output_dir

            _pip = (os.environ.get(pc.ENV_PLAN_INPUT_PATH) or "").strip()
            _out_dir = resolve_result_dispatch_table_output_dir(_pip)
            _written = Path(_out_dir) / pc.RESULT_DISPATCH_TABLE_JSON_FILENAME if _out_dir else None
            if _written is not None and _written.is_file() and _written.resolve() != path.resolve():
                path.write_text(_written.read_text(encoding="utf-8"), encoding="utf-8")
                print(f"[dispatch trial] 結果_配台表.json を入力パスへ同期: {_written} -> {path}", flush=True)
        except Exception as _sync_e:
            print(f"[dispatch trial] 結果_配台表.json 同期スキップ: {_sync_e}", flush=True)
        return 0, shortage_path
    except Exception as e:
        if type(e).__name__ == "PlanningValidationError":
            msg = str(e).strip() or "PlanningValidationError"
            print(msg, flush=True)
            try:
                pc._write_stage2_blocking_message(msg)
            except Exception:
                pass
            try:
                snap = pc.interactive_trial_shortages_snapshot()
                shortage_path.write_text(
                    json.dumps(
                        {
                            "format_version": 2,
                            "source_json": str(path),
                            "note": "validation failed before/during stage2",
                            "error": msg,
                            "op_shortage": snap["op_shortage"],
                            "as_shortage": snap["as_shortage"],
                            "dispatch_qty_shortfall": [],
                        },
                        ensure_ascii=False,
                        indent=2,
                    )
                    + "\n",
                    encoding="utf-8",
                )
            except Exception:
                pass
            return 3, shortage_path
        print(f"dispatch trial failed: {e}", flush=True)
        traceback.print_exc()
        return 1, None
    finally:
        _flush_dispatch_rule_trace_sidecar()
