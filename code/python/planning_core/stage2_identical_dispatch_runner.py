# -*- coding: utf-8 -*-
"""
段階2／段階3（インタラクティブ配台試行）の共通オーケストレーション。

- 段階2の正本: ``master`` の工場時間コンテキストのうえで ``_generate_plan_impl()`` を
  オーバーライド無しで実行する（従来 ``generate_plan()`` が担っていた処理）。
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


def run_stage2_generate_plan() -> None:
    """
    段階2標準: マスタの工場時間オーバーライドのうえで ``_generate_plan_impl()`` を 1 回実行する。
    """
    from planning_core import _core as pc

    master_abs = pc._master_workbook_path_resolved()
    with pc._override_default_factory_hours_from_master(master_abs):
        pc._generate_plan_impl()


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

    shortage_path = path.with_name("dispatch_trial_shortages.json")

    try:
        print("[dispatch trial] 計画タスクを読み込み、表データをマージ中…", flush=True)
        tasks_df = pc.load_planning_tasks_df()
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
        dispatch_qty_shortfall = pc.compute_interactive_trial_dispatch_qty_shortfall(
            targets if targets else None,
            md_snap if md_snap else None,
        )
        shortage_payload: dict = {
            "format_version": 2,
            "source_json": str(path),
            "note": "interactive trial via stage2_identical_dispatch_runner",
            "op_shortage": snap["op_shortage"],
            "as_shortage": snap["as_shortage"],
            "dispatch_qty_shortfall": dispatch_qty_shortfall,
        }
        if isinstance(paths, dict):
            shortage_payload["production_plan"] = str(paths.get("production_plan") or "")
            shortage_payload["member_schedule"] = str(paths.get("member_schedule") or "")
        shortage_path.write_text(
            json.dumps(shortage_payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        print("[dispatch trial] 不足情報JSONを書き出しました。", flush=True)
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
