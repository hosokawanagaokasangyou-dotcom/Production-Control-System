# -*- coding: utf-8 -*-
"""
インタラクティブ配台試行: 結果_配台表.json を tasks_df にマージし、段階2と同じ _generate_plan_impl を実行。

- 段階2の配台ループは変更せず、試行専用の環境・カレンダー解釈・結果表上書きで差し替える。
- 配台試行順は入力 JSON を正とする。結果_配台表は timeline の暦日集約を基準とし、入力が依頼×機械あたり 1 行のときも潰さない（planning_core）。
- 機械カレンダーは * / ＊ / ※ のセルのみ占有。列0にスロット行が無い時刻は配台不可。空セルは行がある時間帯で配台可。
  工場枠は master A12/B12 開始・同日 23:59 まで延長可。加工が暦日をまたぐ場合は PlanningValidationError で中止。
- 指定数量はタイムライン集計と突き合わせ、依頼NO×機械の合計不一致時は PlanningValidationError。
- 人員不足は interactive_trial_shortages_snapshot の op_shortage（フォーム0件）/ as_shortage（人数は足りるが割当不可）に記録。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import traceback
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
# python -P / PYTHONSAFEPATH ではスクリプト所在ディレクトリが sys.path に入らない。
sys.path.insert(0, str(SCRIPT_DIR))
os.chdir(str(SCRIPT_DIR))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def main() -> int:
    if len(sys.argv) < 2:
        print(
            "usage: dispatch_interactive_trial.py <path-to-result-dispatch.json>",
            file=sys.stderr,
        )
        return 2
    path = Path(sys.argv[1]).resolve()
    if not path.is_file():
        print(f"not a file: {path}", file=sys.stderr)
        return 1
    print("[dispatch trial] 入力JSONを読み込み中…", flush=True)
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"json read failed: {e}", file=sys.stderr)
        return 1

    rows = raw.get("rows") if isinstance(raw, dict) else None
    if rows is None:
        print("missing rows array", file=sys.stderr)
        return 1
    json_columns = raw.get("columns") if isinstance(raw, dict) else None

    os.environ["PM_AI_INTERACTIVE_DISPATCH_TRIAL"] = "1"

    shortage_path = path.with_name("dispatch_trial_shortages.json")

    # import が途中で失敗すると except PlanningValidationError が UnboundLocalError になるため、
    # 例外クラス名で分岐する（PlanningValidationError は try 内で import しない）。
    pc = None
    try:
        import planning_core as pc

        print("[dispatch trial] 計画タスクを読み込み、表データをマージ中…", flush=True)
        tasks_df = pc.load_planning_tasks_df()
        merged_df, targets = pc.merge_interactive_result_dispatch_json_into_tasks_df(
            tasks_df, rows
        )
        print("[dispatch trial] 段階2（配台計画）を実行中…（時間がかかる場合があります）", flush=True)
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
            "note": "interactive trial via planning_core._generate_plan_impl",
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
    except Exception as e:
        if pc is not None and type(e).__name__ == "PlanningValidationError":
            msg = str(e).strip() or "PlanningValidationError"
            print(msg, file=sys.stderr)
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
            return 3
        print(f"dispatch trial failed: {e}", file=sys.stderr)
        traceback.print_exc()
        return 1

    export_script = SCRIPT_DIR / "export_result_dispatch_from_json.py"
    if export_script.is_file():
        print("[dispatch trial] 結果Excel(xlsx)をエクスポート中…", flush=True)
        py = sys.executable or "python3"
        try:
            subprocess.run(
                [py, str(export_script), str(path)],
                cwd=str(SCRIPT_DIR),
                check=True,
                timeout=600,
            )
            print("[dispatch trial] xlsx エクスポート完了。", flush=True)
        except Exception as e:
            print(f"xlsx export warning: {e}", file=sys.stderr)

    print(str(shortage_path), flush=True)
    return 0


if __name__ == "__main__":
    try:
        import workbook_env_bootstrap as _wbe_exit

        sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
    except ImportError:
        sys.exit(main())
