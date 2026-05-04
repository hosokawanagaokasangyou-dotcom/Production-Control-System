# -*- coding: utf-8 -*-
"""
インタラクティブ配台試行: 結果_配台表.json を tasks_df にマージし、段階2と同じ _generate_plan_impl を実行。

- 段階2の配台ループは変更せず、試行専用の環境・カレンダー解釈・結果表上書きで差し替える。
- 配台試行順・JSON の日別数量は入力を正とする（planning_core 内で results 表を上書き）。
- 機械カレンダーは * / ＊ / ※ のセルのみ占有。工場枠は master A12/B12 開始・同日 23:59 まで延長可。
  加工が暦日をまたぐ場合は PlanningValidationError で中止。
- 人員不足は interactive_trial_shortages_snapshot の op_shortage / as_shortage に記録。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import traceback
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
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

    try:
        import planning_core as pc
        from planning_core.bootstrap import PlanningValidationError

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
        shortage_payload: dict = {
            "format_version": 2,
            "source_json": str(path),
            "note": "interactive trial via planning_core._generate_plan_impl",
            "op_shortage": snap["op_shortage"],
            "as_shortage": snap["as_shortage"],
        }
        if isinstance(paths, dict):
            shortage_payload["production_plan"] = str(paths.get("production_plan") or "")
            shortage_payload["member_schedule"] = str(paths.get("member_schedule") or "")
        shortage_path.write_text(
            json.dumps(shortage_payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        # #region agent log
        try:
            import time as _t_y5

            _lp_y5 = (
                (os.environ.get("CURSOR_DEBUG_LOG") or os.environ.get("PM_AI_DEBUG_LOG") or "").strip()
                or "/mnt/c/工程管理AIプロジェクト_JAVA/.cursor/debug-7a6e73.log"
            )
            _all_s = list(shortage_payload.get("op_shortage") or []) + list(
                shortage_payload.get("as_shortage") or []
            )
            _y5 = [
                x
                for x in _all_s
                if isinstance(x, dict) and str(x.get("task_id", "")).strip() == "Y5-14"
            ]
            _rec_y5 = {
                "sessionId": "7a6e73",
                "timestamp": int(_t_y5.time() * 1000),
                "location": "dispatch_interactive_trial.main",
                "message": "trial_shortage_y5_14",
                "hypothesisId": "H_Y5",
                "data": {
                    "y5_14_entry_count": len(_y5),
                    "production_plan": str(shortage_payload.get("production_plan") or "")[:400],
                    "source_json": str(shortage_payload.get("source_json") or "")[:400],
                },
            }
            with open(_lp_y5, "a", encoding="utf-8") as _f_y5:
                _f_y5.write(json.dumps(_rec_y5, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        print("[dispatch trial] 不足情報JSONを書き出しました。", flush=True)
    except PlanningValidationError as e:
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
    except Exception as e:
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
