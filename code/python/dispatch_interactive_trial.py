# -*- coding: utf-8 -*-
"""Interactive dispatch trial: merge result-dispatch JSON into plan tasks_df and run stage2."""
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

    # #region agent log
    def _agent_ndjson_append(payload: dict) -> None:
        p = (os.environ.get("CURSOR_DEBUG_LOG") or os.environ.get("PM_AI_DEBUG_LOG") or "").strip()
        if not p:
            return
        try:
            import time

            payload.setdefault("sessionId", "e6dc4e")
            payload.setdefault("timestamp", int(time.time() * 1000))
            with open(p, "a", encoding="utf-8") as _af:
                _af.write(json.dumps(payload, ensure_ascii=False) + "\n")
        except Exception:
            pass

    y514_rows = []
    for _r in rows or []:
        if not isinstance(_r, dict):
            continue
        if (
            str(_r.get("依頼NO", "")).strip() == "Y5-14"
            and str(_r.get("工程名", "")).strip() == "SEC"
        ):
            y514_rows.append(
                {
                    "配台試行順番": _r.get("配台試行順番"),
                    "機械名": _r.get("機械名"),
                    "配台日": _r.get("配台日"),
                    "当日配台数量": _r.get("当日配台数量"),
                }
            )
    _agent_ndjson_append(
        {
            "location": "dispatch_interactive_trial.py",
            "hypothesisId": "H1",
            "message": "argv_json_Y5-14_SEC_rows",
            "data": {"count": len(y514_rows), "rows": y514_rows},
        }
    )
    # #endregion

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
        # #region agent log
        y514_t = {}
        for _k, _v in (targets or {}).items():
            try:
                if _k and len(_k) >= 1 and str(_k[0]).strip() == "Y5-14":
                    y514_t[str(_k)] = float(_v)
            except Exception:
                pass
        _agent_ndjson_append(
            {
                "location": "dispatch_interactive_trial.py",
                "hypothesisId": "H4",
                "message": "merged_targets_Y5-14_keys",
                "data": y514_t,
            }
        )
        # #endregion
        print("[dispatch trial] 段階2（配台計画）を実行中…（時間がかかる場合があります）", flush=True)
        paths = pc._generate_plan_impl(
            tasks_df_override=merged_df,
            return_output_paths=True,
            interactive_relax_intraday=True,
            interactive_dispatch_targets=targets if targets else None,
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
