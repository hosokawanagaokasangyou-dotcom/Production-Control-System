#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
段階1→段階2→アラジン整列→段階3 を実行し、指定依頼NOの段階3配台照合を判定する。

用法:
  py -3 scripts/dispatch_balance_auto_iterate.py --task Y5-24
  py -3 scripts/dispatch_balance_auto_iterate.py --task Y6-4 --skip-stage1 --skip-stage2

終了コード: 0=照合OK, 2=入力不足, 3=段階3致命, 4=照合NG, 5=その他失敗
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PY_DIR = ROOT / "code" / "python"
DISPATCH_JSON = ROOT / "code" / "output" / "結果_配台表.json"
ALADDIN_JSON = ROOT / "code" / "output" / "shaped_aladdin_plan.json"
SHORTAGES_JSON = ROOT / "code" / "output" / "dispatch_trial_shortages.json"
SUMMARY_JSON = ROOT / "code" / "output" / "dispatch_balance_iterate_summary.json"


def apply_env() -> None:
    os.environ["PM_AI_REPO_ROOT"] = str(ROOT)
    os.environ["PM_AI_WORKSPACE"] = str(ROOT)
    os.environ["PM_AI_OUTPUT_DIR"] = str(ROOT / "output")
    os.environ["PM_AI_CODE_PYTHON_DIR"] = str(PY_DIR)
    os.environ["PM_AI_CODE_DIR"] = str(ROOT / "code")
    os.environ["PM_AI_RESULT_DISPATCH_TABLE_DIR"] = str(ROOT / "code" / "output")
    os.environ["PM_AI_PLAN_INPUT_PATH"] = str(ROOT / "output" / "plan_input_tasks.xlsx")
    os.environ["PM_AI_SKIP_WORKBOOK_ENV_SHEET"] = "1"
    os.environ["PM_AI_CMD_PAUSE_ON_ERROR"] = "0"
    os.environ["PM_AI_STAGE2_WRITE_EXCEL"] = "1"
    os.environ["PM_AI_STAGE2_SKIP_TODAY_DISPATCH"] = "1"
    os.environ["PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH"] = "0"
    os.environ["PM_AI_INTERACTIVE_DISPATCH_TRIAL"] = "1"
    os.environ["PM_AI_INTERACTIVE_TRIAL_STAGE2_PARITY"] = "1"
    cache_csv = ROOT / ".pm-ai-cache" / "network-source" / "task-input-newest.csv"
    if cache_csv.is_file():
        os.environ.setdefault("PM_AI_TASK_INPUT_SOURCE_DIR", str(cache_csv.parent))
    master = ROOT / "code" / "master.xlsm"
    if master.is_file():
        os.environ["PM_AI_MASTER_WORKBOOK"] = str(master)
    os.environ["PM_AI_EXCLUDE_RULES_JSON"] = str(ROOT / "code" / "json" / "stage1_exclude_rules.json")
    for key, rel in {
        "RAW_FABRIC_WIDTH_TABLE_PATH": "code/使用原反, 加工幅.txt",
        "ROLL_UNIT_BY_USED_RAW_TABLE_PATH": "code/使用原反,ロール単位の長さ.txt",
        "PRODUCT_WIDTH_TABLE_PATH": "code/製品名, 製品幅.txt",
        "PRODUCT_LENGTH_TABLE_PATH": "code/製品名,製品長.txt",
        "PRODUCT_THICKNESS_TABLE_PATH": "code/製品名,製品厚み.txt",
    }.items():
        os.environ.setdefault(key, str(ROOT / rel))
    os.chdir(str(PY_DIR))
    sys.path.insert(0, str(PY_DIR))


def run_stage1() -> int:
    import planning_core as pc

    print("[balance-iterate] 段階1…", flush=True)
    pc.run_stage1_extract()
    return 0


def run_stage2() -> int:
    from planning_core.stage2_identical_dispatch_runner import run_stage2_generate_plan

    print("[balance-iterate] 段階2…", flush=True)
    run_stage2_generate_plan()
    if not DISPATCH_JSON.is_file():
        out_alt = ROOT / "output" / "結果_配台表.json"
        if out_alt.is_file():
            DISPATCH_JSON.parent.mkdir(parents=True, exist_ok=True)
            DISPATCH_JSON.write_text(out_alt.read_text(encoding="utf-8"), encoding="utf-8")
    return 0 if DISPATCH_JSON.is_file() else 5


def run_aladdin_align(*, from_tomorrow: bool = False, align_all_days: bool = False) -> int:
    from planning_core.dispatch_aladdin_align_json import (
        align_dispatch_json_from_aladdin,
        resolve_align_from_day,
    )

    if not DISPATCH_JSON.is_file():
        return 5
    payload = json.loads(DISPATCH_JSON.read_text(encoding="utf-8"))
    if align_all_days:
        align_from = None
    elif from_tomorrow:
        align_from = date.today() + timedelta(days=1)
    else:
        align_from = resolve_align_from_day()
    new_payload, n_changed = align_dispatch_json_from_aladdin(
        payload, ALADDIN_JSON, align_from_day=align_from
    )
    DISPATCH_JSON.write_text(
        json.dumps(new_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"[balance-iterate] アラジン整列: {n_changed} プロファイル変更", flush=True)
    return 0


def sync_dispatch_json_from_python_output_dir() -> bool:
    """段階3試行の書き出し先（多くは output/）を code/output の正本へ揃える。"""
    from planning_core.dispatch_workspace import resolve_result_dispatch_table_output_dir

    pip = (os.environ.get("PM_AI_PLAN_INPUT_PATH") or "").strip()
    src_dir = resolve_result_dispatch_table_output_dir(pip)
    if not src_dir:
        return False
    src = Path(src_dir) / DISPATCH_JSON.name
    if not src.is_file():
        alt = ROOT / "output" / DISPATCH_JSON.name
        if alt.is_file():
            src = alt
        else:
            return False
    try:
        if src.resolve() == DISPATCH_JSON.resolve():
            return True
        DISPATCH_JSON.parent.mkdir(parents=True, exist_ok=True)
        DISPATCH_JSON.write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
        print(f"[balance-iterate] 配台表JSON同期: {src} -> {DISPATCH_JSON}", flush=True)
        return True
    except OSError as e:
        print(f"[balance-iterate] 配台表JSON同期失敗: {e}", file=sys.stderr)
        return False


def run_stage3() -> int:
    from planning_core.stage2_identical_dispatch_runner import (
        run_interactive_dispatch_trial_from_result_dispatch_json,
    )

    print(f"[balance-iterate] 段階3… input={DISPATCH_JSON}", flush=True)
    code, _ = run_interactive_dispatch_trial_from_result_dispatch_json(DISPATCH_JSON)
    sync_dispatch_json_from_python_output_dir()
    return code


def check_balance(task_id: str, process: str | None) -> dict:
    from planning_core.dispatch_balance_check import check_task_balance

    payload = json.loads(DISPATCH_JSON.read_text(encoding="utf-8"))
    rows = payload.get("rows") or []
    has_actual = any("実配台数量" in (payload.get("columns") or []) for _ in [0]) or bool(
        rows and "実配台数量" in rows[0]
    )
    if payload.get("columns") and "実配台数量" in payload["columns"]:
        has_actual = True
    result = check_task_balance(rows, task_id, process=process, has_actual_col=has_actual)
    if result is None:
        return {"task_id": task_id, "found": False, "check": "NOT_FOUND"}
    out = {
        "task_id": result.task_id,
        "process": result.process,
        "machine": result.machine,
        "found": True,
        "qty_converted": result.qty_converted,
        "actual_processed": result.actual_processed,
        "plan_total": result.plan_total,
        "actual_dispatch_total": result.actual_dispatch_total,
        "expected": result.expected,
        "check": result.check,
        "ok": result.ok,
        "rows": result.rows,
    }
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--task", required=True, help="依頼NO (例 Y5-24)")
    parser.add_argument("--process", default=None, help="工程名で絞る (例 SEC)")
    parser.add_argument("--skip-stage1", action="store_true")
    parser.add_argument("--skip-stage2", action="store_true")
    parser.add_argument("--skip-align", action="store_true")
    parser.add_argument("--skip-stage3", action="store_true")
    parser.add_argument(
        "--align-from-tomorrow",
        action="store_true",
        help="アラジン整列を翌日以降に限定（既定は定常開始前なら当日含む）",
    )
    parser.add_argument(
        "--align-all-days",
        action="store_true",
        help="アラジン整列を全日（操作日以前も含む）対象にする",
    )
    args = parser.parse_args()

    apply_env()
    summary: dict = {"task": args.task, "steps": []}

    try:
        if not args.skip_stage1:
            run_stage1()
            summary["steps"].append("stage1:ok")
    except Exception as e:
        summary["steps"].append(f"stage1:fail:{e}")
        print(f"段階1失敗: {e}", file=sys.stderr)
        SUMMARY_JSON.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        return 5

    try:
        if not args.skip_stage2:
            rc = run_stage2()
            summary["steps"].append(f"stage2:exit{rc}")
            if rc != 0:
                SUMMARY_JSON.write_text(
                    json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                return rc
    except Exception as e:
        summary["steps"].append(f"stage2:fail:{e}")
        print(f"段階2失敗: {e}", file=sys.stderr)
        SUMMARY_JSON.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        return 5

    if not DISPATCH_JSON.is_file():
        print(f"missing {DISPATCH_JSON}", file=sys.stderr)
        return 2

    if not args.skip_align:
        run_aladdin_align(
            from_tomorrow=args.align_from_tomorrow,
            align_all_days=args.align_all_days,
        )
        summary["steps"].append("align:ok")

    if not args.skip_stage3:
        code = run_stage3()
        summary["steps"].append(f"stage3:exit{code}")
        if code != 0:
            summary["balance"] = check_balance(args.task, args.process)
            SUMMARY_JSON.write_text(
                json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            return 3 if code == 3 else code

    balance = check_balance(args.task, args.process)
    summary["balance"] = balance
    if SHORTAGES_JSON.is_file():
        try:
            summary["shortages"] = json.loads(SHORTAGES_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass
    SUMMARY_JSON.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))

    if not balance.get("found"):
        return 2
    if balance.get("ok"):
        return 0
    return 4


if __name__ == "__main__":
    sys.exit(main())
