#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""段階3配台試行を CLI で実行し、未配台・未達を要約する（自動デバッグ用）。"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PY_DIR = ROOT / "code" / "python"
DISPATCH_JSON = ROOT / "code" / "output" / "結果_配台表.json"
SHORTAGES_JSON = ROOT / "code" / "output" / "dispatch_trial_shortages.json"
DEBUG_LOG = ROOT / ".cursor" / "debug-a15218.log"


def _win_to_wsl(p: str) -> str:
    p = p.strip()
    if not p:
        return p
    if p.startswith("C:\\") or p.startswith("c:\\"):
        rest = p[3:].replace("\\", "/")
        return f"/mnt/c/{rest}"
    return p


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
    os.environ["PM_AI_AGENT_DEBUG_SESSION"] = "a15218"
    os.environ["PM_AI_DEBUG_LOG"] = str(DEBUG_LOG)
    os.environ["CURSOR_DEBUG_LOG"] = str(DEBUG_LOG)
    master_local = ROOT / "code" / "master.xlsm"
    if master_local.is_file():
        os.environ["PM_AI_MASTER_WORKBOOK"] = str(master_local)
    os.environ["MASTER_WORKBOOK_FILE"] = "master.xlsm"
    for key in (
        "RAW_FABRIC_WIDTH_TABLE_PATH",
        "ROLL_UNIT_BY_USED_RAW_TABLE_PATH",
        "PRODUCT_WIDTH_TABLE_PATH",
        "PRODUCT_LENGTH_TABLE_PATH",
        "PRODUCT_THICKNESS_TABLE_PATH",
    ):
        rel = {
            "RAW_FABRIC_WIDTH_TABLE_PATH": "code/使用原反, 加工幅.txt",
            "ROLL_UNIT_BY_USED_RAW_TABLE_PATH": "code/使用原反,ロール単位の長さ.txt",
            "PRODUCT_WIDTH_TABLE_PATH": "code/製品名, 製品幅.txt",
            "PRODUCT_LENGTH_TABLE_PATH": "code/製品名,製品長.txt",
            "PRODUCT_THICKNESS_TABLE_PATH": "code/製品名,製品厚み.txt",
        }[key]
        os.environ[key] = str(ROOT / rel)
    os.environ["PM_AI_EXCLUDE_RULES_JSON"] = str(ROOT / "code" / "json" / "stage1_exclude_rules.json")
    os.chdir(str(PY_DIR))
    sys.path.insert(0, str(PY_DIR))


def summarize_shortages(path: Path) -> dict:
    if not path.is_file():
        return {"ok": False, "reason": "shortages json missing"}
    data = json.loads(path.read_text(encoding="utf-8"))
    out = {
        "ok": True,
        "format_version": data.get("format_version"),
        "error": data.get("error"),
        "note": data.get("note"),
        "dispatch_qty_shortfall_count": len(data.get("dispatch_qty_shortfall") or []),
        "op_shortage_count": len(data.get("op_shortage") or []),
        "as_shortage_count": len(data.get("as_shortage") or []),
        "shortfalls": [],
        "task_ids": set(),
    }
    for row in data.get("dispatch_qty_shortfall") or []:
        tid = str(row.get("task_id") or "").strip()
        out["task_ids"].add(tid)
        out["shortfalls"].append(
            {
                "task_id": tid,
                "machine": row.get("machine_name"),
                "date": row.get("dispatch_date"),
                "shortfall_m": row.get("shortfall_m"),
                "note": row.get("note"),
            }
        )
    out["task_ids"] = sorted(t for t in out["task_ids"] if t)
    return out


def main() -> int:
    apply_env()
    if not DISPATCH_JSON.is_file():
        print(f"missing dispatch json: {DISPATCH_JSON}", file=sys.stderr)
        return 2
    from planning_core.stage2_identical_dispatch_runner import (
        run_interactive_dispatch_trial_from_result_dispatch_json,
    )

    print(f"[iterate] input={DISPATCH_JSON}", flush=True)
    code, shortage_path = run_interactive_dispatch_trial_from_result_dispatch_json(
        DISPATCH_JSON
    )
    summary = summarize_shortages(SHORTAGES_JSON)
    out_path = ROOT / "code" / "output" / "stage3_iterate_summary.json"
    out_path.write_text(
        json.dumps({"exit_code": code, "summary": summary}, ensure_ascii=False, indent=2)
        + "\n",
        encoding="utf-8",
    )
    print(json.dumps({"exit_code": code, "summary": summary}, ensure_ascii=False, indent=2))
    if code != 0:
        return code
    if summary.get("error"):
        return 3
    if summary.get("dispatch_qty_shortfall_count", 0) > 0:
        return 4
    if summary.get("op_shortage_count", 0) > 0 or summary.get("as_shortage_count", 0) > 0:
        return 5
    return 0


if __name__ == "__main__":
    sys.exit(main())
