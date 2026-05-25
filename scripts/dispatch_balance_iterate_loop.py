#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""段階3配台照合が OK になるまで最大 N 回パイプラインを実行（各回は手動修正前提）。"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "dispatch_balance_auto_iterate.py"
SUMMARY = ROOT / "code" / "output" / "dispatch_balance_iterate_summary.json"


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--task", required=True)
    p.add_argument("--process", default=None)
    p.add_argument("--max-iter", type=int, default=10)
    p.add_argument("--skip-stage1", action="store_true")
    p.add_argument("--skip-stage2", action="store_true")
    args = p.parse_args()

    for n in range(1, args.max_iter + 1):
        print(f"\n========== 反復 {n}/{args.max_iter} task={args.task} ==========", flush=True)
        cmd = [
            sys.executable,
            str(SCRIPT),
            "--task",
            args.task,
        ]
        if args.process:
            cmd.extend(["--process", args.process])
        if args.skip_stage1:
            cmd.append("--skip-stage1")
        if args.skip_stage2:
            cmd.append("--skip-stage2")
        proc = subprocess.run(cmd, cwd=str(ROOT / "code" / "python"))
        if SUMMARY.is_file():
            summary = json.loads(SUMMARY.read_text(encoding="utf-8"))
            bal = summary.get("balance") or {}
            print(f"照合: {bal.get('check', '?')}", flush=True)
            if bal.get("ok"):
                print(f"{args.task} 段階3配台照合 OK（反復 {n}）", flush=True)
                return 0
        if proc.returncode == 0:
            return 0
        if n >= args.max_iter:
            print(f"{args.task} は {args.max_iter} 反復後も NG", flush=True)
            return proc.returncode or 4
        print("NG — ロジック修正後に再実行してください", flush=True)
        return proc.returncode or 4
    return 4


if __name__ == "__main__":
    sys.exit(main())
