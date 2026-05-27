#!/usr/bin/env python3
"""Run dispatch rule simulation for Java test lab (stdin JSON → stdout JSON)."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--stdin", action="store_true", help="Read request JSON from stdin")
    parser.add_argument("--request", type=Path, help="Request JSON file")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    if args.stdin:
        payload = json.loads(sys.stdin.read())
    elif args.request:
        payload = json.loads(args.request.read_text(encoding="utf-8"))
    else:
        print("usage: simulate_dispatch_rules.py --stdin", file=sys.stderr)
        return 2
    document = payload.get("document") or {}
    task_row = payload.get("task_row") or {}
    rule_id = payload.get("rule_id")
    overrides = payload.get("context_overrides")
    sec_task_row = payload.get("sec_task_row")
    all_rolls = bool(payload.get("all_rolls"))
    from planning_core.dispatch_rules.migrations import migrate_document
    from planning_core.dispatch_rules.simulation import (
        simulate_connection_sec_pipeline,
        simulate_task,
        simulate_task_all_rolls,
    )

    migrated = migrate_document(document)
    # planning_core bootstrap attaches a stdout log handler; keep simulate JSON parseable.
    import logging

    for handler in list(logging.getLogger().handlers):
        if isinstance(handler, logging.StreamHandler) and handler.stream is sys.stdout:
            logging.getLogger().removeHandler(handler)
    if all_rolls and sec_task_row:
        result = simulate_connection_sec_pipeline(
            migrated,
            task_row,
            sec_task_row,
            rule_id=rule_id,
            context_overrides=overrides,
        )
    elif all_rolls:
        result = simulate_task_all_rolls(
            migrated,
            task_row,
            rule_id=rule_id,
            context_overrides=overrides,
        )
    else:
        result = simulate_task(
            migrated,
            task_row,
            rule_id=rule_id,
            context_overrides=overrides,
        )
    out = {
        "final_blocked": result.final_blocked,
        "summary_ja": result.summary_ja,
        "roll_total": result.roll_total,
        "blocked_at_roll": result.blocked_at_roll,
        "steps": [asdict(s) for s in result.steps],
    }
    print(json.dumps(out, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
