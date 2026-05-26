#!/usr/bin/env python3
"""Validate dispatch special rules JSON."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path)
    parser.add_argument("--conflicts", action="store_true")
    args = parser.parse_args()
    json_path = args.path.resolve()
    raw = json.loads(json_path.read_text(encoding="utf-8"))
    root = Path(__file__).resolve().parents[1]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    from planning_core.dispatch_rules.conflict_checker import check_rule_conflicts
    from planning_core.dispatch_rules.migrations import migrate_document

    migrated = migrate_document(raw)
    if args.conflicts:
        report = check_rule_conflicts(migrated)
        print(
            json.dumps(
                {
                    "errors": report.error_count,
                    "warnings": report.warning_count,
                    "conflicts": [
                        {
                            "kind": c.kind,
                            "severity": c.severity,
                            "rule_ids": c.rule_ids,
                            "message": c.message,
                        }
                        for c in report.conflicts
                    ],
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    print("ok schemaVersion", migrated.get("schemaVersion"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
