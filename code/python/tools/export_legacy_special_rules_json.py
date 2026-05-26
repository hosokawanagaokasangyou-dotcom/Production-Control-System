#!/usr/bin/env python3
"""Export legacy special rules to dispatch_special_rules.json seed."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

TEMPLATE = Path(__file__).resolve().parents[2] / "json" / "dispatch_special_rules" / "dispatch_special_rules.json"

LEGACY_RULE_IDS = [
    ("L2", 10, "スライス3名"),
    ("L3", 20, "スライス4名"),
    ("L4", 30, "SEC 製品幅935 速度20"),
    ("L5", 40, "L5 速度"),
    ("L6", 50, "L6 速度"),
    ("L7", 60, "NEED 3名"),
    ("L8", 70, "L8 速度"),
    ("L9", 80, "L9 優先"),
    ("L10", 90, "スリットSEC WIP"),
    ("L11", 100, "L11 WIP"),
    ("L12", 110, "L12 優先"),
    ("L13", 120, "接続→SEC WIP"),
]


def build_export() -> dict:
    if TEMPLATE.is_file():
        base = json.loads(TEMPLATE.read_text(encoding="utf-8"))
    else:
        base = {"schemaVersion": 1, "engineMinVersion": "1.0.0", "rules": []}
    existing = {r["id"]: r for r in base.get("rules") or []}
    rules = []
    for rid, order, name in LEGACY_RULE_IDS:
        if rid in existing:
            rule = dict(existing[rid])
        else:
            rule = {
                "id": rid,
                "name": name,
                "enabled": True,
                "applyOrder": order,
                "executionMode": "legacy",
                "legacyFallback": True,
                "graph": {"nodes": [], "edges": []},
            }
        rule["applyOrder"] = order
        rule.setdefault("executionMode", "legacy")
        rules.append(rule)
    base["rules"] = rules
    return base


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    doc = build_export()
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"wrote {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
