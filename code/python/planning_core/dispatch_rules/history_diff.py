"""Diff summary between rule documents."""

from __future__ import annotations

import json
from pathlib import Path


def diff_summary(before_path: Path, after_path: Path) -> str:
    try:
        before = json.loads(before_path.read_text(encoding="utf-8"))
        after = json.loads(after_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return "変更"
    parts: list[str] = []
    before_rules = {r.get("id"): r for r in before.get("rules") or [] if isinstance(r, dict)}
    after_rules = {r.get("id"): r for r in after.get("rules") or [] if isinstance(r, dict)}
    for rid, ar in after_rules.items():
        br = before_rules.get(rid)
        if br is None:
            parts.append(f"{rid} 追加")
            continue
        if br.get("enabled") != ar.get("enabled"):
            parts.append(f"{rid} {'ON' if ar.get('enabled') else 'OFF'}")
        if br.get("applyOrder") != ar.get("applyOrder"):
            parts.append(f"{rid} 順序変更")
    if not parts:
        return "変更なし"
    return "、".join(parts[:5])
