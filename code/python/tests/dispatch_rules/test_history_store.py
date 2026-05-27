"""Tests for history_store."""

import json
from pathlib import Path

from planning_core.dispatch_rules.history_store import append_snapshot, restore_snapshot


def test_append_and_restore_snapshot(tmp_path):
    work = tmp_path / "dispatch_special_rules" / "dispatch_special_rules.json"
    work.parent.mkdir(parents=True)
    work.write_text(
        json.dumps({"schemaVersion": 1, "rules": []}, ensure_ascii=False),
        encoding="utf-8",
    )
    entry = append_snapshot(work, kind="auto_save", label="保存")
    assert entry["id"]
    work.write_text(
        json.dumps({"schemaVersion": 1, "rules": [{"id": "X"}]}, ensure_ascii=False),
        encoding="utf-8",
    )
    restore_snapshot(work, entry["id"])
    restored = json.loads(work.read_text(encoding="utf-8"))
    assert restored["rules"] == []
