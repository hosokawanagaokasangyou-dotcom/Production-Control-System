"""Edit history snapshots."""

from __future__ import annotations

import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from planning_core.dispatch_rules.history_diff import diff_summary
from planning_core.dispatch_rules.paths import DISPATCH_SPECIAL_RULES_DIR, default_work_json_path


def _history_dir(work_json: Path) -> Path:
    return work_json.parent / "history"


def _index_path(work_json: Path) -> Path:
    return _history_dir(work_json) / "index.json"


def _snapshots_dir(work_json: Path) -> Path:
    return _history_dir(work_json) / "snapshots"


def _backups_dir(work_json: Path) -> Path:
    return work_json.parent / "backups"


def max_entries() -> int:
    raw = os.environ.get("PM_AI_DISPATCH_RULE_HISTORY_MAX", "50").strip()
    try:
        return max(1, int(raw))
    except ValueError:
        return 50


def append_snapshot(work_json: Path, *, kind: str = "auto_save", label: str = "保存") -> dict[str, Any]:
    work_json = work_json.resolve()
    hist = _history_dir(work_json)
    snaps = _snapshots_dir(work_json)
    hist.mkdir(parents=True, exist_ok=True)
    snaps.mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone.utc)
    entry_id = now.strftime("%Y%m%d-%H%M%S")
    snap_name = f"{entry_id}_{kind}.json"
    snap_path = snaps / snap_name
    shutil.copy2(work_json, snap_path)
    content = json.loads(work_json.read_text(encoding="utf-8"))
    summary = diff_summary(work_json, snap_path)
    entry = {
        "id": entry_id,
        "kind": kind,
        "label": label,
        "savedAt": now.isoformat(),
        "schemaVersion": content.get("schemaVersion", 1),
        "snapshotFile": f"snapshots/{snap_name}",
        "summary": summary,
    }
    index_path = _index_path(work_json)
    if index_path.is_file():
        index = json.loads(index_path.read_text(encoding="utf-8"))
    else:
        index = {"version": 1, "maxEntries": max_entries(), "entries": []}
    entries = index.get("entries") or []
    entries.insert(0, entry)
    limit = int(index.get("maxEntries") or max_entries())
    while len(entries) > limit:
        old = entries.pop()
        old_file = hist / str(old.get("snapshotFile", ""))
        if old_file.is_file():
            old_file.unlink(missing_ok=True)
    index["entries"] = entries
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    return entry


def restore_snapshot(work_json: Path, entry_id: str) -> None:
    work_json = work_json.resolve()
    index = json.loads(_index_path(work_json).read_text(encoding="utf-8"))
    entry = next((e for e in index.get("entries") or [] if e.get("id") == entry_id), None)
    if entry is None:
        raise FileNotFoundError(entry_id)
    backups = _backups_dir(work_json)
    backups.mkdir(parents=True, exist_ok=True)
    guard_name = f"restore_guard_{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')}.json"
    if work_json.is_file():
        shutil.copy2(work_json, backups / guard_name)
        append_snapshot(work_json, kind="auto_restore_guard", label="復元前の自動退避")
    snap = _history_dir(work_json) / str(entry["snapshotFile"])
    shutil.copy2(snap, work_json)
