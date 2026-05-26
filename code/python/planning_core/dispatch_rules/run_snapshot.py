"""Run snapshot capture for stage 1-3.5."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from planning_core.dispatch_rules.paths import default_work_json_path


def _snapshots_dir(work_json: Path) -> Path:
    return work_json.parent / "run_snapshots"


def _index_path(work_json: Path) -> Path:
    return _snapshots_dir(work_json) / "index.json"


def max_snapshots() -> int:
    raw = os.environ.get("PM_AI_DISPATCH_RULE_RUN_SNAPSHOT_MAX", "20").strip()
    try:
        return max(1, int(raw))
    except ValueError:
        return 20


def capture_run_snapshot(stage: str, work_path: Path | None = None) -> dict[str, Any]:
    work = work_path or default_work_json_path()
    if work is None or not work.is_file():
        return {"run_id": "", "path": ""}
    work = work.resolve()
    snaps = _snapshots_dir(work)
    snaps.mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone.utc)
    digest = hashlib.sha256(work.read_bytes()).hexdigest()[:8]
    run_id = f"{stage}_{now.strftime('%Y%m%d-%H%M%S')}_{digest}"
    target = snaps / f"{run_id}.json"
    shutil.copy2(work, target)
    entry = {
        "run_id": run_id,
        "stage": stage,
        "capturedAt": now.isoformat(),
        "sourceHash": digest,
        "path": str(target),
    }
    index_path = _index_path(work)
    if index_path.is_file():
        index = json.loads(index_path.read_text(encoding="utf-8"))
    else:
        index = {"version": 1, "entries": []}
    entries = index.get("entries") or []
    entries.insert(0, entry)
    while len(entries) > max_snapshots():
        old = entries.pop()
        old_path = Path(str(old.get("path", "")))
        if old_path.is_file():
            old_path.unlink(missing_ok=True)
    index["entries"] = entries
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    os.environ["PM_AI_DISPATCH_SPECIAL_RULES_JSON"] = str(target)
    from planning_core.dispatch_rules import trace_recorder

    trace_recorder.set_active_run_snapshot(run_id)
    return entry
