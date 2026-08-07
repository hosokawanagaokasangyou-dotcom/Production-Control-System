# -*- coding: utf-8 -*-
"""machine-calendar-data.json の世代管理。"""

from __future__ import annotations

import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from planning_core.core.machine_calendar_paths import (
    ENV_MACHINE_CALENDAR_HISTORY_DIR,
    ENV_MACHINE_CALENDAR_HISTORY_MAX,
    MACHINE_CALENDAR_HISTORY_DIR_NAME,
    machine_calendar_data_json_path,
)

DEFAULT_MAX_ENTRIES = 20
INDEX_VERSION = 1


def max_entries() -> int:
    raw = os.environ.get(ENV_MACHINE_CALENDAR_HISTORY_MAX, "").strip()
    if not raw:
        return DEFAULT_MAX_ENTRIES
    try:
        return max(1, min(20, int(raw)))
    except ValueError:
        return DEFAULT_MAX_ENTRIES


def history_root(json_path: Path | None = None) -> Path:
    explicit = os.environ.get(ENV_MACHINE_CALENDAR_HISTORY_DIR, "").strip()
    if explicit:
        return Path(explicit).resolve()
    jp = (json_path or machine_calendar_data_json_path()).resolve()
    return jp.parent / MACHINE_CALENDAR_HISTORY_DIR_NAME


def _index_path(json_path: Path) -> Path:
    return history_root(json_path) / "index.json"


def _snapshots_dir(json_path: Path) -> Path:
    return history_root(json_path) / "snapshots"


def _read_index(json_path: Path) -> dict[str, Any]:
    index_path = _index_path(json_path)
    if index_path.is_file():
        data = json.loads(index_path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    return {"version": INDEX_VERSION, "maxEntries": max_entries(), "entries": []}


def _write_index(json_path: Path, index: dict[str, Any]) -> None:
    hist = history_root(json_path)
    hist.mkdir(parents=True, exist_ok=True)
    _index_path(json_path).write_text(
        json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def append_machine_calendar_snapshot(
    json_path: Path,
    *,
    kind: str,
    label: str,
) -> None:
    jp = json_path.resolve()
    if not jp.is_file():
        return
    hist = history_root(jp)
    hist.mkdir(parents=True, exist_ok=True)
    _snapshots_dir(jp).mkdir(parents=True, exist_ok=True)
    index = _read_index(jp)
    entry_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%f")
    snap_name = f"{entry_id}.json"
    shutil.copy2(jp, _snapshots_dir(jp) / snap_name)
    entries = index.setdefault("entries", [])
    entries.insert(
        0,
        {
            "id": entry_id,
            "kind": kind,
            "label": label,
            "savedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "snapshot": snap_name,
        },
    )
    max_n = int(index.get("maxEntries") or max_entries())
    while len(entries) > max_n:
        old = entries.pop()
        old_path = _snapshots_dir(jp) / str(old.get("snapshot") or "")
        if old_path.is_file():
            old_path.unlink()
    index["maxEntries"] = max_n
    _write_index(jp, index)
