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


def _backups_dir(json_path: Path) -> Path:
    return history_root(json_path) / "pre-restore-guards"


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


def _entry_summary_from_store(content: dict[str, Any]) -> dict[str, Any]:
    meta = content.get("meta") or {}
    return {
        "revision": int(meta.get("revision") or 0),
        "column_count": len(content.get("columns") or []),
        "occupancy_slot_count": len(content.get("occupancy") or {}),
        "updated_at": meta.get("updated_at"),
    }


def _snapshot_path(json_path: Path, entry: dict[str, Any]) -> Path:
    snap_name = str(entry.get("snapshot") or entry.get("snapshotFile") or "").strip()
    if not snap_name:
        raise FileNotFoundError("snapshot")
    if "/" in snap_name or "\\" in snap_name:
        return history_root(json_path) / snap_name
    return _snapshots_dir(json_path) / snap_name


def append_machine_calendar_snapshot(
    json_path: Path,
    *,
    kind: str,
    label: str,
) -> dict[str, Any] | None:
    jp = json_path.resolve()
    if not jp.is_file():
        return None
    hist = history_root(jp)
    hist.mkdir(parents=True, exist_ok=True)
    _snapshots_dir(jp).mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone.utc)
    entry_id = now.strftime("%Y%m%dT%H%M%S%f")
    snap_name = f"{entry_id}.json"
    shutil.copy2(jp, _snapshots_dir(jp) / snap_name)

    try:
        content = json.loads(jp.read_text(encoding="utf-8"))
        if not isinstance(content, dict):
            content = {}
    except (OSError, json.JSONDecodeError):
        content = {}

    index = _read_index(jp)
    entries = index.setdefault("entries", [])
    entry: dict[str, Any] = {
        "id": entry_id,
        "kind": kind,
        "label": label,
        "savedAt": now.isoformat(timespec="seconds"),
        "snapshot": snap_name,
        "format_version": content.get("format_version", 1),
        **(_entry_summary_from_store(content)),
    }
    entries.insert(0, entry)
    max_n = int(index.get("maxEntries") or max_entries())
    max_n = min(20, max(1, max_n))
    while len(entries) > max_n:
        old = entries.pop()
        old_path = _snapshots_dir(jp) / str(old.get("snapshot") or "")
        if old_path.is_file():
            old_path.unlink()
    index["maxEntries"] = max_n
    _write_index(jp, index)
    return entry


def list_machine_calendar_history(json_path: Path | None = None) -> dict[str, Any]:
    jp = (json_path or machine_calendar_data_json_path()).resolve()
    index = _read_index(jp)
    limit = min(20, max(1, int(index.get("maxEntries") or max_entries())))
    entries = list(index.get("entries") or [])
    return {
        "history_dir": str(history_root(jp)),
        "max_entries": limit,
        "entries": entries,
    }


def restore_machine_calendar_snapshot(
    entry_id: str,
    json_path: Path | None = None,
) -> Path:
    jp = (json_path or machine_calendar_data_json_path()).resolve()
    index = _read_index(jp)
    entry = next(
        (e for e in index.get("entries") or [] if e.get("id") == entry_id),
        None,
    )
    if entry is None:
        raise FileNotFoundError(entry_id)

    snap = _snapshot_path(jp, entry)
    if not snap.is_file():
        raise FileNotFoundError(str(snap))

    backups = _backups_dir(jp)
    backups.mkdir(parents=True, exist_ok=True)
    guard_name = (
        f"restore_guard_{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')}.json"
    )
    if jp.is_file():
        shutil.copy2(jp, backups / guard_name)
        append_machine_calendar_snapshot(
            jp,
            kind="auto_restore_guard",
            label="復元前の自動退避",
        )

    jp.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(snap, jp)
    return jp
