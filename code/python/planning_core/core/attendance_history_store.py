# -*- coding: utf-8 -*-
"""勤怠正本 JSON（attendance-data.json）の世代管理（専用フォルダ・最大20世代）。"""

from __future__ import annotations

import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from planning_core.core.attendance_paths import (
    ATTENDANCE_HISTORY_DIR_NAME,
    ENV_ATTENDANCE_HISTORY_DIR,
    ENV_ATTENDANCE_HISTORY_MAX,
    attendance_data_json_path,
)

DEFAULT_MAX_ENTRIES = 20
INDEX_VERSION = 1


def max_entries() -> int:
    raw = os.environ.get(ENV_ATTENDANCE_HISTORY_MAX, "").strip()
    if not raw:
        return DEFAULT_MAX_ENTRIES
    try:
        return max(1, min(20, int(raw)))
    except ValueError:
        return DEFAULT_MAX_ENTRIES


def attendance_history_root(json_path: Path | None = None) -> Path:
    explicit = os.environ.get(ENV_ATTENDANCE_HISTORY_DIR, "").strip()
    if explicit:
        return Path(explicit).resolve()
    jp = (json_path or attendance_data_json_path()).resolve()
    return jp.parent / ATTENDANCE_HISTORY_DIR_NAME


def _history_dir(json_path: Path) -> Path:
    return attendance_history_root(json_path)


def _index_path(json_path: Path) -> Path:
    return _history_dir(json_path) / "index.json"


def _snapshots_dir(json_path: Path) -> Path:
    return _history_dir(json_path) / "snapshots"


def _backups_dir(json_path: Path) -> Path:
    return _history_dir(json_path) / "pre-restore-guards"


def _read_index(json_path: Path) -> dict[str, Any]:
    index_path = _index_path(json_path)
    if index_path.is_file():
        data = json.loads(index_path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    return {
        "version": INDEX_VERSION,
        "maxEntries": max_entries(),
        "entries": [],
    }


def _write_index(json_path: Path, index: dict[str, Any]) -> None:
    hist = _history_dir(json_path)
    hist.mkdir(parents=True, exist_ok=True)
    index_path = _index_path(json_path)
    index_path.write_text(
        json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _entry_summary_from_store(content: dict[str, Any]) -> dict[str, Any]:
    meta = content.get("meta") or {}
    return {
        "company_calendar_revision": int(meta.get("company_calendar_revision") or 0),
        "member_attendance_revision": int(meta.get("member_attendance_revision") or 0),
        "updated_at": meta.get("updated_at"),
    }


def append_attendance_snapshot(
    json_path: Path | None = None,
    *,
    kind: str = "auto_save",
    label: str = "保存",
) -> dict[str, Any] | None:
    """
    現行 attendance-data.json を世代フォルダへ退避する。
    ファイルが無いときは何もしない。
    """
    jp = (json_path or attendance_data_json_path()).resolve()
    if not jp.is_file():
        return None

    hist = _history_dir(jp)
    snaps = _snapshots_dir(jp)
    hist.mkdir(parents=True, exist_ok=True)
    snaps.mkdir(parents=True, exist_ok=True)

    now = datetime.now(timezone.utc)
    entry_id = now.strftime("%Y%m%d-%H%M%S-%f")
    safe_kind = (kind or "auto_save").replace("/", "_").replace("\\", "_")[:32]
    snap_name = f"{entry_id}_{safe_kind}.json"
    snap_path = snaps / snap_name
    shutil.copy2(jp, snap_path)

    try:
        content = json.loads(jp.read_text(encoding="utf-8"))
        if not isinstance(content, dict):
            content = {}
    except (OSError, json.JSONDecodeError):
        content = {}

    entry: dict[str, Any] = {
        "id": entry_id,
        "kind": kind,
        "label": label,
        "savedAt": now.isoformat(),
        "format_version": content.get("format_version", 1),
        "snapshotFile": f"snapshots/{snap_name}",
        **(_entry_summary_from_store(content)),
    }

    index = _read_index(jp)
    limit = int(index.get("maxEntries") or max_entries())
    limit = min(20, max(1, limit))
    index["maxEntries"] = limit
    entries = list(index.get("entries") or [])
    entries.insert(0, entry)
    while len(entries) > limit:
        old = entries.pop()
        old_file = hist / str(old.get("snapshotFile", ""))
        if old_file.is_file():
            old_file.unlink(missing_ok=True)
    index["entries"] = entries
    _write_index(jp, index)
    return entry


def list_attendance_history(json_path: Path | None = None) -> dict[str, Any]:
    jp = (json_path or attendance_data_json_path()).resolve()
    index = _read_index(jp)
    limit = min(20, max(1, int(index.get("maxEntries") or max_entries())))
    entries = list(index.get("entries") or [])
    return {
        "history_dir": str(_history_dir(jp)),
        "max_entries": limit,
        "entries": entries,
    }


def restore_attendance_snapshot(
    entry_id: str,
    json_path: Path | None = None,
) -> Path:
    jp = (json_path or attendance_data_json_path()).resolve()
    index = _read_index(jp)
    entry = next(
        (e for e in index.get("entries") or [] if e.get("id") == entry_id),
        None,
    )
    if entry is None:
        raise FileNotFoundError(entry_id)

    hist = _history_dir(jp)
    snap = hist / str(entry["snapshotFile"])
    if not snap.is_file():
        raise FileNotFoundError(str(snap))

    backups = _backups_dir(jp)
    backups.mkdir(parents=True, exist_ok=True)
    guard_name = (
        f"restore_guard_{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')}.json"
    )
    if jp.is_file():
        shutil.copy2(jp, backups / guard_name)
        append_attendance_snapshot(
            jp,
            kind="auto_restore_guard",
            label="復元前の自動退避",
        )

    jp.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(snap, jp)
    return jp
