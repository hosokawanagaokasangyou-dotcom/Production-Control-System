# -*- coding: utf-8 -*-
"""勤怠カレンダー.xlsx の世代管理（正本 xlsx と同階層・最大20世代）。"""

from __future__ import annotations

import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from planning_core.core.attendance_paths import (
    ATTENDANCE_CALENDAR_XLSX_HISTORY_MAX_DEFAULT,
    ENV_ATTENDANCE_CALENDAR_XLSX_HISTORY_DIR,
    ENV_ATTENDANCE_CALENDAR_XLSX_HISTORY_MAX,
    attendance_calendar_xlsx_history_root,
    attendance_calendar_xlsx_path,
)

DEFAULT_MAX_ENTRIES = ATTENDANCE_CALENDAR_XLSX_HISTORY_MAX_DEFAULT
INDEX_VERSION = 1


def max_entries() -> int:
    raw = os.environ.get(ENV_ATTENDANCE_CALENDAR_XLSX_HISTORY_MAX, "").strip()
    if not raw:
        return DEFAULT_MAX_ENTRIES
    try:
        return max(1, min(20, int(raw)))
    except ValueError:
        return DEFAULT_MAX_ENTRIES


def _history_dir(xlsx_path: Path) -> Path:
    return attendance_calendar_xlsx_history_root(xlsx_path)


def _index_path(xlsx_path: Path) -> Path:
    return _history_dir(xlsx_path) / "index.json"


def _snapshots_dir(xlsx_path: Path) -> Path:
    return _history_dir(xlsx_path) / "snapshots"


def _backups_dir(xlsx_path: Path) -> Path:
    return _history_dir(xlsx_path) / "pre-restore-guards"


def _read_index(xlsx_path: Path) -> dict[str, Any]:
    index_path = _index_path(xlsx_path)
    if index_path.is_file():
        data = json.loads(index_path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    return {
        "version": INDEX_VERSION,
        "maxEntries": max_entries(),
        "entries": [],
    }


def _write_index(xlsx_path: Path, index: dict[str, Any]) -> None:
    hist = _history_dir(xlsx_path)
    hist.mkdir(parents=True, exist_ok=True)
    index_path = _index_path(xlsx_path)
    index_path.write_text(
        json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def append_calendar_xlsx_snapshot(
    xlsx_path: Path | None = None,
    *,
    kind: str = "export_calendar",
    label: str = "勤怠カレンダー出力",
    store_meta: dict[str, Any] | None = None,
) -> dict[str, Any] | None:
    """現行 勤怠カレンダー.xlsx を世代フォルダへ退避する。ファイルが無いときは何もしない。"""
    xp = (xlsx_path or attendance_calendar_xlsx_path()).resolve()
    if not xp.is_file():
        return None

    hist = _history_dir(xp)
    snaps = _snapshots_dir(xp)
    hist.mkdir(parents=True, exist_ok=True)
    snaps.mkdir(parents=True, exist_ok=True)

    now = datetime.now(timezone.utc)
    entry_id = now.strftime("%Y%m%d-%H%M%S-%f")
    safe_kind = (kind or "export_calendar").replace("/", "_").replace("\\", "_")[:32]
    snap_name = f"{entry_id}_{safe_kind}.xlsx"
    snap_path = snaps / snap_name
    shutil.copy2(xp, snap_path)

    meta = store_meta or {}
    entry: dict[str, Any] = {
        "id": entry_id,
        "kind": kind,
        "label": label,
        "savedAt": now.isoformat(),
        "snapshotFile": f"snapshots/{snap_name}",
        "company_calendar_revision": int(meta.get("company_calendar_revision") or 0),
        "member_attendance_revision": int(meta.get("member_attendance_revision") or 0),
        "calendar_xlsx_export_at": meta.get("calendar_xlsx_export_at"),
    }

    index = _read_index(xp)
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
    _write_index(xp, index)
    return entry


def list_calendar_xlsx_history(xlsx_path: Path | None = None) -> dict[str, Any]:
    xp = (xlsx_path or attendance_calendar_xlsx_path()).resolve()
    index = _read_index(xp)
    limit = min(20, max(1, int(index.get("maxEntries") or max_entries())))
    entries = list(index.get("entries") or [])
    return {
        "calendar_xlsx_path": str(xp),
        "history_dir": str(_history_dir(xp)),
        "max_entries": limit,
        "entries": entries,
    }


def restore_calendar_xlsx_snapshot(
    entry_id: str,
    xlsx_path: Path | None = None,
) -> Path:
    xp = (xlsx_path or attendance_calendar_xlsx_path()).resolve()
    index = _read_index(xp)
    entry = next(
        (e for e in index.get("entries") or [] if e.get("id") == entry_id),
        None,
    )
    if entry is None:
        raise FileNotFoundError(entry_id)

    hist = _history_dir(xp)
    snap = hist / str(entry["snapshotFile"])
    if not snap.is_file():
        raise FileNotFoundError(str(snap))

    backups = _backups_dir(xp)
    backups.mkdir(parents=True, exist_ok=True)
    guard_name = (
        f"restore_guard_{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')}.xlsx"
    )
    if xp.is_file():
        shutil.copy2(xp, backups / guard_name)
        append_calendar_xlsx_snapshot(
            xp,
            kind="auto_restore_guard",
            label="復元前の自動退避",
        )

    xp.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(snap, xp)
    return xp
