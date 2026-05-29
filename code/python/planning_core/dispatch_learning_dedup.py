# -*- coding: utf-8 -*-
"""学習アーカイブ・速度観測の重複排除レジストリ。"""
from __future__ import annotations

import hashlib
import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

REGISTRY_FILENAME = "dedup_registry.json"
SEEN_OBSERVATION_IDS_FILENAME = "seen_observation_ids.json"


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def registry_path(archive_root: str | Path) -> Path:
    return Path(archive_root) / REGISTRY_FILENAME


def seen_observation_ids_path(archive_root: str | Path) -> Path:
    return Path(archive_root) / "speed-distributions" / SEEN_OBSERVATION_IDS_FILENAME


def _empty_registry() -> dict[str, Any]:
    return {
        "version": 1,
        "jobs": {},
        "archive_fingerprints": {},
        "speed_source_files": {},
    }


def load_registry(archive_root: str | Path) -> dict[str, Any]:
    p = registry_path(archive_root)
    if not p.is_file():
        return _empty_registry()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            data.setdefault("version", 1)
            data.setdefault("jobs", {})
            data.setdefault("archive_fingerprints", {})
            data.setdefault("speed_source_files", {})
            return data
    except (OSError, json.JSONDecodeError):
        pass
    return _empty_registry()


def save_registry(archive_root: str | Path, data: dict[str, Any]) -> None:
    p = registry_path(archive_root)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp, p)


def content_fingerprint(*parts: bytes | str) -> str:
    h = hashlib.sha256()
    for part in parts:
        if isinstance(part, str):
            part = part.encode("utf-8")
        h.update(part)
        h.update(b"\x1e")
    return "sha256:" + h.hexdigest()


def is_job_archived(archive_root: str | Path, job_id: str) -> bool:
    if not job_id:
        return False
    reg = load_registry(archive_root)
    return job_id in (reg.get("jobs") or {})


def is_fingerprint_archived(archive_root: str | Path, fingerprint: str) -> bool:
    if not fingerprint:
        return False
    reg = load_registry(archive_root)
    return fingerprint in (reg.get("archive_fingerprints") or {})


def register_job(
    archive_root: str | Path,
    job_id: str,
    *,
    folder: str,
    fingerprint: str | None = None,
) -> None:
    reg = load_registry(archive_root)
    jobs = reg.setdefault("jobs", {})
    jobs[job_id] = {
        "folder": folder,
        "fingerprint": fingerprint or "",
        "archived_at": _utc_now_iso(),
    }
    if fingerprint:
        reg.setdefault("archive_fingerprints", {})[fingerprint] = job_id
    save_registry(archive_root, reg)


def load_seen_observation_ids(archive_root: str | Path) -> set[str]:
    p = seen_observation_ids_path(archive_root)
    if not p.is_file():
        return set()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return {str(x) for x in data}
        if isinstance(data, dict) and "ids" in data:
            return {str(x) for x in (data.get("ids") or [])}
    except (OSError, json.JSONDecodeError):
        pass
    return set()


def save_seen_observation_ids(archive_root: str | Path, ids: set[str]) -> None:
    p = seen_observation_ids_path(archive_root)
    p.parent.mkdir(parents=True, exist_ok=True)
    payload = {"version": 1, "ids": sorted(ids)}
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp, p)


def is_observation_seen(archive_root: str | Path, observation_id: str) -> bool:
    return observation_id in load_seen_observation_ids(archive_root)


def register_observation_ids(archive_root: str | Path, observation_ids: list[str]) -> int:
    seen = load_seen_observation_ids(archive_root)
    before = len(seen)
    seen.update(observation_ids)
    save_seen_observation_ids(archive_root, seen)
    return len(seen) - before


def observation_id_from_fields(
    task_id: str,
    process: str,
    machine: str,
    start_iso: str,
    end_iso: str,
    actual_qty_m: float,
) -> str:
    qty = round(float(actual_qty_m), 4)
    raw = "|".join(
        [
            (task_id or "").strip(),
            (process or "").strip(),
            (machine or "").strip(),
            (start_iso or "").strip(),
            (end_iso or "").strip(),
            str(qty),
        ]
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def speed_source_fingerprint(path: str | Path) -> str:
    p = Path(path)
    try:
        st = p.stat()
        raw = f"{p.resolve()}|{st.st_size}|{int(st.st_mtime_ns)}"
    except OSError:
        raw = str(path)
    return content_fingerprint(raw)
