# -*- coding: utf-8 -*-
"""配台試行など Python 子プロセスからの NDJSON 追記。

Java の ``MainShellController#snapshotDispatchTrialPythonEnv`` が付与する
``PM_AI_DEBUG_LOG`` / ``PM_AI_AGENT_DEBUG_SESSION`` と揃え、
``AgentDebugLog.appendStructured`` と同一形式の 1 行 JSON を追記する。

パス解決は ``AgentDebugLog.resolveNdjsonPath`` と同趣旨（環境変数 → リポジトリ
``.cursor/debug-<session>.log``）。OS 固定の ``/mnt/c/...`` は使わない。
"""

from __future__ import annotations

import json
import os
import time
from pathlib import Path
from typing import Any

_ENV_DEBUG_LOG_KEYS = ("PM_AI_DEBUG_LOG", "CURSOR_DEBUG_LOG")
_ENV_SESSION = "PM_AI_AGENT_DEBUG_SESSION"
_ENV_REPO_ROOT = "PM_AI_REPO_ROOT"
_ENV_WORKSPACE = "PM_AI_WORKSPACE"
_DEFAULT_SESSION_ID = "e04a1d"
_NESTED_REPO_LEAF = "production-control-system"


def _log_path() -> str | None:
    for key in _ENV_DEBUG_LOG_KEYS:
        p = (os.environ.get(key) or "").strip()
        if p:
            return p
    return None


def session_id() -> str:
    for key in (_ENV_SESSION, "CURSOR_DEBUG_SESSION_ID"):
        s = (os.environ.get(key) or "").strip()
        if s:
            return s
    return _DEFAULT_SESSION_ID


def _repo_root_candidates() -> list[Path]:
    out: list[Path] = []
    repo = (os.environ.get(_ENV_REPO_ROOT) or "").strip()
    if repo:
        out.append(Path(repo).resolve())
    try:
        # .../code/python/planning_core/agent_debug_ndjson.py → repo root
        out.append(Path(__file__).resolve().parents[3])
    except (IndexError, OSError):
        pass
    dedup: list[Path] = []
    seen: set[str] = set()
    for p in out:
        key = str(p)
        if key in seen:
            continue
        seen.add(key)
        dedup.append(p)
    return dedup


def resolve_ndjson_path() -> str | None:
    """Java ``AgentDebugLog.resolveNdjsonPath`` と同趣旨のパス（書き込み先候補）。"""
    explicit = _log_path()
    if explicit:
        return explicit

    sid = session_id()
    file_name = f"debug-{sid}.log"
    candidates: list[str] = []

    ws = (os.environ.get(_ENV_WORKSPACE) or "").strip()
    if ws:
        candidates.append(str(Path(ws).resolve() / ".cursor" / file_name))

    for repo in _repo_root_candidates():
        if repo.name.lower() == _NESTED_REPO_LEAF:
            parent = repo.parent
            if parent is not None:
                candidates.append(str(parent / ".cursor" / file_name))
        candidates.append(str(repo / ".cursor" / file_name))

    for c in candidates:
        parent_dir = os.path.dirname(c)
        if parent_dir and os.path.isdir(parent_dir):
            return c
    if candidates:
        return candidates[0]
    return None


def append_structured(
    hypothesis_id: str,
    location: str,
    message: str,
    data: dict[str, Any] | None = None,
) -> None:
    path = resolve_ndjson_path()
    if not path:
        return
    line = {
        "sessionId": session_id(),
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data or {},
        "timestamp": int(time.time() * 1000),
    }
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(line, ensure_ascii=False) + "\n")
    except OSError:
        pass


def y5_21_json_row_sample(rows: list | None, limit: int = 12) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        blob = json.dumps(r, ensure_ascii=False)
        if "Y5-21" not in blob:
            continue
        slim = {k: r.get(k) for k in list(r.keys())[:24]}
        out.append(slim)
        if len(out) >= limit:
            break
    return out


def y5_21_targets_sample(targets: dict | None, limit: int = 24) -> dict[str, float]:
    out: dict[str, float] = {}
    for k, v in (targets or {}).items():
        if "Y5-21" not in str(k):
            continue
        try:
            out[str(k)] = float(v)
        except (TypeError, ValueError):
            out[str(k)] = float("nan")
        if len(out) >= limit:
            break
    return out
