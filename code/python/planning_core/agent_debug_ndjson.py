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
import sys
import time
from pathlib import Path
from typing import Any

_ENV_DEBUG_LOG_KEYS = ("PM_AI_DEBUG_LOG", "CURSOR_DEBUG_LOG")
_ENV_CURSOR_DEBUG_LOG = "PM_AI_CURSOR_DEBUG_LOG"
_ENV_MIRROR = "PM_AI_DEBUG_LOG_MIRROR"
_ENV_SESSION = "PM_AI_AGENT_DEBUG_SESSION"
_ENV_REPO_ROOT = "PM_AI_REPO_ROOT"
_ENV_WORKSPACE = "PM_AI_WORKSPACE"
_DEFAULT_SESSION_ID = "e04a1d"


def _log_path() -> str | None:
    for key in (_ENV_CURSOR_DEBUG_LOG,) + _ENV_DEBUG_LOG_KEYS:
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


def _cursor_debug_directory_root() -> Path | None:
    """Java ``AgentDebugLog.resolveCursorDebugDirectoryRoot`` と同趣旨。"""
    ws = (os.environ.get(_ENV_WORKSPACE) or "").strip()
    if ws:
        try:
            p = Path(ws).resolve()
            if p.is_dir():
                return p
        except OSError:
            pass
    for repo in _repo_root_candidates():
        leaf = repo.name.lower()
        if leaf in ("code_java", "production-control-system") and repo.parent is not None:
            return repo.parent.resolve()
        return repo.resolve()
    return None


def resolve_ndjson_path() -> str | None:
    """Java ``AgentDebugLog.resolveNdjsonPath`` と同趣旨のパス（書き込み先候補）。"""
    explicit = _log_path()
    if explicit:
        return explicit

    sid = session_id()
    file_name = f"debug-{sid}.log"
    cursor_root = _cursor_debug_directory_root()
    candidates: list[str] = []
    if cursor_root is not None:
        candidates.append(str(cursor_root / ".cursor" / file_name))
    for repo in _repo_root_candidates():
        candidates.append(str(repo / ".cursor" / file_name))

    for c in candidates:
        parent_dir = os.path.dirname(c)
        if not parent_dir:
            continue
        try:
            os.makedirs(parent_dir, exist_ok=True)
            return c
        except OSError:
            continue
    return None


def _write_targets() -> list[str]:
    """一次パス・ミラー・リポジトリ固定サイドカー（エージェントが必ず読める経路）。"""
    out: list[str] = []
    seen: set[str] = set()

    def _add(path: str | None) -> None:
        if not path:
            return
        p = os.path.abspath(path)
        if p in seen:
            return
        seen.add(p)
        out.append(p)

    _add(resolve_ndjson_path())
    _add((os.environ.get(_ENV_MIRROR) or "").strip() or None)

    sid = session_id()
    for repo in _repo_root_candidates():
        _add(str(repo / ".cursor" / f"debug-{sid}.log"))
        _add(str(repo / "log" / "agent_debug_latest.ndjson"))
        _add(str(repo / "log" / f"agent_debug_{sid}.ndjson"))

    return out


def _append_line(path: str, json_line: str) -> bool:
    parent_dir = os.path.dirname(path)
    if not parent_dir:
        return False
    try:
        os.makedirs(parent_dir, exist_ok=True)
        with open(path, "a", encoding="utf-8", newline="\n") as f:
            f.write(json_line)
        return True
    except OSError:
        return False


def append_structured(
    hypothesis_id: str,
    location: str,
    message: str,
    data: dict[str, Any] | None = None,
) -> None:
    targets = _write_targets()
    payload = dict(data or {})
    if targets and "ndjson_path" not in payload:
        payload["ndjson_path"] = targets[0]
    payload.setdefault("write_targets", targets[:8])
    line_obj = {
        "sessionId": session_id(),
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": payload,
        "timestamp": int(time.time() * 1000),
    }
    json_line = json.dumps(line_obj, ensure_ascii=False) + "\n"
    wrote_any = False
    for path in targets:
        if _append_line(path, json_line):
            wrote_any = True
    if not wrote_any:
        try:
            print(
                f"[agent-debug-write-failed] {message} targets={targets!r}",
                file=sys.stderr,
                flush=True,
            )
        except Exception:
            pass
