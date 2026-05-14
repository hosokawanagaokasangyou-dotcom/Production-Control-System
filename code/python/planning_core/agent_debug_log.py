# -*- coding: utf-8 -*-
"""
Agent デバッグ用 NDJSON 追記（Python 側）。

Java の ``jp.co.pm.ai.desktop.debug.AgentDebugLog`` と同一の解決順・フォールバック方針に揃える。
正本ルール: ``.cursor/rules/agent-debug-ndjson-logging.mdc`` / ``agent-debug-wsl-windows-mirror.mdc``。
"""

from __future__ import annotations

import json
import os
import pathlib
import re
import sys
import time
from typing import Any, Mapping

# Java AgentDebugLog.DEFAULT_SESSION_ID と一致
DEFAULT_AGENT_DEBUG_SESSION_ID = "e04a1d"

KEY_PM_AI_REPO_ROOT = "PM_AI_REPO_ROOT"
KEY_PM_AI_CURSOR_DEBUG_LOG = "PM_AI_CURSOR_DEBUG_LOG"
KEY_PM_AI_DEBUG_LOG_MIRROR = "PM_AI_DEBUG_LOG_MIRROR"

_PRODUCTION_CONTROL_SYSTEM = "Production-Control-System"


def _trim(s: str | None) -> str:
    return (s or "").strip()


def ui_env_from_os() -> dict[str, str]:
    """Java の ``collectUiEnv()`` 相当に近い最小マップ（ログパス解決用）。"""
    keys = (
        KEY_PM_AI_REPO_ROOT,
        KEY_PM_AI_CURSOR_DEBUG_LOG,
        KEY_PM_AI_DEBUG_LOG_MIRROR,
    )
    return {k: _trim(os.environ.get(k)) for k in keys if _trim(os.environ.get(k))}


def resolve_agent_debug_session_id() -> str:
    for k in ("PM_AI_AGENT_DEBUG_SESSION_ID", "CURSOR_DEBUG_SESSION_ID"):
        v = _trim(os.environ.get(k))
        if v:
            return v
    return DEFAULT_AGENT_DEBUG_SESSION_ID


def _resolve_repo_root(ui: Mapping[str, str] | None) -> pathlib.Path:
    u = dict(ui or {})
    r = _trim(u.get(KEY_PM_AI_REPO_ROOT))
    if r:
        return pathlib.Path(r).resolve()
    here = pathlib.Path(__file__).resolve()
    code = here.parents[2]
    repo = code.parent
    return repo if repo is not None else code


def _is_production_control_system_repo_leaf(repo: pathlib.Path) -> bool:
    try:
        return repo.name.lower() == _PRODUCTION_CONTROL_SYSTEM.lower()
    except Exception:
        return False


def resolve_ndjson_path(ui: Mapping[str, str] | None, session_id: str | None) -> pathlib.Path:
    """
    ``AgentDebugLog.resolveNdjsonPath`` と同一の first-hit 順。

    1. ``CURSOR_DEBUG_LOG`` / ``PM_AI_DEBUG_LOG``（ログファイルへの絶対パス）
    2. ``ui`` / 環境の ``PM_AI_CURSOR_DEBUG_LOG``
    3. ネストクローン: ``parent(repo)/.cursor/debug-<session>.log``
    4. フラット: ``repo/.cursor/debug-<session>.log``
    """
    sid = _trim(session_id) or DEFAULT_AGENT_DEBUG_SESSION_ID
    name = f"debug-{sid}.log"
    v = _trim(os.environ.get("CURSOR_DEBUG_LOG"))
    if not v:
        v = _trim(os.environ.get("PM_AI_DEBUG_LOG"))
    if v:
        return pathlib.Path(v).resolve()
    u = dict(ui or {})
    ui_path = _trim(u.get(KEY_PM_AI_CURSOR_DEBUG_LOG))
    if not ui_path:
        ui_path = _trim(os.environ.get(KEY_PM_AI_CURSOR_DEBUG_LOG))
    if ui_path:
        return pathlib.Path(ui_path).resolve()
    repo = _resolve_repo_root(u)
    if _is_production_control_system_repo_leaf(repo):
        par = repo.parent
        if par is not None:
            return (par / ".cursor" / name).resolve()
    return (repo / ".cursor" / name).resolve()


def _write_utf8_append(path: pathlib.Path, line: str) -> bool:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(line)
        return True
    except OSError:
        return False


def _build_wsl_unc_path(windows_abs: str, distro: str, unc_root: str) -> str | None:
    if not distro or len(windows_abs) < 3:
        return None
    dl = windows_abs[0]
    if not (dl.isalpha() and windows_abs[1] == ":"):
        return None
    tail = windows_abs[2:].replace("/", "\\")
    if not tail.startswith("\\"):
        tail = "\\" + tail
    root = unc_root if unc_root.endswith("\\") else unc_root + "\\"
    return f"{root}{distro.strip()}\\mnt\\{dl.lower()}{tail}"


def _wsl_unc_mirror_enabled() -> bool:
    v = _trim(os.environ.get("PM_AI_DEBUG_LOG_WSL_UNC"))
    if not v:
        return True
    return v not in ("0", "false", "False", "off", "OFF")


def _resolve_wsl_distro_name() -> str | None:
    d = _trim(os.environ.get("PM_AI_WSL_DISTRO")) or _trim(os.environ.get("WSL_DISTRO_NAME"))
    if d:
        return d
    if sys.platform != "win32":
        return None
    try:
        root = pathlib.Path("//wsl$/")
        if not root.is_dir():
            return None
        names: list[str] = []
        for p in root.iterdir():
            if not p.is_dir():
                continue
            n = p.name
            low = n.lower()
            if low == "docker-desktop" or low.startswith("rancher-desktop") or "podman" in low:
                continue
            names.append(n)
        names.sort()
        for n in names:
            if "ubuntu" in n.lower():
                return n
        return names[0] if names else None
    except OSError:
        return None


def _mirror_targets(primary: pathlib.Path, ui: Mapping[str, str] | None) -> list[pathlib.Path]:
    out: list[pathlib.Path] = []
    u = dict(ui or {})
    m = _trim(os.environ.get(KEY_PM_AI_DEBUG_LOG_MIRROR)) or _trim(
        u.get(KEY_PM_AI_DEBUG_LOG_MIRROR)
    )
    if m:
        out.append(pathlib.Path(m).resolve())
    if sys.platform != "win32" or not _wsl_unc_mirror_enabled():
        return out
    distro = _resolve_wsl_distro_name()
    if not distro:
        return out
    s = str(primary.resolve())
    for unc_root in ("\\\\wsl$\\", "\\\\wsl.localhost\\"):
        unc = _build_wsl_unc_path(s, distro, unc_root)
        if unc:
            try:
                out.append(pathlib.Path(unc))
            except Exception:
                pass
    return out


def _append_mirrors(primary: pathlib.Path, line: str, ui: Mapping[str, str] | None) -> None:
    w = primary.resolve()
    for m in _mirror_targets(primary, ui):
        try:
            n = m.resolve()
        except Exception:
            continue
        if n == w:
            continue
        try:
            if w.exists() and n.exists() and w.samefile(n):
                continue
        except OSError:
            pass
        _write_utf8_append(n, line)


def append_ndjson_line(
    ui: Mapping[str, str] | None,
    session_id: str | None,
    json_line: str,
) -> pathlib.Path | None:
    """
    1 行 UTF-8 追記。Java ``AgentDebugLog.appendNdjsonLine`` 相当。
    主パス失敗時は ``user.home/.cursor/debug-<session>.log`` にフォールバック。
    """
    line = json_line if json_line.endswith("\n") else json_line + "\n"
    sid = _trim(session_id) or DEFAULT_AGENT_DEBUG_SESSION_ID
    primary = resolve_ndjson_path(ui, sid)
    if _write_utf8_append(primary, line):
        _append_mirrors(primary, line, ui)
        return primary
    home = pathlib.Path.home() / ".cursor" / f"debug-{sid}.log"
    if _write_utf8_append(home, line):
        _append_mirrors(home, line, ui)
        return home
    return None


def append_structured(
    ui: Mapping[str, str] | None,
    session_id: str | None,
    hypothesis_id: str,
    location: str,
    message: str,
    data: Mapping[str, Any] | None,
) -> pathlib.Path | None:
    """Java ``AgentDebugLog.appendStructured`` と同形の 1 行 NDJSON。"""
    sid = _trim(session_id) or DEFAULT_AGENT_DEBUG_SESSION_ID
    try:
        payload: dict[str, Any] = {
            "sessionId": sid,
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": dict(data) if data is not None else {},
            "timestamp": int(time.time() * 1000),
        }
        return append_ndjson_line(ui, sid, json.dumps(payload, ensure_ascii=False))
    except Exception:
        return None


def append_interactive_core_probe(
    hypothesis_id: str,
    location: str,
    message: str,
    data: Mapping[str, Any] | None = None,
) -> pathlib.Path | None:
    """
    ``planning_core._core`` インタラクティブ試行デバッグ用。
    ``PM_AI_AGENT_DEBUG_SESSION_ID`` / ``CURSOR_DEBUG_SESSION_ID`` で sessionId を上書き可。
    """
    return append_structured(
        ui_env_from_os(),
        resolve_agent_debug_session_id(),
        hypothesis_id,
        location,
        message,
        data,
    )
