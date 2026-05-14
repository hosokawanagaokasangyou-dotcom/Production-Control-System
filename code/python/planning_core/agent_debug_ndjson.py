# -*- coding: utf-8 -*-
"""NDJSON 追記: Java ``AgentDebugLog`` と同一のパス解決・フォールバック・ミラー方針（Python 子プロセス用）。

正本の説明は ``.cursor/rules/agent-debug-ndjson-logging.mdc`` および
``.cursor/rules/agent-debug-wsl-windows-mirror.mdc``。
"""

from __future__ import annotations

import json
import logging
import os
import time
from pathlib import Path
from typing import Any, Mapping

_LOG = logging.getLogger(__name__)

DEFAULT_SESSION_ID = "e04a1d"

KEY_PM_AI_CODE_PYTHON_DIR = "PM_AI_CODE_PYTHON_DIR"
KEY_PM_AI_REPO_ROOT = "PM_AI_REPO_ROOT"
KEY_PM_AI_CURSOR_DEBUG_LOG = "PM_AI_CURSOR_DEBUG_LOG"
KEY_PM_AI_DEBUG_LOG_MIRROR = "PM_AI_DEBUG_LOG_MIRROR"


def _trim(s: str | None) -> str:
    return (s or "").strip()


def _is_plausible_wsl_distro(name: str) -> bool:
    if not name:
        return False
    lower = name.lower()
    if lower == "docker-desktop" or lower.startswith("rancher-desktop") or "podman" in lower:
        return False
    return True


def _discover_wsl_distro_name() -> str | None:
    if os.name != "nt":
        return None
    wsl_root = Path("//wsl$/")
    try:
        if not wsl_root.is_dir():
            return None
        names = sorted(
            p.name
            for p in wsl_root.iterdir()
            if p.is_dir() and _is_plausible_wsl_distro(p.name)
        )
        if not names:
            return None
        ubuntu = [n for n in names if "ubuntu" in n.lower()]
        return min(ubuntu) if ubuntu else names[0]
    except OSError:
        return None


def _resolve_wsl_distro_name() -> str | None:
    d = _trim(os.environ.get("PM_AI_WSL_DISTRO"))
    if d:
        return d
    d = _trim(os.environ.get("WSL_DISTRO_NAME"))
    if d:
        return d
    return _discover_wsl_distro_name()


def _wsl_unc_mirror_enabled() -> bool:
    v = _trim(os.environ.get("PM_AI_DEBUG_LOG_WSL_UNC"))
    if not v:
        return True
    return v not in ("0", "false", "False", "off", "OFF")


def _build_wsl_unc_path_string(
    windows_abs: str, distro: str, unc_root_prefix: str
) -> str | None:
    if not distro or len(windows_abs) < 3:
        return None
    dl = windows_abs[0]
    if not dl.isalpha() or windows_abs[1] != ":":
        return None
    tail = windows_abs[2:].replace("/", "\\")
    if not tail.startswith("\\"):
        tail = "\\" + tail
    root = unc_root_prefix if unc_root_prefix else "\\\\wsl$\\"
    return root + distro.strip() + "\\mnt\\" + dl.lower() + tail


def _mirror_targets(primary_written: Path) -> list[Path]:
    out: list[Path] = []
    seen: set[str] = set()

    def add(p: Path | None) -> None:
        if p is None:
            return
        try:
            k = str(p.resolve())
        except OSError:
            k = str(p)
        if k not in seen:
            seen.add(k)
            out.append(p)

    m = _trim(os.environ.get(KEY_PM_AI_DEBUG_LOG_MIRROR))
    if m:
        add(Path(m))

    if os.name == "nt" and _wsl_unc_mirror_enabled():
        distro = _resolve_wsl_distro_name()
        if distro:
            ws = str(primary_written.resolve())
            for pref in ("\\\\wsl$\\", "\\\\wsl.localhost\\"):
                unc = _build_wsl_unc_path_string(ws, distro, pref)
                if unc:
                    try:
                        add(Path(unc))
                    except OSError:
                        pass
    return out


def _repo_root_from_environ() -> Path:
    rr = _trim(os.environ.get(KEY_PM_AI_REPO_ROOT))
    if rr:
        return Path(rr).resolve()
    cp = _trim(os.environ.get(KEY_PM_AI_CODE_PYTHON_DIR))
    if cp:
        py = Path(cp).resolve()
        if py.name.lower() == "python":
            code = py.parent
            if code.name.lower() == "code" and code.parent is not None:
                return code.parent.resolve()
        return py.resolve()
    # planning_core/agent_debug_ndjson.py -> …/code/python/planning_core/
    here = Path(__file__).resolve().parent
    py_dir = here.parent
    code_dir = py_dir.parent
    if code_dir.name.lower() == "code" and code_dir.parent is not None:
        return code_dir.parent.resolve()
    return code_dir.resolve()


def _is_production_control_system_leaf(repo: Path) -> bool:
    try:
        return repo.name.lower() == "production-control-system"
    except Exception:
        return False


def resolve_ndjson_log_path(session_id: str) -> Path:
    """``AgentDebugLog.resolveNdjsonPath`` と同順（書き込み失敗時の user.home はここでは使わない）。"""
    sid = _trim(session_id) or DEFAULT_SESSION_ID
    file_name = f"debug-{sid}.log"

    for key in ("CURSOR_DEBUG_LOG", "PM_AI_DEBUG_LOG"):
        v = _trim(os.environ.get(key))
        if v:
            return Path(v).resolve()

    ui_path = _trim(os.environ.get(KEY_PM_AI_CURSOR_DEBUG_LOG))
    if ui_path:
        return Path(ui_path).resolve()

    repo = _repo_root_from_environ()
    parent = repo.parent
    if parent is not None and _is_production_control_system_leaf(repo):
        return (parent / ".cursor" / file_name).resolve()
    return (repo / ".cursor" / file_name).resolve()


def _write_utf8_append(path: Path, line: str) -> bool:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(line if line.endswith("\n") else line + "\n")
        return True
    except OSError as e:
        _LOG.debug("agent_debug_ndjson: append failed %s (%s)", path, e)
        return False


def _paths_same_file(a: Path, b: Path) -> bool:
    try:
        return a.exists() and b.exists() and os.path.samefile(a, b)
    except OSError:
        return False


def _append_mirrors(written_primary: Path, line: str) -> None:
    w_norm = written_primary.resolve()
    for mirror in _mirror_targets(written_primary):
        try:
            n = mirror.resolve()
        except OSError:
            n = mirror
        if n == w_norm:
            continue
        if _paths_same_file(w_norm, n):
            continue
        _write_utf8_append(n, line)


def append_ndjson_line(session_id: str, json_line: str) -> Path | None:
    """1 行 NDJSON を追記。成功した ``Path`` を返す（失敗時 ``None``）。"""
    sid = _trim(session_id) or DEFAULT_SESSION_ID
    line = json_line if json_line.endswith("\n") else json_line + "\n"
    primary = resolve_ndjson_log_path(sid)
    if _write_utf8_append(primary, line):
        _append_mirrors(primary, line)
        return primary
    fb = (Path.home() / ".cursor" / f"debug-{sid}.log").resolve()
    if _write_utf8_append(fb, line):
        _append_mirrors(fb, line)
        return fb
    return None


def append_structured(
    session_id: str,
    hypothesis_id: str,
    location: str,
    message: str,
    data: Mapping[str, Any] | None,
    *,
    run_id: str | None = None,
) -> None:
    """``AgentDebugLog.appendStructured`` と同様の 1 行 JSON（任意で ``runId``）。"""
    sid = _trim(session_id) or DEFAULT_SESSION_ID
    line_obj: dict[str, Any] = {
        "sessionId": sid,
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": dict(data) if data is not None else {},
        "timestamp": int(time.time() * 1000),
    }
    rid = _trim(run_id) if run_id is not None else _trim(os.environ.get("PM_AI_DEBUG_RUN_ID"))
    if rid:
        line_obj["runId"] = rid
    try:
        json_line = json.dumps(line_obj, ensure_ascii=False)
    except (TypeError, ValueError):
        return
    append_ndjson_line(sid, json_line)


__all__ = (
    "DEFAULT_SESSION_ID",
    "append_ndjson_line",
    "append_structured",
    "resolve_ndjson_log_path",
)
