# -*- coding: utf-8 -*-
"""postToolUse / afterFileEdit:  ad-hoc NDJSON デバッグ（固定パス・独自ヘルパ）を検知して注意喚起。"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any


def _as_str(v: Any) -> str:
    return v if isinstance(v, str) else ""


def _collect_text(payload: dict[str, Any]) -> tuple[str, str]:
    """Returns (file_path_hint, combined_text_to_scan)."""
    tool = _as_str(payload.get("tool_name"))
    tool_input = payload.get("tool_input")
    if not isinstance(tool_input, dict):
        tool_input = {}

    path = (
        _as_str(tool_input.get("path"))
        or _as_str(tool_input.get("file_path"))
        or _as_str(tool_input.get("target_file"))
        or _as_str(payload.get("file_path"))
        or _as_str(payload.get("path"))
    )
    chunks: list[str] = []
    for key in ("contents", "content", "new_string", "old_string", "patch"):
        part = tool_input.get(key)
        if isinstance(part, str):
            chunks.append(part)
    out = _as_str(payload.get("tool_output"))
    if out:
        chunks.append(out)

    if path and not chunks:
        try:
            p = Path(path)
            if p.is_file() and p.suffix.lower() in {".java", ".py", ".fxml"}:
                chunks.append(p.read_text(encoding="utf-8", errors="replace"))
        except OSError:
            pass

    return path, "\n".join(chunks)


def _violations(path: str, text: str) -> list[str]:
    if not text:
        return []
    low_path = path.lower()
    if low_path and not low_path.endswith((".java", ".py")):
        if "#region agent log" not in text and "debug-" not in text:
            return []
    issues: list[str] = []

    if re.search(r'["\']/mnt/c/[^"\']*debug-[^"\']+\.log', text, re.I):
        issues.append("OS 固定の /mnt/c/.../.cursor/debug-*.log パスが含まれています")

    if re.search(
        r"def\s+_agent_debug_log(?:_path|_session)?_[0-9a-f]+\s*\(",
        text,
        re.I,
    ):
        issues.append(
            "セッション固定の独自 _agent_debug_log_* ヘルパがあります。"
            " planning_core.agent_debug_ndjson または AgentDebugLog に統一してください"
        )

    if "#region agent log" in text or "hypothesisId" in text:
        uses_canonical = (
            "AgentDebugLog.appendStructured" in text
            or "agent_debug_ndjson" in text
            or "append_structured(" in text
        )
        direct_open = bool(
            re.search(
                r'open\s*\(\s*["\'][^"\']*\.cursor[/\\]debug-',
                text,
                re.I,
            )
        )
        if direct_open and not uses_canonical:
            issues.append(
                ".cursor/debug-*.log への open 直書きがあります。"
                " AgentDebugLog / agent_debug_ndjson を使ってください"
            )

    return issues


def main() -> None:
    raw = sys.stdin.read()
    if not raw.strip():
        sys.stdout.write("{}")
        return
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError:
        sys.stdout.write("{}")
        return

    path, text = _collect_text(payload)
    issues = _violations(path, text)
    if not issues:
        sys.stdout.write("{}")
        return

    body = (
        "【agent-debug フック検知】今回の編集に NDJSON デバッグの逸脱がありそうです。\n"
        + "\n".join(f"- {x}" for x in issues)
        + "\n正本: .cursor/rules/agent-debug-ndjson-logging.mdc 、"
        "agent-debug-wsl-windows-mirror.mdc 。Java=AgentDebugLog、Python=agent_debug_ndjson。\n"
    )
    sys.stdout.write(json.dumps({"additional_context": body}, ensure_ascii=False))


if __name__ == "__main__":
    main()
