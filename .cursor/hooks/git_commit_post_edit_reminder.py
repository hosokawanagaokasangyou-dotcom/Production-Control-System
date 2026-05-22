# -*- coding: utf-8 -*-
"""postToolUse / afterFileEdit: 版管理対象の編集後に commit / push 忘れを注意喚起。"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

_SKIP_SUFFIXES = {
    ".log",
    ".tmp",
    ".swp",
}
_SKIP_PARTS = (
    "/.cursor/debug-",
    "/node_modules/",
    "/build/",
    "/target/",
    "/.git/",
)


def _as_str(v: Any) -> str:
    return v if isinstance(v, str) else ""


def _path_from_payload(payload: dict[str, Any]) -> str:
    tool_input = payload.get("tool_input")
    if not isinstance(tool_input, dict):
        tool_input = {}
    return (
        _as_str(tool_input.get("path"))
        or _as_str(tool_input.get("file_path"))
        or _as_str(tool_input.get("target_file"))
        or _as_str(payload.get("file_path"))
        or _as_str(payload.get("path"))
    )


def _is_versioned_target(path: str) -> bool:
    if not path:
        return False
    norm = path.replace("\\", "/")
    low = norm.lower()
    if low.startswith("~$"):
        return False
    for part in _SKIP_PARTS:
        if part in low:
            return False
    suffix = Path(norm).suffix.lower()
    if suffix in _SKIP_SUFFIXES:
        return False
    return True


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

    path = _path_from_payload(payload)
    if not _is_versioned_target(path):
        sys.stdout.write("{}")
        return

    body = (
        "【Git フック】版管理対象を編集しました。"
        " 応答終了前に .cursor/rules/git-commit-push-after-code-changes.mdc に従い、"
        "その依頼で触ったファイルを git add → commit（日本語メッセージ）→ push してください。"
        " 依頼外の差分が混在する場合はコミットを分けるかユーザーに確認し、"
        "応答で git status の要約を報告してください。\n"
    )
    sys.stdout.write(json.dumps({"additional_context": body}, ensure_ascii=False))


if __name__ == "__main__":
    main()
