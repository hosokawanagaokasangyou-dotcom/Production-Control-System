#!/usr/bin/env python3
"""ステージされた *.md が UTF-8 として解釈できることを検証する。

Shift-JIS（CP932）等で保存された .md をコミットから防ぐ。
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def _repo_root() -> Path:
    out = subprocess.check_output(
        ["git", "rev-parse", "--show-toplevel"],
        text=True,
    )
    return Path(out.strip())


def _staged_paths(root: Path) -> list[str]:
    out = subprocess.check_output(
        ["git", "diff", "--cached", "--name-only", "--diff-filter=ACM"],
        cwd=root,
        text=True,
    )
    return [line for line in out.splitlines() if line.strip()]


def _index_blob(root: Path, rel: str) -> bytes | None:
    """インデックス上のパスに対応する blob。取得できなければ None。"""
    try:
        return subprocess.check_output(
            ["git", "show", f":{rel}"],
            cwd=root,
        )
    except subprocess.CalledProcessError:
        return None


def main() -> int:
    root = _repo_root()
    md_rel = [p for p in _staged_paths(root) if p.endswith(".md")]
    bad: list[tuple[str, str]] = []
    for rel in md_rel:
        blob = _index_blob(root, rel)
        if blob is None:
            continue
        try:
            blob.decode("utf-8")
        except UnicodeDecodeError as e:
            bad.append((rel, str(e)))
    if bad:
        print(
            "エラー: 次の Markdown は UTF-8 として解釈できません（Shift-JIS 等での保存は禁止）。",
            file=sys.stderr,
        )
        for rel, msg in bad:
            print(f"  - {rel}: {msg}", file=sys.stderr)
        print(
            "エディタで UTF-8（BOM なし推奨）に変換してから再度コミットしてください。",
            file=sys.stderr,
        )
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
