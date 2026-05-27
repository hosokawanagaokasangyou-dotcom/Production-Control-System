#!/usr/bin/env python3
"""Reject staging paths under log/ (execution_log.txt 等)."""
from __future__ import annotations

import subprocess
import sys


def main() -> int:
    proc = subprocess.run(
        ["git", "diff", "--cached", "--name-only", "--diff-filter=ACMR"],
        check=True,
        capture_output=True,
        text=True,
    )
    blocked: list[str] = []
    for line in proc.stdout.splitlines():
        path = line.strip().replace("\\", "/")
        if not path:
            continue
        if path == "log/execution_log.txt" or path.startswith("log/"):
            blocked.append(path)
    if not blocked:
        return 0
    print(
        "pre-commit: log/ 配下（execution_log.txt 等）は .gitignore 対象のためコミットできません:",
        file=sys.stderr,
    )
    for path in blocked:
        print(f"  - {path}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
