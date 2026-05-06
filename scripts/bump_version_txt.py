#!/usr/bin/env python3
"""Increment version.txt by 0.01 (Decimal). Single-line file, UTF-8."""

from __future__ import annotations

import argparse
import sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path


def main() -> int:
    p = argparse.ArgumentParser(description="Bump version.txt by 0.01")
    p.add_argument("path", type=Path, help="Path to version.txt")
    args = p.parse_args()
    path: Path = args.path
    if not path.is_file():
        print(f"bump_version_txt: file not found: {path}", file=sys.stderr)
        return 1
    raw = path.read_text(encoding="utf-8").strip()
    line = raw.splitlines()[0].strip() if raw else "0.00"
    try:
        cur = Decimal(line)
    except Exception:
        print(f"bump_version_txt: invalid decimal in first line: {line!r}", file=sys.stderr)
        return 1
    nxt = (cur + Decimal("0.01")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    path.write_text(f"{nxt}\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
