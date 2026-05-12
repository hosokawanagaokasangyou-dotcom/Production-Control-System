#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Replace <!-- MANUAL_SNAP:key --> in Markdown per pipeline-manifest.yaml."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    print("PyYAML required: pip install -r manual/requirements.txt", file=sys.stderr)
    sys.exit(1)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--manifest", required=True, type=Path, help="manual/pipeline-manifest.yaml")
    ap.add_argument("--repo-root", type=Path, default=None)
    args = ap.parse_args()

    repo = args.repo_root or Path(__file__).resolve().parent.parent
    manifest_path = args.manifest if args.manifest.is_absolute() else repo / args.manifest
    data = yaml.safe_load(manifest_path.read_text(encoding="utf-8"))
    injections = data.get("injections") or []

    for inj in injections:
        key = inj.get("tab_key")
        md_rel = inj.get("markdown_file")
        caption = inj.get("caption") or key
        if not key or not md_rel:
            continue
        md_path = repo / md_rel
        if not md_path.is_file():
            print(f"[inject] skip (missing file): {md_path}", file=sys.stderr)
            continue
        text = md_path.read_text(encoding="utf-8")
        placeholder = f"<!-- MANUAL_SNAP:{key} -->"
        replacement = f"![{caption}](../images/{key}.png)\n"
        if placeholder not in text:
            print(f"[inject] placeholder not found in {md_path}: {placeholder}", file=sys.stderr)
            continue
        text_new = text.replace(placeholder, replacement, 1)
        md_path.write_text(text_new, encoding="utf-8")
        print(f"[inject] updated {md_path}")


if __name__ == "__main__":
    main()
