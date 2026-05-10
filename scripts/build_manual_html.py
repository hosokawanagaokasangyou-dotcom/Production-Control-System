#!/usr/bin/env python3
"""manual/src ?? Markdown ?? manual/html ???I?T?C?g?????o?????B"""

from __future__ import annotations

import argparse
import html
import os
import re
import shutil
import sys
from pathlib import Path

try:
    import markdown
except ImportError:
    print("markdown ???K?v???: pip install -r manual/requirements.txt", file=sys.stderr)
    sys.exit(1)


def md_link_to_html(text: str) -> str:
    return re.sub(r"(\[[^\]]*\])\(([^)]+\.md)\)", lambda m: f"{m.group(1)}({m.group(2)[:-3]}.html)", text)


def page_title(md_text: str) -> str:
    for line in md_text.splitlines():
        if line.startswith("# "):
            return line[2:].strip()
    return "Manual"


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo-root", type=Path, default=None)
    args = ap.parse_args()

    repo = args.repo_root or Path(__file__).resolve().parent.parent
    src_root = repo / "manual" / "src"
    out_root = repo / "manual" / "html"
    if not src_root.is_dir():
        print(f"[build_manual_html] missing {src_root}", file=sys.stderr)
        sys.exit(1)

    out_root.mkdir(parents=True, exist_ok=True)

    static_src = src_root / "static"
    if static_src.is_dir():
        shutil.copytree(static_src, out_root / "static", dirs_exist_ok=True)

    images_src = src_root / "images"
    if images_src.is_dir():
        shutil.copytree(images_src, out_root / "images", dirs_exist_ok=True)

    md_files = sorted(src_root.rglob("*.md"))
    if not md_files:
        print("[build_manual_html] no .md under manual/src", file=sys.stderr)
        sys.exit(1)

    md_extensions = ["tables", "fenced_code", "nl2br"]

    for md_path in md_files:
        rel = md_path.relative_to(src_root)
        out_html = (out_root / rel).with_suffix(".html")
        out_html.parent.mkdir(parents=True, exist_ok=True)

        raw = md_path.read_text(encoding="utf-8")
        raw = md_link_to_html(raw)
        body = markdown.markdown(raw, extensions=md_extensions)
        title = html.escape(page_title(raw))

        href_css = Path(os.path.relpath(out_root / "static" / "style.css", out_html.parent)).as_posix()

        html_doc = f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>{title}</title>
  <link rel="stylesheet" href="{href_css}"/>
</head>
<body>
{body}
</body>
</html>
"""
        out_html.write_text(html_doc, encoding="utf-8")
        print(f"[build_manual_html] wrote {out_html}")

    print(f"[build_manual_html] done ({len(md_files)} pages)")


if __name__ == "__main__":
    main()
