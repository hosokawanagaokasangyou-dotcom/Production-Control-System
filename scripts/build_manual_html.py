#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build static HTML from manual/src Markdown into manual/html."""

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
    print("markdown required: pip install -r manual/requirements.txt", file=sys.stderr)
    sys.exit(1)

NAV_LINK_RE = re.compile(r"^\s*-\s+\[([^\]]+)\]\(([^)]+\.md)\)\s*$", re.MULTILINE)


def md_link_to_html(text: str) -> str:
    return re.sub(
        r"(\[[^\]]*\])\(([^)]+\.md)\)",
        lambda m: f"{m.group(1)}({m.group(2)[:-3]}.html)",
        text,
    )


def page_title(md_text: str) -> str:
    for line in md_text.splitlines():
        if line.startswith("# "):
            return line[2:].strip()
    return "Manual"


def parse_nav_from_index(index_md: str) -> list[tuple[str, str]]:
    """Return [(label, md_rel_path)] e.g. ('1. はじめに', 'chapters/overview.md')."""
    items: list[tuple[str, str]] = []
    for m in NAV_LINK_RE.finditer(index_md):
        label, md_path = m.group(1), m.group(2)
        items.append((label, md_path))
    return items


def build_nav_html(
    nav_items: list[tuple[str, str]],
    *,
    out_html: Path,
    out_root: Path,
    index_href: str,
) -> str:
    index_html = out_root / "index.html"
    parts = [
        '<nav class="manual-nav" aria-label="目次">',
        '<p class="manual-nav-title">取扱説明書</p>',
        '<ul>',
        f'<li><a href="{html.escape(index_href)}"'
        + (' class="is-current"' if out_html.resolve() == index_html.resolve() else "")
        + ">トップ</a></li>",
    ]
    for label, md_rel in nav_items:
        target = (out_root / Path(md_rel)).with_suffix(".html")
        href = Path(os.path.relpath(target, out_html.parent)).as_posix()
        cls = ' class="is-current"' if target.resolve() == out_html.resolve() else ""
        parts.append(f'<li><a href="{html.escape(href)}"{cls}>{html.escape(label)}</a></li>')
    parts.append("</ul></nav>")
    return "\n".join(parts)


def wrap_page(
    *,
    title: str,
    body_html: str,
    css_href: str,
    nav_html: str,
    is_index: bool,
) -> str:
    header = ""
    if is_index:
        header = """
<header class="manual-header">
  <h1>工程管理 AI デスクトップ</h1>
  <p class="subtitle">取扱説明書 — 操作手順と画面の見方</p>
</header>
"""
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>{title}</title>
  <link rel="stylesheet" href="{css_href}"/>
</head>
<body>
<div class="manual-layout">
{nav_html}
<main class="manual-main">
{header}
<article class="manual-content">
{body_html}
</article>
<footer class="manual-footer">
  工程管理 AI デスクトップ 取扱説明書 — リポジトリ <code>manual/html/</code> より生成
</footer>
</main>
</div>
</body>
</html>
"""


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

    index_path = src_root / "index.md"
    nav_items: list[tuple[str, str]] = []
    if index_path.is_file():
        nav_items = parse_nav_from_index(index_path.read_text(encoding="utf-8"))

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

        css_href = Path(os.path.relpath(out_root / "static" / "style.css", out_html.parent)).as_posix()
        index_href = Path(os.path.relpath(out_root / "index.html", out_html.parent)).as_posix()
        nav_html = build_nav_html(
            nav_items,
            out_html=out_html,
            out_root=out_root,
            index_href=index_href,
        )
        is_index = rel.as_posix() == "index.md"

        html_doc = wrap_page(
            title=title,
            body_html=body,
            css_href=css_href,
            nav_html=nav_html,
            is_index=is_index,
        )
        # fix typo </motion> -> nothing (was meant to be empty)
        out_html.write_text(html_doc, encoding="utf-8")
        print(f"[build_manual_html] wrote {out_html}")

    print(f"[build_manual_html] done ({len(md_files)} pages)")


if __name__ == "__main__":
    main()
