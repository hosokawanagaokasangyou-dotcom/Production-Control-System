#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Manual pipeline: data_prep, inject placeholders, build HTML. Run via Publish-Manual.ps1."""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    print("PyYAML required: pip install -r manual/requirements.txt", file=sys.stderr)
    sys.exit(1)


def repo_root_from_here() -> Path:
    return Path(__file__).resolve().parent.parent


def load_manifest(path: Path) -> dict:
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def run_data_prep(data: dict, root: Path) -> None:
    items = data.get("data_prep") or []
    for item in items:
        src = root / item["src"]
        dst = root / item["dst"]
        dst.parent.mkdir(parents=True, exist_ok=True)
        if src.is_dir():
            shutil.copytree(src, dst, dirs_exist_ok=True)
        else:
            shutil.copy2(src, dst)
        print(f"[pipeline] data_prep {src} -> {dst}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Manual publish pipeline")
    ap.add_argument("--manifest", type=Path, default=Path("manual/pipeline-manifest.yaml"))
    ap.add_argument("--repo-root", type=Path, default=None)
    ap.add_argument("--skip-data-prep", action="store_true")
    ap.add_argument("--skip-inject", action="store_true")
    ap.add_argument("--skip-html", action="store_true")
    args = ap.parse_args()

    root = args.repo_root or repo_root_from_here()
    manifest = args.manifest if args.manifest.is_absolute() else root / args.manifest
    if not manifest.is_file():
        print(f"[pipeline] manifest not found: {manifest}", file=sys.stderr)
        sys.exit(1)

    data = load_manifest(manifest)

    if not args.skip_data_prep:
        run_data_prep(data, root)

    if not args.skip_inject:
        inject_script = root / "scripts" / "inject_manual_images.py"
        subprocess.run(
            [sys.executable, str(inject_script), "--manifest", str(manifest), "--repo-root", str(root)],
            cwd=str(root),
            check=True,
        )

    if not args.skip_html:
        build_script = root / "scripts" / "build_manual_html.py"
        subprocess.run(
            [sys.executable, str(build_script), "--repo-root", str(root)],
            cwd=str(root),
            check=True,
        )

    print("[pipeline] done")


if __name__ == "__main__":
    main()
