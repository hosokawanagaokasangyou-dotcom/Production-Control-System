#!/usr/bin/env python3
"""
マニュアル公開パイプライン（スクショ・画像コピー・MD 注入・HTML 生成）。
PowerShell の Publish-Manual.ps1 から呼ぶ。
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    print("PyYAML が必要です: pip install -r manual/requirements.txt", file=sys.stderr)
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


def run_snap(data: dict, root: Path) -> None:
    keys = data.get("snap_tab_keys") or []
    if not keys:
        print("[pipeline] snap_tab_keys が空のためスナップをスキップします。", file=sys.stderr)
        return

    code_java = root / "code_java"
    mvnw = code_java / ("mvnw.cmd" if os.name == "nt" else "mvnw")
    if not mvnw.is_file():
        print(f"[pipeline] missing {mvnw}", file=sys.stderr)
        sys.exit(1)

    snap_rel = data.get("snap_out_dir") or "manual/snap-out"
    snap_out = (root / snap_rel).resolve()

    key_csv = ",".join(str(k).strip() for k in keys if str(k).strip())

    cmd = [
        str(mvnw),
        "-q",
        f"-Dpm.ai.desktop.manual.snap.tabKeys={key_csv}",
        f"-Dpm.ai.desktop.manual.snap.outputDir={snap_out.as_posix()}",
        "-Dpm.ai.desktop.manual.snap.stageWidth=1800",
        "-Dpm.ai.desktop.manual.snap.stageHeight=900",
        "-Dpm.ai.desktop.manual.snap.pauseMillis=1500",
        "-Dpm.ai.desktop.manual.snap.exitAfter=true",
        "compile",
        "exec:exec@pm-ai-desktop",
    ]

    print("[pipeline] running:", " ".join(cmd))
    subprocess.run(cmd, cwd=str(code_java), check=True)


def sync_snap_to_images(data: dict, root: Path) -> None:
    snap_rel = data.get("snap_out_dir") or "manual/snap-out"
    img_rel = data.get("images_dir") or "manual/src/images"
    snap_dir = root / snap_rel
    img_dir = root / img_rel
    if not snap_dir.is_dir():
        print(f"[pipeline] snap dir missing: {snap_dir}", file=sys.stderr)
        return
    img_dir.mkdir(parents=True, exist_ok=True)
    for png in snap_dir.glob("*.png"):
        shutil.copy2(png, img_dir / png.name)
        print(f"[pipeline] copied {png.name} -> {img_dir}")


def main() -> None:
    ap = argparse.ArgumentParser(description="マニュアル公開パイプライン")
    ap.add_argument("--manifest", type=Path, default=Path("manual/pipeline-manifest.yaml"))
    ap.add_argument("--repo-root", type=Path, default=None)
    ap.add_argument("--skip-data-prep", action="store_true")
    ap.add_argument("--skip-snap", action="store_true")
    ap.add_argument("--skip-sync", action="store_true")
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

    if not args.skip_snap:
        run_snap(data, root)

    if not args.skip_sync:
        sync_snap_to_images(data, root)

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

    print("[pipeline] 完了")


if __name__ == "__main__":
    main()
