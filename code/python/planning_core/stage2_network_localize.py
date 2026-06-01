# -*- coding: utf-8 -*-
"""段階2: UNC 上のマスタ等をローカルスナップショットへ複製し env を差し替える（読込レイテンシ削減）。"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path


def _is_unc_path(path: str) -> bool:
    n = (path or "").strip().replace("/", "\\")
    return n.startswith("\\\\") or n.startswith("//")


def _repo_root_for_cache() -> Path:
    repo = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
    if repo:
        return Path(repo).resolve()
    try:
        return Path(__file__).resolve().parents[3]
    except IndexError:
        return Path.cwd().resolve()


def _should_skip_localize() -> bool:
    v = (os.environ.get("PM_AI_STAGE2_LOCALIZE_UNC_SOURCES") or "1").strip().lower()
    return v in ("0", "false", "no", "off")


def _copy_if_newer(src: str, dest: Path) -> bool:
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        st = os.stat(src)
    except OSError:
        return False
    if dest.is_file():
        try:
            dt = os.stat(dest)
            if int(dt.st_mtime) >= int(st.st_mtime) and int(dt.st_size) == int(st.st_size):
                return True
        except OSError:
            pass
    shutil.copy2(src, dest)
    return True


def localize_unc_paths_for_stage2() -> list[str]:
    """
    ``PM_AI_MASTER_WORKBOOK`` が UNC のときリポジトリ ``.pm-ai-cache/stage2-run-snapshot/`` へ複製する。
    加工計画は Java ``NetworkSourceDirResolver`` がキャッシュパスを渡す想定（ここでは未設定時のみ補完）。

    Returns:
        ログ用メッセージ行
    """
    if _should_skip_localize():
        return ["[stage2-localize] PM_AI_STAGE2_LOCALIZE_UNC_SOURCES=0 のためスキップ"]

    lines: list[str] = []
    snap_dir = _repo_root_for_cache() / ".pm-ai-cache" / "stage2-run-snapshot"

    master_key = "PM_AI_MASTER_WORKBOOK"
    master = (os.environ.get(master_key) or "").strip()
    if master and _is_unc_path(master) and os.path.isfile(master):
        ext = Path(master).suffix or ".xlsm"
        dest = snap_dir / f"master-workbook{ext}"
        if _copy_if_newer(master, dest):
            os.environ[master_key] = str(dest.resolve())
            lines.append(
                f"[stage2-localize] {master_key}: ローカルスナップショット → {dest}"
            )
            try:
                from planning_core import _core as pc

                pc._MASTER_PD_EXCEL_CACHE.clear()
            except Exception:
                pass
        else:
            lines.append(f"[stage2-localize] {master_key}: 複製に失敗（UNCのまま続行）")

    plan_key = "PM_AI_PROCESSING_PLAN_PATH"
    plan = (os.environ.get(plan_key) or "").strip()
    if plan and _is_unc_path(plan) and os.path.isfile(plan):
        ext = Path(plan).suffix or ".xlsx"
        dest = snap_dir / f"processing-plan{ext}"
        if _copy_if_newer(plan, dest):
            os.environ[plan_key] = str(dest.resolve())
            lines.append(
                f"[stage2-localize] {plan_key}: ローカルスナップショット → {dest}"
            )
            try:
                from planning_core import _core as pc

                pc._TABULAR_DF_LOAD_CACHE.clear()
            except Exception:
                pass

    for line in lines:
        print(line, flush=True)
    return lines
