"""PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK → 工場共有 DATA フォルダ（Java AppPaths.summarySharedDataDir 相当）。"""

from __future__ import annotations

import os
from pathlib import Path

from planning_core.core.columns import ENV_SUMMARY_AI_DISPATCH_WORKBOOK

SUMMARY_AI_DISPATCH_XLSX = "サマリ_AI配台.xlsx"

DEFAULT_KONAN_SHARED_DATA_DIR = (
    r"\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\002  加工G\●配台AIシステム\共有DATA"
)
DEFAULT_KOKUBU_DATA_DIR = (
    r"\\192.168.0.101\共有フォルダ\国分工場\国分共有\●配台AIシステム\DATA"
)


def _repo_root() -> Path:
    env = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
    if env:
        return Path(env).resolve()
    return Path.cwd().resolve()


def _resolve_override_path(override: str) -> Path:
    p = Path(override)
    if p.is_absolute():
        return p
    return _repo_root() / "code" / override


def normalize_summary_shared_data_dir(path: Path) -> Path:
    """フォルダパスを正とし、旧設定（.xlsx / .xlsm ファイルパス）は親フォルダへ正規化する。"""
    resolved = path.resolve()
    if resolved.is_dir():
        return resolved
    name = resolved.name.lower()
    if name.endswith((".xlsx", ".xlsm")):
        parent = resolved.parent
        if parent != resolved:
            return parent.resolve()
    return resolved


def resolve_summary_shared_data_dir_from_override(override: str) -> str:
    return str(normalize_summary_shared_data_dir(_resolve_override_path(override)))


def resolve_summary_shared_data_dir() -> str:
    override = (os.environ.get(ENV_SUMMARY_AI_DISPATCH_WORKBOOK) or "").strip()
    if override:
        return str(normalize_summary_shared_data_dir(_resolve_override_path(override)))
    return str((_repo_root() / "code").resolve())


def resolve_summary_ai_dispatch_workbook_path() -> str:
    """レガシー互換: 共有 DATA フォルダ内のサマリ_AI配台.xlsx 絶対パス（ファイル未使用）。"""
    return str(
        (Path(resolve_summary_shared_data_dir()) / SUMMARY_AI_DISPATCH_XLSX).resolve()
    )


def summary_shared_data_sibling_path(filename: str) -> str:
    parent = resolve_summary_shared_data_dir()
    if not parent:
        return ""
    return os.path.normpath(os.path.join(parent, filename))
