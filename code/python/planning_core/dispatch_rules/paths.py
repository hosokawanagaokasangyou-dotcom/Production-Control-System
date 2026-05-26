"""Path resolution for dispatch_special_rules.json."""

from __future__ import annotations

import os
import shutil
from pathlib import Path


DISPATCH_SPECIAL_RULES_DIR = "dispatch_special_rules"
DISPATCH_SPECIAL_RULES_FILENAME = "dispatch_special_rules.json"
ENV_KEY = "PM_AI_DISPATCH_SPECIAL_RULES_JSON"
SUMMARY_ENV_KEY = "PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK"


def _repo_root() -> Path:
    env = os.environ.get("PM_AI_REPO_ROOT", "").strip()
    if env:
        return Path(env).resolve()
    return Path(__file__).resolve().parents[3]


def bundled_template_path() -> Path:
    return _repo_root() / "code" / "json" / "dispatch_special_rules" / DISPATCH_SPECIAL_RULES_FILENAME


def work_dir_from_summary_workbook() -> Path | None:
    summary = os.environ.get(SUMMARY_ENV_KEY, "").strip()
    if not summary:
        return None
    p = Path(summary)
    if not p.is_file():
        return None
    return p.parent / DISPATCH_SPECIAL_RULES_DIR


def default_work_json_path() -> Path | None:
    work_dir = work_dir_from_summary_workbook()
    if work_dir is None:
        return None
    return work_dir / DISPATCH_SPECIAL_RULES_FILENAME


def resolve_dispatch_special_rules_json() -> str | None:
    explicit = os.environ.get(ENV_KEY, "").strip()
    if explicit:
        p = Path(explicit)
        if p.is_file():
            return str(p.resolve())
    default = default_work_json_path()
    if default and default.is_file():
        return str(default.resolve())
    return None


def ensure_work_json_from_repo_template() -> str | None:
    target = default_work_json_path()
    if target is None:
        return None
    if target.is_file():
        return str(target.resolve())
    source = bundled_template_path()
    if not source.is_file():
        return None
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)
    os.environ[ENV_KEY] = str(target.resolve())
    return str(target.resolve())
