# -*- coding: utf-8 -*-
"""Path resolution for attendance-data.json and related files."""

from __future__ import annotations

import os
from pathlib import Path

from planning_core.core.columns import ENV_SUMMARY_AI_DISPATCH_WORKBOOK

ATTENDANCE_DATA_JSON_FILENAME = "attendance-data.json"
ATTENDANCE_VIEW_XLSX_FILENAME = "勤怠_表示用.xlsx"
ENV_ATTENDANCE_JSON = "PM_AI_ATTENDANCE_JSON"
ENV_ATTENDANCE_VIEW_XLSX = "PM_AI_ATTENDANCE_VIEW_XLSX"

APP_MASTER_COMPANY_SHEET = "APP_会社カレンダー"
APP_MASTER_MEMBER_SHEET_PREFIX = "APP_勤怠カレンダー_"

JAPANESE_HOLIDAYS_CACHE_DIR = ".pm-ai-cache/japanese-holidays"


def _repo_root() -> Path:
    env = os.environ.get("PM_AI_REPO_ROOT", "").strip()
    if env:
        return Path(env).resolve()
    return Path(__file__).resolve().parents[4]


def summary_workbook_parent() -> Path | None:
    explicit_json = os.environ.get(ENV_ATTENDANCE_JSON, "").strip()
    if explicit_json:
        p = Path(explicit_json)
        if p.parent.is_dir():
            return p.parent
    summary = os.environ.get(ENV_SUMMARY_AI_DISPATCH_WORKBOOK, "").strip()
    if summary:
        p = Path(summary)
        if p.is_file():
            return p.parent
    try:
        from planning_core.core.task_queue import _resolve_summary_ai_dispatch_workbook_path

        parent = Path(_resolve_summary_ai_dispatch_workbook_path()).parent
        if parent.is_dir():
            return parent
    except Exception:
        pass
    return _repo_root() / "code"


def attendance_data_json_path() -> Path:
    explicit = os.environ.get(ENV_ATTENDANCE_JSON, "").strip()
    if explicit:
        return Path(explicit).resolve()
    parent = summary_workbook_parent()
    return (parent / ATTENDANCE_DATA_JSON_FILENAME).resolve()


def attendance_view_xlsx_path() -> Path:
    explicit = os.environ.get(ENV_ATTENDANCE_VIEW_XLSX, "").strip()
    if explicit:
        return Path(explicit).resolve()
    parent = summary_workbook_parent()
    return (parent / ATTENDANCE_VIEW_XLSX_FILENAME).resolve()


def japanese_holidays_cache_path(year: int) -> Path:
    return (_repo_root() / JAPANESE_HOLIDAYS_CACHE_DIR / f"holidays_{year}.json").resolve()
