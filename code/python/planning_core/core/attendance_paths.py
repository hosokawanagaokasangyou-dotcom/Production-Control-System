# -*- coding: utf-8 -*-
"""Path resolution for attendance-data.json and related files."""

from __future__ import annotations

import os
from pathlib import Path

from planning_core.core.columns import ENV_SUMMARY_AI_DISPATCH_WORKBOOK

ATTENDANCE_DATA_JSON_FILENAME = "attendance-data.json"
ENV_ATTENDANCE_JSON = "PM_AI_ATTENDANCE_JSON"
ENV_ATTENDANCE_HISTORY_DIR = "PM_AI_ATTENDANCE_JSON_HISTORY_DIR"
ENV_ATTENDANCE_HISTORY_MAX = "PM_AI_ATTENDANCE_JSON_HISTORY_MAX"

ATTENDANCE_HISTORY_DIR_NAME = "attendance-json-history"
ATTENDANCE_HISTORY_MAX_DEFAULT = 20

ATTENDANCE_CALENDAR_XLSX_FILENAME = "勤怠・機械カレンダー.xlsx"
ATTENDANCE_CALENDAR_XLSM_FILENAME = "勤怠・機械カレンダー.xlsm"  # 読込フォールバックのみ（新規出力は xlsx）
ATTENDANCE_CALENDAR_XLSX_LEGACY_FILENAME = "勤怠カレンダー.xlsx"
ENV_ATTENDANCE_CALENDAR_XLSX = "PM_AI_ATTENDANCE_CALENDAR_XLSX"
ENV_ATTENDANCE_CALENDAR_XLSX_HISTORY_DIR = "PM_AI_ATTENDANCE_CALENDAR_XLSX_HISTORY_DIR"
ENV_ATTENDANCE_CALENDAR_XLSX_HISTORY_MAX = "PM_AI_ATTENDANCE_CALENDAR_XLSX_HISTORY_MAX"
ATTENDANCE_CALENDAR_XLSX_HISTORY_DIR_NAME = "attendance-calendar-xlsx-history"
ATTENDANCE_CALENDAR_XLSX_HISTORY_MAX_DEFAULT = 20

APP_MASTER_COMPANY_SHEET = "APP_会社カレンダー"
APP_MASTER_MEMBER_SHEET_PREFIX = "APP_勤怠カレンダー_"
APP_MASTER_MACHINE_CALENDAR_SHEET = "APP_機械カレンダー"
APP_MASTER_MACHINE_CALENDAR_DATE_SHEET = "APP_機械カレンダー_日付"
APP_MASTER_MENU_SHEET = "メニュー"

JAPANESE_HOLIDAYS_CACHE_DIR = ".pm-ai-cache/japanese-holidays"


def is_app_master_export_sheet(sheet_name: str) -> bool:
    """アプリが master へ自動出力するシート名（APP_*）。"""
    name = (sheet_name or "").strip()
    if not name:
        return False
    if name == APP_MASTER_COMPANY_SHEET:
        return True
    if name == APP_MASTER_MACHINE_CALENDAR_SHEET:
        return True
    if name == APP_MASTER_MACHINE_CALENDAR_DATE_SHEET:
        return True
    if name == APP_MASTER_MENU_SHEET:
        return True
    return name.startswith(APP_MASTER_MEMBER_SHEET_PREFIX)


def is_legacy_master_calendar_sheet(sheet_name: str) -> bool:
    """
    過去バージョン互換のカレンダーシート（「カレンダー」を含むが APP_* ではない）。
    export では削除・上書きしない。
    """
    name = (sheet_name or "").strip()
    if not name or is_app_master_export_sheet(name):
        return False
    return "カレンダー" in name


def legacy_master_calendar_sheet_names(sheetnames: list[str]) -> list[str]:
    return [n for n in sheetnames if is_legacy_master_calendar_sheet(n)]


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
        from planning_core.core.summary_shared_data_paths import (
            resolve_summary_shared_data_dir_from_override,
        )

        p = Path(resolve_summary_shared_data_dir_from_override(summary))
        if p.is_dir():
            return p
    try:
        from planning_core.core.summary_shared_data_paths import resolve_summary_shared_data_dir

        parent = Path(resolve_summary_shared_data_dir())
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


def attendance_json_history_root(json_path: Path | None = None) -> Path:
    explicit = os.environ.get(ENV_ATTENDANCE_HISTORY_DIR, "").strip()
    if explicit:
        return Path(explicit).resolve()
    jp = json_path or attendance_data_json_path()
    return (jp.resolve().parent / ATTENDANCE_HISTORY_DIR_NAME).resolve()


def _resolve_attendance_calendar_xlsx_in_dir(parent: Path) -> Path:
    canonical = (parent / ATTENDANCE_CALENDAR_XLSX_FILENAME).resolve()
    legacy = (parent / ATTENDANCE_CALENDAR_XLSX_LEGACY_FILENAME).resolve()
    xlsm = (parent / ATTENDANCE_CALENDAR_XLSM_FILENAME).resolve()
    if canonical.is_file():
        return canonical
    if legacy.is_file():
        return legacy
    if xlsm.is_file():
        return xlsm
    return canonical


def attendance_calendar_xlsx_path() -> Path:
    explicit = os.environ.get(ENV_ATTENDANCE_CALENDAR_XLSX, "").strip()
    if explicit:
        return Path(explicit).resolve()
    try:
        from planning_core.core.master_data import _master_workbook_path_resolved

        master = Path(_master_workbook_path_resolved())
        if master.is_file():
            return _resolve_attendance_calendar_xlsx_in_dir(master.parent)
        if master.parent.is_dir():
            return _resolve_attendance_calendar_xlsx_in_dir(master.parent)
    except Exception:
        pass
    parent = summary_workbook_parent()
    if parent is not None and parent.is_dir():
        return _resolve_attendance_calendar_xlsx_in_dir(parent)
    return _resolve_attendance_calendar_xlsx_in_dir(_repo_root() / "code")


def attendance_calendar_xlsx_history_root(xlsx_path: Path | None = None) -> Path:
    explicit = os.environ.get(ENV_ATTENDANCE_CALENDAR_XLSX_HISTORY_DIR, "").strip()
    if explicit:
        return Path(explicit).resolve()
    xp = (xlsx_path or attendance_calendar_xlsx_path()).resolve()
    return (xp.parent / ATTENDANCE_CALENDAR_XLSX_HISTORY_DIR_NAME).resolve()


def japanese_holidays_cache_path(year: int) -> Path:
    return (_repo_root() / JAPANESE_HOLIDAYS_CACHE_DIR / f"holidays_{year}.json").resolve()
