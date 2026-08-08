# -*- coding: utf-8 -*-
"""Path resolution for machine-calendar-data.json."""

from __future__ import annotations

import os
from pathlib import Path

from planning_core.core.attendance_paths import summary_workbook_parent
from planning_core.core.columns import ENV_SUMMARY_AI_DISPATCH_WORKBOOK

MACHINE_CALENDAR_JSON_FILENAME = "machine-calendar-data.json"
ENV_MACHINE_CALENDAR_JSON = "PM_AI_MACHINE_CALENDAR_JSON"
ENV_MACHINE_CALENDAR_HISTORY_DIR = "PM_AI_MACHINE_CALENDAR_JSON_HISTORY_DIR"
ENV_MACHINE_CALENDAR_HISTORY_MAX = "PM_AI_MACHINE_CALENDAR_JSON_HISTORY_MAX"

MACHINE_CALENDAR_HISTORY_DIR_NAME = "machine-calendar-json-history"
MACHINE_CALENDAR_HISTORY_MAX_DEFAULT = 20

APP_MASTER_MACHINE_CALENDAR_SHEET = "APP_機械カレンダー"


def machine_calendar_data_json_path() -> Path:
    explicit = os.environ.get(ENV_MACHINE_CALENDAR_JSON, "").strip()
    if explicit:
        return Path(explicit).resolve()
    att_explicit = os.environ.get("PM_AI_ATTENDANCE_JSON", "").strip()
    if att_explicit:
        return (Path(att_explicit).resolve().parent / MACHINE_CALENDAR_JSON_FILENAME).resolve()
    parent = summary_workbook_parent()
    return (parent / MACHINE_CALENDAR_JSON_FILENAME).resolve()


def machine_calendar_json_history_root(json_path: Path | None = None) -> Path:
    explicit = os.environ.get(ENV_MACHINE_CALENDAR_HISTORY_DIR, "").strip()
    if explicit:
        return Path(explicit).resolve()
    jp = json_path or machine_calendar_data_json_path()
    return (jp.resolve().parent / MACHINE_CALENDAR_HISTORY_DIR_NAME).resolve()
