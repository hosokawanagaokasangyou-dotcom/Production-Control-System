# -*- coding: utf-8 -*-
"""
Optional paths to decouple reads from TASK_INPUT_WORKBOOK (inventory + small resolvers).

Excel layout/formatting in the UI is being phased out; prefer plain files or CSV where listed.
"""

from __future__ import annotations

import os

from .dispatch_workspace import resolve_actual_detail_workbook_path

# --- Environment variable names (also documented in Java EnvVarDocs) ---

ENV_GLOBAL_PRIORITY_OVERRIDE_PATH = "PM_AI_GLOBAL_PRIORITY_OVERRIDE_PATH"
ENV_RESULT_TASK_COLUMN_CONFIG_CSV = "PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV"
ENV_COLUMN_CONFIG_WORKBOOK = "PM_AI_COLUMN_CONFIG_WORKBOOK"
ENV_DATA_EXTRACTION_SOURCE_WORKBOOK = "PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK"
ENV_ACTUALS_DATA_WORKBOOK = "PM_AI_ACTUALS_DATA_WORKBOOK"
ENV_EXCLUDE_RULES_JSON = "PM_AI_EXCLUDE_RULES_JSON"

# Inventory: purpose, _core entry / helper, primary mechanism, env override(s).
EXCEL_IO_INVENTORY: list[dict[str, str]] = [
    {
        "area": "global_priority_override",
        "function": "load_main_sheet_global_priority_override_text",
        "default": "openpyxl scan main sheet for label cell",
        "env": ENV_GLOBAL_PRIORITY_OVERRIDE_PATH,
        "note": "UTF-8 text file replaces Excel scan when path exists",
    },
    {
        "area": "result_task_column_config",
        "function": "load_result_task_column_rows_from_input_workbook",
        "default": "pandas read_excel column_config sheet",
        "env": f"{ENV_RESULT_TASK_COLUMN_CONFIG_CSV} | {ENV_COLUMN_CONFIG_WORKBOOK}",
        "note": "CSV with name/visible columns, or alternate workbook for sheet only",
    },
    {
        "area": "data_extraction_datetime",
        "function": "_extract_data_extraction_datetime",
        "default": "pandas read_excel plan task sheet",
        "env": ENV_DATA_EXTRACTION_SOURCE_WORKBOOK,
        "note": "Alternate xlsx for data-extraction timestamp columns",
    },
    {
        "area": "result_dispatch_table_sidecar",
        "function": "_write_dispatch_table_standalone_json",
        "default": "same dataframe as standalone dispatch-table xlsx",
        "env": "PM_AI_RESULT_DISPATCH_TABLE_JSON",
        "note": "set 0 to skip UTF-8 JSON next to xlsx",
    },
    {
        "area": "machining_actuals",
        "function": "load_machining_actuals_df",
        "default": "pandas read_excel ACTUALS_SHEET_NAME",
        "env": f"{ENV_ACTUALS_DATA_WORKBOOK} | PM_AI_ACTUAL_DETAIL_*",
        "note": "explicit workbook first; else same as resolve_actual_detail (newest under SOURCE_DIR)",
    },
    {
        "area": "machining_actual_detail",
        "function": "load_machining_actual_detail_df",
        "default": "resolve_actual_detail_workbook_path (dispatch_workspace)",
        "env": "PM_AI_ACTUAL_DETAIL_*",
        "note": "Already path-based; see dispatch_workspace",
    },
    {
        "area": "processing_plan_tabular",
        "function": "load_tasks_df / resolve_processing_plan_path_from_env",
        "default": "PM_AI_PROCESSING_PLAN_PATH (tabular); TASK_INPUT_WORKBOOK not used here",
        "env": "PM_AI_PROCESSING_PLAN_PATH | PM_AI_TASK_INPUT_SOURCE_DIR",
        "note": "if path unset or not a file, newest tabular file under PM_AI_TASK_INPUT_SOURCE_DIR",
    },
    {
        "area": "master_skills_need",
        "function": "load_skills_and_needs / pd.ExcelFile(master)",
        "default": "read_excel master workbook (cwd + MASTER_WORKBOOK_FILE)",
        "env": "MASTER_WORKBOOK_FILE | PM_AI_MASTER_WORKBOOK",
        "note": "PM_AI_MASTER_WORKBOOK: absolute path to .xlsm when not beside cwd",
    },
    {
        "area": "exclude_rules",
        "function": "_load_exclude_rules_from_workbook",
        "default": "pandas read_excel exclude-rules sheet (same name as _core EXCLUDE_RULES_SHEET_NAME)",
        "env": ENV_EXCLUDE_RULES_JSON,
        "note": "UTF-8 JSON list or {rules: [...]} skips Excel read and "
        "run_exclude_rules_sheet_maintenance when file exists and parses (stage1 + plan sheet load path)",
    },
    {
        "area": "result_task_sheet_sidecar",
        "function": "read_result_task_dataframe / write_result_task_json_sidecar",
        "default": "same stem as production_plan_*.xlsx + plan_workbook_sidecar RESULT_TASK_JSON_SUFFIX",
        "env": "PM_AI_PLAN_RESULT_TASK_JSON | PM_AI_PLAN_RESULT_TASK_JSON_PATH",
        "note": "0/false skips JSON; explicit path overrides sidecar location",
    },
    {
        "area": "stage2_column_config_xlwings",
        "function": "_stage2_try_copy_column_config_shapes_from_input",
        "default": "xlwings copy shapes after openpyxl save",
        "env": "PM_AI_XLWINGS_STAGE2_DISABLED | STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT",
        "note": "1/true disables xlwings step (Excel automation)",
    },
]


def resolve_column_config_workbook_path(task_input_workbook: str) -> str:
    """Workbook that hosts the result-task column-config sheet (unless CSV is used)."""
    p = (os.environ.get(ENV_COLUMN_CONFIG_WORKBOOK) or "").strip()
    if p and os.path.isfile(p):
        return p
    return (task_input_workbook or "").strip()


def resolve_data_extraction_workbook_path(task_input_workbook: str) -> str:
    """Workbook used to read plan-sheet extraction timestamp columns."""
    p = (os.environ.get(ENV_DATA_EXTRACTION_SOURCE_WORKBOOK) or "").strip()
    if p and os.path.isfile(p):
        return p
    return (task_input_workbook or "").strip()


def resolve_actuals_workbook_path(task_input_workbook: str) -> str:
    """
    Workbook for ACTUALS sheet (load_machining_actuals_df).

    1) PM_AI_ACTUALS_DATA_WORKBOOK if file exists.
    2) Else resolve_actual_detail_workbook_path (newest under SOURCE_DIR, etc.).
    """
    p = (os.environ.get(ENV_ACTUALS_DATA_WORKBOOK) or "").strip()
    if p and os.path.isfile(p):
        return p
    shared = resolve_actual_detail_workbook_path(task_input_workbook)
    if shared and os.path.isfile(shared):
        return shared
    return (task_input_workbook or "").strip()
