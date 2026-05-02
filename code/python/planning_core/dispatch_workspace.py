# -*- coding: utf-8 -*-
"""
Tabular inputs, workspace, and UNC defaults (parity with AppPaths / plan/*.m).
See EnvVarDocs: PM_AI_TASK_INPUT_SOURCE_DIR, PM_AI_ACTUAL_DETAIL_SOURCE_DIR, PM_AI_RESULT_DISPATCH_TABLE_DIR.
"""

from __future__ import annotations

import logging
import os

import pandas as pd

ENV_PM_AI_WORKSPACE = "PM_AI_WORKSPACE"
ENV_PROCESSING_PLAN_PATH = "PM_AI_PROCESSING_PLAN_PATH"
ENV_PLAN_INPUT_PATH = "PM_AI_PLAN_INPUT_PATH"
ENV_PROCESSING_PLAN_SHEET = "PM_AI_PROCESSING_PLAN_SHEET"
ENV_TABULAR_CSV_ENCODING = "PM_AI_TABULAR_CSV_ENCODING"
ENV_ACTUAL_DETAIL_SOURCE_DIR = "PM_AI_ACTUAL_DETAIL_SOURCE_DIR"
ENV_ACTUAL_DETAIL_WORKBOOK = "PM_AI_ACTUAL_DETAIL_WORKBOOK"
ENV_RESULT_DISPATCH_TABLE_DIR = "PM_AI_RESULT_DISPATCH_TABLE_DIR"
ENV_PM_AI_REPO_ROOT = "PM_AI_REPO_ROOT"
ENV_TASK_INPUT_SOURCE_DIR = "PM_AI_TASK_INPUT_SOURCE_DIR"

_LOG = logging.getLogger(__name__)

# CSV / Parquet / Excel task-input files (read_tabular_dataframe).
_PROCESSING_PLAN_FILE_SUFFIXES = (
    ".csv",
    ".parquet",
    ".pq",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
)

# Same logical path as AppPaths.DEFAULT_PQ_A_SOURCE_UNC (PQ 01 / PM_AI_TASK_INPUT_SOURCE_DIR).
DEFAULT_TASK_INPUT_SOURCE_DIR = (
    "\\\\192.168.0.101\\"
    "\u5171\u6709\u30d5\u30a9\u30eb\u30c0\\"
    "\u6e56\u5357\u5de5\u5834\\"
    "\u6e56\u5357\u5171\u6709\\"
    "\u751f\u7523\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\\"
    "\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\\"
    "\u25cfDATA\\"
    "\u751f\u7523\u8a08\u753b\u554f\u5408\u305b"
)

# Same as plan/02 PQ Folder.Files (PM_AI_ACTUAL_DETAIL_SOURCE_DIR).
DEFAULT_ACTUAL_DETAIL_SOURCE_DIR = (
    "\\\\192.168.0.101\\"
    "\u5171\u6709\u30d5\u30a9\u30eb\u30c0\\"
    "\u6e56\u5357\u5de5\u5834\\"
    "\u6e56\u5357\u5171\u6709\\"
    "002  \u52a0\u5de5G\\"
    "\u25cf\u691c\u67fb\u8868\u4f5c\u6210\\"
    "\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30DATA"
)


def pm_ai_workspace_dir() -> str:
    """Return absolute workspace dir if valid, else empty string."""
    p = (os.environ.get(ENV_PM_AI_WORKSPACE) or "").strip()
    return os.path.abspath(p) if p and os.path.isdir(p) else ""


def pick_newest_excel_in_dir(dir_path: str) -> str | None:
    """Newest .xlsx/.xlsm by max(mtime, atime). Skips Excel lock files (~$)."""
    best: str | None = None
    best_key = -1.0
    try:
        for name in os.listdir(dir_path):
            if name.startswith("~$"):
                continue
            low = name.lower()
            if not low.endswith((".xlsx", ".xlsm")):
                continue
            fp = os.path.join(dir_path, name)
            if not os.path.isfile(fp):
                continue
            try:
                st = os.stat(fp)
                t = max(float(st.st_mtime), float(getattr(st, "st_atime", 0.0)))
            except OSError:
                continue
            if t > best_key:
                best_key = t
                best = fp
    except OSError:
        return None
    return best


def pick_newest_processing_plan_file_in_dir(dir_path: str) -> str | None:
    """Newest tabular plan file by max(mtime, atime). Skips Excel lock files (~$)."""
    best: str | None = None
    best_key = -1.0
    try:
        for name in os.listdir(dir_path):
            if name.startswith("~$"):
                continue
            low = name.lower()
            if not low.endswith(_PROCESSING_PLAN_FILE_SUFFIXES):
                continue
            fp = os.path.join(dir_path, name)
            if not os.path.isfile(fp):
                continue
            try:
                st = os.stat(fp)
                t = max(float(st.st_mtime), float(getattr(st, "st_atime", 0.0)))
            except OSError:
                continue
            if t > best_key:
                best_key = t
                best = fp
    except OSError:
        return None
    return best


def resolve_processing_plan_path_from_env() -> str | None:
    """
    Effective PM_AI_PROCESSING_PLAN_PATH:

    1. If PM_AI_PROCESSING_PLAN_PATH points to an existing file -> normalize to absolute and set env.
    2. Else if PM_AI_TASK_INPUT_SOURCE_DIR is a directory -> pick newest tabular file there,
       set PM_AI_PROCESSING_PLAN_PATH to that path (stage1 / load_tasks_df use it).

    Returns the resolved absolute path, or None if no file was resolved.
    """
    explicit = (os.environ.get(ENV_PROCESSING_PLAN_PATH) or "").strip()
    if explicit:
        if os.path.isfile(explicit):
            ap = os.path.normpath(os.path.abspath(explicit))
            os.environ[ENV_PROCESSING_PLAN_PATH] = ap
            return ap
        _LOG.warning(
            "PM_AI_PROCESSING_PLAN_PATH is not a file (%r). Trying PM_AI_TASK_INPUT_SOURCE_DIR.",
            explicit,
        )

    src = (os.environ.get(ENV_TASK_INPUT_SOURCE_DIR) or "").strip()
    if not src:
        return None
    src_abs = os.path.normpath(os.path.abspath(src))
    if not os.path.isdir(src_abs):
        _LOG.warning("PM_AI_TASK_INPUT_SOURCE_DIR is not a directory: %r", src)
        return None

    picked = pick_newest_processing_plan_file_in_dir(src_abs)
    if not picked:
        _LOG.warning(
            "No CSV/Parquet/Excel task-input files under PM_AI_TASK_INPUT_SOURCE_DIR: %s",
            src_abs,
        )
        return None

    ap = os.path.normpath(os.path.abspath(picked))
    os.environ[ENV_PROCESSING_PLAN_PATH] = ap
    _LOG.info(
        "PM_AI_PROCESSING_PLAN_PATH set to newest file under PM_AI_TASK_INPUT_SOURCE_DIR: %s",
        ap,
    )
    return ap


def resolve_actual_detail_workbook_path(task_input_workbook: str) -> str | None:
    """
    Actual-detail workbook path. load_machining_actuals_df reuses this resolution when
    PM_AI_ACTUALS_DATA_WORKBOOK is unset (input_resolution.resolve_actuals_workbook_path).

    Single-file override, then newest workbook under ACTUAL_DETAIL_SOURCE_DIR (or default UNC),
    then TASK_INPUT_WORKBOOK (sheet in macro book).
    """
    wb_explicit = (os.environ.get(ENV_ACTUAL_DETAIL_WORKBOOK) or "").strip()
    if wb_explicit and os.path.isfile(wb_explicit):
        return wb_explicit
    d = (os.environ.get(ENV_ACTUAL_DETAIL_SOURCE_DIR) or "").strip()
    if not d:
        d = DEFAULT_ACTUAL_DETAIL_SOURCE_DIR
    if d and os.path.isdir(d):
        picked = pick_newest_excel_in_dir(d)
        if picked:
            return picked
    tw = (task_input_workbook or "").strip()
    if tw and os.path.isfile(tw):
        return tw
    return None


def resolve_result_dispatch_table_output_dir(task_input_workbook: str) -> str:
    """Output folder for standalone result dispatch table xlsx (fixed filename).

    Priority: PM_AI_RESULT_DISPATCH_TABLE_DIR, dirname(TASK_INPUT_WORKBOOK), PM_AI_REPO_ROOT/code.
    """
    o = (os.environ.get(ENV_RESULT_DISPATCH_TABLE_DIR) or "").strip()
    if o and os.path.isdir(o):
        return os.path.abspath(o)
    wb = (task_input_workbook or "").strip()
    if wb and os.path.isfile(wb):
        return os.path.dirname(os.path.abspath(wb))
    repo = (os.environ.get(ENV_PM_AI_REPO_ROOT) or "").strip()
    if repo:
        cand = os.path.join(os.path.abspath(repo), "code")
        if os.path.isdir(cand):
            return cand
    return ""


def _read_excel_pandas_openpyxl(path: str, sheet_name: str | int) -> pd.DataFrame:
    """
    Load one sheet with pandas/openpyxl.

    Some third-party .xlsx files carry inconsistent style XML; full parse then fails inside
    openpyxl (e.g. IndexError in stylesheet._merge_named_styles). Retrying with
    openpyxl read_only=True avoids loading full styles and usually succeeds for cell values.
    """
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except IndexError as err:
        _LOG.warning(
            "pd.read_excel failed (openpyxl stylesheet); retry read_only: path=%r sheet=%r err=%s",
            path,
            sheet_name,
            err,
        )
        return pd.read_excel(
            path,
            sheet_name=sheet_name,
            engine_kwargs={"read_only": True},
        )


def read_tabular_dataframe(
    path: str,
    *,
    sheet_name: str | None = None,
    csv_encoding: str | None = None,
) -> pd.DataFrame:
    """Load CSV / Parquet / Excel by extension."""
    low = path.lower()
    if low.endswith(".csv"):
        enc = csv_encoding
        if enc is None:
            enc = (os.environ.get(ENV_TABULAR_CSV_ENCODING) or "").strip() or "utf-8-sig"
        return pd.read_csv(path, encoding=enc)
    if low.endswith((".parquet", ".pq")):
        return pd.read_parquet(path)
    sn: str | int = sheet_name if sheet_name else 0
    return _read_excel_pandas_openpyxl(path, sn)
