"""planning_core ?????????????``import planning_core`` ? bootstrap ???????????"""
import pandas as pd
from datetime import datetime, timedelta, time, date
from collections import Counter, defaultdict
import itertools
import csv
import json
import copy
import re

from dispatch_interval_mirror import DispatchIntervalMirror
import traceback
import base64
import hashlib
import unicodedata
import time as time_module
from google import genai
import logging
import calendar
import math
import os
import fnmatch
import shutil
import sys
import ctypes
from contextlib import contextmanager
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .bootstrap import (
    PlanningValidationError,
    _clear_stage2_blocking_message_file,
    _remove_prior_stage2_workbooks_and_prune_empty_dirs,
    _try_remove_path_with_retries,
    _write_stage2_blocking_message,
    api_payment_dir,
    json_data_dir,
    log_dir,
    output_dir,
)

PLAN_DUE_DAY_COMPLETION_TIME = time(16, 0)

# AI ??????????? D?E ? TTL ??????? output/ ?? json/ ????
_ai_remarks_cache_name = "ai_remarks_cache.json"
_ai_cache_legacy = os.path.join(output_dir, _ai_remarks_cache_name)
_ai_cache_new = os.path.join(json_data_dir, _ai_remarks_cache_name)
if os.path.isfile(_ai_cache_legacy) and not os.path.isfile(_ai_cache_new):
    try:
        shutil.move(_ai_cache_legacy, _ai_cache_new)
    except OSError:
        pass
ai_cache_path = _ai_cache_new
# ???_???????????????????????execution_log ????
exclude_rules_sheet_debug_log_path = os.path.join(log_dir, "exclude_rules_sheet_debug.txt")
# region agent log
_AGENT_DEBUG_SESSION = "199241"
_AGENT_DEBUG_LOG_PATH = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__), "..", "..", "..", "..", "..", "debug-199241.log"
    )
)
_AGENT_DEBUG_FOCUS_TID = os.environ.get("AGENT_DEBUG_FOCUS_TID", "W4-13").strip().upper()


def _agent_debug_ndjson(
    *,
    hypothesis_id: str,
    location: str,
    message: str,
    data: dict | None = None,
    run_id: str = "pre-fix",
) -> None:
    try:
        rec = {
            "sessionId": _AGENT_DEBUG_SESSION,
            "runId": run_id,
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data or {},
            "timestamp": int(time_module.time() * 1000),
        }
        with open(_AGENT_DEBUG_LOG_PATH, "a", encoding="utf-8") as _f:
            _f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception:
        pass


# endregion
# ?????? E ????????????????? run_exclude_rules_sheet_maintenance ????????json ?????
EXCLUDE_RULES_E_SIDECAR_FILENAME = "exclude_rules_e_column_pending.json"
# openpyxl ?????? VBA ? E ????????? UTF-8 TSV?Base64??
EXCLUDE_RULES_E_VBA_TSV_FILENAME = "exclude_rules_e_column_vba.tsv"
# openpyxl ?????? VBA ? A?E ??????? UTF-8 TSV????? 5 ??? Base64??
EXCLUDE_RULES_MATRIX_VBA_FILENAME = "exclude_rules_matrix_vba.tsv"
# VBA ???? P ????????? UTF-8 ?????Excel ????? save ??????????
GEMINI_USAGE_SUMMARY_FOR_MAIN_FILE = "gemini_usage_summary_for_main.txt"
# ??????? Gemini ???????????API ???????????? API_Payment ?????
GEMINI_USAGE_CUMULATIVE_JSON_FILE = "gemini_usage_cumulative.json"
# ??????????????? CSV?Excel ???????????
GEMINI_USAGE_BUCKETS_CSV_FILE = "gemini_usage_buckets_for_chart.csv"
# ???????Gemini ?????xlwings: Q?R??????????S?T????????
GEMINI_USAGE_CHART_COL_DATE = 17  # Q
GEMINI_USAGE_CHART_COL_VALUE = 18  # R
GEMINI_USAGE_CHART_COL_TOK_DATE = 19  # S????????????
GEMINI_USAGE_CHART_COL_TOK_VALUE = 20  # T?total_tokens ???
GEMINI_USAGE_CHART_HEADER_ROW = 16
GEMINI_USAGE_CHART_ANCHOR_CELL = "T16"
GEMINI_USAGE_CHART_TOKENS_ANCHOR_CELL = "AA16"
GEMINI_USAGE_CHART_MAX_DAYS = 14
GEMINI_USAGE_CHART_CLEAR_ROWS = 36
# xlwings ??????????????????????????
GEMINI_USAGE_XLW_CHART_NAME = "_GeminiApiDailyTrend"
GEMINI_USAGE_XLW_CHART_TOKENS_NAME = "_GeminiApiDailyTokens"
# ???: EXCLUDE_RULES_TEST_E1234=1 ? EXCLUDE_RULES_SHEET_NAME????_????????? E ?? "1234" ??????????????
# TASK_INPUT_WORKBOOK ??????DATA???????????: ????_AI?????.xlsm?????????
# ?? EXCLUDE_RULES_TEST_E1234_ROW??? 9?2 ??? 9 ??????

# =========================================================
# ????API?? / ????? / ?????
# =========================================================
# Gemini API ??? TASK_INPUT_WORKBOOK ??????????B1 ? JSON ?????????? format_version 2 ??????
# ?????????????? GEMINI_API_KEY ????

GEMINI_MODEL_FLASH = "gemini-2.5-flash"
# ????: USD / 1M tokens???, ???????????????????????
# ???? GEMINI_PRICE_USD_IN_PER_M / GEMINI_PRICE_USD_OUT_PER_M ??????Flash ????
_GEMINI_FLASH_IN_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_IN_PER_M", "0.075") or 0.075
)
_GEMINI_FLASH_OUT_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_OUT_PER_M", "0.30") or 0.30
)
GEMINI_JPY_PER_USD = float(os.environ.get("GEMINI_JPY_PER_USD", "150") or 150)

# ---------------------------------------------------------------------------
# ???????????Excel ?????? 1:1 ???????
# ????????? VBA?????????????????
# ---------------------------------------------------------------------------

MASTER_FILE = "master.xlsm" # skills?attendance(???tasks)?????????
# VBA?master_???????????????1 ???????????2? machine_avail_dt ????
SHEET_MACHINE_CALENDAR = "???????"
# ``generate_plan`` ????????date -> ???? -> [ (start, end), ... ] ???? [start, end)
_MACHINE_CALENDAR_BLOCKS_BY_DATE: dict[
    date, dict[str, list[tuple[datetime, datetime]]]
] = {}

# master.xlsm: ??NO ?????????�?????????????????????????????
SHEET_MACHINE_CHANGEOVER = "??_????????"
SHEET_MACHINE_DAILY_STARTUP = "??_??_??????"
# ``generate_plan`` ??????????????????????????
_STAGE2_MACHINE_CHANGEOVER_BY_EQ: dict[str, tuple[int, int]] = {}
_STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE: dict[str, int] = {}
# master ??? A15???????????????? forward ???? [??, ??+N?) ?????????
_STAGE2_REGULAR_SHIFT_START: time | None = None
# timeline_events ? event_kind????????????
TIMELINE_EVENT_MACHINING = "machining"
TIMELINE_EVENT_MACHINE_DAILY_STARTUP = "machine_daily_startup"
TIMELINE_EVENT_CHANGEOVER_CLEANUP = "changeover_cleanup"
TIMELINE_EVENT_CHANGEOVER_PREP = "changeover_prep"
# VBA?master_???????????????????+????????????
MASTER_SHEET_TEAM_COMBINATIONS = "??????"
# ??????????: master.xlsm ?????????????????
# ??AI???????????? reason??????????????????/????????????????????????
# master ?????????.txt ??: ??=?????????1_????????????14:45?15:00????=?????????1_?????????=??????
# ??????????? AI ???????????is_holiday?????????????????????? AI ????
ATT_COL_LEAVE_TYPE = "????"
ATT_COL_REMARK = "??"
# ??????????master.xlsm?: ?????????????????????????
ATT_COL_OT_END = "????"
# ???? AI ? JSON ??????????????????????????
ATTENDANCE_REMARK_AI_SCHEMA_ID = "v2_haitai_fuka"
# need ???: ??????????A????????????? ???????????????????????
# ?Excel ???? 5 ????????????????????�????
# ? ??????1???????99???????????1?99?
NEED_COL_CONDITION = "??NO??"
NEED_COL_NOTE = "??"
# need??????????????????????????????????????????: 0.05 ? 5%?
SURPLUS_TEAM_MAX_SPEEDUP_RATIO = 0.05
# ???? tasks.xlsx ?????VBA ???? TASK_INPUT_WORKBOOK ??????DATA???
TASKS_INPUT_WORKBOOK = os.environ.get("TASK_INPUT_WORKBOOK", "").strip()
TASKS_SHEET_NAME = "????DATA"

# ????????????? openpyxl ??????????????????load_workbook ??????
OPENPYXL_INCOMPATIBLE_SHEET_MARKER = "??_??????"


def _ooxml_workbook_sheet_names(wb_path: str) -> list[str] | None:
    """ZIP ? xl/workbook.xml ????????????openpyxl ???????"""
    import zipfile
    import xml.etree.ElementTree as ET

    if not wb_path or not os.path.isfile(wb_path):
        return None
    low = wb_path.lower()
    if not low.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        return None
    try:
        with zipfile.ZipFile(wb_path, "r") as zf:
            if "xl/workbook.xml" not in zf.namelist():
                return None
            raw = zf.read("xl/workbook.xml")
    except (OSError, zipfile.BadZipFile, KeyError):
        return None
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return None
    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    tag_sheet = "{%s}sheet" % ns_main
    names: list[str] = []
    for el in root.iter():
        if el.tag == tag_sheet or el.tag.endswith("}sheet"):
            n = el.get("name")
            if n:
                names.append(n)
    return names


def _workbook_should_skip_openpyxl_io(wb_path: str) -> bool:
    """????? OOXML ???????_???????????? True?openpyxl ????????"""
    p = (wb_path or "").strip()
    if not p:
        return False
    names = _ooxml_workbook_sheet_names(p)
    if not names:
        return False
    return OPENPYXL_INCOMPATIBLE_SHEET_MARKER in names


# ??????????B1: ?????? Gemini ?? JSON ???
APP_CONFIG_SHEET_NAME = "??"
# ????? JSON?format_version 2???????????????????????????????????UI ???????
_GEMINI_CREDENTIALS_PASSPHRASE_FIXED = "nagaoka1234"
_GEMINI_CREDENTIALS_PBKDF2_ITERATIONS_DEFAULT = 480_000


def _config_cell_text(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _resolve_path_relative_to_workbook(wb_path: str, user_path: str) -> str:
    p = (user_path or "").strip().strip('"')
    if not p:
        return ""
    if os.path.isabs(p):
        return os.path.normpath(p)
    base = os.path.dirname(os.path.abspath(wb_path))
    return os.path.normpath(os.path.join(base, p))


def _read_gemini_credentials_json_path_from_workbook(wb_path: str) -> str | None:
    """??????? B1 ???? JSON ?????????????? None?"""
    if not wb_path or not os.path.isfile(wb_path):
        return None
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.debug(
            "Gemini: ?????%s?????? openpyxl ??%s?!B1 ???????",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
            APP_CONFIG_SHEET_NAME,
        )
        return None
    try:
        keep_vba = str(wb_path).lower().endswith(".xlsm")
        wb = load_workbook(
            wb_path, read_only=True, data_only=True, keep_vba=keep_vba
        )
        try:
            if APP_CONFIG_SHEET_NAME not in wb.sheetnames:
                return None
            ws = wb[APP_CONFIG_SHEET_NAME]
            b1 = _config_cell_text(ws.cell(row=1, column=2).value)
        finally:
            wb.close()
    except Exception as ex:
        logging.debug(
            "Gemini: ???????%s???%s?!B1 ??????: %s",
            wb_path,
            APP_CONFIG_SHEET_NAME,
            ex,
        )
        return None
    if not b1:
        return None
    return _resolve_path_relative_to_workbook(wb_path, b1) or None


def _read_task_ids_from_config_sheet_column(
    wb_path: str,
    column_index: int,
    log_label: str,
    column_letter_desc: str,
    *,
    openpyxl_skip_hint: str | None = None,
) -> list[str]:
    """
    ??????????????????1=A, 2=B?3 ????????NO????
    ??????????? 30 ??????????? 500 ?????????????
    """
    out: list[str] = []
    if not wb_path or not os.path.isfile(wb_path):
        return out
    if _workbook_should_skip_openpyxl_io(wb_path):
        msg = (
            f"{log_label}: ?????{OPENPYXL_INCOMPATIBLE_SHEET_MARKER}??????"
            f"?{APP_CONFIG_SHEET_NAME}?!{column_letter_desc}3 ??? openpyxl ???????"
        )
        if openpyxl_skip_hint:
            msg += " " + openpyxl_skip_hint.strip()
        logging.info(msg)
        return out
    try:
        keep_vba = str(wb_path).lower().endswith(".xlsm")
        wb = load_workbook(
            wb_path, read_only=True, data_only=True, keep_vba=keep_vba
        )
        try:
            if APP_CONFIG_SHEET_NAME not in wb.sheetnames:
                return out
            ws = wb[APP_CONFIG_SHEET_NAME]
            consecutive_empty = 0
            for r in range(3, 3 + 500):
                t = _config_cell_text(ws.cell(row=r, column=column_index).value)
                if not t:
                    consecutive_empty += 1
                    if consecutive_empty >= 30:
                        break
                    continue
                consecutive_empty = 0
                if "," in t:
                    for part in t.split(","):
                        p = part.strip()
                        if p:
                            out.append(p)
                else:
                    out.append(t)
        finally:
            wb.close()
    except Exception as ex:
        logging.warning(
            "%s: ?%s?!%s3 ?????NO??????????: %s",
            log_label,
            APP_CONFIG_SHEET_NAME,
            column_letter_desc,
            ex,
        )
        return []
    return out


def _read_trace_schedule_task_ids_from_config_sheet(wb_path: str) -> list[str]:
    """
    ????????????? A ?? 3 ?????????????????NO??????
    ??????????? 30 ???????????? 500 ??????
    """
    return _read_task_ids_from_config_sheet_column(
        wb_path,
        1,
        "??????",
        "A",
        openpyxl_skip_hint="?????????????? A ?? openpyxl ????????????",
    )


def _read_debug_dispatch_task_ids_from_config_sheet(wb_path: str) -> list[str]:
    """
    ????????????? B ?? 3 ????????2???????????NO??????
    1 ?????????2??????????????????????? A ?????????
    """
    return _read_task_ids_from_config_sheet_column(
        wb_path,
        2,
        "??????",
        "B",
        openpyxl_skip_hint="?????????????? B ?? openpyxl ??????????????????",
    )


def _show_stage2_debug_dispatch_mode_dialog(task_ids_sorted: list[str]) -> None:
    """????? B3??????????????Windows ?? MessageBox?????? WARNING ???"""
    if not task_ids_sorted:
        return
    preview_lines = task_ids_sorted[:30]
    preview = "\n".join(preview_lines)
    if len(task_ids_sorted) > 30:
        preview += "\n…"
    body = (
        "??????????????\n\n"
        "??????? B3?????????NO????????????\n\n"
        "????NO:\n"
        + preview
    )
    title = "??2????� ???????"
    if sys.platform != "win32":
        logging.warning("%s\n%s", title, body)
        return
    try:
        ctypes.windll.user32.MessageBoxW(0, body, title, 0x00000040)
    except Exception as ex:
        logging.warning(
            "??????: ????????????????? (%s)?%s", ex, body
        )


def _extract_gemini_api_key_from_plain_dict(data: dict, json_path: str) -> str | None:
    key = data.get("gemini_api_key")
    if key is None or (isinstance(key, str) and not key.strip()):
        key = data.get("GEMINI_API_KEY")
    if key is None:
        logging.warning(
            "Gemini: ?????? gemini_api_key???? GEMINI_API_KEY????????%s??",
            json_path,
        )
        return None
    s = str(key).strip()
    return s or None


def _derive_fernet_key_from_passphrase(
    passphrase: str, salt: bytes, iterations: int
) -> bytes:
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.backends import default_backend
    from cryptography.hazmat.primitives import hashes

    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=iterations,
        backend=default_backend(),
    )
    return base64.urlsafe_b64encode(kdf.derive(passphrase.encode("utf-8")))


def _credentials_json_is_encrypted_v2(data: dict) -> bool:
    if data.get("format_version") == 2:
        return True
    return bool(
        data.get("kdf") == "pbkdf2_sha256" and (data.get("fernet_ciphertext") or "").strip()
    )


def _decrypt_gemini_credentials_v2(
    data: dict, passphrase: str, json_path: str
) -> dict | None:
    try:
        from cryptography.fernet import Fernet
    except ImportError:
        logging.warning(
            "Gemini: ????? JSON ?? cryptography ??????pip install cryptography??"
        )
        return None
    token_s = (data.get("fernet_ciphertext") or "").strip()
    if not token_s:
        logging.warning(
            "Gemini: ????? JSON ? fernet_ciphertext ???????%s??",
            json_path,
        )
        return None
    salt_b64 = (data.get("salt_b64") or "").strip()
    if not salt_b64:
        logging.warning(
            "Gemini: ????? JSON ? salt_b64 ???????%s??",
            json_path,
        )
        return None
    try:
        salt = base64.standard_b64decode(salt_b64)
    except Exception as ex:
        logging.warning("Gemini: salt_b64 ???????????%s?: %s", json_path, ex)
        return None
    iterations = int(data.get("iterations") or _GEMINI_CREDENTIALS_PBKDF2_ITERATIONS_DEFAULT)
    kdf_name = (data.get("kdf") or "pbkdf2_sha256").strip()
    if kdf_name != "pbkdf2_sha256":
        logging.warning("Gemini: ???? kdf?%s?: %s", kdf_name, json_path)
        return None
    try:
        fkey = _derive_fernet_key_from_passphrase(passphrase, salt, iterations)
        plain = Fernet(fkey).decrypt(token_s.encode("ascii"))
    except Exception:
        logging.debug("Gemini: ??????????????????%s??", json_path)
        return None
    try:
        inner = json.loads(plain.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as ex:
        logging.warning(
            "Gemini: ???? JSON ??????%s?: %s",
            json_path,
            ex,
        )
        return None
    if not isinstance(inner, dict):
        logging.warning("Gemini: ???? JSON ??????????????????%s??", json_path)
        return None
    return inner


def _load_gemini_api_key_from_credentials_json(
    json_path: str, workbook_path: str | None = None
) -> tuple[str | None, bool]:
    """???: (api_key ??? None, ?????????)?????? _GEMINI_CREDENTIALS_PASSPHRASE_FIXED ??????"""
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
    except OSError as ex:
        logging.warning("Gemini: ?? JSON ??????: %s (%s)", json_path, ex)
        return None, False
    except json.JSONDecodeError as ex:
        logging.warning("Gemini: ?? JSON ????????: %s (%s)", json_path, ex)
        return None, False
    if not isinstance(data, dict):
        logging.warning("Gemini: ?? JSON ???????????????????: %s", json_path)
        return None, False
    if _credentials_json_is_encrypted_v2(data):
        inner = _decrypt_gemini_credentials_v2(
            data, _GEMINI_CREDENTIALS_PASSPHRASE_FIXED, json_path
        )
        if inner is None:
            return None, True
        return _extract_gemini_api_key_from_plain_dict(inner, json_path), True
    return _extract_gemini_api_key_from_plain_dict(data, json_path), False


API_KEY = None
_cred_path = _read_gemini_credentials_json_path_from_workbook(TASKS_INPUT_WORKBOOK)
_used_encrypted_credentials = False
if _cred_path and os.path.isfile(_cred_path):
    API_KEY, _used_encrypted_credentials = _load_gemini_api_key_from_credentials_json(
        _cred_path, workbook_path=TASKS_INPUT_WORKBOOK
    )
    if API_KEY:
        if _used_encrypted_credentials:
            logging.info("Gemini API ??: ????? JSON ??????????")
        else:
            logging.info(
                "Gemini API ??: ???????%s?B1 ?????????????",
                APP_CONFIG_SHEET_NAME,
            )
elif _cred_path:
    logging.warning(
        "Gemini: ?%s?B1 ???????? JSON ?????????",
        APP_CONFIG_SHEET_NAME,
    )

# B1 ???? JSON ????????????? JSON ?????????????????????????????????
_encrypted_json_missing_key = (
    bool(_cred_path)
    and os.path.isfile(_cred_path)
    and _used_encrypted_credentials
    and not API_KEY
)
if _encrypted_json_missing_key:
    logging.error(
        "Gemini: ?%s?B1 ????????? API ???????????"
        " ????????????????????????????????",
        APP_CONFIG_SHEET_NAME,
    )

if not API_KEY:
    logging.warning(
        "Gemini API ?????????????????%s???????? B1 ??? JSON ??????????????"
        " ??? AI ?????????????"
        " ???: gemini_credentials.example.json / encrypt_gemini_credentials.py??????",
        APP_CONFIG_SHEET_NAME,
    )

RESULT_SHEET_GANTT_NAME = "??_?????"

# ???????????????????DATA??
TASK_COL_TASK_ID = "??NO"
TASK_COL_MACHINE = "???"
TASK_COL_MACHINE_NAME = "???"
TASK_COL_QTY = "????"
TASK_COL_ORDER_QTY = "???"
TASK_COL_SPEED = "????"
TASK_COL_PRODUCT = "???"
TASK_COL_ANSWER_DUE = "????"
TASK_COL_SPECIFIED_DUE = "????"
TASK_COL_RAW_INPUT_DATE = "?????"
# ????DATA ???????_????????????????SOURCE_BASE_COLUMNS ???
TASK_COL_STOCK_LOCATION = "????"
# ????NO?????????????????????DATA?????_??????
TASK_COL_PROCESS_CONTENT = "????"
# ???????????????????????????????????/???????????????????????
# ????????????DATA?
TASK_COL_COMPLETION_FLAG = "??????"
TASK_COL_ACTUAL_DONE = "????"   # ?????????????
TASK_COL_ACTUAL_OUTPUT = "????"  # ????????????
TASK_COL_DATA_EXTRACTION_DT = "??????"
AI_CACHE_TTL_SECONDS = 6 * 60 * 60  # 6??
# json/ai_remarks_cache.json ??????????_??????????????? D?E?
AI_CACHE_KEY_PREFIX_EXCLUDE_RULE_DE = "exclude_rule_de_v1"

# ???????????DATA??Power Query ?????????
ACTUALS_SHEET_NAME = "????DATA"
ACT_COL_TASK_ID = "??NO"
ACT_COL_PROCESS = "???"
ACT_COL_OPERATOR = "???"
ACT_COL_START_DT = "????"
ACT_COL_END_DT = "????"
ACT_COL_START_ALT = "????"
ACT_COL_END_ALT = "????"
ACT_COL_DAY = "??"
ACT_COL_TIME_START = "????"
ACT_COL_TIME_END = "????"
ACTUAL_HEADER_CANONICAL = (
    ACT_COL_TASK_ID,
    ACT_COL_PROCESS,
    ACT_COL_OPERATOR,
    ACT_COL_START_DT,
    ACT_COL_END_DT,
    ACT_COL_START_ALT,
    ACT_COL_END_ALT,
    ACT_COL_DAY,
    ACT_COL_TIME_START,
    ACT_COL_TIME_END,
)

# --- 2????: ??1?? ? ????????_???????? ? ??2?? ---
STAGE1_OUTPUT_FILENAME = "plan_input_tasks.xlsx"
PLAN_INPUT_SHEET_NAME = os.environ.get("TASK_PLAN_SHEET", "").strip() or "????_?????"
PLAN_COL_SPEED_OVERRIDE = "????_???"
# ???????????????????DATA ?????????????????????????????_???????????????????
PLAN_COL_RAW_INPUT_DATE_OVERRIDE = "?????_???"
PLAN_COL_PREFERRED_OP = "??OP_??"
PLAN_COL_SPECIAL_REMARK = "????_??"
# ??????????????????????????????????????
# ?????????????: Excel ? TRUE / ?? 1 / ????????yes??true????????????
# ??FALSE?0???????????????? _plan_row_exclude_from_assignment?
PLAN_COL_EXCLUDE_FROM_ASSIGNMENT = "????"
PLAN_COL_AI_PARSE = "AI????_??"
PLAN_COL_PROCESS_FACTOR = "??????????????"
# 1??????????m??????_???????????????DATA ?????????????????
PLAN_COL_ROLL_UNIT_LENGTH = "???????"
DEBUG_TASK_ID = os.environ.get("DEBUG_TASK_ID", "Y3-26").strip()
# ?: set TRACE_TEAM_ASSIGN_TASK_ID=W3-14 � ?????????????????????? INFO ?????
TRACE_TEAM_ASSIGN_TASK_ID = os.environ.get("TRACE_TEAM_ASSIGN_TASK_ID", "").strip()
# ?????????????????????? A ? 3 ???????generate_plan ?????????????????
TRACE_SCHEDULE_TASK_IDS: frozenset[str] = frozenset()
# ??2??????: ????B ? 3 ???????NO???????????????????generate_plan ?????????????
DEBUG_DISPATCH_ONLY_TASK_IDS: frozenset[str] = frozenset()
# ????????????????0=????????????? while ?????????????????????????????
DISPATCH_TRACE_OUTER_ROUND: int = 0


def _trace_schedule_task_enabled(task_id) -> bool:
    if not TRACE_SCHEDULE_TASK_IDS:
        return False
    return str(task_id or "").strip() in TRACE_SCHEDULE_TASK_IDS


def _sanitize_dispatch_trace_filename_part(task_id: str) -> str:
    """??NO? log ???????????????????Windows ??????????"""
    s = "".join(
        c if (c.isalnum() or c in "-_.") else "_"
        for c in str(task_id or "").strip()
    )
    return s[:120] if s else "task"


def _reset_dispatch_trace_per_task_logfiles() -> None:
    """
    ??2??????1??log ?? dispatch_trace_*.txt ???????????????????????
    ????????????? generate_plan ? while ??? _dispatch_trace_begin_outer_round ?????????????
    execution_log.txt ??????????? [?????? task=�] ?? _log_dispatch_trace_schedule ???
    ??????????????????????????????
    """
    if not TRACE_SCHEDULE_TASK_IDS:
        return
    try:
        os.makedirs(log_dir, exist_ok=True)
    except OSError:
        return
    try:
        for _name in os.listdir(log_dir):
            if not (
                str(_name).startswith("dispatch_trace_") and str(_name).endswith(".txt")
            ):
                continue
            _p = os.path.join(log_dir, _name)
            try:
                os.unlink(_p)
            except OSError:
                pass
    except OSError:
        pass


def _dispatch_trace_begin_outer_round(round_n: int) -> None:
    """???????????????????????????? dispatch_trace_*_rNN.txt ?????1??????"""
    global DISPATCH_TRACE_OUTER_ROUND
    DISPATCH_TRACE_OUTER_ROUND = max(0, int(round_n))
    if not TRACE_SCHEDULE_TASK_IDS:
        return
    try:
        os.makedirs(log_dir, exist_ok=True)
    except OSError:
        return
    for tid in TRACE_SCHEDULE_TASK_IDS:
        t = str(tid or "").strip()
        if not t:
            continue
        safe = _sanitize_dispatch_trace_filename_part(t)
        path = os.path.join(
            log_dir,
            f"dispatch_trace_{safe}_r{DISPATCH_TRACE_OUTER_ROUND:02d}.txt",
        )
        if os.path.exists(path):
            continue
        try:
            with open(path, "w", encoding="utf-8", newline="\n") as f:
                f.write(
                    "# ?????????NO???????????????? log/execution_log.txt ?????????\n"
                    f"# task_id={t}  outer_round={DISPATCH_TRACE_OUTER_ROUND}  "
                    "# ?0=???????????????????????? +1?\n\n"
                )
        except OSError as ex:
            logging.warning("dispatch_trace ?????????: %s (%s)", path, ex)


def _log_dispatch_trace_schedule(task_id, msg: str, *args) -> None:
    """[?????? task=�] ? execution_log ??????????NO?????????????"""
    t = str(task_id or "").strip()
    body_raw = msg % args if args else msg
    body = body_raw
    if t and t in TRACE_SCHEDULE_TASK_IDS:
        body = f"[outer_round={DISPATCH_TRACE_OUTER_ROUND:02d}] {body_raw}"
    logging.info(body)
    if not t or t not in TRACE_SCHEDULE_TASK_IDS:
        return
    safe = _sanitize_dispatch_trace_filename_part(t)
    path = os.path.join(
        log_dir,
        f"dispatch_trace_{safe}_r{DISPATCH_TRACE_OUTER_ROUND:02d}.txt",
    )
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S,%f")[:-3]
        line = f"{ts} - INFO - {body}\n"
        with open(path, "a", encoding="utf-8", newline="\n") as f:
            f.write(line)
    except OSError as ex:
        try:
            logging.warning("dispatch_trace ????????????: %s (%s)", path, ex)
        except Exception:
            pass


# True: ????????????? (-??, ??, -???, ?????)?False ????????????????
TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF = os.environ.get(
    "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF", "0"
).strip().lower() not in ("0", "false", "no", "off", "???")


def _team_assign_start_slack_wait_minutes() -> int:
    """???????????????????????????????????0 ????"""
    raw = os.environ.get("TEAM_ASSIGN_START_SLACK_WAIT_MINUTES", "60").strip()
    try:
        v = int(raw)
    except ValueError:
        v = 60
    return max(0, v)


TEAM_ASSIGN_START_SLACK_WAIT_MINUTES = _team_assign_start_slack_wait_minutes()

# True ??? need ???????????????????????????????req_num??????????????????
TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW = (
    os.environ.get("TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW", "0")
    .strip()
    .lower()
    in ("1", "true", "yes", "on", "??")
)

# True: ??????????????????? req_num?req_num+???????????
# False????: ???? req_num ????????????????????????????????
#     ??????????????????? skills ?????????????append_surplus_staff_after_main_dispatch??
TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS = (
    os.environ.get("TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS", "")
    .strip()
    .lower()
    in ("1", "true", "yes", "on", "??")
)

# True????: ??????????? need????????????????
# False ?????????? AI ? required_op ??????????????????????????????
TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY = (
    os.environ.get("TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "???")
)
# True????: master?????????????????+??????????????
# ??????????????????????????? itertools ??????
TEAM_ASSIGN_USE_MASTER_COMBO_SHEET = (
    os.environ.get("TEAM_ASSIGN_USE_MASTER_COMBO_SHEET", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "???")
)

# �B-2 ???????????????????????1?????????????????????????
# 0 / false / no / off ?????????????????????????????????????: W3-14 ?????????
PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE = (
    os.environ.get("PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "???", "??")
)

# �B-2 / �B-3 ????? EC ????????????????????????
# 0 / false / no / off / ??? / ?? ??????????????????????????????????????
PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS = (
    os.environ.get("PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "???", "??")
)

# ?????????_???????: ???? openpyxl save ???? xlwings ???Save?Excel ???? openpyxl ??????????????? TSV?VBA ???
# ?????? openpyxl ?????? EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1?
EXCLUDE_RULES_SHEET_NAME = "??_??????"
EXCLUDE_RULES_SKIP_OPENPYXL_SAVE = os.environ.get(
    "EXCLUDE_RULES_TRY_OPENPYXL_SAVE", ""
).strip().lower() not in ("1", "true", "yes", "on")
EXCLUDE_RULE_COL_PROCESS = "???"
EXCLUDE_RULE_COL_MACHINE = "???"
EXCLUDE_RULE_COL_FLAG = "????"
EXCLUDE_RULE_COL_LOGIC_JA = "????????"
EXCLUDE_RULE_COL_LOGIC_JSON = "?????"
# ????????????????????????????????????????
_exclude_rules_effective_read_path: str | None = None
# ??? apply_exclude_rules?????????: VBA ????? E ?????????
_exclude_rules_rules_snapshot: list | None = None
_exclude_rules_snapshot_wb: str | None = None
# ??? JSON ? conditions ????????AI ????????????????
EXCLUDE_RULE_ALLOWED_COLUMNS = frozenset(
    {
        TASK_COL_TASK_ID,
        TASK_COL_MACHINE,
        TASK_COL_MACHINE_NAME,
        TASK_COL_QTY,
        TASK_COL_ORDER_QTY,
        TASK_COL_SPEED,
        TASK_COL_PRODUCT,
        TASK_COL_ANSWER_DUE,
        TASK_COL_SPECIFIED_DUE,
        TASK_COL_RAW_INPUT_DATE,
        TASK_COL_STOCK_LOCATION,
        TASK_COL_PROCESS_CONTENT,
        TASK_COL_COMPLETION_FLAG,
        TASK_COL_ACTUAL_DONE,
        TASK_COL_ACTUAL_OUTPUT,
        TASK_COL_DATA_EXTRACTION_DT,
        PLAN_COL_SPEED_OVERRIDE,
        PLAN_COL_RAW_INPUT_DATE_OVERRIDE,
        PLAN_COL_PREFERRED_OP,
        PLAN_COL_SPECIAL_REMARK,
        PLAN_COL_PROCESS_FACTOR,
        PLAN_COL_ROLL_UNIT_LENGTH,
    }
)

# ??????????_?????????????????????????????????
RESULT_TASK_SHEET_NAME = "??_?????"
RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME = "??_???????"
# ???????????????????_?????????? 10 ???????
TEMP_EQUIPMENT_SCHEDULE_SHEET_NAME = "TEMP_???????"
# ?????????????? 10 ?????????????
BLOCK_TABLE_SHEET_NAME = "????????"
# ???+??????????????????????NO?????????
RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME = "??_???????_????"
# master ??? A15/B15 ??????????????????_??????????_??????
RESULT_OUTSIDE_REGULAR_TIME_FILL = "FCE4D6"
# ??_???????_????: ??????NO??????????????
# ??_??????????? TEMP?: ??????????????????????
RESULT_DISPATCHED_REQUEST_FILL = "C6EFCE"
# ??_???????: master????????????????????10???
RESULT_MACHINE_CALENDAR_BLOCK_FILL = "D4B3E8"
# ??_?????: ????????????????????? B?E ?????????????????
RESULT_EQUIP_GANTT_MACHINE_GROUP_FILL_COLORS = (
    "E8F4FC",
    "FCE8F0",
    "E8F8E8",
    "FFF0D8",
    "EDE8FC",
    "E0F8F4",
    "F8E8E0",
    "E8ECF8",
    "F5F5E0",
    "F0E8E8",
)
# ???????????????? task_queue.sort ????????1 ??????????
RESULT_TASK_COL_DISPATCH_TRIAL_ORDER = "??????"
# ??_??????????+16:00?????????+16:00????????????????
RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16 = "??_????16???"
# ??? skills ???+?????? OP/AS ???????????????????????????
RESULT_MEMBER_PRIORITY_SHEET_NAME = "??_???????"
COLUMN_CONFIG_SHEET_NAME = "???_??_?????"
COLUMN_CONFIG_HEADER_COL = "??"
COLUMN_CONFIG_VISIBLE_COL = "??"
# ??2??? xlsx ??????????????????????????????? xlwings ???????? ON????? STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT=0?
STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT = os.environ.get(
    "STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT", "1"
).strip().lower() in ("1", "true", "yes", "on")
# ??_??????????yyyy/mm/dd ??????????????????????????
RESULT_TASK_DATE_STYLE_HEADERS = frozenset(
    {
        "????",
        "????",
        "??????",
        TASK_COL_RAW_INPUT_DATE,
        "?????",
        "??_????",
        "??_????",
    }
)

SOURCE_BASE_COLUMNS = [
    TASK_COL_TASK_ID, TASK_COL_MACHINE, TASK_COL_MACHINE_NAME, TASK_COL_QTY, TASK_COL_ORDER_QTY, TASK_COL_SPEED, TASK_COL_PRODUCT,
    TASK_COL_ANSWER_DUE, TASK_COL_SPECIFIED_DUE, TASK_COL_RAW_INPUT_DATE, TASK_COL_STOCK_LOCATION,
    TASK_COL_PROCESS_CONTENT,
    TASK_COL_COMPLETION_FLAG, TASK_COL_ACTUAL_DONE, TASK_COL_ACTUAL_OUTPUT,
]
PLAN_OVERRIDE_COLUMNS = [
    PLAN_COL_EXCLUDE_FROM_ASSIGNMENT,
    PLAN_COL_SPEED_OVERRIDE,
    PLAN_COL_RAW_INPUT_DATE_OVERRIDE,
    PLAN_COL_PREFERRED_OP,
    PLAN_COL_SPECIAL_REMARK,
    PLAN_COL_AI_PARSE,
]
# ???????????????????????????????
PLAN_CONFLICT_STYLABLE_COLS = tuple(PLAN_OVERRIDE_COLUMNS)
# ??1????????????_??????????????AI??????????????
PLAN_STAGE1_MERGE_COLUMNS = tuple(c for c in PLAN_OVERRIDE_COLUMNS if c != PLAN_COL_AI_PARSE)
# ????????????????????????????????????
# ???????????????? fill_plan_dispatch_trial_order_column_stage1 ???????????
PLAN_STAGE1_MERGE_EXTRA_COLUMNS = (PLAN_COL_ROLL_UNIT_LENGTH,)
# openpyxl ?????????????????VBA ?????????????????????????
PLANNING_CONFLICT_SIDECAR = "planning_conflict_highlight.tsv"
# ????_?????????????????????????????????1????????????
# ? ??????: load_planning_tasks_df ?????????????????????????????????1???????????????
PLAN_SHEET_GLOBAL_PARSE_LABEL_COL = 50  # AX
PLAN_SHEET_GLOBAL_PARSE_VALUE_COL = 51  # AY
PLAN_SHEET_GLOBAL_PARSE_MAX_ROWS = 42


def plan_reference_column_name(override_col: str) -> str:
    """????????????????????????????????????"""
    return f"???{override_col}"


def plan_input_sheet_column_order():
    """
    ????_???????????1?????2???????

    0. ?????????1????????????2??????????2?????????????????
    1. ???????????
    2. ????DATA ???SOURCE_BASE_COLUMNS?� ??NO?????????????????????????????????????
    3. ??????????????
    4. ????� ??????????�?????AI????_??????????
       ???????? ?????_??? ???????????????????????

    global_speed_rules ????????????????????????????????
    """
    cols = [RESULT_TASK_COL_DISPATCH_TRIAL_ORDER, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT]
    for c in SOURCE_BASE_COLUMNS:
        cols.append(c)
        if c == TASK_COL_PRODUCT:
            cols.append(PLAN_COL_ROLL_UNIT_LENGTH)
    cols.append(PLAN_COL_PROCESS_FACTOR)
    for c in PLAN_OVERRIDE_COLUMNS:
        if c == PLAN_COL_EXCLUDE_FROM_ASSIGNMENT:
            continue
        if c == PLAN_COL_AI_PARSE:
            cols.append(c)
        else:
            cols.append(plan_reference_column_name(c))
            cols.append(c)
    return cols


def _format_paren_ref_scalar(val):
    """?????: ?????????????????"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "???"
    if isinstance(val, datetime):
        d = val.date() if hasattr(val, "date") else val
        if isinstance(d, date):
            return f"?{d.year}/{d.month}/{d.day}?"
    if isinstance(val, date):
        return f"?{val.year}/{val.month}/{val.day}?"
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none"):
        return "???"
    return f"?{s}?"


def _reference_text_for_override_row(row, override_col: str, req_map: dict, need_rules: list) -> str:
    """1???????????????????????"""
    _ = (req_map, need_rules)  # ???????????????????????????????????????
    if override_col == PLAN_COL_SPEED_OVERRIDE:
        v = row.get(TASK_COL_SPEED)
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "???"
        try:
            x = float(v)
            if abs(x - round(x)) < 1e-9:
                return f"?{int(round(x))}?"
            return f"?{x}?"
        except (TypeError, ValueError):
            return _format_paren_ref_scalar(v)
    if override_col in (PLAN_COL_PREFERRED_OP, PLAN_COL_SPECIAL_REMARK):
        return "???"
    if override_col == PLAN_COL_RAW_INPUT_DATE_OVERRIDE:
        return _format_paren_ref_scalar(
            parse_optional_date(_planning_df_cell_scalar(row, TASK_COL_RAW_INPUT_DATE))
        )
    return "???"


def _refresh_plan_reference_columns(df, req_map: dict, need_rules: list):
    """????DATA?need ????????�??????????????????"""
    if df is None or df.empty:
        return df
    need_rules = need_rules or []
    req_map = req_map or {}
    for i in df.index:
        row = df.loc[i]
        for oc in PLAN_OVERRIDE_COLUMNS:
            if oc == PLAN_COL_AI_PARSE:
                continue
            if oc == PLAN_COL_EXCLUDE_FROM_ASSIGNMENT:
                continue
            ref_col = plan_reference_column_name(oc)
            if ref_col not in df.columns:
                continue
            df.at[i, ref_col] = _reference_text_for_override_row(row, oc, req_map, need_rules)
    return df


def _apply_plan_input_visual_format(path: str, sheet_name: str = "?????"):
    """???????????????????????AI????????"""
    # ???????????????????????????1????plan_input_sheet_column_order??
    # 1-based ???????to_excel ???????????
    fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    order = plan_input_sheet_column_order()
    col_1based = {name: i + 1 for i, name in enumerate(order)}
    if _workbook_should_skip_openpyxl_io(path):
        logging.info(
            "?????????: ?????%s?????? openpyxl ??????????????",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return
    wb = load_workbook(path)
    try:
        if sheet_name not in wb.sheetnames:
            return
        ws = wb[sheet_name]
        last_row = ws.max_row or 1
        if last_row < 2:
            return
        for oc in PLAN_OVERRIDE_COLUMNS:
            if oc == PLAN_COL_AI_PARSE:
                continue
            ci = col_1based.get(oc)
            if not ci:
                continue
            for r in range(2, last_row + 1):
                ws.cell(row=r, column=ci).fill = fill_yellow
        _ci_rul = col_1based.get(PLAN_COL_ROLL_UNIT_LENGTH)
        if _ci_rul:
            for r in range(2, last_row + 1):
                ws.cell(row=r, column=_ci_rul).fill = fill_yellow
        wb.save(path)
    finally:
        wb.close()


def _planning_conflict_sidecar_path():
    return os.path.join(log_dir, PLANNING_CONFLICT_SIDECAR)


def _remove_planning_conflict_sidecar_safe():
    try:
        os.remove(_planning_conflict_sidecar_path())
    except OSError:
        pass


def write_planning_conflict_highlight_sidecar(sheet_name, num_data_rows, conflicts_by_row):
    """
    Excel ???????????????????????VBA ?? TSV ? log ????
    ??: V1 / ???? / ????? / ????????? / ?? ???\\t??
    """
    path = _planning_conflict_sidecar_path()
    clear_cols = "\t".join(PLAN_CONFLICT_STYLABLE_COLS)
    lines = ["V1", sheet_name, str(int(num_data_rows)), clear_cols]
    for r in sorted(conflicts_by_row.keys()):
        for name in sorted(conflicts_by_row[r]):
            lines.append(f"{int(r)}\t{name}")
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")

# ??1???????????? Excel ???????(??????) ????
STAGE1_SHEET_DATEONLY_HEADERS = frozenset(
    {
        TASK_COL_ANSWER_DUE,
        TASK_COL_SPECIFIED_DUE,
        TASK_COL_RAW_INPUT_DATE,
        PLAN_COL_RAW_INPUT_DATE_OVERRIDE,
    }
)


def _result_font(**kwargs):
    """?????? Font??????? name/size ??????"""
    return Font(**kwargs)


def _output_book_font(bold=False):
    return _result_font(bold=bold)


def _apply_output_font_to_result_sheet(ws):
    """??_* ??????????: ???????1?????????? VBA AutoFit??"""
    base = _output_book_font(bold=False)
    hdr = _output_book_font(bold=True)
    mr, mc = ws.max_row or 1, ws.max_column or 1
    for row in ws.iter_rows(min_row=1, max_row=mr, min_col=1, max_col=mc):
        for cell in row:
            cell.font = base
    for cell in ws[1]:
        cell.font = hdr


def _apply_excel_date_columns_date_only_display(path, sheet_name, header_names=None):
    """openpyxl: ???????? yyyy/mm/dd ???????????????????"""
    from openpyxl import load_workbook

    headers = header_names or STAGE1_SHEET_DATEONLY_HEADERS
    if _workbook_should_skip_openpyxl_io(path):
        logging.info(
            "???????: ?????%s?????? openpyxl ??????????????",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return
    wb = load_workbook(path)
    try:
        ws = wb[sheet_name] if isinstance(sheet_name, str) else wb.worksheets[int(sheet_name)]
        cmap = {}
        for cell in ws[1]:
            if cell.value is not None:
                cmap[str(cell.value).strip()] = cell.column
        fmt = "yyyy/mm/dd"
        for h in headers:
            col = cmap.get(h)
            if not col:
                continue
            for r in range(2, ws.max_row + 1):
                c = ws.cell(row=r, column=col)
                v = c.value
                if v is None:
                    continue
                if isinstance(v, datetime):
                    c.value = v.date()
                elif isinstance(v, date):
                    pass
                else:
                    try:
                        d0 = pd.to_datetime(v, errors="coerce")
                        if pd.isna(d0):
                            continue
                        c.value = d0.date()
                    except Exception:
                        continue
                c.number_format = fmt
        wb.save(path)
    finally:
        wb.close()


def _extract_data_extraction_datetime():
    """
    `????DATA` ???? `??????` ?? datetime ??????
    """
    try:
        if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
            return None
        df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=TASKS_SHEET_NAME)
        df.columns = df.columns.str.strip()
        if TASK_COL_DATA_EXTRACTION_DT not in df.columns:
            return None
        s = df[TASK_COL_DATA_EXTRACTION_DT]
        first = None
        for v in s:
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            first = v
            break
        if first is None:
            return None
        dt = pd.to_datetime(first, errors="coerce")
        if pd.isna(dt):
            return None
        if isinstance(dt, pd.Timestamp):
            return dt.to_pydatetime()
        return dt if isinstance(dt, datetime) else None
    except Exception:
        return None


def _extract_data_extraction_datetime_str():
    """
    `????DATA` ???? `??????` ????????????????????????
    """
    try:
        dt = _extract_data_extraction_datetime()
        if dt is None:
            return "�"
        return dt.strftime("%Y/%m/%d %H:%M:%S")
    except Exception:
        return "�"


def _weekday_jp(d):
    return "???????"[d.weekday()]


# ????????????????????????????????????????????????
_GANTT_BAR_FILLS_PRINT_SAFE = (
    "E8E8E8",
    "D8E4EF",
    "E6E2DB",
    "DEEADF",
    "E8E0E8",
    "EAE8D8",
    "DDE6EA",
    "E5DCE5",
)

# ????????????????????????????
_GANTT_BAR_FILLS_ACTUAL = (
    "D4E4D4",
    "C9DDE8",
    "DED8CC",
    "D2E5CD",
    "DAD2D9",
    "E0DCCF",
    "CDE2E8",
    "DCD2DC",
)

# ?????: ???????machine_daily_startup?????????
_GANTT_DAILY_STARTUP_FILL = "FFEB9C"


def _gantt_bar_fill_for_task_id(task_id):
    """??NO???????????1??RRGGBB????????????????????"""
    h = hashlib.md5(str(task_id).encode("utf-8")).hexdigest()
    i = int(h[:8], 16) % len(_GANTT_BAR_FILLS_PRINT_SAFE)
    return _GANTT_BAR_FILLS_PRINT_SAFE[i]


def _gantt_bar_fill_actual_for_task_id(task_id):
    h = hashlib.md5(str(task_id).encode("utf-8")).hexdigest()
    i = int(h[:8], 16) % len(_GANTT_BAR_FILLS_ACTUAL)
    return _GANTT_BAR_FILLS_ACTUAL[i]


# ?????????????????: ?????? new ???
_GANTT_TIMELINE_CELL_ALIGNMENT = Alignment(
    horizontal="left",
    vertical="center",
    wrap_text=False,
    shrink_to_fit=False,
    indent=1,
)
# ???????????????? PatternFill ? hex ??????openpyxl ?????????????
_GANTT_TASK_PATTERN_FILL_BY_HEX: dict[str, PatternFill] = {}


def _gantt_cached_pattern_fill(hex_rrggbb: str) -> PatternFill:
    fi = _GANTT_TASK_PATTERN_FILL_BY_HEX.get(hex_rrggbb)
    if fi is None:
        fi = PatternFill(fill_type="solid", start_color=hex_rrggbb, end_color=hex_rrggbb)
        _GANTT_TASK_PATTERN_FILL_BY_HEX[hex_rrggbb] = fi
    return fi


def _gantt_slot_state_tuple(evlist, slot_mid, task_fill_fn=None):
    """????????????1???????('idle',) | ('break',) | ('daily_startup', fill_hex) | ('task', tid, fill_hex, pct)"""
    fill_fn = task_fill_fn or _gantt_bar_fill_for_task_id
    active = None
    for e in evlist:
        if e["start_dt"] <= slot_mid < e["end_dt"]:
            active = e
            break
    if active is None:
        return ("idle",)
    if _timeline_event_kind(active) == TIMELINE_EVENT_MACHINE_DAILY_STARTUP:
        return ("daily_startup", _GANTT_DAILY_STARTUP_FILL)
    if any(b_s <= slot_mid < b_e for b_s, b_e in active.get("breaks") or ()):
        return ("break",)
    tid = str(active["task_id"])
    gh = fill_fn(active["task_id"])
    pct = None
    try:
        # ?????????????????pct_macro ? timeline_event ??????
        if active.get("pct_macro") is not None:
            pct = int(round(parse_float_safe(active.get("pct_macro"), 0.0)))
            pct = max(0, min(100, pct))
        else:
            # ??????????????????
            tot = parse_float_safe(active.get("total_units"), 0.0)
            done = parse_float_safe(active.get("already_done_units"), 0.0) + parse_float_safe(
                active.get("units_done"), 0.0
            )
            if tot > 0:
                pct = max(0, min(100, int(round((done / tot) * 100))))
    except Exception:
        pct = None
    return ("task", tid, gh, pct)


def _gantt_timeline_same_segment(st_a, st_b) -> bool:
    """????????????????? tuple ?????????"""
    if st_a[0] != st_b[0]:
        return False
    if st_a[0] == "idle" or st_a[0] == "break":
        return True
    # daily_startup: [1]=fill / task: [1]=task_id
    return st_a[1] == st_b[1]


def _paint_gantt_timeline_row_merged(
    ws,
    row,
    n_fixed,
    slots,
    slot_mins,
    evlist,
    idle_fill,
    break_fill,
    gantt_label_font,
    grid_border,
    task_fill_fn=None,
    label_font=None,
):
    """
    ??????????????????????????????????????
    ?????????????15????????????????????????????
    """
    bar_label_font = label_font or gantt_label_font
    n_slots = len(slots)
    states = []
    for slot_start in slots:
        mid = slot_start + timedelta(minutes=slot_mins / 2)
        states.append(_gantt_slot_state_tuple(evlist, mid, task_fill_fn))
    tcol0 = n_fixed + 1
    i = 0
    while i < n_slots:
        st0 = states[i]
        j = i + 1
        while j < n_slots and _gantt_timeline_same_segment(st0, states[j]):
            j += 1
        col_s = tcol0 + i
        col_e = tcol0 + j - 1
        for col in range(col_s, col_e + 1):
            c = ws.cell(row=row, column=col)
            c.border = grid_border
            c.alignment = _GANTT_TIMELINE_CELL_ALIGNMENT
            if st0[0] == "idle":
                c.fill = idle_fill
                c.value = None
            elif st0[0] == "break":
                c.fill = break_fill
                c.value = None
            elif st0[0] == "daily_startup":
                _, gh_ds = st0
                c.fill = _gantt_cached_pattern_fill(gh_ds)
                if col == col_s:
                    c.value = "(??????)"
                    c.font = bar_label_font
                else:
                    c.value = None
            else:
                _, tid, gh, pct = st0
                c.fill = _gantt_cached_pattern_fill(gh)
                if col == col_s:
                    c.value = f"{tid[:9]} {pct}%" if pct is not None else tid[:9]
                    c.font = bar_label_font
                else:
                    c.value = None
        i = j


def _time_intervals_overlap_half_open(
    a_start: time, a_end: time, b_start: time, b_end: time
) -> bool:
    """???? [a_start, a_end) ? [b_start, b_end) ????????????"""

    def _sec(t: time) -> int:
        return t.hour * 3600 + t.minute * 60 + t.second

    return _sec(a_start) < _sec(b_end) and _sec(a_end) > _sec(b_start)


def _parse_equipment_schedule_time_band_cell(val) -> tuple[time | None, time | None]:
    """??_???????????????? 08:45-09:00?????"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None, None
    s = str(val).strip()
    if not s or "?" in s:
        return None, None
    for sep in ("-", "?", "~", "?"):
        if sep in s:
            left, right = s.split(sep, 1)
            left = left.strip().replace("?", ":")
            right = right.strip().replace("?", ":")
            t0 = parse_time_str(left, None)
            t1 = parse_time_str(right, None)
            if t0 is not None and t1 is not None and t0 < t1:
                return t0, t1
            return None, None
    return None, None


def _apply_equipment_schedule_outside_regular_fill(
    ws, reg_start: time, reg_end: time
) -> None:
    """????????? [reg_start, reg_end) ??????????????"""
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_OUTSIDE_REGULAR_TIME_FILL,
        end_color=RESULT_OUTSIDE_REGULAR_TIME_FILL,
    )
    col_idx = None
    for i, c in enumerate(ws[1], start=1):
        if c.value is not None and str(c.value).strip() == "???":
            col_idx = i
            break
    if col_idx is None:
        return
    mr = ws.max_row or 1
    for r in range(2, mr + 1):
        cell = ws.cell(row=r, column=col_idx)
        t0, t1 = _parse_equipment_schedule_time_band_cell(cell.value)
        if t0 is None or t1 is None:
            continue
        if not _time_intervals_overlap_half_open(t0, t1, reg_start, reg_end):
            cell.fill = fill


def _apply_equipment_schedule_prep_cleanup_fill(ws) -> None:
    """
    ??????????????????????????????????????????????????????
    ??_??????? / TEMP_??????? ? equip ????????????????
    """
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_DISPATCHED_REQUEST_FILL,
        end_color=RESULT_DISPATCHED_REQUEST_FILL,
    )
    markers = ("(??????)", "(?????)", "(???????)")
    col_tb = None
    equip_cols: list[int] = []
    for i, c in enumerate(ws[1], start=1):
        if c.value is None:
            continue
        h = str(c.value).strip()
        if h == "???":
            col_tb = i
            continue
        if h.endswith("??"):
            continue
        equip_cols.append(i)
    if col_tb is None or not equip_cols:
        return
    mr = ws.max_row or 1
    for r in range(2, mr + 1):
        for ci in equip_cols:
            cell = ws.cell(row=r, column=ci)
            val = cell.value
            if val is None or (isinstance(val, float) and pd.isna(val)):
                continue
            s = str(val).strip().replace("\r", "").replace("\n", "")
            if any(m in s for m in markers):
                cell.fill = fill


def _parse_equipment_schedule_day_header_date(val) -> date | None:
    """???????? YYYY/MM/DD � ??????????"""
    if val is None:
        return None
    s = str(val).strip()
    m = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})", s)
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def _machine_calendar_intervals_for_equipment_line(
    day_blocks: dict[str, list[tuple[datetime, datetime]]],
    eq_line: str,
    day_d: date,
) -> list[tuple[datetime, datetime]]:
    """?????????????????????????????????????????"""
    if not day_blocks:
        return []
    ek = str(eq_line or "").strip()
    blocks: list[tuple[datetime, datetime]] | None = None
    if ek in day_blocks:
        blocks = day_blocks[ek]
    else:
        pk = (
            _normalize_equipment_match_key(ek.split("+", 1)[1])
            if "+" in ek
            else _normalize_equipment_match_key(ek)
        )
        if pk and pk in day_blocks:
            blocks = day_blocks[pk]
        else:
            nk = _normalize_equipment_match_key(ek)
            for k, iv in day_blocks.items():
                if _normalize_equipment_match_key(str(k)) == nk:
                    blocks = iv
                    break
    if not blocks:
        return []
    w0 = datetime.combine(day_d, DEFAULT_START_TIME)
    w1 = datetime.combine(day_d, DEFAULT_END_TIME)
    return _clip_machine_busy_blocks_to_planning_window(blocks, w0, w1)


def _apply_equipment_schedule_machine_calendar_fill(
    ws,
    equipment_list: list,
    calendar_blocks_by_date: dict[date, dict[str, list[tuple[datetime, datetime]]]],
) -> None:
    """
    ??_???????: ???????????????????????????????
    10 ??????? [slot_start, slot_end) ??? [bs, be) ????????
    """
    if not calendar_blocks_by_date or not equipment_list:
        return
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_MACHINE_CALENDAR_BLOCK_FILL,
        end_color=RESULT_MACHINE_CALENDAR_BLOCK_FILL,
    )
    col_tb = None
    for i, c in enumerate(ws[1], start=1):
        if c.value is not None and str(c.value).strip() == "???":
            col_tb = i
            break
    if col_tb is None:
        return
    eq_col_indices: list[int] = [
        col_tb + 1 + 2 * idx for idx in range(len(equipment_list))
    ]
    mr = ws.max_row or 1
    current_d: date | None = None
    for r in range(2, mr + 1):
        tb_cell = ws.cell(row=r, column=col_tb)
        tv = tb_cell.value
        d_hdr = _parse_equipment_schedule_day_header_date(tv)
        if d_hdr is not None:
            current_d = d_hdr
            continue
        t0, t1 = _parse_equipment_schedule_time_band_cell(tv)
        if t0 is None or t1 is None or current_d is None:
            continue
        slot_a = datetime.combine(current_d, t0)
        slot_b = datetime.combine(current_d, t1)
        if slot_b <= slot_a:
            continue
        day_blocks = calendar_blocks_by_date.get(current_d)
        if not day_blocks:
            continue
        for col_idx, eq_line in zip(eq_col_indices, equipment_list):
            blocks_c = _machine_calendar_intervals_for_equipment_line(
                day_blocks, eq_line, current_d
            )
            if not blocks_c:
                continue
            for bs, be in blocks_c:
                if slot_a < be and bs < slot_b:
                    ws.cell(row=r, column=col_idx).fill = fill
                    break


def _apply_equipment_by_machine_dispatched_request_fill(ws) -> None:
    """
    ??_???????_?????????????NO?????????????????
    ????????????????????????????????
    """
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_DISPATCHED_REQUEST_FILL,
        end_color=RESULT_DISPATCHED_REQUEST_FILL,
    )
    col_tb = None
    for i, c in enumerate(ws[1], start=1):
        if c.value is not None and str(c.value).strip() == "???":
            col_tb = i
            break
    if col_tb is None:
        return
    mr = ws.max_row or 1
    mc = ws.max_column or col_tb
    for r in range(2, mr + 1):
        for c in range(col_tb + 1, mc + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val is None or (isinstance(val, float) and pd.isna(val)):
                continue
            s = str(val).strip().replace("\r", "").replace("\n", "")
            if not s or s == "????":
                continue
            cell.fill = fill


def _equipment_gantt_fills_by_machine_name(equipment_list) -> dict[str, PatternFill]:
    """
    ??_??????????B?E?A ?????????equipment_list ??????+ ?????????????????
    ?????????????????? PatternFill ??????
    """
    order: list[str] = []
    seen: set[str] = set()
    for eq in equipment_list or []:
        _, mn = _split_equipment_line_process_machine(eq)
        key = (mn or "").strip() or "�"
        if key not in seen:
            seen.add(key)
            order.append(key)
    palette = RESULT_EQUIP_GANTT_MACHINE_GROUP_FILL_COLORS
    if not palette:
        fb = "F5F5F5"
        return {k: PatternFill(fill_type="solid", start_color=fb, end_color=fb) for k in order}
    out: dict[str, PatternFill] = {}
    n = len(palette)
    for i, key in enumerate(order):
        hx = palette[i % n]
        out[key] = PatternFill(fill_type="solid", start_color=hx, end_color=hx)
    return out


def _write_results_equipment_gantt_sheet(
    writer,
    timeline_events,
    equipment_list,
    sorted_dates,
    attendance_data,
    data_extract_dt_str,
    base_now_dt=None,
    actual_timeline_events=None,
    regular_shift_times: tuple[time | None, time | None] | None = None,
):
    """
    ??_???????????????timeline_events??????
    ??�?????????????????????
    ???15??????????????????????????????????
    actual_timeline_events ??????????????????????????
    """
    wb = writer.book
    try:
        insert_at = wb.sheetnames.index("??_???????") + 1
    except ValueError:
        insert_at = len(wb.sheetnames)
    ws = wb.create_sheet("??_?????", insert_at)
    try:
        ws.sheet_properties.tabColor = "7F7F7F"
    except Exception:
        pass

    events_by_date = defaultdict(list)
    for e in timeline_events:
        events_by_date[e["date"]].append(e)

    by_dm = defaultdict(lambda: defaultdict(list))
    for e in timeline_events:
        by_dm[e["date"]][e["machine"]].append(e)
    for d0 in by_dm:
        for mk in by_dm[d0]:
            by_dm[d0][mk].sort(key=lambda x: x["start_dt"])

    by_dm_actual = defaultdict(lambda: defaultdict(list))
    show_actual_rows = bool(actual_timeline_events)
    actual_events_by_date = defaultdict(list)
    if show_actual_rows:
        for e in actual_timeline_events:
            actual_events_by_date[e["date"]].append(e)
            by_dm_actual[e["date"]][e["machine"]].append(e)
        for d0 in by_dm_actual:
            for mk in by_dm_actual[d0]:
                by_dm_actual[d0][mk].sort(key=lambda x: x["start_dt"])

    slot_mins = 15
    hdr_font = _result_font(bold=True, color="000000", size=12)
    hdr_fill = PatternFill(fill_type="solid", start_color="D9D9D9", end_color="D9D9D9")
    hdr_time_font = _result_font(bold=True, color="000000", size=11)
    title_font = _result_font(bold=True, size=24, color="1A1A1A")
    title_fill = PatternFill(fill_type="solid", start_color="DDDDDD", end_color="DDDDDD")
    meta_font = _result_font(size=11, color="333333")
    meta_fill = PatternFill(fill_type="solid", start_color="F3F3F3", end_color="F3F3F3")
    day_banner_font = _result_font(bold=True, size=13, color="1A1A1A")
    day_banner_fill = PatternFill(fill_type="solid", start_color="D0D0D0", end_color="D0D0D0")
    accent_left = Side(style="thick", color="2B2B2B")
    banner_sep = Side(style="thin", color="7A7A7A")
    thin = Side(style="thin", color="666666")
    grid_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    idle_fill = PatternFill(fill_type="solid", start_color="FFFFFF", end_color="FFFFFF")
    break_fill = PatternFill(fill_type="solid", start_color="B8B8B8", end_color="B8B8B8")
    gantt_label_font = _result_font(size=10, bold=True, color="000000")
    gantt_label_font_actual = _result_font(size=10, bold=True, color="000000", italic=True)
    hdr_fill_outside_regular = PatternFill(
        fill_type="solid",
        start_color=RESULT_OUTSIDE_REGULAR_TIME_FILL,
        end_color=RESULT_OUTSIDE_REGULAR_TIME_FILL,
    )
    rs, re_ = (regular_shift_times or (None, None))

    # ??(10???)??????????slot_times ?????
    base_dt = base_now_dt if isinstance(base_now_dt, datetime) else datetime.now()
    dummy_d = sorted_dates[0] if sorted_dates else base_dt.date()
    d_start0 = datetime.combine(dummy_d, DEFAULT_START_TIME)
    d_end0 = datetime.combine(dummy_d, DEFAULT_END_TIME)
    slot_times = []
    t0 = d_start0
    while t0 < d_end0:
        slot_times.append(t0.time())
        t0 += timedelta(minutes=slot_mins)

    n_slots = len(slot_times)
    n_fixed = 5  # A=??????????????/ B?E=?????????????????
    last_col = n_fixed + n_slots
    fills_by_mach = _equipment_gantt_fills_by_machine_name(equipment_list)
    fb_gantt = "F5F5F5"
    fill_gantt_fallback = PatternFill(fill_type="solid", start_color=fb_gantt, end_color=fb_gantt)

    # ??????????????
    create_ts = base_dt.strftime("%Y/%m/%d %H:%M:%S")
    master_path = os.path.join(os.getcwd(), MASTER_FILE) if MASTER_FILE else ""

    def _fmt_mtime(p):
        try:
            if p and os.path.exists(p):
                return datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y/%m/%d %H:%M:%S")
        except Exception:
            pass
        return "�"

    master_mtime = _fmt_mtime(master_path)

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    tcell = ws.cell(row=row, column=1, value="???? ????")
    tcell.font = title_font
    tcell.fill = title_fill
    # ???????????????????????
    tcell.alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=False,
        shrink_to_fit=False,
        indent=1,
    )
    tcell.border = Border(left=accent_left, bottom=banner_sep)
    ws.row_dimensions[row].height = 40
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    meta_line = (
        f"???{create_ts}"
        f"??????????{data_extract_dt_str or '�'}"
        f"???????master.xlsm??{master_mtime}"
    )
    mtop = ws.cell(row=row, column=1, value=meta_line)
    mtop.font = meta_font
    mtop.fill = meta_fill
    mtop.alignment = Alignment(
        horizontal="left",
        vertical="center",
        indent=1,
        wrap_text=False,
        shrink_to_fit=False,
    )
    mtop.border = Border(left=accent_left, bottom=banner_sep)
    ws.row_dimensions[row].height = 26
    row += 1

    dates_to_show: list = []
    for d0 in sorted_dates:
        evs0 = events_by_date.get(d0, [])
        a_evs0 = actual_events_by_date.get(d0, []) if show_actual_rows else []
        is_anyone_working0 = any(
            attendance_data[d0][mm]["is_working"] for mm in attendance_data[d0] if mm in attendance_data[d0]
        )
        if not evs0 and not a_evs0 and not is_anyone_working0:
            continue
        dates_to_show.append(d0)

    hdr_row = row
    fixed_hdr = ["??", "???", "???", "???", "?????"]
    for ci, h in enumerate(fixed_hdr, 1):
        c = ws.cell(row=hdr_row, column=ci, value=h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
    slots_hdr = [datetime.combine(dummy_d, tm) for tm in slot_times]
    for si, st in enumerate(slots_hdr):
        c = ws.cell(row=hdr_row, column=n_fixed + 1 + si, value=st.strftime("%H:%M"))
        c.font = hdr_time_font
        slot_end_t = (st + timedelta(minutes=slot_mins)).time()
        hdr_use = hdr_fill
        if rs is not None and re_ is not None:
            if not _time_intervals_overlap_half_open(st.time(), slot_end_t, rs, re_):
                hdr_use = hdr_fill_outside_regular
        c.fill = hdr_use
        c.alignment = Alignment(horizontal="center", vertical="bottom", textRotation=90)
    ws.row_dimensions[hdr_row].height = 44
    # ????????????????F4????????1?3??A?E????
    ws.freeze_panes = f"{get_column_letter(n_fixed + 1)}{hdr_row + 1}"
    row = hdr_row + 1

    sep_fill = PatternFill(fill_type="solid", start_color="000000", end_color="000000")
    no_border = Border()

    for di, d in enumerate(dates_to_show):
        evs = events_by_date.get(d, [])
        a_evs_day = actual_events_by_date.get(d, []) if show_actual_rows else []

        slots = [datetime.combine(d, tm) for tm in slot_times]

        day_start = row
        for eq in equipment_list:
            proc_nm, mach_nm = _split_equipment_line_process_machine(eq)
            mk_key = (mach_nm or "").strip() or "�"
            lab_fill = fills_by_mach.get(mk_key) or fill_gantt_fallback
            evlist = by_dm[d].get(eq, [])
            if evlist:
                tids: list[str] = []
                seen_tid: set[str] = set()
                for e in evlist:
                    tid = str(e.get("task_id") or "").strip()
                    if tid and tid not in seen_tid:
                        seen_tid.add(tid)
                        tids.append(tid)
                task_sum = " ".join(tids) if tids else "�"
                member_disp = _gantt_row_member_names(evlist)
            else:
                task_sum = "�"
                member_disp = "�"

            c1 = ws.cell(row=row, column=2, value=mach_nm if mach_nm else "�")
            c2 = ws.cell(row=row, column=3, value=proc_nm if proc_nm else "�")
            c3 = ws.cell(row=row, column=4, value=member_disp)
            c4 = ws.cell(row=row, column=5, value=task_sum)
            for c in (c1, c2, c3, c4):
                c.font = _result_font(size=12, color="000000")
                c.fill = lab_fill
                c.border = grid_border
            c1.font = _result_font(size=12, bold=True, color="000000")
            c1.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
            c2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
            c3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c4.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            _paint_gantt_timeline_row_merged(
                ws,
                row,
                n_fixed,
                slots,
                slot_mins,
                evlist,
                idle_fill,
                break_fill,
                gantt_label_font,
                grid_border,
            )

            ws.row_dimensions[row].height = 52
            row += 1

            if show_actual_rows:
                evlist_a = by_dm_actual[d].get(eq, [])
                if evlist_a:
                    tids_a: list[str] = []
                    seen_aid: set[str] = set()
                    for e_a in evlist_a:
                        tid = str(e_a.get("task_id") or "").strip()
                        if tid and tid not in seen_aid:
                            seen_aid.add(tid)
                            tids_a.append(tid)
                    task_sum_a = " ".join(tids_a) if tids_a else "�"
                    member_disp_a = _gantt_row_member_names(evlist_a)
                else:
                    task_sum_a = "�"
                    member_disp_a = "�"

                lab_fill_a = fills_by_mach.get(mk_key) or fill_gantt_fallback

                if mach_nm:
                    act_mach = f"{mach_nm}????"
                elif proc_nm:
                    act_mach = "????"
                else:
                    act_mach = "�"
                ca1 = ws.cell(row=row, column=2, value=act_mach)
                ca2 = ws.cell(row=row, column=3, value=proc_nm if proc_nm else "�")
                ca3 = ws.cell(row=row, column=4, value=member_disp_a)
                ca4 = ws.cell(row=row, column=5, value=task_sum_a)
                for c in (ca1, ca2, ca3, ca4):
                    c.font = _result_font(size=12, color="000000")
                    c.fill = lab_fill_a
                    c.border = grid_border
                ca1.font = _result_font(size=12, bold=True, color="000000", italic=True)
                ca1.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
                ca2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
                ca3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                ca4.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                _paint_gantt_timeline_row_merged(
                    ws,
                    row,
                    n_fixed,
                    slots,
                    slot_mins,
                    evlist_a,
                    idle_fill,
                    break_fill,
                    gantt_label_font_actual,
                    grid_border,
                    task_fill_fn=_gantt_bar_fill_actual_for_task_id,
                    label_font=gantt_label_font_actual,
                )

                ws.row_dimensions[row].height = 52
                row += 1

        day_end = row - 1
        if day_end >= day_start:
            ws.merge_cells(start_row=day_start, start_column=1, end_row=day_end, end_column=1)
            ban = ws.cell(
                row=day_start,
                column=1,
                value=f"?{d.strftime('%Y/%m/%d')}?",
            )
            ban.font = day_banner_font
            ban.fill = day_banner_fill
            ban.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=False,
                textRotation=90,
            )
            ban.border = Border(left=accent_left, top=thin, bottom=thin, right=thin)

        if di < len(dates_to_show) - 1 and day_end >= day_start:
            for cc in range(1, last_col + 1):
                sc = ws.cell(row=row, column=cc)
                sc.value = None
                sc.fill = sep_fill
                sc.border = no_border
            ws.row_dimensions[row].height = 5
            row += 1

    # ????????????????????????/????????
    # ???????? VBA ????????_?????_?????????

    try:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToHeight = False
        ws.page_setup.fitToWidth = 1
        # A3?openpyxl ?? paperSize=8 ? A3 ???
        ws.page_setup.paperSize = 8
        # ???????1???????????????: ????
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.2
        ws.page_margins.bottom = 0.2
        # ?????????????????????
        ws.print_options.horizontalCentered = False
        ws.print_options.verticalCentered = False
        ws.print_options.gridLines = False
    except Exception:
        pass


def row_has_completion_keyword(row):
    """????????????????????????????????"""
    v = row.get(TASK_COL_COMPLETION_FLAG)
    if v is None or pd.isna(v):
        return False
    return "??" in str(v)


def _plan_row_exclude_from_assignment(row) -> bool:
    """
    ???????????????????????????????_??? AI ?????????

    ?????????: ??? True??? 1?????NFKC ??????
      true / 1 / yes / on / y / t / ?? / ? / ? / ?
    ???????: ??None?False?0?no / off / false / ??? / ? ?
    ???????????????????????????????? TRUE/FALSE ??? 1/0?
    """
    v = row.get(PLAN_COL_EXCLUDE_FROM_ASSIGNMENT)
    if v is True:
        return True
    if v is False:
        return False
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            iv = int(v)
            if iv == 1:
                return True
            if iv == 0:
                return False
        except (TypeError, ValueError):
            pass
    s = unicodedata.normalize("NFKC", str(v).strip()).lower()
    if not s or s in ("nan", "none", "false", "0", "no", "off", "???", "?"):
        return False
    if s in ("true", "1", "yes", "on", "??", "y", "t", "?", "?", "?"):
        return True
    return False


def _coerce_plan_exclude_column_value_for_storage(v):
    """
    ???????????????StringDtype ????????????????????
    Excel ????? True / 1 / False / 0 ?????????_plan_row_exclude_from_assignment ??????
    """
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    if v is True:
        return "yes"
    if v is False:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            iv = int(v)
            if iv == 1:
                return "yes"
            if iv == 0:
                return ""
        except (TypeError, ValueError):
            pass
    return str(v).strip()


def parse_float_safe(val, default=0.0):
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return default
        return float(val)
    except (TypeError, ValueError):
        return default


def calc_done_qty_equivalent_from_row(row):
    """
    ??????????????????

    ???:
      ???? � (??? � ????)
    = ???? * ???? / ???

    ??????/???????????????????????????????
    """
    qty_total = parse_float_safe(row.get(TASK_COL_QTY), 0.0)
    order_qty = parse_float_safe(row.get(TASK_COL_ORDER_QTY), 0.0)
    actual_output = parse_float_safe(row.get(TASK_COL_ACTUAL_OUTPUT), 0.0)
    legacy_done = parse_float_safe(row.get(TASK_COL_ACTUAL_DONE), 0.0)

    if qty_total <= 0:
        return max(0.0, legacy_done)

    if order_qty > 0 and actual_output >= 0:
        done_qty = actual_output * qty_total / order_qty
        return max(0.0, done_qty)

    return max(0.0, legacy_done)


def parse_optional_int(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None
    try:
        return int(round(float(s)))
    except (TypeError, ValueError):
        return None


def parse_optional_date(val):
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return None
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None


def _planning_df_cell_scalar(row, col_name):
    """
    iterrows() 1??????????????????????? row.get ? Series ????
    str?to_datetime ???????????????????????????????
    """
    v = row.get(col_name) if hasattr(row, "get") else None
    if isinstance(v, pd.Series):
        for x in v:
            if x is None or (isinstance(x, float) and pd.isna(x)):
                continue
            return x
        return None
    return v


def load_ai_cache():
    try:
        if os.path.exists(ai_cache_path):
            with open(ai_cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    # ????????????6???
                    now_ts = time_module.time()
                    cleaned = {}
                    expired_count = 0
                    for k, v in data.items():
                        # ???: {"ts": epoch_seconds, "data": {...}}
                        if isinstance(v, dict) and "ts" in v and "data" in v:
                            ts = parse_float_safe(v.get("ts"), 0.0)
                            if ts > 0 and (now_ts - ts) <= AI_CACHE_TTL_SECONDS:
                                cleaned[k] = v
                            else:
                                expired_count += 1
                        # ???: ????AI??dict???????????????????????
                        else:
                            cleaned[k] = {"ts": now_ts, "data": v}
                    if expired_count > 0:
                        logging.info(f"AI????????????: {expired_count}?")
                    return cleaned
    except Exception as e:
        logging.warning(f"AI???????????: {e}")
    return {}

def save_ai_cache(cache_obj):
    try:
        with open(ai_cache_path, "w", encoding="utf-8") as f:
            json.dump(cache_obj, f, ensure_ascii=False)
    except Exception as e:
        logging.warning(f"AI?????????: {e}")

def get_cached_ai_result(cache_obj, cache_key, content_key=None):
    """
    content_key: ??????????????????????????????????????????????
    ?????? content_key ?????? SHA256 ????????????????????
    """
    entry = cache_obj.get(cache_key)
    if not isinstance(entry, dict):
        return None
    ts = parse_float_safe(entry.get("ts"), 0.0)
    if ts <= 0:
        return None
    if (time_module.time() - ts) > AI_CACHE_TTL_SECONDS:
        return None
    if content_key is not None:
        stored_ck = entry.get("content_key")
        if stored_ck is not None and stored_ck != content_key:
            logging.info(
                "AI?????: ????????? content_key ??????????????????"
            )
            return None
    data = entry.get("data")
    if isinstance(data, dict):
        return data
    return None

def put_cached_ai_result(cache_obj, cache_key, parsed_obj, content_key=None):
    payload = {"ts": time_module.time(), "data": parsed_obj}
    if content_key is not None:
        payload["content_key"] = content_key
    cache_obj[cache_key] = payload

def extract_retry_seconds(err_text):
    # ?: "Please retry in 57.089735313s."
    m = re.search(r"retry in ([0-9]+(?:\.[0-9]+)?)s", err_text, re.IGNORECASE)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    # ?: "'retryDelay': '57s'"
    m = re.search(r"retryDelay'\s*:\s*'([0-9]+)s'", err_text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return None


def infer_unit_m_from_product_name(product_name, fallback_unit):
    """
    ????????????(m)???????????
    ?: 15020-JX5R- 770X300F-A   R -> 300
    ? ????????????????????????????????????
    """
    if product_name is None or pd.isna(product_name):
        return fallback_unit
    s = str(product_name)
    # "770X300..." ?????????? X ?????????????????X????
    matches = re.findall(r"[xX]\s*(\d{2,6})", s)
    if matches:
        try:
            v = int(matches[-1])
            if v > 0:
                return v
        except ValueError:
            pass
    return fallback_unit

def load_tasks_df():
    """
    ???????????tasks.xlsx ????????
    ??: ???? TASK_INPUT_WORKBOOK ???????????????VBA ????
         ????????DATA????????????????????????????????
    """
    if not TASKS_INPUT_WORKBOOK:
        raise FileNotFoundError(
            "TASK_INPUT_WORKBOOK ???????VBA ? RunPython ?????????????????????"
        )
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        raise FileNotFoundError(f"TASK_INPUT_WORKBOOK ???????: {TASKS_INPUT_WORKBOOK}")
    df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=TASKS_SHEET_NAME)
    df.columns = df.columns.str.strip()
    logging.info(f"?????: '{TASKS_INPUT_WORKBOOK}' ? '{TASKS_SHEET_NAME}' ?????????")
    return df


def _nfkc_column_aliases(canonical_name):
    """????????????????????????????????"""
    return unicodedata.normalize("NFKC", str(canonical_name).strip())


def _align_dataframe_headers_to_canonical(df, canonical_names):
    """??? NFKC ??? canonical ?????Excel ???? '_' ???????????"""
    key_to_canonical = {_nfkc_column_aliases(c): c for c in canonical_names}
    rename_map = {}
    for col in df.columns:
        k = _nfkc_column_aliases(col)
        if k in key_to_canonical:
            target = key_to_canonical[k]
            if col != target:
                rename_map[col] = target
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def _normalize_equipment_match_key(val):
    """
    ???????????????
    NFKC???????????NBSP/???????????????????
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    t = unicodedata.normalize("NFKC", str(val))
    t = t.replace("\u00a0", " ").replace("\u3000", " ")
    t = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _equipment_line_key_to_physical_occupancy_key(eq_line: str) -> str:
    """????????+?? ????????????????????????????????"""
    s = str(eq_line or "").strip()
    if not s:
        return ""
    nk = _normalize_equipment_match_key(s)
    if "+" in nk:
        return _normalize_equipment_match_key(nk.split("+", 1)[1])
    return nk


def _physical_machine_occupancy_key_for_task(task: dict) -> str:
    """
    ?????????machine_avail_dt??????????????
    ????????? equipment_line_key ????+???????????
    ??????+??????? **machine_name ????** ???????????????
    ?machine_name ???????????????????? ????????????????????
    ???????????? machine_name ????????? equipment_line_key / machine ???????
    machine_name ????+???????????????????????+ ?????????
    ?????????? NFKC ?????+????????????????????
    """
    ek = str(task.get("equipment_line_key") or "").strip()
    nek = _normalize_equipment_match_key(ek)
    if nek and "+" in nek:
        pk = _equipment_line_key_to_physical_occupancy_key(ek)
        if pk:
            return pk
    mn = str(task.get("machine_name") or "").strip()
    if mn:
        nk = _normalize_equipment_match_key(mn)
        if "+" in nk:
            return _normalize_equipment_match_key(nk.split("+", 1)[1])
        return nk
    return _equipment_line_key_to_physical_occupancy_key(
        str(task.get("equipment_line_key") or task.get("machine") or "")
    )


def _machine_occupancy_key_resolve(task: dict, eq_line: str) -> str:
    """
    machine_avail_dt?????????????????????: ???????
    task ????????? eq_line???+???????????????????? eq_line?
    ?� or eq_line??????+???????????????????????????????????
    """
    occ = (_physical_machine_occupancy_key_for_task(task) or "").strip()
    if occ:
        return occ
    ek = str(eq_line or "").strip()
    if not ek:
        return ""
    pk = (_equipment_line_key_to_physical_occupancy_key(ek) or "").strip()
    return pk or ek


def _equipment_lookup_normalized_to_canonical(equipment_list):
    """????? ? master ???????????canonical ????"""
    lookup = {}
    for eq in equipment_list:
        k = _normalize_equipment_match_key(eq)
        if k and k not in lookup:
            lookup[k] = eq
    # ?????????????DATA??: ???????????+???????
    for eq in equipment_list:
        s = str(eq).strip()
        if "+" not in s:
            continue
        p, _rest = s.split("+", 1)
        pk = _normalize_equipment_match_key(p)
        if pk and pk not in lookup:
            lookup[pk] = eq
    return lookup


def _equipment_schedule_header_labels(equipment_list: list) -> list:
    """
    ??_??????????_??????????????
    ????????+?????????????????????????????????
    """
    raw = []
    for eq in equipment_list:
        s = str(eq).strip()
        if "+" in s:
            mpart = s.split("+", 1)[1].strip()
            raw.append(mpart if mpart else s)
        else:
            raw.append(s)
    counts = {}
    for r in raw:
        counts[r] = counts.get(r, 0) + 1
    out = []
    for eq, r in zip(equipment_list, raw):
        if counts.get(r, 0) > 1:
            s = str(eq).strip()
            if "+" in s:
                p = s.split("+", 1)[0].strip()
                out.append(f"{r}?{p}?" if p else r)
            else:
                out.append(r)
        else:
            out.append(r)
    return out


def _split_equipment_line_process_machine(eq_line: str) -> tuple[str, str]:
    """
    ????????????+???? (???, ???) ??????
    '+' ????????????????????????
    """
    s = str(eq_line).strip()
    if not s:
        return ("", "")
    if "+" in s:
        p, m = s.split("+", 1)
        return (p.strip(), m.strip())
    return ("", s)


def _gantt_member_label_surname_only(raw: str) -> str:
    """
    ???????????????????????????????????????????
    ????1???????????????????????????NFKC???/???????????
    """
    sei, mei = _split_person_sei_mei(raw)
    if not sei:
        return ""
    n = _normalize_sei_for_match(sei)
    return n if n else sei


def _gantt_row_member_names(evlist) -> str:
    """???????: ???(op)???(sub)??????????????????+??????????"""
    raw_names: list[str] = []
    seen_raw: set[str] = set()
    for e in evlist or []:
        op = str(e.get("op") or "").strip()
        if op and op not in seen_raw:
            seen_raw.add(op)
            raw_names.append(op)
        sub_raw = str(e.get("sub") or "").strip()
        if not sub_raw:
            continue
        for seg in re.split(r"[,?]", sub_raw):
            t = seg.strip()
            if t and t not in seen_raw:
                seen_raw.add(t)
                raw_names.append(t)
    labels: list[str] = []
    seen_label: set[str] = set()
    for raw in raw_names:
        lab = _gantt_member_label_surname_only(raw)
        if lab and lab not in seen_label:
            seen_label.add(lab)
            labels.append(lab)
    return ", ".join(labels) if labels else "�"


def _resolve_equipment_line_key_for_task(task: dict, equipment_list: list | None) -> str:
    """
    ?????????????????skills / need ??????+???????????
    ?????????????????1????????????????
    """
    p = str(task.get("machine") or "").strip()
    mn = str(task.get("machine_name") or "").strip()
    cand = f"{p}+{mn}" if (p and mn) else (p or mn)
    elist = [str(x).strip() for x in (equipment_list or []) if str(x).strip()]
    if cand in elist:
        return cand
    if mn:
        return cand
    if not p:
        return cand
    exact_p = [x for x in elist if x == p]
    if len(exact_p) == 1:
        return exact_p[0]
    prefixed = [x for x in elist if x.startswith(p + "+")]
    if len(prefixed) == 1:
        return prefixed[0]
    return p



def _apply_planning_sheet_post_load_mutations(
    df: "pd.DataFrame", wb_path: str, log_prefix: str
) -> None:
    """
    ????_?????? DataFrame ???????????
    ??2? ``load_planning_tasks_df`` ???????????????????????
    """
    try:
        _pairs_lr = []
        _seen_lr = set()
        for _, _row_lr in df.iterrows():
            _p = str(_row_lr.get(TASK_COL_MACHINE, "") or "").strip()
            _m = str(_row_lr.get(TASK_COL_MACHINE_NAME, "") or "").strip()
            if not _p:
                continue
            _k = (
                _normalize_process_name_for_rule_match(_p),
                _normalize_equipment_match_key(_m),
            )
            if _k in _seen_lr:
                continue
            _seen_lr.add(_k)
            _pairs_lr.append((_p, _m))
        run_exclude_rules_sheet_maintenance(wb_path, _pairs_lr, log_prefix)
    except Exception:
        logging.exception("%s: ??_????????????????", log_prefix)
    try:
        _apply_auto_exclude_bunkatsu_duplicate_machine(df, log_prefix=log_prefix)
    except Exception as ex:
        logging.warning(
            "%s: ???????????????????: %s",
            log_prefix,
            ex,
        )
    try:
        apply_exclude_rules_config_to_plan_df(df, wb_path, log_prefix)
    except Exception as ex:
        logging.warning(
            "%s: ?????????????????????: %s",
            log_prefix,
            ex,
        )

def load_planning_tasks_df():
    """
    2????: ?????????????_???????????????

    ???OP_?????????????? AI ?? preferred_operator ???? OP ???????skills ???????????????
    ????????????? task_preferred_operators ? generate_plan ????????????
    ??????????TRUE/1/?? ???????????
    ??????????NO???????????????????????????????=yes???1?????
    ???_??????????+?????????C/D/E ???????????????????? VBA??
    """
    if not TASKS_INPUT_WORKBOOK:
        raise FileNotFoundError(
            "TASK_INPUT_WORKBOOK ???????VBA ? RunPython ?????????????????????"
        )
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        raise FileNotFoundError(f"TASK_INPUT_WORKBOOK ???????: {TASKS_INPUT_WORKBOOK}")
    df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=PLAN_INPUT_SHEET_NAME)
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(
        df, plan_input_sheet_column_order()
    )
    for c in plan_input_sheet_column_order():
        if c not in df.columns:
            df[c] = ""
    _apply_planning_sheet_post_load_mutations(df, TASKS_INPUT_WORKBOOK, "???????")
    logging.info(
        f"???????: '{TASKS_INPUT_WORKBOOK}' ? '{PLAN_INPUT_SHEET_NAME}' ?????????"
    )
    return df


def _main_sheet_cell_is_global_comment_label(val) -> bool:
    """?????????????????????????????????"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    s = unicodedata.normalize("NFKC", str(val).strip())
    if not s:
        return False
    if _nfkc_column_aliases(s) == _nfkc_column_aliases("?????????"):
        return True
    if "?????" in s and "????" in s:
        return True
    return False


def load_main_sheet_global_priority_override_text() -> str:
    """
    TASK_INPUT_WORKBOOK ??????????????????????????? **??** ????
    ????: ?????????_??Main??????????????????????VBA GetMainWorksheet ??????

    ??? **Gemini ?????**?`analyze_global_priority_override_comment`???????????????????????? JSON ????
    API ??????????????????????? `parse_factory_closure_dates_from_global_comment` ??????
    """
    wb_path = TASKS_INPUT_WORKBOOK.strip() if TASKS_INPUT_WORKBOOK else ""
    if not wb_path or not os.path.exists(wb_path):
        return ""
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "????????: ?????%s?????? openpyxl ?????????????????",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return ""
    try:
        wb = load_workbook(wb_path, data_only=True, read_only=False)
    except Exception as e:
        logging.warning("????????: ????????????: %s", e)
        return ""
    try:
        ws = None
        for name in ("???", "???_", "Main"):
            if name in wb.sheetnames:
                ws = wb[name]
                break
        if ws is None:
            for sn in wb.sheetnames:
                if "???" in sn:
                    ws = wb[sn]
                    break
        if ws is None:
            return ""
        max_r = min(ws.max_row or 0, 400)
        max_c = min(ws.max_column or 0, 40)
        if max_r < 1 or max_c < 1:
            return ""
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cell = ws.cell(row=r, column=c)
                if not _main_sheet_cell_is_global_comment_label(cell.value):
                    continue
                below = ws.cell(row=r + 1, column=c).value
                if below is None or (isinstance(below, float) and pd.isna(below)):
                    return ""
                return str(below).strip()
        return ""
    finally:
        pass


def _global_comment_chunk_implies_factory_closure(chunk: str) -> bool:
    """
    ????????????????????????????????????????????????????
    """
    c = unicodedata.normalize("NFKC", str(chunk or ""))
    if not c.strip():
        return False
    if re.search(r"??\s*??", c):
        return True
    if "??" in c:
        return True
    if re.search(r"??", c) and re.search(r"?|??|??|??|???", c):
        return True
    if re.search(r"(?:??|??|???).{0,15}(?:?|??|??)", c):
        return True
    if re.search(r"(?:??|??|???).{0,12}(?:??|??|??)", c):
        return True
    if re.search(r"??.{0,15}(?:???|??|??|???)", c):
        return True
    if "??" in c and re.search(
        r"(?:??|??|??|??|??|??|??|???)", c
    ):
        return True
    return False


def _md_slash_is_likely_fraction_not_date(t: str, start: int, end: int, mo: int, day: int) -> bool:
    """
    ??????1/3?????? 1/3 ? 1?3? ???????
    ?4/1???????? 4/1 ?????????????????????????
    """
    if mo <= 0 or day <= 0:
        return True
    before = t[max(0, start - 32) : start]
    after = t[end : min(len(t), end + 14)]
    after_st = after.lstrip()
    if after_st.startswith("?"):
        return False
    if re.search(
        r"(?:????|??\s*????|??|??|????|??|??)(?:\s*?)?\s*$",
        before,
    ):
        return True
    # 1/2?1/3?2/3 ? + ?????????� ??????????3/1?????????????????? ??/??? ??????
    frac_pat = re.compile(
        r"^(?:?????|???|?|??|???|???|???|??|???|???)"
    )
    if mo <= 12 and day <= 12 and frac_pat.match(after_st):
        if mo <= 2 or (mo == 3 and day <= 3):
            return True
    # ?1/2????1/10??????????????? 1/ ???
    if (
        mo == 1
        and 2 <= day <= 12
        and re.match(r"^??|???\b", after_st)
    ):
        return True
    return False


def _extract_calendar_dates_from_text(s: str, default_year: int) -> list[date]:
    """???????????????? date ????????????????"""
    t = unicodedata.normalize("NFKC", str(s or ""))
    found: list[date] = []
    seen: set[date] = set()

    def add(y: int, mo: int, d: int) -> None:
        try:
            dd = date(y, mo, d)
        except ValueError:
            return
        if dd not in seen:
            seen.add(dd)
            found.append(dd)

    for m in re.finditer(
        r"(\d{4})\s*?\s*(\d{1,2})\s*?\s*(\d{1,2})\s*??",
        t,
    ):
        add(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    for m in re.finditer(
        r"(\d{4})\s*[/\-\.?]\s*(\d{1,2})\s*[/\-\.?]\s*(\d{1,2})",
        t,
    ):
        add(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    for m in re.finditer(r"(\d{1,2})\s*?\s*(\d{1,2})\s*?", t):
        add(int(default_year), int(m.group(1)), int(m.group(2)))
    for m in re.finditer(
        r"(?<!\d)(\d{1,2})\s*[/?]\s*(\d{1,2})(?!\d)",
        t,
    ):
        mo_i, d_i = int(m.group(1)), int(m.group(2))
        if _md_slash_is_likely_fraction_not_date(t, m.start(), m.end(), mo_i, d_i):
            continue
        add(int(default_year), mo_i, d_i)
    return found


def _split_global_comment_into_chunks(blob: str) -> list[str]:
    """
    ?????????????????????????
    ???Excel ? Alt+Enter?Unicode ????????????????? ?;? ???????
    """
    t = unicodedata.normalize("NFKC", str(blob or "").strip())
    if not t:
        return []
    lines = [ln.strip() for ln in re.split(r"[\n\r\v\f\u2028\u2029]+", t) if ln.strip()]
    if not lines:
        return []
    chunks: list[str] = []
    for line in lines:
        subs = [c.strip() for c in re.split(r"[?;?]+", line) if c.strip()]
        if subs:
            chunks.extend(subs)
        else:
            chunks.append(line)
    return chunks


def parse_factory_closure_dates_from_global_comment(
    text: str, default_year: int
) -> set[date]:
    """
    ????????????????????????????????????????
    ?????????????????????????????????????
    """
    blob = unicodedata.normalize("NFKC", str(text or "").strip())
    if not blob:
        return set()
    chunks = _split_global_comment_into_chunks(blob)
    if not chunks:
        chunks = [blob]
    out: set[date] = set()
    y0 = int(default_year)
    for ch in chunks:
        if not _global_comment_chunk_implies_factory_closure(ch):
            continue
        for d in _extract_calendar_dates_from_text(ch, y0):
            out.add(d)
    if not out and _global_comment_chunk_implies_factory_closure(blob):
        for d in _extract_calendar_dates_from_text(blob, y0):
            out.add(d)
    return out


def apply_factory_closure_dates_to_attendance(
    attendance_data: dict, members: list, closure_dates: set[date]
) -> None:
    """?????: ?????? is_working=False ?????????????????"""
    if not closure_dates or not attendance_data:
        return
    tag = "???????????????????"
    for d in sorted(closure_dates):
        if d not in attendance_data:
            logging.warning(
                "??????????????? %s ???????????????"
                " ???????????????????????????????",
                d,
            )
            continue
        day = attendance_data[d]
        for m in members:
            if m not in day:
                continue
            ent = day[m]
            ent["is_working"] = False
            ent["eligible_for_assignment"] = False
            prev = str(ent.get("reason") or "").strip()
            ent["reason"] = f"{tag} {prev}".strip() if prev else tag


def _apply_global_priority_abolish_heuristic(blob: str, coerced: dict) -> dict:
    """
    ???????????????: ???????????????????????abolish_all_scheduling_limits??
    """
    b = unicodedata.normalize("NFKC", str(blob or ""))
    strong = (
        "????",
        "?????",
        "??????",
        "?????",
        "??????",
        "??????",
        "??????",
        "?????",
        "????",
        "???",
    )
    if any(k in b for k in strong):
        out = dict(coerced)
        out["abolish_all_scheduling_limits"] = True
        out["ignore_skill_requirements"] = True
        out["ignore_need_minimum"] = True
        logging.warning(
            "????????: ??????????????????????????????????????"
        )
        return out
    return coerced


def _maybe_fill_global_speed_rules_from_scheduler_notes(coerced: dict) -> dict:
    """
    AI ? global_speed_rules ?????? scheduler_notes ????????????????
    ???????????????1/3 ?????
    """
    if not isinstance(coerced, dict):
        return coerced
    if coerced.get("global_speed_rules"):
        return coerced
    sn = str(coerced.get("scheduler_notes_ja") or "")
    t = unicodedata.normalize("NFKC", sn)
    if "???" not in t or "??" not in t:
        return coerced
    if not re.search(r"(?:1\s*/\s*3|?\s*/\s*3|????|3\s*??\s*1)", t):
        return coerced
    out = dict(coerced)
    out["global_speed_rules"] = [
        {
            "process_contains": "???",
            "machine_contains": "??",
            "speed_multiplier": 1.0 / 3.0,
        }
    ]
    logging.info(
        "????????: scheduler_notes_ja ?? global_speed_rules ???????????1/3?"
    )
    return out


def _finalize_global_priority_override(blob: str, coerced: dict) -> dict:
    """???????abolish ? true ??????????????"""
    coerced = _maybe_fill_global_speed_rules_from_scheduler_notes(dict(coerced))
    coerced = _apply_global_priority_solo_heuristic(blob, coerced)
    coerced = _apply_global_priority_abolish_heuristic(blob, coerced)
    if coerced.get("abolish_all_scheduling_limits"):
        out = dict(coerced)
        out["ignore_skill_requirements"] = True
        out["ignore_need_minimum"] = True
        return out
    return coerced


def _apply_global_priority_solo_heuristic(blob: str, coerced: dict) -> dict:
    """
    ????????????????????????????????????????????
    ??????????????????
    """
    if not coerced.get("ignore_need_minimum") or coerced.get("ignore_skill_requirements"):
        return coerced
    b = unicodedata.normalize("NFKC", str(blob or ""))
    solo_kw = ("??", "???", "??", "??", "1?", "??", "??")
    if any(k in b for k in solo_kw):
        out = dict(coerced)
        out["ignore_skill_requirements"] = True
        logging.info(
            "????????: ??????????? ignore_skill_requirements ????? true ??????"
        )
        return out
    return coerced


def _coerce_task_preferred_operators_dict(raw_val) -> dict:
    """AI ? task_preferred_operators ? {??NO: ??} ?????"""
    out = {}
    if not isinstance(raw_val, dict):
        return out
    for k, v in raw_val.items():
        ks = str(k).strip()
        if not ks:
            continue
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue
        vs = str(v).strip()
        if vs and vs.lower() not in ("nan", "none", "null"):
            out[ks] = vs
    return out


def _normalize_factory_closure_dates_iso_list(val, default_year: int) -> list[str]:
    """
    AI ????????????????? YYYY-MM-DD ???????????????
    ??? ISO ????Excel ????4/1??????????
    """
    y0 = int(default_year)
    seen: set[str] = set()
    out: list[str] = []
    if not isinstance(val, list):
        return out
    for item in val:
        if item is None or (isinstance(item, float) and pd.isna(item)):
            continue
        d = parse_optional_date(item)
        if d is not None:
            iso = d.isoformat()
            if iso not in seen:
                seen.add(iso)
                out.append(iso)
            continue
        s = unicodedata.normalize("NFKC", str(item).strip())
        if not s:
            continue
        for d2 in _extract_calendar_dates_from_text(s, y0):
            iso = d2.isoformat()
            if iso not in seen:
                seen.add(iso)
                out.append(iso)
    return sorted(out)


def _coerce_global_speed_rules(raw_val) -> list[dict]:
    """
    Gemini ? global_speed_rules ?????
    ???: process_contains / machine_contains??????????????, speed_multiplier?????????0??10????
    """
    out: list[dict] = []
    if not isinstance(raw_val, list):
        return out
    for item in raw_val:
        if not isinstance(item, dict):
            continue
        sm = item.get("speed_multiplier")
        if sm is None:
            sm = item.get("relative_speed")
        try:
            mult = float(sm)
        except (TypeError, ValueError):
            continue
        if mult <= 0 or mult > 10.0:
            continue
        pps = unicodedata.normalize("NFKC", str(item.get("process_contains") or "")).strip()
        mms = unicodedata.normalize("NFKC", str(item.get("machine_contains") or "")).strip()
        if not pps and not mms:
            continue
        out.append(
            {
                "process_contains": pps,
                "machine_contains": mms,
                "speed_multiplier": mult,
            }
        )
    return out


def _global_speed_rule_substring_matches_row(pnorm: str, mnorm: str, sub_nfkc: str) -> bool:
    """sub ?????????????????????????????? True?"""
    if not sub_nfkc:
        return True
    return sub_nfkc in pnorm or sub_nfkc in mnorm


def _global_speed_multiplier_for_row(process_name: str, machine_name: str, rules: list) -> float:
    """
    ???????????????? speed_multiplier ????????????? 1.0??

    process_contains / machine_contains ????? **??????????????** ?????????
    ?????? AND??: ???????????????????????????????
    ???????=?????=???? ????????????????????????????
    """
    if not rules:
        return 1.0
    pnorm = unicodedata.normalize("NFKC", str(process_name or "")).strip()
    mnorm = unicodedata.normalize("NFKC", str(machine_name or "")).strip()
    combined = 1.0
    for r in rules:
        if not isinstance(r, dict):
            continue
        pc = unicodedata.normalize(
            "NFKC", str(r.get("process_contains") or "").strip()
        )
        mc = unicodedata.normalize(
            "NFKC", str(r.get("machine_contains") or "").strip()
        )
        if not pc and not mc:
            continue
        if pc and not _global_speed_rule_substring_matches_row(pnorm, mnorm, pc):
            continue
        if mc and not _global_speed_rule_substring_matches_row(pnorm, mnorm, mc):
            continue
        try:
            m = float(r.get("speed_multiplier", 1.0))
        except (TypeError, ValueError):
            continue
        if m <= 0:
            continue
        combined *= m
    if combined <= 0:
        return 1.0
    return combined


def _infer_global_day_process_rules_from_free_text(text: str, ref_y: int) -> list[dict]:
    """
    Gemini ? task_preferred_operators ???????????????
    ???????? global_day_process_operator_rules ?????????????
    ?: ?2026/4/4 ???:EC ?????????
    """
    t = unicodedata.normalize("NFKC", str(text or "")).strip()
    if len(t) < 6:
        return []
    dates = _extract_calendar_dates_from_text(t, int(ref_y))
    if not dates:
        return []
    d0 = dates[0]
    proc_m = re.search(
        r"????\s*[:?]?\s*([A-Za-z0-9?-?????]+)",
        t,
    )
    pc = proc_m.group(1).strip() if proc_m else ""
    if not pc:
        m2 = re.search(r"([\dA-Za-z?-?????]{1,12})\s*??", t)
        pc = m2.group(1).strip() if m2 else ""
    if not pc:
        return []
    names: list[str] = []
    for m in re.finditer(
        r"([\u3040-\u9FFF???A-Za-z???]{1,16}?)\s*?\s*([\u3040-\u9FFF???A-Za-z???]{1,16}?)\s*??\s*(?:??|??|???|?????)",
        t,
    ):
        a, b = m.group(1).strip(), m.group(2).strip()
        if a:
            names.append(a)
        if b:
            names.append(b)
    if len(names) < 2:
        return []
    return [
        {
            "date": d0.isoformat(),
            "process_contains": pc,
            "operator_names": names[:12],
        }
    ]


def _salvage_malformed_global_priority_gemini_dict(raw: dict, ref_y: int) -> dict:
    """
    Gemini ? task_preferred_operators ? **??**???????workstation_id ?????????
    ???? global_day_process_operator_rules / scheduler_notes_ja ??????
    """
    out = dict(raw)
    tpo = out.get("task_preferred_operators")
    if not isinstance(tpo, list) or not tpo:
        return out
    narratives: list[str] = []
    extra_rule_objs: list[dict] = []
    for item in tpo:
        if not isinstance(item, dict):
            continue
        onames = item.get("operator_names")
        if isinstance(onames, list) and (
            item.get("date") is not None or item.get("process_contains")
        ):
            extra_rule_objs.append(item)
            continue
        for key in ("workstation_id", "schedule_notes_ai", "schedule_notes", "note", "text"):
            s = str(item.get(key) or "").strip()
            if len(s) >= 12:
                narratives.append(s[:800])
        for _k, v in item.items():
            if _k in (
                "factory_closure_dates",
                "operator_names",
                "date",
                "process_contains",
            ):
                continue
            if isinstance(v, str) and len(v) > 35 and ("?" in v or "??" in v):
                narratives.append(v[:800])
    out["task_preferred_operators"] = {}
    gdp_existing = out.get("global_day_process_operator_rules")
    gdp_list: list = list(gdp_existing) if isinstance(gdp_existing, list) else []
    gdp_list.extend(extra_rule_objs)
    seen_n: set[str] = set()
    for nb in narratives:
        if nb in seen_n:
            continue
        seen_n.add(nb)
        gdp_list.extend(_infer_global_day_process_rules_from_free_text(nb, ref_y))
    out["global_day_process_operator_rules"] = gdp_list
    if narratives:
        sn0 = str(out.get("scheduler_notes_ja") or "").strip()
        add = " | ".join(n[:280] for n in narratives[:4])
        out["scheduler_notes_ja"] = (sn0 + " " + add).strip()[:600]
    return out


def _coerce_global_priority_override_dict(raw, reference_year: int | None = None) -> dict:
    """Gemini ???????????????????????"""
    y0 = int(reference_year) if reference_year is not None else date.today().year

    def as_bool(v):
        if v is True:
            return True
        if v is False:
            return False
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        s = unicodedata.normalize("NFKC", str(v).strip()).lower()
        return s in ("true", "1", "yes", "??", "on")

    base = {
        "ignore_skill_requirements": False,
        "ignore_need_minimum": False,
        "abolish_all_scheduling_limits": False,
        "task_preferred_operators": {},
        "interpretation_ja": "",
        "factory_closure_dates": [],
        "scheduler_notes_ja": "",
        "global_speed_rules": [],
        "global_day_process_operator_rules": [],
    }
    if not isinstance(raw, dict):
        return base
    raw = _salvage_malformed_global_priority_gemini_dict(raw, y0)
    base["ignore_skill_requirements"] = as_bool(raw.get("ignore_skill_requirements"))
    base["ignore_need_minimum"] = as_bool(raw.get("ignore_need_minimum"))
    base["abolish_all_scheduling_limits"] = as_bool(
        raw.get("abolish_all_scheduling_limits")
    )
    base["task_preferred_operators"] = _coerce_task_preferred_operators_dict(
        raw.get("task_preferred_operators")
    )
    ij = raw.get("interpretation_ja")
    if ij is not None and not (isinstance(ij, float) and pd.isna(ij)):
        base["interpretation_ja"] = str(ij).strip()[:800]
    base["factory_closure_dates"] = _normalize_factory_closure_dates_iso_list(
        raw.get("factory_closure_dates"), y0
    )
    sn = raw.get("scheduler_notes_ja")
    if sn is not None and not (isinstance(sn, float) and pd.isna(sn)):
        base["scheduler_notes_ja"] = str(sn).strip()[:600]
    base["global_speed_rules"] = _coerce_global_speed_rules(raw.get("global_speed_rules"))
    base["global_day_process_operator_rules"] = _coerce_global_day_process_operator_rules(
        raw.get("global_day_process_operator_rules")
    )
    return base


def _parse_global_priority_override_gemini_response(res):
    """Gemini ???? JSON ??????1???????```json ???????????"""
    raw = (_gemini_result_text(res) or "").strip()
    if not raw:
        return None
    candidate = None
    fence = re.search(
        r"```(?:json)?\s*(\{.*\})\s*```",
        raw,
        re.DOTALL | re.IGNORECASE,
    )
    if fence:
        candidate = fence.group(1).strip()
    elif raw.startswith("{"):
        candidate = raw
    else:
        loose = re.search(r"\{.*\}", raw, re.DOTALL)
        candidate = loose.group(0).strip() if loose else None
    if not candidate:
        return None
    try:
        parsed = json.loads(candidate)
    except json.JSONDecodeError:
        return None
    return parsed if isinstance(parsed, dict) else None


def _apply_regex_factory_closure_fallback(coerced: dict, blob: str, ref_y: int) -> dict:
    """Gemini ???????????: ???????????????????????"""
    out = dict(coerced)
    rx = parse_factory_closure_dates_from_global_comment(blob, ref_y)
    out["factory_closure_dates"] = sorted({d.isoformat() for d in rx})
    return out


def analyze_global_priority_override_comment(
    text: str, members: list, reference_year: int, ai_sheet_sink: dict | None = None
) -> dict:
    """
    ??????????????????UI ???????? **Gemini ?????**??????? JSON ?????
    ????????????????????? AI ????????????????????????

    - factory_closure_dates: **????**????????????????? YYYY-MM-DD ???????????? []?
    - ignore_skill_requirements / ignore_need_minimum / abolish_all_scheduling_limits / task_preferred_operators: ??????
    - global_speed_rules: **???????**?????????????? **???????????**????????????????????? **??**?????????????? []?
    - global_day_process_operator_rules: **??????????????????**?????????????**???????????**???????????? []?
    - scheduler_notes_ja: ????????????????????????? global_speed_rules ?????

    API ?????JSON ?????: ???????????????????????????????????
    """
    ref_y = int(reference_year) if reference_year is not None else date.today().year
    empty = _coerce_global_priority_override_dict({}, ref_y)
    if not text or not str(text).strip():
        if ai_sheet_sink is not None:
            ai_sheet_sink["????????_AI_API"] = "?????????????"
        return empty
    blob = str(text).strip()
    mem_sig = ",".join(sorted(str(m).strip() for m in (members or []) if m))
    cache_fingerprint = f"{GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX}{ref_y}\n{blob}\n{mem_sig}"
    cache_key = hashlib.sha256(cache_fingerprint.encode("utf-8")).hexdigest()
    ai_cache = load_ai_cache()
    cached = get_cached_ai_result(ai_cache, cache_key, content_key=cache_fingerprint)
    if cached is not None:
        logging.info("????????: ?????????Gemini ????????")
        if ai_sheet_sink is not None:
            ai_sheet_sink["????????_AI_API"] = "???????????"
        return _finalize_global_priority_override(
            blob, _coerce_global_priority_override_dict(cached, ref_y)
        )

    if not API_KEY:
        logging.info("GEMINI_API_KEY ??????????????? AI ????????????")
        if ai_sheet_sink is not None:
            ai_sheet_sink["????????_AI_API"] = "???API??????????????????"
        coerced = _apply_regex_factory_closure_fallback(
            _coerce_global_priority_override_dict({}, ref_y), blob, ref_y
        )
        return _finalize_global_priority_override(blob, coerced)

    member_sample = ", ".join(str(m) for m in (members or [])[:80])
    if len(members or []) > 80:
        member_sample += " �"

    prompt = f"""?????????????????????????
Excel ??????? **???????????**???????????? **??** ????????????? JSON ?1?????????

????
???????????????????????????**???????**???????? **??????????** ????????????
??????? true ????????????????? true?

?????
???????????????need??????????_??? AI ??????????????????????

????????
?????? **???????????** ?????????1????????**?????? 1/3??????????????**???

????????

A) **factory_closure_dates** ???????
   - **????**?????????????????????????????????? **YYYY-MM-DD** ????????
   - **?????????????**??????? **????**?[]??
   - ??????? **??? []**?????????
   - ???????????? {ref_y} ???????

B) **ignore_skill_requirements** / **ignore_need_minimum** / **abolish_all_scheduling_limits** / **task_preferred_operators**
   - ?????????????????1??????????NO????OP?????????? false ??? {{}}?

C) **global_speed_rules** ???????
   - ??? **???**?Excel???????? **???**????????????**?????????????** ????????????????
   - ??????????:
     - "process_contains": ?????????**??????????????**? **????**?NFKC ????
     - "machine_contains": ?????????**??????????????**? **????**?
     - "speed_multiplier": ????**1/3 ???**??? **0.333333**????? � ?????**2??**?? 2.0?
   - **?????? AND**?2?????????**????**??????????????????????: ??=?????=???? ?????=??????=????? ?????????
   - ????????????????????????????????????????????
   - ????????? **??? []**?
   - ?: ???????????????1/3??
     [{{"process_contains":"???","machine_contains":"??","speed_multiplier":0.333333}}]
     ????????????????????????????????????1/3????

D) **scheduler_notes_ja** ????????
   - ?????????????????? **global_speed_rules ???????????????????**???????????????????? ""?

E) **interpretation_ja** ????????
   - ??????1??200??????

F) **global_day_process_operator_rules** ???????
   - **??????**?? **????Excel????????????** ??????????????
     ???? **?????????????????** ????**OP/AS ??????????**?????? **??OP?????**??
   - **??NO????????1???**? **task_preferred_operators** ????????? **?????????????????** ???? **?????????**???? **???**?????
   - ??????????:
     - "date": **YYYY-MM-DD**?????????????????
     - "process_contains": ???? **????**?NFKC ?????: "EC"
     - "operator_names": ???????: ["??", "?????"]?
   - ????????? **??? []**?

??????
??? {{ ????? }} ? **JSON ??????1???**???????????????

??????:
- "factory_closure_dates": string ????YYYY-MM-DD?
- "ignore_skill_requirements": true ??? false
- "ignore_need_minimum": true ??? false
- "abolish_all_scheduling_limits": true ??? false
- "task_preferred_operators": **JSON ????????**???=??NO??=???????**??????????**?????? {{}}
- "global_speed_rules": ??????????????? []?
- "global_day_process_operator_rules": ??????????????? []?
- "scheduler_notes_ja": ???
- "interpretation_ja": ???

????? ?????????? {ref_y} ???????????

?????????????????JSON ?????????
{member_sample}

??????????????
{blob}
"""
    try:
        ppath = os.path.join(log_dir, "ai_global_priority_override_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("????????: ??????? ? %s", ppath)
    except OSError as ex:
        logging.warning("????????: ?????????: %s", ex)

    client = genai.Client(api_key=API_KEY)
    try:
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        parsed = _parse_global_priority_override_gemini_response(res)
        if parsed is None:
            logging.warning(
                "????????: AI ???? JSON ??????????????????????????????"
            )
            try:
                rpath = os.path.join(log_dir, "ai_global_priority_override_last_response.txt")
                with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                    rf.write(_gemini_result_text(res) or "")
            except OSError:
                pass
            if ai_sheet_sink is not None:
                ai_sheet_sink["????????_AI_API"] = "???JSON????????????????"
            coerced = _apply_regex_factory_closure_fallback(
                _coerce_global_priority_override_dict({}, ref_y), blob, ref_y
            )
            return _finalize_global_priority_override(blob, coerced)
        coerced = _coerce_global_priority_override_dict(parsed, ref_y)
        coerced = _finalize_global_priority_override(blob, coerced)
        try:
            rpath = os.path.join(log_dir, "ai_global_priority_override_last_response.txt")
            with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                rf.write(_gemini_result_text(res) or "")
        except OSError:
            pass
        put_cached_ai_result(ai_cache, cache_key, coerced, content_key=cache_fingerprint)
        save_ai_cache(ai_cache)
        _tpo = coerced.get("task_preferred_operators") or {}
        _fcd = coerced.get("factory_closure_dates") or []
        _gsr = coerced.get("global_speed_rules") or []
        _gdp = coerced.get("global_day_process_operator_rules") or []
        logging.info(
            "????????: AI ?? factory??=%s? ?????=%s? ?�?????=%s? skill=%s need1=%s abolish=%s task_pref=%s? � %s",
            len(_fcd),
            len(_gsr),
            len(_gdp),
            coerced["ignore_skill_requirements"],
            coerced["ignore_need_minimum"],
            coerced.get("abolish_all_scheduling_limits"),
            len(_tpo),
            coerced.get("interpretation_ja", "")[:100],
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["????????_AI_API"] = "??"
        return coerced
    except Exception as e:
        logging.warning("????????: Gemini ??????: %s", e)
        if ai_sheet_sink is not None:
            ai_sheet_sink["????????_AI_API"] = f"??: {e}"[:500]
        coerced = _apply_regex_factory_closure_fallback(
            _coerce_global_priority_override_dict({}, ref_y), blob, ref_y
        )
        return _finalize_global_priority_override(blob, coerced)


def default_result_task_sheet_column_order(max_history_len: int) -> list:
    """??_?????????????????????????"""
    hist = [f"??{i+1}" for i in range(max_history_len)]
    return [
        "?????",
        "???ID",
        "???",
        "???",
        "???",
        RESULT_TASK_COL_DISPATCH_TRIAL_ORDER,
        *hist,
        "??OP(??)",
        "?????",
        "????",
        "??????",
        "??OP??",
        "????",
        "????",
        "??????",
        TASK_COL_RAW_INPUT_DATE,
        "????",
        "?????",
        "??_????",
        "??_????",
        RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16,
        "????",
        "????",
        "???(????)",
        "????_AI",
    ]


def _task_date_key_for_result_sheet_sort(val):
    """??_???????????????????????date.max??"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return date.max
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        ts = pd.Timestamp(val)
        if pd.isna(ts):
            return date.max
        return ts.date()
    except Exception:
        return date.max


def _coerce_planning_date_for_deadline(d) -> date | None:
    """???????????? date ???????? None??"""
    if d is None:
        return None
    if isinstance(d, datetime):
        return d.date()
    if isinstance(d, date):
        return d
    return None


def _result_task_plan_end_within_answer_or_spec_16_label(
    plan_window: list | None, answer_due, specified_due
) -> str:
    """
    ??_??????: ???_??????????????
    ??????? + PLAN_DUE_DAY_COMPLETION_TIME??? 16:00????????
    ???????????????? + 16:00 ????
    ??????????????
    """
    if not plan_window or len(plan_window) < 2:
        return "???"
    _pe = plan_window[1]
    if _pe is None:
        return "???"
    dd = _coerce_planning_date_for_deadline(answer_due)
    if dd is None:
        dd = _coerce_planning_date_for_deadline(specified_due)
    if dd is None:
        return "????"
    try:
        deadline_dt = datetime.combine(dd, PLAN_DUE_DAY_COMPLETION_TIME)
        if _pe <= deadline_dt:
            return "??"
        return "???"
    except Exception:
        return "????"


def _result_task_sheet_sort_key(t: dict):
    """
    ??_??????????????????generate_plan ???????????? 1..n????
    ??????????????????NO?????????????????????
    """
    _dto = t.get("dispatch_trial_order")
    try:
        trial_k = int(_dto) if _dto is not None else 10**9
    except (TypeError, ValueError):
        trial_k = 10**9
    return (
        trial_k,
        str(t.get("task_id", "")).strip(),
        str(t.get("machine", "")).strip(),
        _task_date_key_for_result_sheet_sort(t.get("start_date_req")),
        _task_date_key_for_result_sheet_sort(t.get("answer_due_date")),
        _task_date_key_for_result_sheet_sort(t.get("specified_due_date")),
    )


def _is_result_task_history_expand_token(cell_val) -> bool:
    """???????????1???????1?n ???????????"""
    if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
        return False
    s = unicodedata.normalize("NFKC", str(cell_val).strip())
    return s in ("??", "??*")


def _result_task_column_alias_map(df_columns) -> dict:
    """???? NFKC ????? ? DataFrame ??????"""
    m = {}
    for c in df_columns:
        m[_nfkc_column_aliases(str(c).strip())] = c
    return m


def _resolve_result_task_column_label(label, col_by_norm: dict):
    if label is None or (isinstance(label, float) and pd.isna(label)):
        return None
    s = unicodedata.normalize("NFKC", str(label).strip())
    if not s or s.lower() in ("nan", "none"):
        return None
    nk = _nfkc_column_aliases(s)
    resolved = col_by_norm.get(nk)
    if resolved is not None:
        return resolved
    # ??????????????? ??_????16???
    if nk == _nfkc_column_aliases("??_??16???"):
        return col_by_norm.get(
            _nfkc_column_aliases(RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16)
        )
    return None


def _parse_column_visible_cell(val) -> bool:
    """???: ?????? True?????FALSE/0/??? ??????"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return True
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if val == 0:
            return False
        if val == 1:
            return True
    s = unicodedata.normalize("NFKC", str(val).strip()).lower()
    if s in ("", "true", "1", "??", "yes", "on", "??", "?"):
        return True
    if s in ("false", "flase", "0", "???", "no", "off", "???", "??", "�"):
        return False
    return True


def parse_result_task_column_config_dataframe(
    df_cfg: pd.DataFrame | None, max_history_len: int
) -> list | None:
    """
    ????_??_????????? DataFrame ?? (????, ??) ???????
    ???????????????????????? True??
    ???????*??1????1???n ????????????????????
    ?????NFKC???????????????????????????????????????
    """
    if df_cfg is None or df_cfg.empty:
        return None
    df_cfg = df_cfg.dropna(how="all")
    if df_cfg.empty:
        return None

    name_col = None
    for c in df_cfg.columns:
        if _nfkc_column_aliases(str(c).strip()) == _nfkc_column_aliases(COLUMN_CONFIG_HEADER_COL):
            name_col = c
            break
    if name_col is None:
        name_col = df_cfg.columns[0]

    vis_col = None
    for c in df_cfg.columns:
        if _nfkc_column_aliases(str(c).strip()) == _nfkc_column_aliases(COLUMN_CONFIG_VISIBLE_COL):
            vis_col = c
            break

    seen_norm: set[str] = set()
    out: list[tuple[str, bool]] = []

    def _try_add(label: str, vis: bool) -> None:
        lab = str(label).strip()
        if not lab:
            return
        nk = _nfkc_column_aliases(unicodedata.normalize("NFKC", lab))
        if nk in seen_norm:
            logging.warning(
                "????%s?: ?????%s???????????????????",
                COLUMN_CONFIG_SHEET_NAME,
                lab,
            )
            return
        seen_norm.add(nk)
        out.append((lab, vis))

    for i in range(len(df_cfg)):
        raw = df_cfg[name_col].iloc[i]
        vis = _parse_column_visible_cell(df_cfg[vis_col].iloc[i] if vis_col is not None else None)
        if _is_result_task_history_expand_token(raw):
            for j in range(max_history_len):
                _try_add(f"??{j+1}", vis)
            continue
        if raw is None or (isinstance(raw, float) and pd.isna(raw)):
            continue
        s = unicodedata.normalize("NFKC", str(raw).strip())
        if not s or s.lower() in ("nan", "none"):
            continue
        _try_add(s, vis)
    return out or None


def _xlwings_write_column_config_sheet_ab(xw_sheet, rows: list[tuple[str, bool]]) -> None:
    """??????? A:B ? ????? ???????1???????????"""
    mat = [[COLUMN_CONFIG_HEADER_COL, COLUMN_CONFIG_VISIBLE_COL]]
    for lab, vis in rows:
        mat.append([lab, bool(vis)])
    n_r = len(mat)
    try:
        ur = xw_sheet.used_range
        lim_r = max(ur.row + ur.rows.count - 1, n_r, 2)
        xw_sheet.range((1, 1), (lim_r, 2)).clear_contents()
    except Exception:
        try:
            xw_sheet.range((1, 1)).resize(max(n_r, 50), 2).clear_contents()
        except Exception:
            pass
    xw_sheet.range((1, 1)).resize(n_r, 2).value = mat


def load_result_task_column_rows_from_input_workbook(max_history_len: int) -> list | None:
    """
    TASK_INPUT_WORKBOOK ?????_??_??????????? (????, ??) ???????
    """
    wb = TASKS_INPUT_WORKBOOK
    if not wb or not os.path.exists(wb):
        return None
    if _workbook_should_skip_openpyxl_io(wb):
        logging.info(
            "???: ?????%s?????? pandas(openpyxl) ???%s????????????????????",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
            COLUMN_CONFIG_SHEET_NAME,
        )
        return None
    try:
        df_cfg = pd.read_excel(wb, sheet_name=COLUMN_CONFIG_SHEET_NAME, header=0)
    except ValueError:
        return None
    except Exception as e:
        logging.warning(
            "????%s?: ????????????????????? (%s)",
            COLUMN_CONFIG_SHEET_NAME,
            e,
        )
        return None
    return parse_result_task_column_config_dataframe(df_cfg, max_history_len)


def apply_result_task_sheet_column_order(
    df: pd.DataFrame,
    max_history_len: int,
    *,
    config_dataframe: pd.DataFrame | None = None,
):
    """
    ?????????????????????????????????????? True??
    config_dataframe ????????????????????????????xlwings ??????
    ???: (????? DataFrame, ????????, ??????????, ?????bool)
    """
    default_order = default_result_task_sheet_column_order(max_history_len)
    if config_dataframe is not None:
        user_rows = parse_result_task_column_config_dataframe(config_dataframe, max_history_len)
    else:
        user_rows = load_result_task_column_rows_from_input_workbook(max_history_len)
    if user_rows:
        primary = user_rows
        source = (
            f"???????{COLUMN_CONFIG_SHEET_NAME}?"
            if config_dataframe is None
            else f"????{COLUMN_CONFIG_SHEET_NAME}?????????"
        )
    else:
        primary = [(n, True) for n in default_order]
        source = "??"

    actual = list(df.columns)
    actual_set = set(actual)
    col_by_norm = _result_task_column_alias_map(actual)
    vis_map = {c: True for c in actual}

    seen = set()
    ordered = []
    unknown = []

    for item, vis in primary:
        resolved = _resolve_result_task_column_label(item, col_by_norm)
        if resolved and resolved not in seen:
            ordered.append(resolved)
            seen.add(resolved)
            vis_map[resolved] = vis
        elif not resolved:
            if item is not None and not (isinstance(item, float) and pd.isna(item)):
                lab = str(item).strip()
                if lab and lab not in unknown:
                    unknown.append(lab)

    for name in default_order:
        if name in actual_set and name not in seen:
            ordered.append(name)
            seen.add(name)
    for name in actual:
        if name not in seen:
            ordered.append(name)
            seen.add(name)

    if unknown:
        logging.warning(
            "???: ?????????????????20??: %s",
            ", ".join(unknown[:20]) + (" �" if len(unknown) > 20 else ""),
        )
    logging.info("??_???????????: %s?%s ??", source, len(ordered))
    if not user_rows and config_dataframe is None:
        logging.info(
            "????????????: ?????????????%s?????"
            " ????%s??%s?� ??? FALSE ?????????????"
            " 1????????1?n ????VBA ?????_??_?????_???????????????????????????????",
            COLUMN_CONFIG_SHEET_NAME,
            COLUMN_CONFIG_HEADER_COL,
            COLUMN_CONFIG_VISIBLE_COL,
        )
    return df[ordered], ordered, source, vis_map


def _xlwings_sheet_to_matrix(sheet) -> list:
    """xlwings Sheet ? UsedRange ???? list[list] ????1?????2????"""
    ur = sheet.used_range
    if ur is None:
        return []
    raw = ur.options(ndim=2).value
    if raw is None:
        return []
    if not isinstance(raw, list):
        return [[raw]]
    if len(raw) == 0:
        return []
    if not isinstance(raw[0], list):
        return [raw]
    return raw


def _matrix_to_dataframe_header_first(matrix: list) -> pd.DataFrame | None:
    """1????????? DataFrame ??????? None?"""
    if not matrix or not matrix[0]:
        return None
    header = []
    for x in matrix[0]:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            header.append("")
        else:
            header.append(str(x).strip())
    if not any(h for h in header):
        return None
    body = matrix[1:] if len(matrix) > 1 else []
    return pd.DataFrame(body, columns=header)


def _max_history_len_from_result_task_df_columns(columns) -> int:
    """??_?????????n???? n ??????????? 1??"""
    imax = 0
    for c in columns:
        m = re.match(r"^??(\d+)$", str(c).strip())
        if m:
            imax = max(imax, int(m.group(1)))
    return max(imax, 1)


def apply_result_task_column_layout_via_xlwings(workbook_path: str | None = None) -> bool:
    """
    Excel ?????????????????
    ????_??_?????????????????_????????????????????
    ???????????????? Excel ?????????????xlwings ???????
    """
    path = (workbook_path or "").strip() or TASKS_INPUT_WORKBOOK.strip()
    if not path:
        logging.error("??_????? ???: ??????????TASK_INPUT_WORKBOOK ???????????")
        return False
    try:
        import xlwings as xw
    except ImportError:
        logging.error("??_????? ???: xlwings ? import ??????pip install xlwings ??????????")
        return False

    try:
        wb = xw.Book(path)
    except Exception as e:
        logging.error("??_????? ???: ???????????: %s", e)
        return False

    try:
        ws_res = wb.sheets[RESULT_TASK_SHEET_NAME]
        ws_cfg = wb.sheets[COLUMN_CONFIG_SHEET_NAME]
    except Exception as e:
        logging.error("??_????? ???: ?????????????: %s", e)
        return False

    mat_res = _xlwings_sheet_to_matrix(ws_res)
    mat_cfg = _xlwings_sheet_to_matrix(ws_cfg)
    df_res = _matrix_to_dataframe_header_first(mat_res)
    df_cfg = _matrix_to_dataframe_header_first(mat_cfg)
    if df_res is None or df_res.empty:
        logging.error("??_????? ???: ?%s????????????", RESULT_TASK_SHEET_NAME)
        return False
    if df_cfg is None:
        logging.error("??_????? ???: ?%s????????????", COLUMN_CONFIG_SHEET_NAME)
        return False

    max_h = _max_history_len_from_result_task_df_columns(df_res.columns)
    rows_cfg = parse_result_task_column_config_dataframe(df_cfg, max_h)
    if not rows_cfg:
        logging.error(
            "??_????? ???: ?%s???????????????",
            COLUMN_CONFIG_SHEET_NAME,
        )
        return False
    _xlwings_write_column_config_sheet_ab(ws_cfg, rows_cfg)
    df_cfg_clean = pd.DataFrame(
        rows_cfg, columns=[COLUMN_CONFIG_HEADER_COL, COLUMN_CONFIG_VISIBLE_COL]
    )
    df_out, ordered, source, vis_map = apply_result_task_sheet_column_order(
        df_res, max_h, config_dataframe=df_cfg_clean
    )

    df_write = df_out.astype(object).where(pd.notna(df_out), None)
    headers = [str(h) for h in df_write.columns.tolist()]
    body = df_write.values.tolist()
    out_matrix = [headers] + body
    nrows = len(out_matrix)
    ncols = len(headers)
    if ncols == 0:
        return False

    try:
        ur_old = ws_res.used_range
        if ur_old is not None:
            ws_res.range((ur_old.row, ur_old.column)).resize(
                ur_old.rows.count, ur_old.columns.count
            ).clear_contents()
    except Exception:
        try:
            ws_res.used_range.clear_contents()
        except Exception:
            pass

    ws_res.range((1, 1)).resize(nrows, ncols).value = out_matrix

    for ci in range(1, ncols + 1):
        try:
            ws_res.range((1, ci)).api.EntireColumn.Hidden = False
        except Exception:
            pass

    for ci, col_name in enumerate(ordered, 1):
        if not vis_map.get(col_name, True):
            try:
                ws_res.range((1, ci)).api.EntireColumn.Hidden = True
            except Exception as e:
                logging.warning("?????????%s %s?: %s", ci, col_name, e)

    try:
        wb.save()
    except Exception as e:
        logging.warning("??_????? ???: ????????????????????????: %s", e)

    logging.info(
        "??_????? ?????: %s?%s ?????=%s?",
        source,
        len(ordered),
        sum(1 for c in ordered if not vis_map.get(c, True)),
    )
    return True


def apply_result_task_column_layout_only() -> bool:
    """???? TASK_INPUT_WORKBOOK ????????????????VBA ??????"""
    p = os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK
    return apply_result_task_column_layout_via_xlwings(p)

_PLAN_INPUT_XLWINGS_ORIG_ROW = "__orig_sheet_row__"


def _plan_input_dispatch_trial_order_local_only_from_env() -> bool:
    """???? PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY ???? post_load ????????"""
    v = (os.environ.get("PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY") or "").strip().lower()
    return v in ("1", "true", "yes", "on", "y")


def refresh_plan_input_dispatch_trial_order_via_xlwings(
    workbook_path: str | None = None,
    *,
    apply_post_load_mutations: bool = True,
) -> bool:
    """
    Excel ?????????????????_???????????
    ??2 ??? ``fill_plan_dispatch_trial_order_column_stage1`` ??????????
    ???????1 ??????????????????
    ????????? xlwings ???????? read_excel ??????

    apply_post_load_mutations=False ???? ``_apply_planning_sheet_post_load_mutations`` ??????
    ????DATA ???????????????True ??????2 ??????
    ??_?????????????????????????????????????????
    ????????????????????_?????????????????????????????? False?
    """
    path = (workbook_path or "").strip() or os.environ.get(
        "TASK_INPUT_WORKBOOK", ""
    ).strip() or TASKS_INPUT_WORKBOOK.strip()
    if not path:
        logging.error("????????: ??????????")
        return False
    try:
        import xlwings as xw
    except ImportError:
        logging.error("????????: xlwings ???????")
        return False
    try:
        wb = xw.Book(path)
        ws = wb.sheets[PLAN_INPUT_SHEET_NAME]
    except Exception as e:
        logging.error("????????: ????????: %s", e)
        return False

    mat = _xlwings_sheet_to_matrix(ws)
    df = _matrix_to_dataframe_header_first(mat)
    if df is None or df.empty:
        logging.warning("????????: ???????????")
        return False

    df = df.copy()
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(df, plan_input_sheet_column_order())
    for c in plan_input_sheet_column_order():
        if c not in df.columns:
            df[c] = ""

    df.insert(0, _PLAN_INPUT_XLWINGS_ORIG_ROW, range(len(df)))

    if apply_post_load_mutations:
        _apply_planning_sheet_post_load_mutations(df, path, "????????")
    else:
        logging.info(
            "????????: ????????????????????????????????????"
        )

    dto_col = RESULT_TASK_COL_DISPATCH_TRIAL_ORDER
    if dto_col not in df.columns:
        logging.error("????????: ??%s????????", dto_col)
        return False

    _dto_loc = df.columns.get_loc(dto_col)
    if isinstance(_dto_loc, slice):
        logging.error("????????: ??%s?????????", dto_col)
        return False
    # Excel ????? float64 ??? "" ????? pandas ???????????? NaN
    df[dto_col] = float("nan")

    data_extract_dt = _extract_data_extraction_datetime()
    base_now_dt = data_extract_dt if data_extract_dt is not None else datetime.now()
    run_date = base_now_dt.date()

    try:
        (
            _sd,
            _mem,
            equipment_list,
            req_map,
            need_rules,
            _sm,
            need_combo_col_index,
        ) = load_skills_and_needs()
    except Exception as e:
        logging.exception("????????: master ?????: %s", e)
        return False

    try:
        fill_plan_dispatch_trial_order_column_stage1(
            df,
            run_date,
            req_map,
            need_rules,
            need_combo_col_index,
            equipment_list,
        )
    except Exception as e:
        logging.exception("????????: ????????: %s", e)
        return False

    df_sorted = _sort_stage1_plan_df_by_dispatch_trial_order_asc(df)
    orig_list = [int(x) for x in df_sorted[_PLAN_INPUT_XLWINGS_ORIG_ROW].tolist()]
    df_sorted = df_sorted.drop(columns=[_PLAN_INPUT_XLWINGS_ORIG_ROW])

    header_row = mat[0] if mat else []
    n_hdr = len(header_row)
    if n_hdr == 0:
        return False

    def _pad_row(r, n):
        r = list(r) if r is not None else []
        if len(r) < n:
            r = r + [None] * (n - len(r))
        return r

    new_mat = [_pad_row(header_row, n_hdr)]
    for i in range(len(df_sorted)):
        orig = orig_list[i]
        src_row = mat[orig + 1] if orig + 1 < len(mat) else []
        src_row = _pad_row(src_row, n_hdr)
        out_row = []
        for j in range(n_hdr):
            h_cell = header_row[j]
            if h_cell is None or (isinstance(h_cell, float) and pd.isna(h_cell)):
                hname = ""
            else:
                hname = str(h_cell).strip()
            if hname and hname in df_sorted.columns:
                v = df_sorted.iat[i, df_sorted.columns.get_loc(hname)]
                if pd.isna(v):
                    out_row.append(None)
                else:
                    out_row.append(v)
            else:
                out_row.append(src_row[j])
        new_mat.append(out_row)

    try:
        n_r = len(new_mat)
        ws.range((1, 1)).resize(n_r, n_hdr).value = new_mat
    except Exception as e:
        logging.exception("????????: ????????: %s", e)
        return False

    try:
        wb.save()
    except Exception as e:
        logging.warning("????????: Save ??: %s", e)

    logging.info(
        "????????: ?%s?? %s ?????????",
        PLAN_INPUT_SHEET_NAME,
        len(df_sorted),
    )
    return True


def refresh_plan_input_dispatch_trial_order_only() -> bool:
    """TASK_INPUT_WORKBOOK ??????????????VBA / cmd ?????????
    ???? PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY=1 ???????????????????
    """
    p = os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK
    local = _plan_input_dispatch_trial_order_local_only_from_env()
    return refresh_plan_input_dispatch_trial_order_via_xlwings(
        p, apply_post_load_mutations=not local
    )


def apply_plan_input_column_layout_only() -> bool:
    """
    ????_?????????????????????VBA ???
    ???????????1???????????????????
    """
    logging.warning("apply_plan_input_column_layout_only: not implemented")
    return False



def dedupe_result_task_column_config_sheet_via_xlwings(workbook_path: str | None = None) -> bool:
    """
    ????_??_??????? A:B ????????????????????????????
    ???_?????????????????????????????????
    """
    path = (workbook_path or "").strip() or TASKS_INPUT_WORKBOOK.strip()
    if not path:
        logging.error("??? ????: ??????????")
        return False
    try:
        import xlwings as xw
    except ImportError:
        logging.error("??? ????: xlwings ? import ??????")
        return False
    try:
        wb = xw.Book(path)
        ws_cfg = wb.sheets[COLUMN_CONFIG_SHEET_NAME]
    except Exception as e:
        logging.error("??? ????: ?????????????: %s", e)
        return False

    max_h = 1
    try:
        ws_res = wb.sheets[RESULT_TASK_SHEET_NAME]
        df_r = _matrix_to_dataframe_header_first(_xlwings_sheet_to_matrix(ws_res))
        if df_r is not None and not df_r.empty:
            max_h = _max_history_len_from_result_task_df_columns(df_r.columns)
    except Exception:
        pass

    df_cfg = _matrix_to_dataframe_header_first(_xlwings_sheet_to_matrix(ws_cfg))
    if df_cfg is None:
        logging.error("??? ????: ?%s????????????", COLUMN_CONFIG_SHEET_NAME)
        return False
    rows = parse_result_task_column_config_dataframe(df_cfg, max_h)
    if not rows:
        logging.warning("??? ????: ??????????????")
        return False
    _xlwings_write_column_config_sheet_ab(ws_cfg, rows)
    try:
        wb.save()
    except Exception as e:
        logging.warning("??? ????: ????: %s", e)
    logging.info(
        "????%s????????? %s ???????????????????",
        COLUMN_CONFIG_SHEET_NAME,
        len(rows),
    )
    return True


def dedupe_result_task_column_config_sheet_only() -> bool:
    """???? TASK_INPUT_WORKBOOK ??????????????????VBA ???"""
    p = os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK
    return dedupe_result_task_column_config_sheet_via_xlwings(p)


def _apply_result_task_sheet_column_visibility(worksheet, column_names: list, vis_map: dict):
    """??_???????vis_map ? False ??????????"""
    for idx, col_name in enumerate(column_names, 1):
        if not vis_map.get(col_name, True):
            worksheet.column_dimensions[get_column_letter(idx)].hidden = True


def _norm_history_member_label(name: str) -> str:
    """?????????????????1?????trim??????????"""
    t = str(name or "").replace("\u3000", " ").strip()
    return " ".join(t.split())


def _history_team_text_main_assignment_only(h: dict) -> str:
    """
    ???????????: ??????????????????????????
    append_surplus ?? h['team'] ?? post_dispatch_surplus_names ??????
    """
    raw = (h.get("team") or "").strip()
    if not raw:
        return ""
    ps = h.get("post_dispatch_surplus_names") or []
    if not ps:
        return raw
    ps_set = {
        _norm_history_member_label(x)
        for x in ps
        if x and str(x).strip()
    }
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    kept = [p for p in parts if _norm_history_member_label(p) not in ps_set]
    return ", ".join(kept) if kept else raw


def _format_result_task_history_cell(task: dict, h: dict) -> str:
    """??_??????????????????????ID?????????????????????"""
    um = task.get("unit_m") or 0
    try:
        done_r = int(h["done_m"] / um) if um else 0
    except (TypeError, ValueError, ZeroDivisionError):
        done_r = 0
    dm = h.get("done_m", 0)
    parts_out: list[str] = [f"??{h.get('date', '')}??{done_r}R ({dm}m)"]
    cid = h.get("combo_sheet_row_id")
    if cid is not None:
        try:
            parts_out.append(f"????#{int(cid)}")
        except (TypeError, ValueError):
            parts_out.append(f"????#{cid}")
    parts_out.append(f"??[{_history_team_text_main_assignment_only(h)}]")
    sm = h.get("surplus_member_names") or []
    if sm:
        parts_out.append(f"??[{','.join(str(x) for x in sm)}]")
    ps = h.get("post_dispatch_surplus_names") or []
    if ps:
        parts_out.append(f"????[{','.join(str(x) for x in ps)}]")
    return " ".join(parts_out)


_RESULT_TASK_HISTORY_RICH_HEAD_RE = re.compile(r"^?(?[^?]*?)(.*)$", re.DOTALL)


def _apply_result_task_history_rich_text(worksheet, column_names: list):
    """
    ???: ???????�??????????????????????
    openpyxl 3.1 ???????????????????
    """
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        from openpyxl.styles.colors import Color
    except ImportError:
        return

    hist_cols = [
        i + 1 for i, c in enumerate(column_names) if str(c).startswith("??")
    ]
    if not hist_cols:
        return

    _plain_kw: dict = {}
    _blue_kw: dict = {"color": Color(rgb="FF0070C0")}
    plain_if = InlineFont(**_plain_kw)
    blue_if = InlineFont(**_blue_kw)
    top = Alignment(wrap_text=False, vertical="top")

    for r in range(2, worksheet.max_row + 1):
        for ci in hist_cols:
            cell = worksheet.cell(row=r, column=ci)
            v = cell.value
            if not isinstance(v, str) or not v.startswith("??"):
                continue
            m = _RESULT_TASK_HISTORY_RICH_HEAD_RE.match(v)
            if not m:
                continue
            bracketed, rest = m.group(1), m.group(2)
            cell.value = CellRichText(
                TextBlock(plain_if, "?"),
                TextBlock(blue_if, bracketed),
                TextBlock(plain_if, rest),
            )
            cell.alignment = top


def _apply_result_task_date_columns_blue_font(worksheet, column_names: list):
    """
    ??_?????: ??????????????????????????????????????
    ?????????? _apply_result_task_history_rich_text ???? 0070C0 ????
    """
    blue = _result_font(color="0070C0")
    top = Alignment(wrap_text=False, vertical="top")
    for col_idx, col_name in enumerate(column_names, 1):
        if str(col_name) not in RESULT_TASK_DATE_STYLE_HEADERS:
            continue
        for r in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=r, column=col_idx)
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str) and not str(v).strip():
                continue
            cell.font = blue
            cell.alignment = top


def _apply_result_task_history_need_surplus_highlight(
    worksheet, column_names: list, sorted_tasks: list
):
    """
    need??????????????????????????????????
    ??????????????????????????????n??????????
    """
    hist_cols: list[tuple[int, int]] = []
    for col_idx, col_name in enumerate(column_names, 1):
        m = re.match(r"^??(\d+)$", str(col_name).strip())
        if m:
            hist_cols.append((int(m.group(1)), col_idx))
    hist_cols.sort(key=lambda x: x[0])
    if not hist_cols or worksheet.max_row < 2:
        return
    fill_surplus = PatternFill(
        start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"
    )
    n_tasks = len(sorted_tasks)
    for r in range(2, worksheet.max_row + 1):
        ti = r - 2
        if ti < 0 or ti >= n_tasks:
            continue
        ah = sorted_tasks[ti].get("assigned_history") or []
        for ord1, cidx in hist_cols:
            i = ord1 - 1
            if i < 0 or i >= len(ah):
                continue
            if not ah[i].get("need_surplus_assigned"):
                continue
            worksheet.cell(row=r, column=cidx).fill = fill_surplus


def _apply_result_task_task_id_content_mismatch_highlight(
    worksheet, column_names: list, sorted_tasks: list
):
    """
    ????????????????????ID????????????????????????????
    """
    task_id_col_idx = None
    for col_idx, col_name in enumerate(column_names, 1):
        if str(col_name) == "???ID":
            task_id_col_idx = col_idx
            break
    if task_id_col_idx is None or worksheet.max_row < 2:
        return
    fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    font_white = _result_font(color="FFFFFF")
    top = Alignment(wrap_text=False, vertical="top")
    n_data = worksheet.max_row - 1
    for i in range(min(len(sorted_tasks), n_data)):
        if not sorted_tasks[i].get("process_content_mismatch"):
            continue
        cell = worksheet.cell(row=i + 2, column=task_id_col_idx)
        cell.fill = fill_red
        cell.font = font_white
        cell.alignment = top


def _apply_result_task_plan_end_answer_spec_16_no_highlight(
    worksheet, column_names: list
):
    """
    ????_????16????????????????????????????
    ?????????_??16????????????????
    """
    target_names = frozenset(
        {
            RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16,
            "??_??16???",
        }
    )
    col_idx = None
    for ci, col_name in enumerate(column_names, 1):
        if str(col_name) in target_names:
            col_idx = ci
            break
    if col_idx is None or worksheet.max_row < 2:
        return
    fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    font_white_bold = _result_font(color="FFFFFF", bold=True)
    top = Alignment(wrap_text=False, vertical="top")
    for r in range(2, worksheet.max_row + 1):
        cell = worksheet.cell(row=r, column=col_idx)
        v = cell.value
        if v is None:
            continue
        s = str(v).strip()
        if s != "???":
            continue
        cell.fill = fill_red
        cell.font = font_white_bold
        cell.alignment = top


def _apply_result_task_id_hyperlinks_to_equipment_schedule(
    worksheet_tasks,
    column_names: list,
    sorted_tasks_for_row_order: list,
    task_id_to_schedule_cell: dict[str, str],
    schedule_sheet_name: str,
) -> None:
    """
    ??_??????????ID???????_???????????????????????????????????????
    ??????????????????????????
    """
    if not task_id_to_schedule_cell or worksheet_tasks.max_row < 2:
        return
    task_id_col_idx = None
    for col_idx, col_name in enumerate(column_names, 1):
        if str(col_name) == "???ID":
            task_id_col_idx = col_idx
            break
    if task_id_col_idx is None:
        return
    esc = schedule_sheet_name.replace("'", "''")
    loc_prefix = f"#'{esc}'!"
    font_link = Font(color="0563C1", underline="single")
    font_link_on_red = Font(color="FFFFFF", underline="single")
    top = Alignment(wrap_text=False, vertical="top")
    n_tasks = len(sorted_tasks_for_row_order)
    for r in range(2, worksheet_tasks.max_row + 1):
        cell = worksheet_tasks.cell(row=r, column=task_id_col_idx)
        raw = cell.value
        if raw is None:
            continue
        tid = str(raw).strip()
        if not tid:
            continue
        addr = task_id_to_schedule_cell.get(tid)
        if not addr:
            continue
        cell.hyperlink = loc_prefix + addr
        row_i = r - 2
        mismatch = (
            row_i < n_tasks
            and bool(sorted_tasks_for_row_order[row_i].get("process_content_mismatch"))
        )
        cell.font = font_link_on_red if mismatch else font_link
        cell.alignment = top


def _add_column_config_sheet_helpers(ws_cfg, num_data_rows: int):
    """???? TRUE/FALSE ???????????????????????"""
    last_r = max(num_data_rows + 1, 2)
    cap = max(last_r + 50, 500)
    dv = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws_cfg.add_data_validation(dv)
    dv.add(f"B2:B{cap}")


def _stage2_try_copy_column_config_shapes_from_input(
    result_path: str,
    input_path: str | None,
) -> None:
    """
    pandas/openpyxl ???????????????????????
    ??????????? 0/false/no/off ?????????????
    ????_??_???????? Shapes ??????????????????
    ???? Left/Top/Width/Height????????? Placement????????????
    openpyxl ???????????????????????????
    """
    if not STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT:
        return
    rp = (result_path or "").strip()
    ip = (input_path or "").strip()
    if not rp or not os.path.isfile(rp):
        logging.warning(
            "???????????: ???????????????????"
        )
        return
    if not ip or not os.path.isfile(ip):
        logging.warning(
            "???????????: TASK_INPUT_WORKBOOK ???????????????"
        )
        return
    try:
        import xlwings as xw
    except ImportError:
        logging.warning(
            "???????????: xlwings ? import ??????"
        )
        return
    app = None
    wb_out = None
    wb_in = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        wb_out = app.books.open(os.path.abspath(rp), update_links=False)
        wb_in = app.books.open(os.path.abspath(ip), read_only=True, update_links=False)
        try:
            ws_out = wb_out.sheets[COLUMN_CONFIG_SHEET_NAME]
        except Exception:
            logging.warning(
                "???????????: ??????????%s????????",
                COLUMN_CONFIG_SHEET_NAME,
            )
            return
        try:
            ws_in = wb_in.sheets[COLUMN_CONFIG_SHEET_NAME]
        except Exception:
            logging.warning(
                "???????????: ??????????%s????????",
                COLUMN_CONFIG_SHEET_NAME,
            )
            return
        n_shapes = int(ws_in.api.Shapes.Count)
        if n_shapes <= 0:
            logging.info(
                "???????????: ???????????????????"
            )
            return
        ws_out.activate()
        api_in = ws_in.api
        api_out = ws_out.api
        for i in range(1, n_shapes + 1):
            src = api_in.Shapes(i)
            left = float(src.Left)
            top = float(src.Top)
            width = float(src.Width)
            height = float(src.Height)
            placement = None
            try:
                placement = int(src.Placement)
            except Exception:
                pass
            src.Copy()
            api_out.Paste()
            dst = api_out.Shapes(int(api_out.Shapes.Count))
            try:
                dst.LockAspectRatio = False
            except Exception:
                pass
            if placement is not None:
                try:
                    dst.Placement = placement
                except Exception:
                    pass
            dst.Left = left
            dst.Top = top
            dst.Width = width
            dst.Height = height
        wb_out.save()
        logging.info(
            "???????????: ???? %s ??????????????????",
            n_shapes,
        )
    except Exception as e:
        logging.warning(
            "???????????: ???????%s??Excel ???COM ??????????????",
            e,
        )
    finally:
        for _wb in (wb_in, wb_out):
            if _wb is not None:
                try:
                    _wb.close()
                except Exception:
                    pass
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


def _coerce_actual_sheet_datetime(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, date) and not isinstance(val, datetime):
        return datetime.combine(val, time(0, 0))
    try:
        ts = pd.to_datetime(val, errors="coerce")
        if pd.isna(ts) or ts is pd.NaT:
            return None
        if isinstance(ts, pd.Timestamp):
            return ts.to_pydatetime()
        return ts if isinstance(ts, datetime) else None
    except Exception:
        return None


def _actual_row_time_bounds(row):
    """????DATA ?1??? (??, ??) ?????????? (None, None)?"""
    s_dt = _coerce_actual_sheet_datetime(row.get(ACT_COL_START_DT))
    e_dt = _coerce_actual_sheet_datetime(row.get(ACT_COL_END_DT))
    if s_dt and e_dt and s_dt < e_dt:
        return s_dt, e_dt
    s_dt = _coerce_actual_sheet_datetime(row.get(ACT_COL_START_ALT))
    e_dt = _coerce_actual_sheet_datetime(row.get(ACT_COL_END_ALT))
    if s_dt and e_dt and s_dt < e_dt:
        return s_dt, e_dt

    d_date = parse_optional_date(row.get(ACT_COL_DAY))
    if not d_date:
        cd = _coerce_actual_sheet_datetime(row.get(ACT_COL_DAY))
        if isinstance(cd, datetime):
            d_date = cd.date()
        elif isinstance(cd, date):
            d_date = cd
    if not d_date:
        return None, None

    ts_s = row.get(ACT_COL_TIME_START)
    ts_e = row.get(ACT_COL_TIME_END)
    if ts_s is None or pd.isna(ts_s) or ts_e is None or pd.isna(ts_e):
        return None, None

    if isinstance(ts_s, time):
        t0 = ts_s
    elif isinstance(ts_s, datetime):
        t0 = ts_s.time()
    else:
        t0 = parse_time_str(ts_s, None)

    if isinstance(ts_e, time):
        t1 = ts_e
    elif isinstance(ts_e, datetime):
        t1 = ts_e.time()
    else:
        t1 = parse_time_str(ts_e, None)

    if t0 is None or t1 is None or t0 >= t1:
        return None, None
    return datetime.combine(d_date, t0), datetime.combine(d_date, t1)


def load_machining_actuals_df():
    """
    ????????????DATA?????????? DataFrame??
    Power Query ?????????????
    """
    if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
        return pd.DataFrame()
    try:
        df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=ACTUALS_SHEET_NAME)
    except ValueError:
        logging.info(
            f"????{ACTUALS_SHEET_NAME}??????????????????????"
        )
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(df, ACTUAL_HEADER_CANONICAL)
    logging.info(
        f"????: '{TASKS_INPUT_WORKBOOK}' ? '{ACTUALS_SHEET_NAME}' ? {len(df)} ??????"
    )
    return df


def build_actual_timeline_events(df, equipment_list, sorted_dates):
    """
    ?????????????????????
    ??????sorted_dates???????????????????????
    ???? NFKC??????????????????????
    ??? DEFAULT_START_TIME / DEFAULT_END_TIME ?????????
    """
    if df is None or len(df) == 0:
        return []
    equip_lookup = _equipment_lookup_normalized_to_canonical(equipment_list)
    date_ok = set(sorted_dates)
    events = []
    bad_eq = 0
    bad_time = 0
    no_plan_overlap = 0
    mismatch_norm_samples = []

    for _, row in df.iterrows():
        tid = row.get(ACT_COL_TASK_ID)
        if tid is None or pd.isna(tid):
            continue
        tid_s = str(tid).strip()
        if not tid_s:
            continue
        proc = row.get(ACT_COL_PROCESS)
        if proc is None or pd.isna(proc):
            continue
        proc_key = _normalize_equipment_match_key(proc)
        mach = equip_lookup.get(proc_key)
        if not mach:
            bad_eq += 1
            if len(mismatch_norm_samples) < 12 and proc_key:
                if proc_key not in mismatch_norm_samples:
                    mismatch_norm_samples.append(proc_key)
            continue
        start_dt, end_dt = _actual_row_time_bounds(row)
        if not start_dt or not end_dt or start_dt >= end_dt:
            bad_time += 1
            continue
        op_val = row.get(ACT_COL_OPERATOR)
        op_s = ""
        if op_val is not None and not pd.isna(op_val):
            op_s = str(op_val).strip()

        before = len(events)
        for d in sorted_dates:
            if d not in date_ok:
                continue
            day_start = datetime.combine(d, DEFAULT_START_TIME)
            day_end = datetime.combine(d, DEFAULT_END_TIME)
            if end_dt <= day_start or start_dt >= day_end:
                continue
            s_clip = max(start_dt, day_start)
            e_clip = min(end_dt, day_end)
            if s_clip >= e_clip:
                continue
            events.append(
                {
                    "date": d,
                    "task_id": tid_s,
                    "machine": mach,
                    "op": op_s,
                    "sub": "",
                    "start_dt": s_clip,
                    "end_dt": e_clip,
                    "breaks": [],
                    "units_done": 0,
                    "already_done_units": 0,
                    "total_units": 0,
                    "eff_time_per_unit": 0.0,
                    "unit_m": 0.0,
                }
            )
        if len(events) == before:
            no_plan_overlap += 1

    if bad_eq:
        logging.warning(
            f"????DATA: ????????????????? {bad_eq} ?????????????????????"
        )
        if mismatch_norm_samples:
            logging.info(
                "  ???????????????????: "
                + " | ".join(mismatch_norm_samples[:12])
            )
    if bad_time:
        logging.info(
            f"????DATA: ??/????????????? {bad_time} ??????????"
        )
    if no_plan_overlap and sorted_dates:
        logging.info(
            f"????DATA: ?????????????????????????�{DEFAULT_START_TIME}?{DEFAULT_END_TIME}????????? {no_plan_overlap} ???????"
        )
    if not events and len(df) > 0:
        logging.info(
            "????DATA: ??????????0???????????????????????sorted_dates?????????????????"
        )
    logging.info(f"????DATA ??????????? {len(events)} ?????????")
    return events


TASK_SPECIAL_AI_LAST_RESPONSE_FILE = "ai_task_special_remark_last.txt"
# ????????????????????SHA????????????????????????????
TASK_SPECIAL_CACHE_KEY_PREFIX = "TASK_SPECIAL_v3|"
# ??????????????????????? ? Gemini ?????????????????
GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX = "GLOBAL_PRIO_v8|"


def _normalize_special_task_id_for_ai(val):
    """
    ??NO??????????????????????
    Excel ?????? float ???????? 12345.0 ? \"12345\" ?????
    ???? NFKC????????????????????????API??????????
    """
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except TypeError:
        pass
    if isinstance(val, bool):
        return None
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if math.isnan(val):
            return None
        if val.is_integer():
            return str(int(val))
        s = str(val).strip()
        if not s or s.lower() in ("nan", "none", "null"):
            return None
        return unicodedata.normalize("NFKC", s).strip() or None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return None
    s = unicodedata.normalize("NFKC", s).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return None
    # ??????? "20010.0" ??Excel?CSV?????????NO????
    if re.fullmatch(r"-?\d+\.0+", s):
        try:
            return str(int(float(s)))
        except ValueError:
            pass
    return s or None


def planning_task_id_str_from_scalar(val) -> str:
    """?????1???????????????NO?????????? \"\"??"""
    return _normalize_special_task_id_for_ai(val) or ""


def planning_task_id_str_from_plan_row(row) -> str:
    """????????????????????NO? planning_task_id_str_from_scalar ????"""
    return planning_task_id_str_from_scalar(_planning_df_cell_scalar(row, TASK_COL_TASK_ID))


def _cell_text_task_special_remark(val):
    """
    ????_???????????????????? **strip ??**
    ?????????Excel ??????????????????????????????????
    """
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except TypeError:
        pass
    if isinstance(val, bool):
        s = str(val)
    elif isinstance(val, float):
        if math.isnan(val):
            return ""
        # ????????????????????????
        if val.is_integer():
            s = str(int(val))
        else:
            s = str(val)
    elif isinstance(val, int):
        s = str(val)
    else:
        s = str(val)
    s = s.strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s


def _task_special_prompt_lines(tasks_df):
    """??????????????????????????????????"""
    lines = []
    for _, row in tasks_df.iterrows():
        if _plan_row_exclude_from_assignment(row):
            continue
        tid = planning_task_id_str_from_plan_row(row)
        rem = _cell_text_task_special_remark(row.get(PLAN_COL_SPECIAL_REMARK))
        if not tid or not rem:
            continue
        proc = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        macn = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        proc_disp = proc if proc else "???"
        macn_disp = macn if macn else "???"
        lines.append(
            f"- ??NO?{tid}?| ????{proc_disp}? | ????{macn_disp}? | ????: {rem}"
        )
    return lines


def _repair_task_special_ai_wrong_top_level_keys(parsed: dict, tasks_df) -> dict:
    """
    ?????????????: 20010 ?????????????????????? JSON ??????
    ????????????NO?�???????????????????????????NO???????
    """
    if not isinstance(parsed, dict) or not parsed or tasks_df is None or getattr(tasks_df, "empty", True):
        return parsed
    valid_tids: set[str] = set()
    remark_by_tid: dict[str, list[str]] = {}
    for _, row in tasks_df.iterrows():
        if _plan_row_exclude_from_assignment(row):
            continue
        tid = planning_task_id_str_from_plan_row(row)
        rem = _cell_text_task_special_remark(row.get(PLAN_COL_SPECIAL_REMARK))
        if not tid or not rem:
            continue
        valid_tids.add(tid)
        remark_by_tid.setdefault(tid, []).append(rem)

    for bad_key in list(parsed.keys()):
        sk = str(bad_key).strip()
        if sk in valid_tids:
            continue
        if not re.fullmatch(r"\d{4,}", sk):
            continue
        hits = [
            tid
            for tid, rems in remark_by_tid.items()
            if any(
                r.startswith(sk)
                or r.startswith(sk + " ")
                or r.startswith(sk + "\u3000")
                or r.startswith(sk + "-")
                or r.startswith(sk + "?")
                for r in rems
            )
        ]
        if len(hits) != 1:
            continue
        target = hits[0]
        val = parsed.pop(bad_key, None)
        if val is None:
            continue
        if target not in parsed:
            parsed[target] = val
            logging.info(
                "???????: JSON ???????????%r ???NO???? ? %r?",
                bad_key,
                target,
            )
        else:
            parsed[bad_key] = val
    return parsed


def _normalize_task_special_scope_str(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return unicodedata.normalize("NFKC", str(s).strip())


def _task_special_scope_matches_row_field(row_val, restrict_val) -> bool:
    """
    restrict ????????????True??
    ???? Excel ??????????????????
    """
    if restrict_val is None:
        return True
    r = _normalize_task_special_scope_str(restrict_val)
    if not r:
        return True
    v = _normalize_task_special_scope_str(row_val)
    if not v:
        return False
    if v == r:
        return True
    if r in v or v in r:
        return True
    return False


def _ai_remark_entry_applies_to_row(entry: dict, row) -> bool:
    """restrict_to_* ??????????NO???????"""
    if not isinstance(entry, dict):
        return False
    rp = row.get(TASK_COL_MACHINE, "")
    rm = row.get(TASK_COL_MACHINE_NAME, "")
    if not _task_special_scope_matches_row_field(rp, entry.get("restrict_to_process_name")):
        return False
    if not _task_special_scope_matches_row_field(rm, entry.get("restrict_to_machine_name")):
        return False
    return True


def _row_matches_remark_source_row(entry: dict, row) -> bool:
    """
    JSON ? process_name / machine_name ???? Excel ????????????????
    ????????????????????????????????????
    """
    if not isinstance(entry, dict):
        return False
    rp = _normalize_task_special_scope_str(row.get(TASK_COL_MACHINE))
    rm = _normalize_task_special_scope_str(row.get(TASK_COL_MACHINE_NAME))
    ep = _normalize_task_special_scope_str(entry.get("process_name"))
    em = _normalize_task_special_scope_str(entry.get("machine_name"))
    proc_ok = (not ep) or (not rp) or ep == rp or ep in rp or rp in ep
    mac_ok = (not em) or (not rm) or em == rm or em in rm or rm in em
    return proc_ok and mac_ok


def _entry_is_global_task_special_scope(entry: dict) -> bool:
    """restrict_to_* ??????????NO?????????????"""
    if not isinstance(entry, dict):
        return False
    a = _normalize_task_special_scope_str(entry.get("restrict_to_process_name"))
    b = _normalize_task_special_scope_str(entry.get("restrict_to_machine_name"))
    return not a and not b


def _select_ai_task_special_entry_for_tid_value(val, row):
    """1??NO?????? dict ??? dict ???????????????????"""
    if val is None:
        return None
    if isinstance(val, list):
        for item in val:
            if (
                isinstance(item, dict)
                and _ai_remark_entry_applies_to_row(item, row)
                and _row_matches_remark_source_row(item, row)
            ):
                return item
        for item in val:
            if (
                isinstance(item, dict)
                and _ai_remark_entry_applies_to_row(item, row)
                and _entry_is_global_task_special_scope(item)
            ):
                return item
        for item in val:
            if isinstance(item, dict) and _ai_remark_entry_applies_to_row(item, row):
                return item
        return None
    if isinstance(val, dict):
        if _ai_remark_entry_applies_to_row(val, row):
            return val
        return None
    return None


def _ai_task_special_entry_for_row(ai_by_tid, row):
    """
    analyze_task_special_remarks ?????????????????
    ???????????????NO????Excel ? 12345.0 ????????
    restrict_to_process_name / restrict_to_machine_name ?????????
    ????NO???????????????????????
    """
    if not isinstance(ai_by_tid, dict) or not ai_by_tid:
        return None
    tid_norm = planning_task_id_str_from_plan_row(row)
    tid_raw = str(_planning_df_cell_scalar(row, TASK_COL_TASK_ID) or "").strip()

    def try_val(v):
        return _select_ai_task_special_entry_for_tid_value(v, row)

    if tid_norm and tid_norm in ai_by_tid:
        hit = try_val(ai_by_tid[tid_norm])
        if hit is not None:
            return hit
    if tid_raw:
        for key in (tid_raw, str(tid_raw)):
            if key in ai_by_tid:
                hit = try_val(ai_by_tid[key])
                if hit is not None:
                    return hit
    if tid_norm:
        for k, v in ai_by_tid.items():
            if str(k).strip() == tid_norm:
                hit = try_val(v)
                if hit is not None:
                    return hit
    if tid_raw:
        for k, v in ai_by_tid.items():
            if str(k).strip() == tid_raw:
                hit = try_val(v)
                if hit is not None:
                    return hit
    return None


def _gemini_result_text(res):
    try:
        return (res.text or "").strip()
    except Exception:
        return ""


# 1 ?? Python ?????1 ??? ??2??????????
_gemini_usage_session: dict[str, dict[str, int]] = {}


def reset_gemini_usage_tracker() -> None:
    global _gemini_usage_session
    _gemini_usage_session = {}


def _gemini_cumulative_json_path() -> str:
    path = os.path.join(api_payment_dir, GEMINI_USAGE_CUMULATIVE_JSON_FILE)
    legacy = os.path.join(log_dir, GEMINI_USAGE_CUMULATIVE_JSON_FILE)
    if os.path.isfile(legacy) and not os.path.isfile(path):
        try:
            shutil.move(legacy, path)
        except OSError:
            pass
    return path


def _load_gemini_cumulative_payload() -> dict:
    """API_Payment ???? JSON ????????????????????"""
    path = _gemini_cumulative_json_path()
    default: dict = {
        "version": 1,
        "updated_at": None,
        "calls_total": 0,
        "prompt_total": 0,
        "candidates_total": 0,
        "thoughts_total": 0,
        "total_tokens_reported": 0,
        "estimated_cost_usd_total": 0.0,
        "by_model": {},
    }
    if not os.path.isfile(path):
        _gemini_buckets_ensure_structure(default)
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict) or int(data.get("version") or 0) != 1:
            _gemini_buckets_ensure_structure(default)
            return default
        data.setdefault("by_model", {})
        _gemini_buckets_ensure_structure(data)
        for k in (
            "calls_total",
            "prompt_total",
            "candidates_total",
            "thoughts_total",
            "total_tokens_reported",
        ):
            data[k] = int(data.get(k) or 0)
        data["estimated_cost_usd_total"] = float(data.get("estimated_cost_usd_total") or 0.0)
        return data
    except Exception:
        _gemini_buckets_ensure_structure(default)
        return default


def _save_gemini_cumulative_payload(data: dict) -> None:
    path = _gemini_cumulative_json_path()
    try:
        os.makedirs(api_payment_dir, exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except OSError as ex:
        logging.debug("Gemini ?? JSON ??????: %s", ex)


def _gemini_buckets_ensure_structure(data: dict) -> None:
    """?? JSON ???????????????????? v1 ??????????"""
    b = data.setdefault("buckets", {})
    if not isinstance(b, dict):
        b = {}
        data["buckets"] = b
    for sub in ("by_year", "by_month", "by_week", "by_day", "by_hour"):
        x = b.setdefault(sub, {})
        if not isinstance(x, dict):
            b[sub] = {}
    b.setdefault(
        "timezone_note",
        "period_key ? PC ???????datetime.now?????? PC ???????????????",
    )


def _gemini_time_bucket_keys(dt: datetime) -> tuple[str, str, str, str, str]:
    """????ISO????? ???????????????????????"""
    iy, iw, _ = dt.isocalendar()
    y = dt.strftime("%Y")
    ym = dt.strftime("%Y-%m")
    wk = f"{iy}-W{iw:02d}"
    d = dt.strftime("%Y-%m-%d")
    h = dt.strftime("%Y-%m-%dT%H")
    return y, ym, wk, d, h


def _gemini_bucket_add_one_call(
    buckets_root: dict,
    pt: int,
    ct: int,
    th: int,
    tt: int,
    inc_usd: float | None,
    *,
    when: datetime | None = None,
) -> None:
    """1 ?? API ?????????????????????????"""
    dt = when or datetime.now()
    y, ym, wk, d, h = _gemini_time_bucket_keys(dt)
    pairs = (
        ("by_year", y),
        ("by_month", ym),
        ("by_week", wk),
        ("by_day", d),
        ("by_hour", h),
    )
    for sub, pk in pairs:
        subd = buckets_root.setdefault(sub, {})
        ent = subd.setdefault(
            pk,
            {
                "calls": 0,
                "prompt": 0,
                "candidates": 0,
                "thoughts": 0,
                "total_tokens": 0,
                "estimated_cost_usd": 0.0,
            },
        )
        ent["calls"] = int(ent.get("calls") or 0) + 1
        ent["prompt"] = int(ent.get("prompt") or 0) + pt
        ent["candidates"] = int(ent.get("candidates") or 0) + ct
        ent["thoughts"] = int(ent.get("thoughts") or 0) + th
        ent["total_tokens"] = int(ent.get("total_tokens") or 0) + tt
        if inc_usd is not None:
            ent["estimated_cost_usd"] = float(
                ent.get("estimated_cost_usd") or 0.0
            ) + float(inc_usd)


def _append_gemini_cumulative_one_call(
    model_id: str, pt: int, ct: int, th: int, tt: int
) -> None:
    """1 ?? API ?????? JSON ????????????????????"""
    mid = str(model_id).strip()
    data = _load_gemini_cumulative_payload()
    data["calls_total"] = int(data["calls_total"]) + 1
    data["prompt_total"] = int(data["prompt_total"]) + pt
    data["candidates_total"] = int(data["candidates_total"]) + ct
    data["thoughts_total"] = int(data["thoughts_total"]) + th
    data["total_tokens_reported"] = int(data["total_tokens_reported"]) + tt
    bm: dict = data["by_model"]
    if mid not in bm or not isinstance(bm[mid], dict):
        bm[mid] = {
            "calls": 0,
            "prompt": 0,
            "candidates": 0,
            "thoughts": 0,
            "total": 0,
            "estimated_cost_usd": 0.0,
        }
    m = bm[mid]
    m["calls"] = int(m.get("calls") or 0) + 1
    m["prompt"] = int(m.get("prompt") or 0) + pt
    m["candidates"] = int(m.get("candidates") or 0) + ct
    m["thoughts"] = int(m.get("thoughts") or 0) + th
    m["total"] = int(m.get("total") or 0) + tt
    inc_usd = _gemini_estimate_cost_usd(mid, pt, ct, th)
    if inc_usd is not None:
        m["estimated_cost_usd"] = float(m.get("estimated_cost_usd") or 0.0) + float(inc_usd)
        data["estimated_cost_usd_total"] = float(
            data.get("estimated_cost_usd_total") or 0.0
        ) + float(inc_usd)
    _gemini_buckets_ensure_structure(data)
    _gemini_bucket_add_one_call(
        data["buckets"], pt, ct, th, tt, inc_usd, when=datetime.now()
    )
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    _save_gemini_cumulative_payload(data)


def record_gemini_response_usage(res, model_id: str) -> None:
    """generate_content ????? usage_metadata ?????????????? JSON??"""
    global _gemini_usage_session
    if res is None or not str(model_id or "").strip():
        return
    um = getattr(res, "usage_metadata", None)
    if um is None:
        return

    def _iv(name: str) -> int:
        v = getattr(um, name, None)
        try:
            return int(v) if v is not None else 0
        except (TypeError, ValueError):
            return 0

    pt = _iv("prompt_token_count")
    ct = _iv("candidates_token_count")
    tt = _iv("total_token_count")
    th = _iv("thoughts_token_count")
    if tt <= 0 and (pt > 0 or ct > 0 or th > 0):
        tt = pt + ct + th
    mid = str(model_id).strip()
    b = _gemini_usage_session.setdefault(
        mid,
        {"prompt": 0, "candidates": 0, "total": 0, "thoughts": 0, "calls": 0},
    )
    b["prompt"] += pt
    b["candidates"] += ct
    b["total"] += tt
    b["thoughts"] += th
    b["calls"] += 1
    try:
        _append_gemini_cumulative_one_call(mid, pt, ct, th, tt)
    except Exception as ex:
        logging.debug("Gemini ????????????: %s", ex)


def _gemini_estimate_cost_usd(
    model_id: str, prompt_tok: int, cand_tok: int, thoughts_tok: int
) -> float | None:
    m = str(model_id).strip().lower()
    rin, rout = None, None
    if "flash" in m:
        rin, rout = _GEMINI_FLASH_IN_PER_M, _GEMINI_FLASH_OUT_PER_M
    elif "pro" in m:
        # ???????????????????
        rin, rout = 1.25, 5.0
    if rin is None:
        return None
    out_equiv = cand_tok + thoughts_tok
    return (prompt_tok / 1_000_000.0) * rin + (out_equiv / 1_000_000.0) * rout


def _gemini_daily_trend_series(
    cum: dict, *, max_days: int | None = None
) -> tuple[list[str], list[float], str] | None:
    """?? JSON ? by_day ??????????????????????? None?"""
    lim = GEMINI_USAGE_CHART_MAX_DAYS if max_days is None else max_days
    b = cum.get("buckets")
    if not isinstance(b, dict):
        return None
    subd = b.get("by_day")
    if not isinstance(subd, dict) or not subd:
        return None
    keys = sorted(subd.keys())
    keys = keys[-max(1, lim) :]
    usds: list[float] = []
    calls: list[int] = []
    for pk in keys:
        ent = subd.get(pk)
        if isinstance(ent, dict):
            usds.append(float(ent.get("estimated_cost_usd") or 0.0))
            calls.append(int(ent.get("calls") or 0))
        else:
            usds.append(0.0)
            calls.append(0)
    use_calls = sum(usds) <= 0.0 and sum(calls) > 0
    series = [float(c) for c in calls] if use_calls else usds
    label = "?????" if use_calls else "??USD"
    return (keys, series, label)


def _gemini_daily_total_tokens_for_days(cum: dict, day_keys: list[str]) -> list[int]:
    """by_day ?????????total_tokens????? prompt+candidates+thoughts?????"""
    b = cum.get("buckets")
    if not isinstance(b, dict):
        return [0] * len(day_keys)
    subd = b.get("by_day")
    if not isinstance(subd, dict):
        return [0] * len(day_keys)
    out: list[int] = []
    for pk in day_keys:
        ent = subd.get(pk)
        if not isinstance(ent, dict):
            out.append(0)
            continue
        tt = int(ent.get("total_tokens") or 0)
        if tt <= 0:
            tt = (
                int(ent.get("prompt") or 0)
                + int(ent.get("candidates") or 0)
                + int(ent.get("thoughts") or 0)
            )
        out.append(tt)
    return out


def _gemini_usage_trend_caption_lines(cum: dict) -> list[str]:
    """???????????? CSV ?????ASCII ??????????????"""
    ser = _gemini_daily_trend_series(cum)
    if ser is None:
        return []
    keys, _, label = ser
    b = cum.get("buckets")
    lines = [
        "?????????????: Q?R ???????: S?T ???????????????",
        f"  ??1: ?? {label}?{keys[0]} ? {keys[-1]}?",
        "  ??2: ?? ???????API ?? total ????????",
        f"  ????????????: log\\{GEMINI_USAGE_BUCKETS_CSV_FILE}?Excel ??????",
    ]
    if isinstance(b, dict):
        note = b.get("timezone_note")
        if note:
            lines.append(f"  ?{note}?")
    return lines


def _gemini_resolve_main_sheet_xlwings(book) -> object | None:
    """xlwings Book ?????????????????? None?"""
    for name in ("???", "???_", "Main"):
        try:
            return book.sheets[name]
        except Exception:
            continue
    try:
        for sht in book.sheets:
            try:
                if "???" in str(sht.name):
                    return sht
            except Exception:
                continue
    except Exception:
        pass
    return None


def _strip_gemini_usage_charts_xlwings(ws) -> None:
    """???????????????????????????????"""
    managed_names = (
        GEMINI_USAGE_XLW_CHART_NAME,
        GEMINI_USAGE_XLW_CHART_TOKENS_NAME,
    )
    title_markers = (
        "Gemini API ????",
        "Gemini API ??????",
    )
    try:
        charts_iter = list(ws.charts)
    except Exception:
        return
    for ch in charts_iter:
        try:
            if str(getattr(ch, "name", "") or "") in managed_names:
                ch.delete()
                continue
        except Exception:
            pass
        try:
            ca = ch.api
            if bool(ca.HasTitle):
                cap = getattr(ca.ChartTitle, "Caption", None)
                txt = getattr(ca.ChartTitle, "Text", None)
                title_s = str(cap or txt or "")
                for mk in title_markers:
                    if mk in title_s:
                        ch.delete()
                        break
        except Exception:
            pass


def _apply_main_sheet_gemini_usage_chart_xlwings(ws, cum: dict) -> None:
    """???????? Q?R?S?T ??????????? 2 ??????xlwings??"""
    hr = GEMINI_USAGE_CHART_HEADER_ROW
    cdt = GEMINI_USAGE_CHART_COL_DATE
    cvl = GEMINI_USAGE_CHART_COL_VALUE
    cts = GEMINI_USAGE_CHART_COL_TOK_DATE
    ctv = GEMINI_USAGE_CHART_COL_TOK_VALUE
    nclear = GEMINI_USAGE_CHART_CLEAR_ROWS
    try:
        block = ws.range((hr, cdt), (hr + nclear - 1, ctv))
        block.clear_contents()
    except Exception:
        for i in range(nclear):
            r = hr + i
            for c in (cdt, cvl, cts, ctv):
                try:
                    ws.range((r, c)).clear_contents()
                except Exception:
                    pass

    _strip_gemini_usage_charts_xlwings(ws)
    ser = _gemini_daily_trend_series(cum)
    if ser is None:
        return
    day_keys, values, val_label = ser
    n = len(day_keys)
    if n <= 0:
        return

    ws.range((hr, cdt)).value = "??"
    ws.range((hr, cvl)).value = val_label
    for i, (dk, val) in enumerate(zip(day_keys, values)):
        r = hr + 1 + i
        ws.range((r, cdt)).value = dk
        ws.range((r, cvl)).value = val
    try:
        vrng = ws.range((hr + 1, cvl), (hr + n, cvl))
        vrng.number_format = "0.000000" if val_label == "??USD" else "0"
    except Exception:
        pass

    try:
        anchor = ws.range(GEMINI_USAGE_CHART_ANCHOR_CELL)
        left = float(anchor.left)
        top = float(anchor.top)
    except Exception:
        left, top = 0.0, 0.0
    chart = ws.charts.add(left=left, top=top, width=410, height=220)
    try:
        chart.name = GEMINI_USAGE_XLW_CHART_NAME
    except Exception:
        pass
    data_rng = ws.range((hr, cdt), (hr + n, cvl))
    chart.set_source_data(data_rng)
    try:
        chart.chart_type = "line"
    except Exception:
        try:
            chart.api.ChartType = 4
        except Exception:
            pass
    try:
        ca = chart.api
        ca.HasTitle = True
        ca.ChartTitle.Text = "Gemini API ????"
        ca.HasLegend = False
    except Exception:
        pass

    tok_vals = _gemini_daily_total_tokens_for_days(cum, day_keys)
    if not tok_vals or max(tok_vals) <= 0:
        return

    tok_label = "??????"
    ws.range((hr, cts)).value = "??"
    ws.range((hr, ctv)).value = tok_label
    for i, dk in enumerate(day_keys):
        r = hr + 1 + i
        ws.range((r, cts)).value = dk
        ws.range((r, ctv)).value = int(tok_vals[i])
    try:
        ws.range((hr + 1, ctv), (hr + n, ctv)).number_format = "#,##0"
    except Exception:
        pass

    try:
        anchor2 = ws.range(GEMINI_USAGE_CHART_TOKENS_ANCHOR_CELL)
        left2 = float(anchor2.left)
        top2 = float(anchor2.top)
    except Exception:
        left2, top2 = left + 420.0, top
    chart2 = ws.charts.add(left=left2, top=top2, width=410, height=220)
    try:
        chart2.name = GEMINI_USAGE_XLW_CHART_TOKENS_NAME
    except Exception:
        pass
    data_rng2 = ws.range((hr, cts), (hr + n, ctv))
    chart2.set_source_data(data_rng2)
    try:
        chart2.chart_type = "line"
    except Exception:
        try:
            chart2.api.ChartType = 4
        except Exception:
            pass
    try:
        ca2 = chart2.api
        ca2.HasTitle = True
        ca2.ChartTitle.Text = "Gemini API ??????"
        ca2.HasLegend = False
    except Exception:
        pass


def _write_main_sheet_gemini_usage_via_xlwings(
    macro_wb_path: str, text: str, log_prefix: str
) -> bool:
    """Excel ???????????????? P ??Q?T?????????2??? xlwings ????? Save?"""
    attached = _xlwings_attach_open_macro_workbook(macro_wb_path, log_prefix)
    if attached is None:
        logging.info(
            "%s: xlwings ????????????????? AI ?????????????",
            log_prefix,
        )
        return False
    xw_book, info = attached
    ok = False
    try:
        try:
            xw_book.app.display_alerts = False
        except Exception:
            pass
        ws_main = _gemini_resolve_main_sheet_xlwings(xw_book)
        if ws_main is None:
            logging.info(
                "%s: ??????????? xlwings ?? AI ?????????????",
                log_prefix,
            )
            return False
        start_r, col_p, clear_n = 16, 16, 120
        _perf_snap = _xlwings_app_save_perf_state_push(xw_book.app)
        try:
            p_rng = ws_main.range((start_r, col_p)).resize(clear_n, 1)
            p_rng.clear_contents()
            lines_list = text.split("\n") if (text or "").strip() else []
            p_vals = [
                [lines_list[i] if i < len(lines_list) else None]
                for i in range(clear_n)
            ]
            p_rng.value = p_vals
            try:
                p_rng.api.WrapText = True
                p_rng.api.VerticalAlignment = -4160
            except Exception:
                pass
            _apply_main_sheet_gemini_usage_chart_xlwings(
                ws_main, _load_gemini_cumulative_payload()
            )
            xw_book.save()
            ok = True
            logging.info(
                "%s: ?????? P%d ???Gemini ????????/?????????? xlwings ????????",
                log_prefix,
                start_r,
            )
        finally:
            _xlwings_app_save_perf_state_pop(xw_book.app, _perf_snap)
    except Exception as ex:
        logging.warning(
            "%s: ??? AI ???? xlwings ?????: %s", log_prefix, ex
        )
        ok = False
    finally:
        _xlwings_release_book_after_mutation(xw_book, info, ok)
    return ok


def _gemini_kv_table_lines(title: str, rows: list[tuple[str, str]]) -> list[str]:
    """????????? 2 ?????????????????"""
    out = [title]
    if not rows:
        return out
    lw = min(22, max(len(a) for a, _ in rows))
    sep = "  " + ("?" * (lw + 2 + 28))
    out.append(sep)
    for a, b in rows:
        out.append(f"  {a:<{lw}}  {b}")
    return out


def _export_gemini_buckets_csv_for_charts(cum: dict) -> None:
    """Excel ?????????????? CSV ? log ??????"""
    b = cum.get("buckets")
    if not isinstance(b, dict):
        return
    mapping = (
        ("year", "by_year"),
        ("month", "by_month"),
        ("week_iso", "by_week"),
        ("day", "by_day"),
        ("hour", "by_hour"),
    )
    rows_out: list[dict[str, object]] = []
    for gran_label, sub in mapping:
        subd = b.get(sub)
        if not isinstance(subd, dict):
            continue
        for pk in sorted(subd.keys()):
            ent = subd.get(pk)
            if not isinstance(ent, dict):
                continue
            calls = int(ent.get("calls") or 0)
            pt = int(ent.get("prompt") or 0)
            cc = int(ent.get("candidates") or 0)
            th = int(ent.get("thoughts") or 0)
            tt = int(ent.get("total_tokens") or 0)
            usd = float(ent.get("estimated_cost_usd") or 0.0)
            rows_out.append(
                {
                    "granularity": gran_label,
                    "period_key": pk,
                    "calls": calls,
                    "prompt_tokens": pt,
                    "candidates_tokens": cc,
                    "thoughts_tokens": th,
                    "total_tokens": tt,
                    "estimated_cost_usd": round(usd, 8),
                    "estimated_cost_jpy": round(usd * GEMINI_JPY_PER_USD, 4),
                }
            )
    if not rows_out:
        return
    path = os.path.join(log_dir, GEMINI_USAGE_BUCKETS_CSV_FILE)
    fieldnames = [
        "granularity",
        "period_key",
        "calls",
        "prompt_tokens",
        "candidates_tokens",
        "thoughts_tokens",
        "total_tokens",
        "estimated_cost_usd",
        "estimated_cost_jpy",
    ]
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(rows_out)
    except OSError as ex:
        logging.debug("Gemini ???? CSV ??????: %s", ex)


def build_gemini_usage_summary_text() -> str:
    """???????????????????????????? JSON??"""
    cum = _load_gemini_cumulative_payload()
    ct_tot = int(cum.get("calls_total") or 0)
    if not _gemini_usage_session and ct_tot <= 0:
        return ""

    lines: list[str] = []
    ts = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    if _gemini_usage_session:
        lines.append(f"????: {ts}??????? Gemini API?")
        tot_calls = sum(b["calls"] for b in _gemini_usage_session.values())
        tot_p = sum(b["prompt"] for b in _gemini_usage_session.values())
        tot_c = sum(b["candidates"] for b in _gemini_usage_session.values())
        tot_th = sum(b["thoughts"] for b in _gemini_usage_session.values())
        tot_t = sum(b["total"] for b in _gemini_usage_session.values())
        sess_rows: list[tuple[str, str]] = [
            ("???", f"{tot_calls:,} ?"),
            ("??????", f"{tot_p:,}"),
            ("??????", f"{tot_c:,}"),
        ]
        if tot_th:
            sess_rows.append(("??????", f"{tot_th:,}"))
        sess_rows.append(("total ??", f"{tot_t:,}"))
        lines.extend(_gemini_kv_table_lines("??????", sess_rows))
        grand_usd = 0.0
        any_price = False
        for mid in sorted(_gemini_usage_session.keys()):
            b = _gemini_usage_session[mid]
            mrows: list[tuple[str, str]] = [
                ("???", mid),
                ("???", f"{b['calls']:,} ?"),
                ("??????", f"{b['prompt']:,}"),
                ("??????", f"{b['candidates']:,}"),
            ]
            if b.get("thoughts", 0):
                mrows.append(("??????", f"{b['thoughts']:,}"))
            mrows.append(("total_token_count", f"{b['total']:,}"))
            est = _gemini_estimate_cost_usd(
                mid, b["prompt"], b["candidates"], b.get("thoughts", 0)
            )
            if est is not None:
                any_price = True
                grand_usd += est
                mrows.append(("??USD", f"${est:.6f}"))
                mrows.append(
                    (
                        "??JPY",
                        f"�{est * GEMINI_JPY_PER_USD:.2f}?{GEMINI_JPY_PER_USD:.0f}?/USD?",
                    )
                )
            else:
                mrows.append(("????", "??????????"))
            lines.append("")
            lines.extend(_gemini_kv_table_lines(f"???????????", mrows))
        if any_price:
            lines.append("")
            lines.extend(
                _gemini_kv_table_lines(
                    "?????????????",
                    [
                        ("USD", f"${grand_usd:.6f}"),
                        (
                            "JPY",
                            f"�{grand_usd * GEMINI_JPY_PER_USD:.2f}?{GEMINI_JPY_PER_USD:.0f}?/USD?",
                        ),
                    ],
                )
            )
    else:
        lines.append(f"????: {ts}")
        lines.append("??????? Gemini API ??????????")
    lines.append("? ????? API ? usage_metadata ???????")
    lines.append(
        "? USD ????????????????????? Google ?????????????"
    )
    lines.append(
        "? ? API ????????????????????????? JSON ??????????"
    )

    if ct_tot > 0:
        lines.append("")
        cum_hdr = (
            f"????{GEMINI_USAGE_CUMULATIVE_JSON_FILE} "
            "?API_Payment ?????????????"
        )
        pt0 = int(cum.get("prompt_total") or 0)
        cc0 = int(cum.get("candidates_total") or 0)
        th0 = int(cum.get("thoughts_total") or 0)
        tt0 = int(cum.get("total_tokens_reported") or 0)
        cum_rows: list[tuple[str, str]] = [
            ("????", str(cum.get("updated_at") or "�")),
            ("???", f"{ct_tot:,} ?"),
            ("??????", f"{pt0:,}"),
            ("??????", f"{cc0:,}"),
        ]
        if th0:
            cum_rows.append(("??????", f"{th0:,}"))
        cum_rows.append(("total ??", f"{tt0:,}"))
        usd_all = float(cum.get("estimated_cost_usd_total") or 0.0)
        if usd_all > 0:
            cum_rows.append(("??USD ??", f"${usd_all:.6f}"))
            cum_rows.append(
                (
                    "??JPY ??",
                    f"�{usd_all * GEMINI_JPY_PER_USD:.2f}?{GEMINI_JPY_PER_USD:.0f}?/USD?",
                )
            )
        lines.extend(_gemini_kv_table_lines(cum_hdr, cum_rows))
        bm = cum.get("by_model") or {}
        if isinstance(bm, dict) and bm:
            for mid in sorted(bm.keys()):
                m = bm[mid]
                if not isinstance(m, dict):
                    continue
                mrows2: list[tuple[str, str]] = [
                    ("???", mid),
                    ("???", f"{int(m.get('calls') or 0):,} ?"),
                    (
                        "?? / ??",
                        f"{int(m.get('prompt') or 0):,} / {int(m.get('candidates') or 0):,}",
                    ),
                ]
                if int(m.get("thoughts") or 0):
                    mrows2.append(("??????", f"{int(m.get('thoughts') or 0):,}"))
                mud = float(m.get("estimated_cost_usd") or 0.0)
                if mud > 0:
                    mrows2.append(("??USD ??", f"${mud:.6f}"))
                    mrows2.append(
                        ("??JPY ??", f"�{mud * GEMINI_JPY_PER_USD:.2f}")
                    )
                lines.append("")
                lines.extend(_gemini_kv_table_lines("?????????", mrows2))
        trend = _gemini_usage_trend_caption_lines(cum)
        if trend:
            lines.append("")
            lines.extend(trend)
    return "\n".join(lines)


def write_main_sheet_gemini_usage_summary(wb_path: str, log_prefix: str) -> None:
    """Gemini ?????? log ????xlwings ???? P ???????????????????????"""
    text = build_gemini_usage_summary_text()
    path = os.path.join(log_dir, GEMINI_USAGE_SUMMARY_FOR_MAIN_FILE)
    xw_ok = False
    if wb_path and os.path.isfile(wb_path):
        try:
            xw_ok = _write_main_sheet_gemini_usage_via_xlwings(
                wb_path, text, log_prefix
            )
        except Exception as ex:
            logging.warning(
                "%s: AI ???? xlwings ???????: %s", log_prefix, ex
            )
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
    except OSError:
        pass
    try:
        cum2 = _load_gemini_cumulative_payload()
        if int(cum2.get("calls_total") or 0) > 0:
            _export_gemini_buckets_csv_for_charts(cum2)
    except Exception as ex:
        logging.debug("Gemini ???? CSV ?????????: %s", ex)
    if xw_ok:
        return
    if text.strip():
        logging.info(
            "%s: ??? P ?????? xlwings ????????????"
            " %s ????? ? ??????????_Gemini??????P?????? P ??????????",
            log_prefix,
            path,
        )
    else:
        logging.info(
            "%s: Gemini ???: ?????? %s ????",
            log_prefix,
            path,
        )


def _try_write_main_sheet_gemini_usage_summary(phase: str) -> None:
    try:
        write_main_sheet_gemini_usage_summary(TASKS_INPUT_WORKBOOK, phase)
    except Exception as ex:
        logging.warning(
            "%s: ???????? AI ????????????????: %s", phase, ex
        )


def _plan_sheet_write_global_parse_block_to_ws(
    ws,
    global_priority_override: dict,
    when_str: str,
) -> None:
    """????????????_???????????? AX:AY ????????????????"""
    gpo = global_priority_override or {}
    lc = PLAN_SHEET_GLOBAL_PARSE_LABEL_COL
    vc = PLAN_SHEET_GLOBAL_PARSE_VALUE_COL
    max_r = PLAN_SHEET_GLOBAL_PARSE_MAX_ROWS
    for i in range(max_r):
        ws.cell(row=1 + i, column=lc, value=None)
        ws.cell(row=1 + i, column=vc, value=None)
    align_top = Alignment(wrap_text=True, vertical="top")
    pairs: list[tuple[str, str]] = [
        ("?????????????", "??????2?????"),
        (
            "?????????",
            "??????????????????????????????????"
            "??AX?AY???????????????????????????????"
            "?????????????????",
        ),
        ("??????", (when_str or "").strip() or "?"),
        (
            "?????",
            ", ".join(str(x) for x in (gpo.get("factory_closure_dates") or []))
            if gpo.get("factory_closure_dates")
            else "????",
        ),
        (
            "????????",
            "??" if gpo.get("ignore_skill_requirements") else "???",
        ),
        (
            "need??1??",
            "??" if gpo.get("ignore_need_minimum") else "???",
        ),
        (
            "???????",
            "??" if gpo.get("abolish_all_scheduling_limits") else "???",
        ),
        (
            "?????OP??",
            json.dumps(gpo.get("task_preferred_operators") or {}, ensure_ascii=False)
            if gpo.get("task_preferred_operators")
            else "????",
        ),
        (
            "??�???????",
            json.dumps(
                gpo.get("global_day_process_operator_rules") or [],
                ensure_ascii=False,
            )
            if gpo.get("global_day_process_operator_rules")
            else "????",
        ),
        (
            "??????????",
            json.dumps(gpo.get("global_speed_rules") or [], ensure_ascii=False)
            if gpo.get("global_speed_rules")
            else "????",
        ),
        (
            "?????(AI)",
            str(gpo.get("scheduler_notes_ja") or "").strip() or "????",
        ),
        (
            "AI??",
            str(gpo.get("interpretation_ja") or "").strip() or "????",
        ),
    ]
    for i, (lab, val) in enumerate(pairs):
        if i >= max_r:
            break
        c1 = ws.cell(row=1 + i, column=lc, value=lab)
        c2 = ws.cell(row=1 + i, column=vc, value=val)
        c1.alignment = align_top
        c2.alignment = align_top


def write_plan_sheet_global_comment_parse_block(
    wb_path: str,
    sheet_name: str,
    global_priority_override: dict,
    *,
    when_str: str,
    log_prefix: str = "??2",
) -> bool:
    """
    ?????_???????????????AX:AY???????????????????????
    ??????????????????????????????????????????????
    Excel ??????????????????????????? openpyxl ???????
    """
    if not wb_path or not os.path.isfile(wb_path):
        return False
    gpo = global_priority_override or {}
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = None
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "%s: ?????%s?????? openpyxl ???????????????????????????",
            log_prefix,
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return False
    try:
        wb = load_workbook(
            wb_path, keep_vba=keep_vba, read_only=False, data_only=False
        )
    except Exception as ex:
        logging.info(
            "%s: ???????????????????????????????: %s",
            log_prefix,
            ex,
        )
        return False
    try:
        if sheet_name not in wb.sheetnames:
            logging.info(
                "%s: ??? '%s' ?????????????????????????",
                log_prefix,
                sheet_name,
            )
            return False
        ws = wb[sheet_name]
        _plan_sheet_write_global_parse_block_to_ws(ws, gpo, when_str)
        lc = PLAN_SHEET_GLOBAL_PARSE_LABEL_COL
        vc = PLAN_SHEET_GLOBAL_PARSE_VALUE_COL
        wb.save(wb_path)
        logging.info(
            "%s: ?%s?%s:%s ?????????????????????",
            log_prefix,
            sheet_name,
            get_column_letter(lc),
            get_column_letter(vc),
        )
        return True
    except OSError as ex:
        logging.warning(
            "%s: ?????????????????????????????Excel ????????: %s",
            log_prefix,
            ex,
        )
        return False
    except Exception as ex:
        logging.warning(
            "%s: ??????????????????????: %s", log_prefix, ex
        )
        return False
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def _try_write_plan_sheet_global_comment_parse_block(
    global_priority_override: dict,
    when_str: str,
) -> None:
    try:
        write_plan_sheet_global_comment_parse_block(
            TASKS_INPUT_WORKBOOK,
            PLAN_INPUT_SHEET_NAME,
            global_priority_override,
            when_str=when_str,
            log_prefix="??2",
        )
    except Exception as ex:
        logging.warning(
            "??2: ?????????????????????????????: %s",
            ex,
        )


def _try_write_plan_input_global_parse_and_conflicts_one_save(
    global_priority_override: dict,
    when_str: str,
    num_data_rows: int,
    conflicts_by_row,
) -> None:
    try:
        write_plan_sheet_global_parse_and_conflict_styles_one_io(
            TASKS_INPUT_WORKBOOK,
            PLAN_INPUT_SHEET_NAME,
            global_priority_override,
            when_str=when_str,
            num_data_rows=num_data_rows,
            conflicts_by_row=conflicts_by_row,
            log_prefix="??2",
        )
    except Exception as ex:
        logging.warning(
            "??2: ????????????????????1???????????: %s",
            ex,
        )


def _log_task_special_ai_response(raw_text, parsed, extracted_json_str, prompt_text=None):
    """????_???? Gemini ???????????????JSON???????1????????"""
    path = os.path.join(log_dir, TASK_SPECIAL_AI_LAST_RESPONSE_FILE)
    try:
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            if prompt_text is not None and str(prompt_text).strip():
                f.write("=== Gemini ?????????????? ===\n")
                f.write(str(prompt_text).strip())
                f.write("\n\n")
            f.write("=== Gemini ????????????????? ===\n")
            f.write(raw_text or "")
            f.write(
                "\n\n=== AI ?????????????????????? JSON ??? ===\n"
                "????????????????????????????????????????\n"
            )
            f.write(extracted_json_str if extracted_json_str else "(????)")
            f.write("\n\n=== json.loads ????NO??? ===\n")
            if isinstance(parsed, dict):
                f.write(json.dumps(parsed, ensure_ascii=False, indent=2))
            else:
                f.write("(??????)")
        logging.info(
            "???????: ??????AI????? ? %s",
            path,
        )
    except OSError as ex:
        logging.warning("???????: AI???????????: %s", ex)
    if isinstance(parsed, dict) and parsed:
        logging.info(
            "???????: ???????NO: %s",
            ", ".join(sorted(parsed.keys(), key=lambda x: str(x))),
        )
        for tid_k in sorted(parsed.keys(), key=lambda x: str(x)):
            logging.info(
                "  ??NO [%s] AI???????: %s",
                tid_k,
                json.dumps(parsed[tid_k], ensure_ascii=False),
            )


def _parse_and_log_task_special_gemini_response(res, prompt_text=None):
    """
    API ?????? JSON ????????????????? None?
    ???????????????????????? JSON ??????????????
    """
    raw = _gemini_result_text(res)
    if raw:
        stripped = raw.strip()
        if stripped.startswith("{"):
            try:
                trial = json.loads(stripped)
                if isinstance(trial, dict):
                    _log_task_special_ai_response(raw, trial, stripped, prompt_text)
                    return trial
            except json.JSONDecodeError:
                pass
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if not match:
        _log_task_special_ai_response(raw, {}, None, prompt_text)
        logging.warning(
            "???????: AI???? JSON ??????????????????? 3000 ??:\n%s",
            (raw[:3000] if raw else "(?)"),
        )
        return None
    extracted = match.group(0)
    try:
        parsed = json.loads(extracted)
    except json.JSONDecodeError as je:
        _log_task_special_ai_response(raw, None, extracted, prompt_text)
        logging.warning("???????: JSON ?????: %s", je)
        return None
    if not isinstance(parsed, dict):
        _log_task_special_ai_response(raw, None, extracted, prompt_text)
        logging.warning("???????: ??????? JSON ??????????????")
        return None
    _log_task_special_ai_response(raw, parsed, extracted, prompt_text)
    return parsed


def analyze_task_special_remarks(tasks_df, reference_year=None, ai_sheet_sink: dict | None = None):
    """
    ?????_????????????_???? AI ?????????????????????????
    ???????????????????????API ????????????????
    ??OP?????????????????? preferred_operator ????????????????????????????
    json/ai_remarks_cache.json ? TTL AI_CACHE_TTL_SECONDS ??????????????????? API ???????
    ??NO????????????????????????????????????????????????????????

    ?????: ??NO -> ??????????????NO??????????????????????
      process_name, machine_name � ??????????????????????????????
      restrict_to_process_name, restrict_to_machine_name � ????????????NO?????????????
      ??? required_op, speed_override, task_efficiency, priority, start_date, start_time,
      target_completion_date, ship_by_date, preferred_operator ???
    """
    lines = _task_special_prompt_lines(tasks_df)
    if not lines:
        n_rows = len(tasks_df)
        n_rem_only = 0
        n_tid_raw = 0
        for _, row in tasks_df.iterrows():
            tid = planning_task_id_str_from_plan_row(row)
            rem = _cell_text_task_special_remark(row.get(PLAN_COL_SPECIAL_REMARK))
            if tid:
                n_tid_raw += 1
            if rem:
                n_rem_only += 1
        miss_col = PLAN_COL_SPECIAL_REMARK not in tasks_df.columns
        logging.warning(
            "???????: AI ????????????%s???%s??"
            "???=%s???NO????=%s??????????=%s?"
            "??2????????????????%s???????????????????",
            PLAN_COL_SPECIAL_REMARK,
            "???????" if miss_col else "??????????",
            n_rows,
            n_tid_raw,
            n_rem_only,
            PLAN_COL_SPECIAL_REMARK,
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["??????_AI_API"] = "???????????"
        return {}

    blob = "\n".join(sorted(lines))
    ref_y = int(reference_year) if reference_year is not None else date.today().year
    cache_fingerprint = f"{ref_y}\n{blob}"
    cache_key_input = f"{TASK_SPECIAL_CACHE_KEY_PREFIX}{cache_fingerprint}"
    cache_key = hashlib.sha256(cache_key_input.encode("utf-8")).hexdigest()
    ai_cache = load_ai_cache()
    cached_parsed = get_cached_ai_result(
        ai_cache, cache_key, content_key=cache_fingerprint
    )
    if cached_parsed is not None:
        logging.info(
            "???????: ?????????%s ?????=%s??Gemini ???????",
            len(lines),
            ref_y,
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["??????_AI_API"] = "???????????"
        out = copy.deepcopy(cached_parsed)
        if isinstance(out, dict):
            _repair_task_special_ai_wrong_top_level_keys(out, tasks_df)
        return out

    logging.info(
        "???????: ????????Gemini ? %s ??????????????=%s??",
        len(lines),
        ref_y,
    )

    if not API_KEY:
        logging.info("GEMINI_API_KEY ??????????????AI????????????")
        if ai_sheet_sink is not None:
            ai_sheet_sink["??????_AI_API"] = "???API??????"
        return {}

    prompt = f"""
???????????????Excel?????_??????????????????????????????????????????????

?????
1) ????????????????????????????? **????????????????**?????????????????????????**??????????????????????????**
2) ??????? **1?? JSON ????????**???? {{ ???? }} ????????????????????????
3) JSON ?????????????? **??NO???????????** ? **????** ??????**????**???????????????????: 20010 ??????????????????????????????????????????????????NO??????

???JSON???????????????????
? ??????
- ??: ??????????? **??NO?�?????** ????? **????**?????????????????????????????????????????
- ?: ???????
  (A) **JSON??????1?** � ????NO?????????? **1???** ????
  (B) **JSON??**???????????� ????NO???????????????? **??** ?????????????????????????

? process_name??????machine_name?????� **??**
- ???????????????? **????�?**?**????�?** ??? **??** ????????????????? ""??
- ??????????????

? restrict_to_process_name??????restrict_to_machine_name?????� **??**
- **?????????????????????????????????????** ?????
- **????????????????**????????????????? **??????** ??? **???? ""** ????
- ???????????? **????NO?????: ???????????????????** ???
- ?????????????????????Excel ???????????????????

? preferred_operator?????� ????**??**
- **????**: ????????????**??????????????OP???????????**?? **?????** ????????
  ?: ???????????????????????????OP?????????????????**?????????**???????????
- **???????????**: ????????????????????????**??** ?? `preferred_operator` ?????? **???????** ??????? **process_name / machine_name ???**??: `{{"process_name":"�","machine_name":"�","preferred_operator":"�"}}`??
- **????**: ??????? **????????1??**?????????????????????????????????????????????:???????????????? `"??"`?
- **???????????**: ?????????? **????** ???????NO?? `preferred_operator` ????? **??** ??????????????

? ?????????required_op, speed_override, task_efficiency, priority, start_date, start_time, target_completion_date, ship_by_date?
- ???? **???** ???????????????????????? **??**??????????

?????NO????????
??NO Y4-2 ??????????????????????4/5??????????????????????:
- process_name / machine_name ? **????????** ???????
- restrict_to_* ? **??????** ???**????????????** ???????????????????

?????????????
?4/5??4/5???????? **????** ????? **?? {ref_y} ?** ???YYYY-MM-DD ????

???????????????
- process_name, machine_name: ?????????????????
- restrict_to_process_name, restrict_to_machine_name: ????????????
- preferred_operator: ????????????
- required_op: ????
- speed_override: ????m/??
- task_efficiency: 0?1
- priority: ?????????????
- start_date: YYYY-MM-DD / start_time: HH:MM
- target_completion_date, ship_by_date: YYYY-MM-DD

???????
- ????????????????? priority ?????????????????????? target_completion_date ??? ship_by_date ?????
- ?????? **????** ? preferred_operator ??????????????????????
- ??????????????
- **???????????????????????**?restrict_to_* ?????????NO???????????????

??????????????????? JSON ??????
- ????????? **??** ???????????????? **process_name** ? **machine_name** ?????
- ????NO????????? **??** ????1???????????????????????????restrict ??????????
- ????OP?????????? **??? preferred_operator** ?????

???????????NO?????????????????
{{
  "W3-14": {{
    "process_name": "??",
    "machine_name": "???A",
    "preferred_operator": "??"
  }},
  "Y3-26": {{
    "process_name": "??????",
    "machine_name": "",
    "priority": 1,
    "ship_by_date": "{ref_y}-04-05",
    "target_completion_date": "{ref_y}-04-05"
  }},
  "Y4-2": {{
    "process_name": "????",
    "machine_name": "E1",
    "priority": 2,
    "restrict_to_process_name": "",
    "restrict_to_machine_name": ""
  }}
}}

?????????Excel ???????1????NO???????
{blob}
"""
    try:
        ppath = os.path.join(log_dir, "ai_task_special_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("???????: ?? Gemini ??????????? ? %s", ppath)
    except OSError as ex:
        logging.warning("???????: ?????????: %s", ex)

    client = genai.Client(api_key=API_KEY)
    try:
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        parsed = _parse_and_log_task_special_gemini_response(res, prompt_text=prompt)
        if parsed is not None:
            _repair_task_special_ai_wrong_top_level_keys(parsed, tasks_df)
            put_cached_ai_result(
                ai_cache, cache_key, parsed, content_key=cache_fingerprint
            )
            save_ai_cache(ai_cache)
            logging.info("???????: AI??????????")
            if ai_sheet_sink is not None:
                ai_sheet_sink["??????_AI_API"] = "??"
            return parsed
        if ai_sheet_sink is not None:
            ai_sheet_sink["??????_AI_API"] = "???JSON?????"
        return {}
    except Exception as e:
        err_text = str(e)
        is_quota = ("429" in err_text) or ("RESOURCE_EXHAUSTED" in err_text)
        is_unavailable = ("503" in err_text) or ("UNAVAILABLE" in err_text)
        retry_sec = extract_retry_seconds(err_text) if is_quota else None
        if is_quota and retry_sec is not None:
            wait_sec = min(max(retry_sec, 1.0), 90.0)
            logging.warning(f"??????? AI 429?{wait_sec:.1f}????????????")
            time_module.sleep(wait_sec)
            try:
                res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
                record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                parsed = _parse_and_log_task_special_gemini_response(res, prompt_text=prompt)
                if parsed is not None:
                    _repair_task_special_ai_wrong_top_level_keys(parsed, tasks_df)
                    put_cached_ai_result(
                        ai_cache, cache_key, parsed, content_key=cache_fingerprint
                    )
                    save_ai_cache(ai_cache)
                    if ai_sheet_sink is not None:
                        ai_sheet_sink["??????_AI_API"] = "???429?????"
                    return parsed
            except Exception as e2:
                logging.warning(f"??????? AI ?????: {e2}")
        elif is_unavailable:
            wait_sec = 8.0
            logging.warning(
                f"??????? AI 503/UNAVAILABLE?{wait_sec:.1f}????????????"
            )
            time_module.sleep(wait_sec)
            try:
                res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
                record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                parsed = _parse_and_log_task_special_gemini_response(res, prompt_text=prompt)
                if parsed is not None:
                    _repair_task_special_ai_wrong_top_level_keys(parsed, tasks_df)
                    put_cached_ai_result(
                        ai_cache, cache_key, parsed, content_key=cache_fingerprint
                    )
                    save_ai_cache(ai_cache)
                    logging.info("???????: AI??????????????")
                    if ai_sheet_sink is not None:
                        ai_sheet_sink["??????_AI_API"] = "???503?????"
                    return parsed
                logging.warning("??????? AI 503???: JSON ??????????")
            except Exception as e2:
                logging.warning(f"??????? AI 503?????: {e2}")
        else:
            logging.warning(f"??????? AI ???: {e}")
        logging.warning(
            "???????: AI????????????????????_??????/?????????????"
            "????????_????????_?????????????????????? AI ??????????????"
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["??????_AI_API"] = f"??: {e}"[:500]
        return {}


def _merge_preferred_operator_cell_and_ai(row, ai_for_tid):
    """Excel???OP_??????????? AI ? preferred_operator?"""
    ai = ai_for_tid if isinstance(ai_for_tid, dict) else {}
    v = row.get(PLAN_COL_PREFERRED_OP)
    if v is not None and not (isinstance(v, float) and pd.isna(v)):
        s = str(v).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            return s
    a = ai.get("preferred_operator")
    if a is not None:
        s = str(a).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            return s
    return ""


def _global_override_preferred_operator_for_task(tpref, task_id) -> str | None:
    """
    ????????????? task_preferred_operators?
    ?????NO???????????????
    """
    if not isinstance(tpref, dict) or not tpref:
        return None
    tid = str(task_id).strip()
    if not tid:
        return None
    tlo = tid.lower()
    for k, v in tpref.items():
        if str(k).strip().lower() != tlo:
            continue
        s = str(v).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            return s
        return None
    return None


def _merge_task_row_with_ai(
    row, ai_for_tid, *, allow_ai_dispatch_priority_from_remark: bool = True
):
    """
    ?????????_?????????_???????????????????????? AI ???
    allow_ai_dispatch_priority_from_remark ? False ????AI ? required_op / task_efficiency / priority /
    start_date / start_time ???????????????????????
    """
    ai = ai_for_tid if isinstance(ai_for_tid, dict) else {}

    def first_float_pos_cell_or_ai(cell, ai_key):
        v = parse_float_safe(row.get(cell), None)
        if v is not None and (not isinstance(v, float) or not pd.isna(v)) and float(v) > 0:
            return float(v)
        a = ai.get(ai_key)
        try:
            if a is not None and float(a) > 0:
                return float(a)
        except (TypeError, ValueError):
            pass
        return None

    if allow_ai_dispatch_priority_from_remark:
        req_op = parse_optional_int(ai.get("required_op"))
    else:
        req_op = None
    if req_op is not None and req_op < 1:
        req_op = None

    if allow_ai_dispatch_priority_from_remark:
        te = None
        a = ai.get("task_efficiency")
        try:
            if a is not None and float(a) > 0:
                te = float(a)
        except (TypeError, ValueError):
            te = None
        if te is None or te <= 0:
            te = 1.0
    else:
        te = 1.0

    if allow_ai_dispatch_priority_from_remark:
        pri = parse_optional_int(ai.get("priority"))
    else:
        pri = None
    if pri is None:
        pri = 999

    st_date = None
    if allow_ai_dispatch_priority_from_remark and ai.get("start_date"):
        st_date = parse_optional_date(ai.get("start_date"))

    st_time = None
    if allow_ai_dispatch_priority_from_remark and ai.get("start_time"):
        st_time = parse_time_str(str(ai.get("start_time")), None)

    speed_ov = first_float_pos_cell_or_ai(PLAN_COL_SPEED_OVERRIDE, "speed_override")

    return req_op, speed_ov, te, pri, st_date, st_time, ai


def _plan_row_cell_nonempty(row, col_name):
    v = row.get(col_name)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none"):
        return False
    return True


def _ai_float_for_conflict(ai, key):
    if not ai or ai.get(key) is None:
        return None
    try:
        f = float(ai.get(key))
        return f if f > 0 else None
    except (TypeError, ValueError):
        return None


def detect_planning_remark_ai_conflicts(row, ai_for_tid):
    """
    ????_????? AI ???????????????????????????
    ???AI??????????????
    """
    remark = str(row.get(PLAN_COL_SPECIAL_REMARK, "") or "").strip()
    if not remark or remark.lower() in ("nan", "none"):
        return set()
    ai = ai_for_tid if isinstance(ai_for_tid, dict) else {}
    if not ai:
        return set()
    out = set()

    if _plan_row_cell_nonempty(row, PLAN_COL_SPEED_OVERRIDE):
        cv = parse_float_safe(row.get(PLAN_COL_SPEED_OVERRIDE), None)
        if cv is not None and cv > 0:
            av = _ai_float_for_conflict(ai, "speed_override")
            if av is not None and abs(cv - av) > 1e-5:
                out.add(PLAN_COL_SPEED_OVERRIDE)

    if _plan_row_cell_nonempty(row, PLAN_COL_PREFERRED_OP):
        cv = _normalize_person_name_for_match(row.get(PLAN_COL_PREFERRED_OP))
        av = _normalize_person_name_for_match(ai.get("preferred_operator"))
        if cv and av and cv != av:
            out.add(PLAN_COL_PREFERRED_OP)

    if out:
        out.add(PLAN_COL_SPECIAL_REMARK)
    return out


def collect_planning_conflicts_by_excel_row(tasks_df, ai_by_tid):
    """Excel ???(1????????=1??) -> ???????????"""
    res = {}
    for i, (_, row) in enumerate(tasks_df.iterrows()):
        if _plan_row_exclude_from_assignment(row):
            continue
        ai_one = _ai_task_special_entry_for_row(ai_by_tid, row)
        cset = detect_planning_remark_ai_conflicts(row, ai_one)
        if cset:
            res[i + 2] = cset
    return res


def _plan_sheet_apply_conflict_styles_to_ws(ws, num_data_rows: int, conflicts_by_row) -> None:
    """???????????????????????????????????????????????"""
    header_map = {}
    for col_idx in range(1, ws.max_column + 1):
        v = ws.cell(1, col_idx).value
        if v is not None:
            header_map[str(v).strip()] = col_idx

    last_row = max(2, 1 + int(num_data_rows))
    clear_fill = PatternFill(fill_type=None)
    yellow_input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    conflict_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    conflict_font = Font(color="FFFFFF", bold=True)

    for r in range(2, last_row + 1):
        for name in PLAN_CONFLICT_STYLABLE_COLS:
            ci = header_map.get(name)
            if not ci:
                continue
            cell = ws.cell(row=r, column=ci)
            if name == PLAN_COL_AI_PARSE:
                cell.fill = clear_fill
            else:
                cell.fill = yellow_input_fill
            # ????????????????????????????

    for r, colnames in conflicts_by_row.items():
        if r < 2:
            continue
        for name in colnames:
            ci = header_map.get(name)
            if not ci:
                continue
            cell = ws.cell(row=r, column=ci)
            cell.fill = conflict_fill
            cell.font = conflict_font


def write_plan_sheet_global_parse_and_conflict_styles_one_io(
    wb_path: str,
    sheet_name: str,
    global_priority_override: dict,
    *,
    when_str: str,
    num_data_rows: int,
    conflicts_by_row,
    log_prefix: str = "??2",
) -> bool:
    """
    ??2??: ????????????????AX:AY?????????? **1?? load/save** ??????
    ???????????2?????????.xlsm ???????????????????
    """
    if not wb_path or not os.path.isfile(wb_path):
        return False
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "%s: ?????%s?????? openpyxl ???????????????????????",
            log_prefix,
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return False
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = None
    try:
        wb = load_workbook(
            wb_path, keep_vba=keep_vba, read_only=False, data_only=False
        )
    except Exception as ex:
        logging.info(
            "%s: ?????????????????????: %s",
            log_prefix,
            ex,
        )
        return False
    try:
        if sheet_name not in wb.sheetnames:
            logging.info(
                "%s: ??? '%s' ???????????????????????",
                log_prefix,
                sheet_name,
            )
            return False
        ws = wb[sheet_name]
        _plan_sheet_write_global_parse_block_to_ws(ws, global_priority_override or {}, when_str)
        _plan_sheet_apply_conflict_styles_to_ws(ws, num_data_rows, conflicts_by_row or {})
        lc = PLAN_SHEET_GLOBAL_PARSE_LABEL_COL
        vc = PLAN_SHEET_GLOBAL_PARSE_VALUE_COL
        try:
            wb.save(wb_path)
        except OSError as e:
            write_planning_conflict_highlight_sidecar(
                sheet_name, num_data_rows, conflicts_by_row or {}
            )
            logging.warning(
                "%s: ???????????????Excel ?????????"
                " ???????? '%s' ?????????????????????????????? (%s)",
                log_prefix,
                _planning_conflict_sidecar_path(),
                e,
            )
            return False
        _remove_planning_conflict_sidecar_safe()
        _n_conf = len(conflicts_by_row) if conflicts_by_row else 0
        if _n_conf:
            logging.info(
                "%s: ?%s?%s:%s ??????????????"
                "????_??????? %s ?????????????????",
                log_prefix,
                sheet_name,
                get_column_letter(lc),
                get_column_letter(vc),
                _n_conf,
            )
        else:
            logging.info(
                "%s: ?%s?%s:%s ????????????????????????",
                log_prefix,
                sheet_name,
                get_column_letter(lc),
                get_column_letter(vc),
            )
        return True
    except OSError as ex:
        logging.warning(
            "%s: ?????????? OSError: %s",
            log_prefix,
            ex,
        )
        return False
    except Exception as ex:
        logging.warning(
            "%s: ??????????????????????????: %s",
            log_prefix,
            ex,
        )
        return False
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def apply_planning_sheet_conflict_styles(wb_path, sheet_name, num_data_rows, conflicts_by_row):
    """
    ????_??????????????????????????????
    ???????????????1???????????????????????????
    AI????????????1??????????
    .xlsm ? keep_vba=True ??????
    """
    if not wb_path or not os.path.exists(wb_path):
        return
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "????: ?????%s?????? openpyxl ?????????????????",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = load_workbook(wb_path, keep_vba=keep_vba)
    try:
        if sheet_name not in wb.sheetnames:
            logging.warning(f"????: ??? '{sheet_name}' ?????????")
            return
        ws = wb[sheet_name]
        _plan_sheet_apply_conflict_styles_to_ws(ws, num_data_rows, conflicts_by_row)

        try:
            wb.save(wb_path)
        except OSError as e:
            write_planning_conflict_highlight_sidecar(sheet_name, num_data_rows, conflicts_by_row)
            logging.warning(
                "??????????????????????????????Excel ?????????????"
                " '%s' ????????????????????????????? (%s)",
                _planning_conflict_sidecar_path(),
                e,
            )
        else:
            _remove_planning_conflict_sidecar_safe()
            if conflicts_by_row:
                logging.info(
                    f"????_???????: {len(conflicts_by_row)} ?? '{sheet_name}' ???????????"
                )
    finally:
        wb.close()


def _ai_planning_target_due_date(ai_dict):
    """AI JSON ??????????????????1???????????????????????"""
    if not isinstance(ai_dict, dict):
        return None
    dates = []
    for k in ("target_completion_date", "ship_by_date", "latest_ship_date", "due_date"):
        d = parse_optional_date(ai_dict.get(k))
        if d is not None:
            dates.append(d)
    if not dates:
        return None
    return min(dates)


def _special_remark_implies_due_related_dispatch_priority(remark_raw: str) -> bool:
    """
    ????_???????????????????????????????????? True?
    ?????????????? True ?????AI ?????????????????????
    """
    if not remark_raw:
        return False
    s = str(remark_raw).strip()
    if not s or s.lower() in ("nan", "none"):
        return False
    n = unicodedata.normalize("NFKC", s)
    n_lower = n.casefold()
    needles = (
        "??",
        "????",
        "????",
        "????",
        "??",
        "??",
        "????",
        "??",
        "???",
        "??",
        "??",
        "???",
        "??",
        "??",
        "??",
        "??",
        "deadline",
        "??????",
        "???",
        "???",
        "??",
        "???",
        "????",
        "????",
        "????",
        "????",
        "????",
        "????",
        "????",
        "???",
        "????",
    )
    return any(w.casefold() in n_lower for w in needles)


def _task_id_same_machine_due_tiebreak_key(task_id) -> tuple:
    """
    ?????????????????????????
    Y3-24 ???????Y4-1-1 ?????????2???????????? - ????????????
    """
    s = str(task_id or "").strip()
    if not s:
        return (2, 10**9, "")
    parts = s.split("-", 1)
    if len(parts) < 2:
        return (2, 10**9, s)
    rest = parts[1]
    if "-" in rest:
        first_seg = rest.split("-", 1)[0]
        try:
            return (0, int(first_seg), s)
        except ValueError:
            return (1, 10**9, s)
    tail = rest.strip()
    try:
        return (0, int(tail), s)
    except ValueError:
        return (1, 10**9, s)


# ---------------------------------------------------------------------------
# ?????????
#   ???? DataFrame 1? ? ????????? dict ????????????AI ???????
# ---------------------------------------------------------------------------
def build_task_queue_from_planning_df(
    tasks_df,
    run_date,
    req_map,
    ai_by_tid=None,
    global_priority_override=None,
    equipment_list=None,
):
    """
    ``generate_plan`` ???????????????????????? task_queue ????
    ai_by_tid ? None ???????? analyze_task_special_remarks ??????
    """
    if ai_by_tid is None:
        ai_by_tid = analyze_task_special_remarks(tasks_df, reference_year=run_date.year)
    task_queue = []
    n_exclude_plan = 0
    seq_by_tid = _collect_process_content_order_by_task_id(tasks_df)
    same_tid_line_seq = defaultdict(int)
    # ??NO????????: iterrows ???????0 ?????task_queue.sort ?????
    planning_sheet_row_seq = 0

    for planning_df_iloc, (_, row) in enumerate(tasks_df.iterrows()):
        _tid_dbg = _normalize_task_id_for_dup_grouping(
            planning_task_id_str_from_plan_row(row)
        )
        if _tid_dbg == _AGENT_DEBUG_FOCUS_TID:
            # region agent log
            _agent_debug_ndjson(
                hypothesis_id="H-scan",
                location="build_task_queue_from_planning_df:iter",
                message="focus_row",
                data={
                    "planning_df_iloc": planning_df_iloc,
                    "process": str(row.get(TASK_COL_MACHINE, "") or ""),
                    "machine_name": str(row.get(TASK_COL_MACHINE_NAME, "") or ""),
                    "exclude_cell": str(row.get(PLAN_COL_EXCLUDE_FROM_ASSIGNMENT, "") or ""),
                    "completion_flag": str(row.get(TASK_COL_COMPLETION_FLAG, "") or ""),
                },
            )
            # endregion
        if row_has_completion_keyword(row):
            if _tid_dbg == _AGENT_DEBUG_FOCUS_TID:
                # region agent log
                _agent_debug_ndjson(
                    hypothesis_id="H3",
                    location="build_task_queue_from_planning_df",
                    message="skip_completed",
                    data={"planning_df_iloc": planning_df_iloc},
                )
                # endregion
            continue
        if _plan_row_exclude_from_assignment(row):
            if _tid_dbg == _AGENT_DEBUG_FOCUS_TID:
                # region agent log
                _agent_debug_ndjson(
                    hypothesis_id="H4",
                    location="build_task_queue_from_planning_df",
                    message="skip_exclude_from_assignment",
                    data={
                        "planning_df_iloc": planning_df_iloc,
                        "process": str(row.get(TASK_COL_MACHINE, "") or ""),
                        "exclude_cell": str(row.get(PLAN_COL_EXCLUDE_FROM_ASSIGNMENT, "") or ""),
                    },
                )
                # endregion
            n_exclude_plan += 1
            continue

        task_id = planning_task_id_str_from_plan_row(row)
        machine = str(row.get(TASK_COL_MACHINE, "")).strip()
        machine_name = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        qty_total = parse_float_safe(row.get(TASK_COL_QTY), 0.0)
        done_qty = calc_done_qty_equivalent_from_row(row)
        speed_raw = row.get(TASK_COL_SPEED, 1)
        product_name = row.get(TASK_COL_PRODUCT, None)
        answer_due = parse_optional_date(_planning_df_cell_scalar(row, TASK_COL_ANSWER_DUE))
        specified_due = parse_optional_date(_planning_df_cell_scalar(row, TASK_COL_SPECIFIED_DUE))
        specified_due_ov = None
        # ????: ????????????????????????????_??????????
        specified_basis = specified_due
        due_basis = None
        due_source = "none"
        due_source_rank = 9
        raw_input_sheet = parse_optional_date(
            _planning_df_cell_scalar(row, TASK_COL_RAW_INPUT_DATE)
        )
        raw_input_date_ov = parse_optional_date(
            _planning_df_cell_scalar(row, PLAN_COL_RAW_INPUT_DATE_OVERRIDE)
        )
        raw_input_date = (
            raw_input_date_ov if raw_input_date_ov is not None else raw_input_sheet
        )
        if (
            raw_input_date_ov is not None
            and raw_input_sheet is not None
            and raw_input_date_ov != raw_input_sheet
        ):
            logging.info(
                "?????_??????: ??NO=%s ????????=%s ???=%s",
                task_id,
                raw_input_sheet,
                raw_input_date_ov,
            )

        qty = max(0.0, qty_total - done_qty)
        speed = parse_float_safe(speed_raw, 1.0)
        if speed <= 0:
            speed = 1.0

        if qty <= 0 or not machine or not task_id:
            if _normalize_task_id_for_dup_grouping(task_id) == _AGENT_DEBUG_FOCUS_TID:
                # region agent log
                _agent_debug_ndjson(
                    hypothesis_id="H2",
                    location="build_task_queue_from_planning_df",
                    message="skip_qty_or_machine",
                    data={
                        "qty": qty,
                        "qty_total": qty_total,
                        "done_qty": done_qty,
                        "machine": machine,
                        "task_id": task_id,
                    },
                )
                # endregion
            continue

        _line_seq = same_tid_line_seq[task_id]
        same_tid_line_seq[task_id] += 1

        remark_raw = str(row.get(PLAN_COL_SPECIAL_REMARK, "") or "").strip()
        has_special_remark = bool(remark_raw) and remark_raw.lower() not in ("nan", "none")
        remark_implies_due_dispatch_priority = (
            _special_remark_implies_due_related_dispatch_priority(remark_raw)
        )
        in_progress = done_qty > 0.0

        ai_one = _ai_task_special_entry_for_row(ai_by_tid, row)
        req_op, speed_ov, task_eff_factor, priority, start_date_ov, start_time_ov, ai_used = _merge_task_row_with_ai(
            row,
            ai_one,
            allow_ai_dispatch_priority_from_remark=remark_implies_due_dispatch_priority,
        )
        preferred_operator_raw = _merge_preferred_operator_cell_and_ai(row, ai_one)
        gpo = global_priority_override or {}
        gop_name = _global_override_preferred_operator_for_task(
            gpo.get("task_preferred_operators"), task_id
        )
        if gop_name is not None:
            preferred_operator_raw = gop_name
            logging.info(
                "????????: ??NO=%s ???OP???????????? %r??????????AI?????",
                task_id,
                gop_name,
            )

        if answer_due is not None:
            due_basis = answer_due
            due_source = "answer_due"
            due_source_rank = 0
        elif specified_basis is not None:
            due_basis = specified_basis
            due_source = "specified_due"
            due_source_rank = 1
        has_done_deadline_override = False

        if speed_ov is not None:
            speed = speed_ov
        if speed <= 0:
            speed = 1.0

        gsm = _global_speed_multiplier_for_row(
            machine, machine_name, gpo.get("global_speed_rules") or []
        )
        if abs(gsm - 1.0) > 1e-12:
            speed_before_g = speed
            speed = speed * gsm
            if speed <= 0:
                speed = 1.0
            logging.info(
                "????????: ??NO=%s ??=%r ???=%r ? speed_multiplier ??=%s ?????? %s ? %s?",
                task_id,
                machine,
                machine_name,
                gsm,
                speed_before_g,
                speed,
            )

        unit = parse_float_safe(
            _planning_df_cell_scalar(row, PLAN_COL_ROLL_UNIT_LENGTH), 0.0
        )
        if unit <= 0:
            unit = infer_unit_m_from_product_name(
                product_name, fallback_unit=qty_total if qty_total > 0 else qty
            )
        try:
            unit = float(unit)
        except Exception:
            unit = qty
        if unit <= 0:
            unit = qty

        # ???????????????????????????????????????????????
        if due_basis is None:
            due_urgent = False
        else:
            due_urgent = due_basis <= run_date

        # ??????:
        # 1) ????????????????? 13:00 ?????????????????
        #    ?????: max(run_date, raw_input_date)???????: 13:00?
        # 2) ???????/AI????????????????????????????date ??????
        # 3) ???????? run_date
        if raw_input_date:
            effective_start_date = max(run_date, raw_input_date)
        else:
            effective_start_date = run_date
        if start_date_ov is not None:
            effective_start_date = (
                max(start_date_ov, raw_input_date)
                if raw_input_date
                else start_date_ov
            )
            if raw_input_date and start_date_ov < raw_input_date:
                logging.info(
                    "???????????????????: ??NO=%s ?????=%s ?????=%s ?????=%s",
                    task_id,
                    start_date_ov,
                    raw_input_date,
                    effective_start_date,
                )

        same_day_raw_start_limit = (
            time(13, 0)
            if (raw_input_date and effective_start_date == raw_input_date)
            else None
        )

        calc_time_val = qty * speed
        ai_note = ""
        if ai_used:
            try:
                ai_note = json.dumps(ai_used, ensure_ascii=False)[:500]
            except Exception:
                ai_note = str(ai_used)[:500]

        _order_list = seq_by_tid.get(task_id) or []
        _p_rank = _process_sequence_rank_for_machine(machine, _order_list)
        _init_rem = float(qty / unit if unit else 0.0)
        _process_content_mismatch = bool(_order_list) and not _process_name_matches_kakou_content_tokens(
            machine, _order_list
        )

        _dto_from_sheet = None
        if RESULT_TASK_COL_DISPATCH_TRIAL_ORDER in tasks_df.columns:
            _dto_from_sheet = parse_optional_int(
                _planning_df_cell_scalar(row, RESULT_TASK_COL_DISPATCH_TRIAL_ORDER)
            )

        task_queue.append(
            {
                "task_id": task_id,
                "machine": machine,
                "machine_name": machine_name,
                "equipment_line_key": _resolve_equipment_line_key_for_task(
                    {"machine": machine, "machine_name": machine_name},
                    equipment_list,
                ),
                "start_date_req": effective_start_date,
                "answer_due_date": answer_due,
                "specified_due_date": specified_due,
                "specified_due_override": specified_due_ov,
                "due_basis_date": due_basis,
                # ?????????? due_basis_date ??? +1 ??????_???????????????????????????????
                "due_basis_date_result_sheet": due_basis,
                "due_source": due_source,
                "due_source_rank": due_source_rank,
                "due_urgent": due_urgent,
                "raw_input_date": raw_input_date,
                "same_day_raw_start_limit": same_day_raw_start_limit,
                "total_qty_m": int(qty_total),
                "unit_m": int(unit),
                "remaining_units": qty / unit if unit else 0,
                "base_time_per_unit": (qty / speed) / (qty / unit) if unit and speed and qty else 0,
                "assigned_history": [],
                "calc_time_value": calc_time_val,
                "required_op": req_op,
                "task_eff_factor": task_eff_factor,
                "priority": priority,
                "earliest_start_time": start_time_ov,
                "preferred_operator_raw": preferred_operator_raw,
                "task_special_ai_note": ai_note,
                "in_progress": in_progress,
                "has_special_remark": has_special_remark,
                "has_done_deadline_override": has_done_deadline_override,
                "done_qty_reported": done_qty,
                "process_sequence_rank": _p_rank,
                "same_request_line_seq": _line_seq,
                "initial_remaining_units": _init_rem,
                "roll_pipeline_ec": _row_matches_roll_pipeline_ec(machine, machine_name),
                "roll_pipeline_inspection": _row_matches_roll_pipeline_inspection(
                    machine, machine_name
                ),
                "roll_pipeline_rewind": _row_matches_roll_pipeline_rewind(
                    machine, machine_name
                ),
                "process_content_mismatch": _process_content_mismatch,
                "planning_sheet_row_seq": planning_sheet_row_seq,
                "planning_df_iloc": planning_df_iloc,
                "dispatch_trial_order_from_sheet": _dto_from_sheet,
            }
        )
        planning_sheet_row_seq += 1

    logging.info(
        "task_queue ????: total=%s???????????? %s ??",
        len(task_queue),
        n_exclude_plan,
    )
    return task_queue


def _task_id_priority_key(task_id):
    """
    ??NO??????????????
    ?: Y3-24, Y3-34 ?????????????????????????
    """
    s = str(task_id or "").strip()
    if not s:
        return ("", 10**9, "")
    parts = s.rsplit("-", 1)
    if len(parts) == 2:
        head = parts[0].strip()
        tail = parts[1].strip()
        if re.match(r"^\d+$", tail):
            return (head, int(tail), s)
    return (s, 10**9, s)


def _serial_dispatch_order_task_ids(task_queue) -> list:
    """
    ??NO????????????NO???? **??????????** ??????????????
    ????????????????????????????????????????
    ?planning_sheet_row_seq????NO???
    """
    min_dto_by_tid: dict = {}
    first_seq_by_tid: dict = {}
    for t in task_queue:
        tid = str(t.get("task_id", "") or "").strip()
        if not tid:
            continue
        try:
            dto = int(t.get("dispatch_trial_order") or 10**9)
        except (TypeError, ValueError):
            dto = 10**9
        prev_d = min_dto_by_tid.get(tid)
        if prev_d is None or dto < prev_d:
            min_dto_by_tid[tid] = dto
        seq = t.get("planning_sheet_row_seq")
        seq = int(seq) if seq is not None else 10**9
        prev = first_seq_by_tid.get(tid)
        if prev is None or seq < prev:
            first_seq_by_tid[tid] = seq
    return sorted(
        min_dto_by_tid.keys(),
        key=lambda tid: (
            min_dto_by_tid[tid],
            first_seq_by_tid.get(tid, 10**9),
            _task_id_priority_key(tid),
        ),
    )


def _excel_scalar_to_plan_string_cell(v):
    """
    ??????read_excel?????????????? DataFrame ??????StringDtype??
    ????? str ???????Excel ????????????? 1 ? \"1\" ???
    """
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    if isinstance(v, bool):
        return str(v).lower()
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if isinstance(v, float) and math.isfinite(v) and float(int(v)) == v:
            return str(int(v))
        if isinstance(v, float) and math.isfinite(v):
            s = str(v)
            if "." in s:
                s = s.rstrip("0").rstrip(".")
            return s
        return str(int(v))
    if isinstance(v, pd.Timestamp):
        if pd.isna(v):
            return ""
        if v.hour == 0 and v.minute == 0 and v.second == 0 and v.microsecond == 0:
            return v.strftime("%Y/%m/%d")
        return v.strftime("%Y/%m/%d %H:%M")
    if isinstance(v, datetime):
        if v.hour == 0 and v.minute == 0 and v.second == 0 and v.microsecond == 0:
            return v.date().strftime("%Y/%m/%d")
        return v.strftime("%Y/%m/%d %H:%M")
    if isinstance(v, date):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, str):
        return v.strip()
    return str(v).strip()


def _merge_plan_sheet_user_overrides(out_df):
    """
    ??????????_??????????????????????
    ??1?????? (??NO, ???) ????????
    ????????????????????????
    """
    if out_df is None or out_df.empty:
        return out_df
    if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
        return out_df
    try:
        df_old = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=PLAN_INPUT_SHEET_NAME)
    except Exception as e:
        logging.info("??1: ????????????????????????? (%s)", e)
        return out_df
    df_old.columns = df_old.columns.str.strip()
    df_old = _align_dataframe_headers_to_canonical(
        df_old,
        plan_input_sheet_column_order(),
    )
    if TASK_COL_TASK_ID not in df_old.columns or TASK_COL_MACHINE not in df_old.columns:
        return out_df

    lookup = {}
    for _, r in df_old.iterrows():
        tid = planning_task_id_str_from_plan_row(r)
        mach = str(r.get(TASK_COL_MACHINE, "") or "").strip()
        if not tid or not mach:
            continue
        key = (tid, mach)
        bucket = lookup.setdefault(key, {})
        for c in (*PLAN_STAGE1_MERGE_COLUMNS, *PLAN_STAGE1_MERGE_EXTRA_COLUMNS):
            if c not in df_old.columns or c not in out_df.columns:
                continue
            v = r.get(c)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            if isinstance(v, str):
                s = v.strip()
                if not s or s.lower() in ("nan", "none"):
                    continue
            bucket[c] = v

    if not lookup:
        return out_df

    merged_rows = 0
    for i, row in out_df.iterrows():
        tid = planning_task_id_str_from_plan_row(row)
        mach = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        bucket = lookup.get((tid, mach))
        if not bucket:
            continue
        merged_rows += 1
        for c, v in bucket.items():
            if c == PLAN_COL_EXCLUDE_FROM_ASSIGNMENT:
                v = _coerce_plan_exclude_column_value_for_storage(v)
            elif c in out_df.columns and pd.api.types.is_string_dtype(out_df[c].dtype):
                v = _excel_scalar_to_plan_string_cell(v)
            out_df.at[i, c] = v

    if merged_rows:
        logging.info(
            "??1: ???????????? %s ????????????: ??NO+?????",
            merged_rows,
        )
    return out_df


# ---------------------------------------------------------------------------
# ?????2???
#   (A) DataFrame ????? � ????NO�??????????? yes????????????
#   (B) ?????????_???????� ??+????? C/D/E ??Gemini ? D?E?
#       ??????? xlwings ? A:E ???Save ??????????
#   ???? apply_exclude_rules_config_to_plan_df ??? DataFrame ???????
# ---------------------------------------------------------------------------

def _auto_exclude_cell_empty_for_autofill(v) -> bool:
    """?????????????????? yes ??????"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return True
    if isinstance(v, str):
        s = str(v).strip()
        return not s or s.lower() in ("nan", "none")
    return False


def _normalize_task_id_for_dup_grouping(raw) -> str:
    """????NO?????????????????????????"""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    if isinstance(raw, float) and raw == int(raw):
        s = str(int(raw))
    else:
        s = unicodedata.normalize("NFKC", str(raw).strip())
    s = s.strip()
    if not s or s.lower() == "nan":
        return ""
    return s.upper()


def _process_name_is_bunkatsu_for_auto_exclude(raw) -> bool:
    """??????????????NFKC ???"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\s?]+", "", t)
    return t == "??"


def _apply_auto_exclude_bunkatsu_duplicate_machine(
    df: pd.DataFrame, log_prefix: str = "??1"
) -> pd.DataFrame:
    """
    ????NO?2??????????????????2????????????
    ?????????????????? yes ????????????????
    ???? _normalize_equipment_match_key ??????
    """
    if df is None or df.empty:
        return df
    need_cols = (TASK_COL_TASK_ID, TASK_COL_MACHINE, TASK_COL_MACHINE_NAME)
    for c in need_cols:
        if c not in df.columns:
            return df
    if PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns:
        df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = ""
    # read_excel ?? StringDtype ?????????? .at ??? TypeError ????? object ????
    df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT].astype(object)

    by_tid = defaultdict(list)
    for i in df.index:
        tid = _normalize_task_id_for_dup_grouping(df.at[i, TASK_COL_TASK_ID])
        if not tid:
            continue
        by_tid[tid].append(i)

    n_set = 0
    for _tid_key, idx_list in by_tid.items():
        if len(idx_list) < 2:
            continue
        counts = defaultdict(int)
        for i in idx_list:
            mn_key = _normalize_equipment_match_key(df.at[i, TASK_COL_MACHINE_NAME])
            if not mn_key:
                continue
            counts[mn_key] += 1
        if not any(c >= 2 for c in counts.values()):
            continue
        if _tid_key == _AGENT_DEBUG_FOCUS_TID:
            # region agent log
            row_snap = []
            for _ri in idx_list:
                row_snap.append(
                    {
                        "i": int(_ri),
                        "process": str(df.at[_ri, TASK_COL_MACHINE]),
                        "machine_name": str(df.at[_ri, TASK_COL_MACHINE_NAME]),
                        "mn_key": _normalize_equipment_match_key(
                            df.at[_ri, TASK_COL_MACHINE_NAME]
                        ),
                        "exclude_before": str(
                            df.at[_ri, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] or ""
                        ),
                    }
                )
            _agent_debug_ndjson(
                hypothesis_id="H1",
                location="_apply_auto_exclude_bunkatsu_duplicate_machine",
                message="duplicate_machine_group_for_tid",
                data={"tid": _tid_key, "machine_key_counts": dict(counts), "rows": row_snap},
            )
            # endregion
        for i in idx_list:
            if not _process_name_is_bunkatsu_for_auto_exclude(df.at[i, TASK_COL_MACHINE]):
                continue
            if not _auto_exclude_cell_empty_for_autofill(
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT]
            ):
                continue
            # ?? StringDtype ??? int ??? TypeError ????????????_plan_row_exclude_from_assignment ? yes ???????
            df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
            if _tid_key == _AGENT_DEBUG_FOCUS_TID:
                # region agent log
                _agent_debug_ndjson(
                    hypothesis_id="H1",
                    location="_apply_auto_exclude_bunkatsu_duplicate_machine",
                    message="set_exclude_yes_for_bunkatsu",
                    data={
                        "row_index": int(i),
                        "process": str(df.at[i, TASK_COL_MACHINE]),
                        "machine_name": str(df.at[i, TASK_COL_MACHINE_NAME]),
                    },
                )
                # endregion
            n_set += 1

    if n_set:
        logging.info(
            "%s: ????NO???????????????????????????? %s ????????=yes ??????????",
            log_prefix,
            n_set,
        )
    return df


def _normalize_process_name_for_rule_match(raw) -> str:
    """??????????NFKC???????"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\s?]+", "", t)
    return t


def _exclude_rules_sheet_header_map(ws) -> dict:
    """1????? ? ???(1???)?
    openpyxl ????????? max_column ? 0 ????????????????????? return ??????
    ?????? A?E ?????????
    """
    h = {}
    last_col = max(5, int(ws.max_column or 0))
    for col in range(1, last_col + 1):
        v = ws.cell(1, col).value
        if v is not None:
            h[str(v).strip()] = col
    return h


def _ensure_exclude_rules_sheet_headers_and_columns(ws, log_prefix: str) -> tuple[int, int, int, int, int]:
    """
    1????????????????????????????????????????????????
    ???????????????? A1:E1 ?????????????????????
    """
    headers = (
        EXCLUDE_RULE_COL_PROCESS,
        EXCLUDE_RULE_COL_MACHINE,
        EXCLUDE_RULE_COL_FLAG,
        EXCLUDE_RULE_COL_LOGIC_JA,
        EXCLUDE_RULE_COL_LOGIC_JSON,
    )
    hm = _exclude_rules_sheet_header_map(ws)
    if all(hm.get(x) for x in headers):
        return tuple(hm[x] for x in headers)
    for i, name in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=name)
    logging.info(
        "%s: ?%s???????????????????????1???A1:E1?????????",
        log_prefix,
        EXCLUDE_RULES_SHEET_NAME,
    )
    return (1, 2, 3, 4, 5)


def _compact_exclude_rules_data_rows(
    ws,
    c_proc: int,
    c_mach: int,
    c_flag: int,
    c_d: int,
    c_e: int,
    log_prefix: str,
) -> tuple[int, int]:
    """
    2 ????????????????????????????????????
    ??: ????????? A?E ???5???????????
    Returns (????????, ??????).
    """
    max_r = int(ws.max_row or 1)
    if max_r < 2:
        return 0, 0

    old_body = max_r - 1
    cols = (c_proc, c_mach, c_flag, c_d, c_e)
    rows: list[tuple[str, str, object, object, object]] = []
    for r in range(2, max_r + 1):
        pv = ws.cell(row=r, column=c_proc).value
        mv = ws.cell(row=r, column=c_mach).value
        cv = ws.cell(row=r, column=c_flag).value
        dv = ws.cell(row=r, column=c_d).value
        ev = ws.cell(row=r, column=c_e).value
        all_blank = all(
            _cell_is_blank_for_rule(ws.cell(row=r, column=c).value) for c in cols
        )
        p = str(pv).strip() if pv is not None and not (isinstance(pv, float) and pd.isna(pv)) else ""
        m = str(mv).strip() if mv is not None and not (isinstance(mv, float) and pd.isna(mv)) else ""
        if all_blank or not p:
            continue
        rows.append((p, m, cv, dv, ev))

    n_skip = old_body - len(rows)

    if not rows:
        ws.delete_rows(2, old_body)
        if old_body > 0:
            logging.info(
                "%s: ?%s????????????????????? %s ?????????",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                old_body,
            )
        return 0, n_skip

    ws.delete_rows(2, old_body)
    for i, (p, m, cv, dv, ev) in enumerate(rows, start=2):
        ws.cell(row=i, column=c_proc, value=p)
        ws.cell(row=i, column=c_mach, value=m)
        ws.cell(row=i, column=c_flag, value=cv)
        ws.cell(row=i, column=c_d, value=dv)
        ws.cell(row=i, column=c_e, value=ev)

    if n_skip:
        logging.info(
            "%s: ?%s?????? %s ?????%s ????????????????",
            log_prefix,
            EXCLUDE_RULES_SHEET_NAME,
            n_skip,
            len(rows),
        )
    return len(rows), n_skip


def _cell_is_blank_for_rule(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip()
    return not s or s.lower() in ("nan", "none", "null")


def _exclude_rule_c_column_is_yes(v) -> bool:
    """C???????????????+???????????????"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return int(v) == 1
        except (TypeError, ValueError):
            pass
    s = unicodedata.normalize("NFKC", str(v).strip()).lower()
    return s in ("yes", "true", "1", "y", "??", "?", "?", "?")


def _task_row_matches_exclude_rule_target(
    task_proc: str, task_mach: str, rule_proc: str, rule_mach: str
) -> bool:
    if _normalize_process_name_for_rule_match(task_proc) != _normalize_process_name_for_rule_match(
        rule_proc
    ):
        return False
    rm = str(rule_mach or "").strip()
    if not rm:
        return True
    return _normalize_equipment_match_key(task_mach) == _normalize_equipment_match_key(rm)


def _collect_process_machine_pairs_for_exclude_rules(df_src: pd.DataFrame) -> list[tuple[str, str]]:
    """????DATA ?????1???????? (???, ???) ???????????????"""
    out: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for _, row in df_src.iterrows():
        if row_has_completion_keyword(row):
            continue
        task_id = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
        machine = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        machine_name = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        qty_total = parse_float_safe(row.get(TASK_COL_QTY), 0.0)
        done_qty = calc_done_qty_equivalent_from_row(row)
        qty = max(0.0, qty_total - done_qty)
        if qty <= 0 or not machine or not task_id:
            continue
        key = (
            _normalize_process_name_for_rule_match(machine),
            _normalize_equipment_match_key(machine_name),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append((machine, machine_name))
    return out


def _parse_exclude_rule_json_cell(raw) -> dict | None:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    s = str(raw).strip()
    if not s:
        return None
    fence = re.search(
        r"```(?:json)?\s*(\{.*\})\s*```",
        s,
        re.DOTALL | re.IGNORECASE,
    )
    if fence:
        s = fence.group(1).strip()
    try:
        obj = json.loads(s)
    except json.JSONDecodeError:
        return None
    return obj if isinstance(obj, dict) else None


def _validate_exclude_rule_parsed_dict(o: object) -> dict | None:
    """Gemini?E????? dict ???????????????"""
    if not isinstance(o, dict):
        return None
    if int(o.get("version") or 0) != 1:
        return None
    mode = str(o.get("mode") or "").strip().lower()
    if mode not in ("always_exclude", "conditions"):
        return None
    return o


def _exclude_rule_de_cache_key(stripped_blob: str) -> str:
    """??????????????????????? ai_remarks_cache ????"""
    h = hashlib.sha256(stripped_blob.encode("utf-8")).hexdigest()
    return f"{AI_CACHE_KEY_PREFIX_EXCLUDE_RULE_DE}:{h}"


def _cache_get_exclude_rule_de_parsed(cache_obj: dict, blob: str) -> dict | None:
    s = str(blob or "").strip()
    if not s:
        return None
    data = get_cached_ai_result(
        cache_obj, _exclude_rule_de_cache_key(s), content_key=s
    )
    if not isinstance(data, dict):
        return None
    return _validate_exclude_rule_parsed_dict(data)


def _cache_put_exclude_rule_de_parsed(
    cache_obj: dict, blob: str, parsed: dict | None
) -> None:
    if parsed is None:
        return
    s = str(blob or "").strip()
    if not s:
        return
    put_cached_ai_result(
        cache_obj, _exclude_rule_de_cache_key(s), parsed, content_key=s
    )


def _exclude_rule_logic_gemini_schema_instructions() -> str:
    allowed = ", ".join(sorted(EXCLUDE_RULE_ALLOWED_COLUMNS))
    return (
        "????? version ??? 1?\n"
        "1) ???????????????????????:\n"
        '{"version":1,"mode":"always_exclude"}\n\n'
        "2) ??????????????:\n"
        '{"version":1,"mode":"conditions","require_all": true ??? false,"conditions":[ ... ]}\n\n'
        "conditions ????:\n"
        "- {\"column\":\"??\",\"op\":\"empty\"} � ????\n"
        "- {\"column\":\"??\",\"op\":\"not_empty\"}\n"
        "- {\"column\":\"??\",\"op\":\"eq\",\"value\":\"???\"} / ne / contains / not_contains / regex??????\n"
        "- {\"column\":\"??\",\"op\":\"gt\"|\"gte\"|\"lt\"|\"lte\",\"value\":??} � ??????????????\n\n"
        f"??????????????????????:\n{allowed}\n"
    )


def _parse_exclude_rule_json_array_response(text: str) -> list | None:
    """??????? JSON ????????```json ?????????"""
    s = (text or "").strip()
    if not s:
        return None
    fence = re.search(
        r"```(?:json)?\s*(\[.*\])\s*```",
        s,
        re.DOTALL | re.IGNORECASE,
    )
    if fence:
        s = fence.group(1).strip()
    try:
        obj = json.loads(s)
    except json.JSONDecodeError:
        return None
    return obj if isinstance(obj, list) else None


def _row_scalar_for_exclude_rule(row, col_name: str):
    try:
        return row.get(col_name)
    except Exception:
        return None


def _evaluate_exclude_rule_one_condition(cond: dict, row) -> bool:
    if not isinstance(cond, dict):
        return False
    col = cond.get("column")
    if col not in EXCLUDE_RULE_ALLOWED_COLUMNS:
        logging.warning("???????: ???????????????: %s", col)
        return False
    op = str(cond.get("op") or "").strip().lower()
    val = _row_scalar_for_exclude_rule(row, col)
    val_s = "" if val is None or (isinstance(val, float) and pd.isna(val)) else str(val).strip()
    val_s_lower = val_s.lower()

    if op == "empty":
        return val_s == ""
    if op == "not_empty":
        return val_s != ""

    if op in ("contains", "not_contains", "regex", "eq", "ne"):
        rhs = cond.get("value", "")
        pat = "" if rhs is None else str(rhs)
        if op == "contains":
            return pat in val_s
        if op == "not_contains":
            return pat not in val_s
        if op == "regex":
            try:
                return re.search(pat, val_s) is not None
            except re.error:
                return False
        if op == "eq":
            return val_s == pat
        if op == "ne":
            return val_s != pat

    def _num(x):
        try:
            return float(x)
        except (TypeError, ValueError):
            return None

    nv = _num(val)
    cv = _num(cond.get("value"))
    if nv is None or cv is None:
        return False
    if op == "gt":
        return nv > cv
    if op == "gte":
        return nv >= cv
    if op == "lt":
        return nv < cv
    if op == "lte":
        return nv <= cv
    return False


def evaluate_exclude_rule_json_for_row(rule: dict, row) -> bool:
    """
    E?? JSON?version=1??????????????????????? True?
    mode: always_exclude | conditions
    """
    if not isinstance(rule, dict) or int(rule.get("version") or 0) != 1:
        return False
    mode = str(rule.get("mode") or "").strip().lower()
    if mode == "always_exclude":
        return True
    if mode != "conditions":
        return False
    conds = rule.get("conditions")
    if not isinstance(conds, list) or not conds:
        return False
    require_all = bool(rule.get("require_all", True))
    checks = []
    for c in conds:
        if isinstance(c, dict) and c.get("column") in EXCLUDE_RULE_ALLOWED_COLUMNS:
            checks.append(_evaluate_exclude_rule_one_condition(c, row))
    if not checks:
        return False
    return all(checks) if require_all else any(checks)


def _ai_compile_exclude_rule_logic_to_json(natural_language: str) -> dict | None:
    """
    D??????? Gemini ? JSON ?????????? None?
    json/ai_remarks_cache.json ? TTL ??????????????? API ???????
    """
    blob = str(natural_language or "").strip()
    if not blob:
        return None
    ai_cache = load_ai_cache()
    hit = _cache_get_exclude_rule_de_parsed(ai_cache, blob)
    if hit is not None:
        logging.info("???????: AI??????????????????JSON?")
        return hit
    if not API_KEY:
        return None
    schema = _exclude_rule_logic_gemini_schema_instructions()
    prompt = (
        "?????????????????????????????????1???????????????????????\n\n"
        "??????? { ????? } ? JSON ??????1?????????????????\n\n"
        f"{schema}\n"
        f"?????\n{blob}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("???????: ????? ? %s", ppath)
    except OSError as ex:
        logging.warning("???????: ?????????: %s", ex)
    try:
        client = genai.Client(api_key=API_KEY)
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        raw = (_gemini_result_text(res) or "").strip()
        rpath = os.path.join(log_dir, "ai_exclude_rule_logic_last_response.txt")
        try:
            with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                rf.write(raw)
        except OSError:
            pass
        parsed = _validate_exclude_rule_parsed_dict(_parse_exclude_rule_json_cell(raw))
        if parsed:
            _cache_put_exclude_rule_de_parsed(ai_cache, blob, parsed)
            save_ai_cache(ai_cache)
        return parsed
    except Exception as e:
        logging.warning("???????: Gemini ????: %s", e)
        return None


def _ai_compile_exclude_rule_logics_batch(blobs: list[str]) -> list[dict | None]:
    """
    ??? D ???? 1 ?? Gemini ????? JSON ????????????? 1 ????????????
    json/ai_remarks_cache.json ????????? API ??????
    """
    n = len(blobs)
    if n == 0:
        return []
    ai_cache = load_ai_cache()
    out: list[dict | None] = [None] * n
    pend_i: list[int] = []
    pend_b: list[str] = []
    for i, b in enumerate(blobs):
        s = str(b).strip()
        hit = _cache_get_exclude_rule_de_parsed(ai_cache, s) if s else None
        if hit is not None:
            out[i] = hit
        else:
            pend_i.append(i)
            pend_b.append(s)
    if not pend_b:
        logging.info(
            "???????: AI???????? D?E ??? %s ?????API ????????",
            n,
        )
        return out
    if not API_KEY:
        return out
    m = len(pend_b)
    if m == 1:
        out[pend_i[0]] = _ai_compile_exclude_rule_logic_to_json(pend_b[0])
        return out

    schema = _exclude_rule_logic_gemini_schema_instructions()
    numbered = "\n".join(f"[{i + 1}] {str(b).strip()}" for i, b in enumerate(pend_b))
    prompt = (
        "???????????????????? N ??????????????????????? JSON ?????????????\n\n"
        f"????JSON ???????? [ ????? ] ??????? {m}?Markdown???????\n"
        f"???????? [1]?2 ??? [2] � ???????\n\n"
        f"{schema}\n"
        f"?????\n{numbered}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_batch_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("???????(???): ????? ? %s", ppath)
    except OSError as ex:
        logging.warning("???????(???): ?????????: %s", ex)
    try:
        client = genai.Client(api_key=API_KEY)
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        raw = (_gemini_result_text(res) or "").strip()
        rpath = os.path.join(log_dir, "ai_exclude_rule_logic_batch_last_response.txt")
        try:
            with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                rf.write(raw)
        except OSError:
            pass
        arr = _parse_exclude_rule_json_array_response(raw)
        if not isinstance(arr, list) or len(arr) != m:
            logging.warning(
                "???????: ???????????? %s??? %s??1 ??????????",
                len(arr) if isinstance(arr, list) else None,
                m,
            )
            for j, idx in enumerate(pend_i):
                out[idx] = _ai_compile_exclude_rule_logic_to_json(pend_b[j])
            return out
        cache_dirty = False
        for j, item in enumerate(arr):
            parsed = _validate_exclude_rule_parsed_dict(item)
            out[pend_i[j]] = parsed
            if parsed:
                _cache_put_exclude_rule_de_parsed(ai_cache, pend_b[j], parsed)
                cache_dirty = True
        if cache_dirty:
            save_ai_cache(ai_cache)
        return out
    except Exception as e:
        logging.warning("???????: ??? Gemini ?????????????: %s", e)
        for j, idx in enumerate(pend_i):
            out[idx] = _ai_compile_exclude_rule_logic_to_json(pend_b[j])
        return out


def _log_exclude_rules_sheet_debug(
    event: str,
    log_prefix: str,
    summary: str,
    details: str = "",
    exc: BaseException | None = None,
) -> None:
    """
    ???_????????????????????

    ??????????? log/exclude_rules_sheet_debug.txt ?????execution_log ????????????
    event ?: START, OPEN_OK, OPEN_RETRY, OPEN_FAIL, HEADER_FIX, SYNC_ROWS, OPENPYXL_SAVE_OK, OPENPYXL_SAVE_FAIL,
    OPENPYXL_SAVE_SKIPPED_EXCLUDE_RULES_POLICY, OPENPYXL_RETRY_WAIT, OPENPYXL_VBA_FALLBACK, MATRIX_TSV_WRITTEN,
    XLWINGS_UNAVAILABLE, XLWINGS_ATTACH_FAIL, XLWINGS_SYNC_SKIP, XLWINGS_SYNC_OK, XLWINGS_SYNC_FAIL,
    E_SIDECAR_WRITTEN, E_SIDECAR_APPLIED, FALLBACK_FAIL,
    SKIP_NO_PATH, SKIP_NO_FILE, SKIP_NO_SHEET, DATA_COMPACT
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        f"--- {ts} ---",
        f"event={event}",
        f"phase={log_prefix}",
        f"summary={summary}",
    ]
    if details:
        lines.append(f"details={details}")
    if exc is not None:
        lines.append(f"exception={type(exc).__name__}: {exc}")
        lines.append(traceback.format_exc().rstrip())
    block = "\n".join(lines) + "\n\n"
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(exclude_rules_sheet_debug_log_path, "a", encoding="utf-8", newline="\n") as df:
            df.write(block)
    except OSError as wex:
        logging.warning("exclude_rules_sheet_debug.txt ??????: %s", wex)

    tag = "[??_??????]"
    msg = f"{tag} {event} | {log_prefix} | {summary}"
    if details:
        msg += f" | {details}"
    if event in (
        "OPEN_FAIL",
        "SAVE_FAIL",
        "COM_MERGE_FAIL",
        "FALLBACK_FAIL",
        "SKIP_NO_PATH",
        "SKIP_NO_FILE",
        "SKIP_NO_SHEET",
        "FATAL",
    ):
        logging.error(msg)
    elif event in (
        "OPEN_RETRY",
        "SAVE_FAIL_HINT",
        "SAVE_RETRY",
        "COM_SYNC_UNAVAILABLE",
        "COM_ATTACH_OPEN_FAIL",
        "XLWINGS_UNAVAILABLE",
        "XLWINGS_ATTACH_FAIL",
        "XLWINGS_SYNC_SKIP",
        "XLWINGS_SYNC_FAIL",
        "E_SIDECAR_WRITTEN",
        "OPENPYXL_SAVE_FAIL",
        "OPENPYXL_VBA_FALLBACK",
    ):
        logging.warning(msg)
    elif event in (
        "COM_MERGE_SKIP",
        "MATRIX_TSV_WRITTEN",
        "OPENPYXL_SAVE_OK",
        "OPENPYXL_RETRY_WAIT",
        "XLWINGS_SYNC_OK",
    ):
        logging.info(msg)
    else:
        logging.info(msg)


def _xlwings_paths_equivalent(disk_path: str, book_fullname: str) -> bool:
    """??????? xlwings Book.full_name ???????????????????????"""
    try:
        fn = str(book_fullname).strip()
    except Exception:
        return False

    def _norm(p: str) -> str:
        p = os.path.normpath(str(p).strip().replace("/", "\\"))
        return os.path.normcase(os.path.abspath(p))

    try:
        if _norm(disk_path) == _norm(fn):
            return True
    except Exception:
        pass
    try:
        return os.path.samefile(disk_path, fn)
    except Exception:
        pass
    try:
        import win32api  # type: ignore

        a = _norm(win32api.GetLongPathName(disk_path))
        b = _norm(win32api.GetLongPathName(fn))
        if a == b:
            return True
    except Exception:
        pass
    try:
        if os.path.basename(_norm(disk_path)).lower() == os.path.basename(_norm(fn)).lower():
            if _norm(os.path.dirname(disk_path)) == _norm(os.path.dirname(fn)):
                return True
    except Exception:
        pass
    return False


def _xlwings_book_matches_path(book, disk_path: str) -> bool:
    try:
        fn = book.full_name
    except Exception:
        return False
    return _xlwings_paths_equivalent(disk_path, fn)


def _xlwings_find_book_on_running_instances(abs_path: str):
    """???? Excel ???????? xlwings Book ???????? None?"""
    try:
        import xlwings as xw
    except ImportError:
        return None
    target = os.path.abspath(abs_path)
    try:
        for app in list(xw.apps):
            try:
                for book in app.books:
                    try:
                        if _xlwings_book_matches_path(book, target):
                            return book
                    except Exception:
                        continue
            except Exception:
                continue
    except Exception:
        return None
    return None


def _xlwings_try_open_in_running_apps(abs_path: str):
    """??? Excel.App ? Workbooks.Open ??????? Book???? None?"""
    try:
        import xlwings as xw
    except ImportError:
        return None
    path = os.path.abspath(abs_path)
    for app in list(xw.apps):
        try:
            return app.books.open(path, update_links=False)
        except Exception:
            continue
    return None


def _xlwings_release_book_after_mutation(xw_book, info: dict, mutation_ok: bool) -> None:
    """?????? Excel ????????? Excel ??? Open ???????????????"""
    if xw_book is None:
        return
    mode = info.get("mode", "keep")
    opened_here = bool(info.get("opened_wb_here"))
    if mode == "quit_excel":
        try:
            xw_book.close()
        except Exception:
            pass
        try:
            xw_book.app.quit()
        except Exception:
            pass
        return
    if opened_here and not mutation_ok:
        try:
            xw_book.close()
        except Exception:
            pass


def _xlwings_attach_open_macro_workbook(macro_wb_path: str, log_prefix: str):
    """
    ??????? xlwings ????????????????
    ???: (Book, release_info) / ??? None?
    release_info: mode ? keep ??? quit_excel?opened_wb_here ? bool?
    """
    try:
        import xlwings as xw  # noqa: F401
    except ImportError:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_UNAVAILABLE",
            log_prefix,
            "xlwings ? import ??????pip install xlwings ?????",
        )
        return None

    abs_path = os.path.abspath(macro_wb_path)

    book = _xlwings_find_book_on_running_instances(abs_path)
    if book is not None:
        return book, {"mode": "keep", "opened_wb_here": False}

    book = _xlwings_try_open_in_running_apps(abs_path)
    if book is not None:
        return book, {"mode": "keep", "opened_wb_here": True}

    try:
        import xlwings as xw

        app = xw.App(visible=False, add_book=False)
        try:
            app.display_alerts = False
        except Exception:
            pass
        book = app.books.open(abs_path, update_links=False)
        return book, {"mode": "quit_excel", "opened_wb_here": True}
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_ATTACH_FAIL",
            log_prefix,
            "xlwings ??????????????",
            details=f"path={abs_path}",
            exc=ex,
        )
        return None


def _xlwings_attach_workbook_for_tests(
    book_path: str,
    label: str,
    *,
    allow_dispatch_open: bool = False,
):
    """
    ????????: ??????????????????? Excel ????
    ???: (Book, info, ?????) ??? None?
    """
    abs_path = os.path.abspath(book_path)
    book = _xlwings_find_book_on_running_instances(abs_path)
    if book is not None:
        return book, {"mode": "keep", "opened_wb_here": False}, f"{label}:????????"
    if not allow_dispatch_open:
        return None
    try:
        import xlwings as xw

        app = xw.App(visible=True, add_book=False)
        try:
            app.display_alerts = False
        except Exception:
            pass
        book = app.books.open(abs_path, update_links=False)
        return book, {"mode": "keep", "opened_wb_here": True}, f"{label}:dispatch-open"
    except Exception:
        return None


def _xlwings_app_save_perf_state_push(app):
    """VBA ???????????????????????????????????? Excel ???????"""
    snap = {}
    for attr in ("screen_updating", "calculation", "enable_events"):
        try:
            snap[attr] = getattr(app, attr)
        except Exception:
            snap[attr] = None
    try:
        app.screen_updating = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        try:
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass
    try:
        app.enable_events = False
    except Exception:
        pass
    return snap


def _xlwings_app_save_perf_state_pop(app, snap):
    if not snap:
        return
    for attr in ("enable_events", "calculation", "screen_updating"):
        prev = snap.get(attr)
        if prev is None:
            continue
        try:
            setattr(app, attr, prev)
        except Exception:
            pass


def _xlwings_sync_exclude_rules_sheet_from_openpyxl(
    wb_path: str, ws_oxl, log_prefix: str
) -> bool:
    """
    openpyxl ??????????xlwings ????_???????A:E ???????????? Save?

    ???????????? .value ??????????????????D3=true???
    Range ???????????????????????????????? api.Value2 ????
    """
    global _exclude_rules_effective_read_path

    attached = _xlwings_attach_open_macro_workbook(wb_path, log_prefix)
    if attached is None:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_SKIP",
            log_prefix,
            "xlwings ?????????? A:E ????????",
            details=f"path={wb_path}",
        )
        return False

    xw_book, info = attached
    ok = False
    try:
        try:
            xw_book.app.display_alerts = False
        except Exception:
            pass
        # ????????????????? COM ??????D3=true ?? VBA ??????????
        # 1 ????????????????????? 40 ????213s????????????
        try:
            sht = xw_book.sheets[EXCLUDE_RULES_SHEET_NAME]
        except Exception:
            _log_exclude_rules_sheet_debug(
                "XLWINGS_SYNC_SKIP",
                log_prefix,
                f"xlwings ??????{EXCLUDE_RULES_SHEET_NAME}????????",
                details=f"path={wb_path}",
            )
            return False
        max_r = max(1, int(ws_oxl.max_row or 1))
        ncols = EXCLUDE_RULES_SHEET_COM_SYNC_MAX_COL
        data = [
            [ws_oxl.cell(row=r, column=c).value for c in range(1, ncols + 1)]
            for r in range(1, max_r + 1)
        ]
        _perf_snap = _xlwings_app_save_perf_state_push(xw_book.app)
        rng = sht.range((1, 1)).resize(len(data), ncols)
        hid_sheet_for_write = False
        try:
            try:
                if int(sht.api.Visible) == -1:  # xlSheetVisible
                    sht.api.Visible = 0  # xlSheetHidden?????????????????????????
                    hid_sheet_for_write = True
            except Exception:
                pass
            try:
                rng.api.Value2 = data
            except Exception:
                rng.value = data
            xw_book.save()
        finally:
            if hid_sheet_for_write:
                try:
                    sht.api.Visible = -1
                except Exception:
                    pass
            _xlwings_app_save_perf_state_pop(xw_book.app, _perf_snap)
        ok = True
        _exclude_rules_effective_read_path = wb_path
        _clear_exclude_rules_e_apply_files()
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_OK",
            log_prefix,
            "xlwings ???????? A?E ???????????????",
            details=f"path={wb_path} rows={max_r}",
        )
        logging.info(
            "%s: ?????? xlwings ???????????????A?E??",
            log_prefix,
        )
        return True
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_FAIL",
            log_prefix,
            "xlwings ?? A:E ????? Save ????????",
            details=f"path={wb_path}",
            exc=ex,
        )
        return False
    finally:
        _xlwings_release_book_after_mutation(xw_book, info, ok)


# ??????????A?E??xlwings ???VBA ?? TSV ???????
EXCLUDE_RULES_SHEET_COM_SYNC_MAX_COL = 5
EXCLUDE_RULES_MATRIX_CLIP_MAX_COL = 5


def _persist_exclude_rules_workbook(_wb, wb_path: str, ws, log_prefix: str) -> bool:
    """
    ???????????????? xlwings ? A:E ???Save?EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1 ????? openpyxl save ?????
    ????????? log ??? TSV ????VBA???_??????_A??E_TSV???????????

    _wb � ???? openpyxl ????openpyxl ????? save ?????
    """
    global _exclude_rules_effective_read_path

    def _openpyxl_persist_ok(which: str) -> bool:
        try:
            _wb.save(wb_path)
        except Exception as ex:
            _log_exclude_rules_sheet_debug(
                "OPENPYXL_SAVE_FAIL",
                log_prefix,
                f"openpyxl ?? .xlsm ????????? {which}?Excel ?????????????????",
                details=f"path={wb_path}",
                exc=ex,
            )
            return False
        _exclude_rules_effective_read_path = wb_path
        _clear_exclude_rules_e_apply_files()
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_OK",
            log_prefix,
            "openpyxl ????????????????????A?E??",
            details=f"path={wb_path} {which}",
        )
        logging.info(
            "%s: ?????? openpyxl ???????????????%s",
            log_prefix,
            which,
        )
        return True

    saved_openpyxl = False
    if EXCLUDE_RULES_SKIP_OPENPYXL_SAVE:
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_SKIPPED_EXCLUDE_RULES_POLICY",
            log_prefix,
            "??_??????????? openpyxl save ????????xlwings ?????????????? EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1??",
            details=f"path={wb_path}",
        )
        logging.info(
            "%s: ??_??????? openpyxl ???? xlwings ???Save ?????????? VBA ??? TSV??",
            log_prefix,
        )
    elif not _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "%s: ??_??????? openpyxl ????????????? xlwings ???Save???????? VBA ??? TSV??",
            log_prefix,
        )
        labels = ("(1/4)", "(2/4)", "(3/4)", "(4/4)")
        for i, label in enumerate(labels):
            if i:
                _log_exclude_rules_sheet_debug(
                    "OPENPYXL_RETRY_WAIT",
                    log_prefix,
                    f"openpyxl ????? 2 ????? {label}?",
                    details=f"path={wb_path}",
                )
                time_module.sleep(2.0)
            if _openpyxl_persist_ok(label):
                saved_openpyxl = True
                break
    else:
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_SKIPPED_INCOMPATIBLE_SHEET",
            log_prefix,
            f"?????{OPENPYXL_INCOMPATIBLE_SHEET_MARKER}?????? openpyxl ???????????",
            details=f"path={wb_path}",
        )
        logging.info(
            "%s: ?????%s?????? openpyxl save ???????xlwings ????? TSV ????????",
            log_prefix,
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )

    if saved_openpyxl:
        return True

    if _xlwings_sync_exclude_rules_sheet_from_openpyxl(wb_path, ws, log_prefix):
        return True

    if _write_exclude_rules_matrix_vba_tsv(wb_path, ws, log_prefix):
        logging.warning(
            "%s: ?????? log\\%s ????????"
            " Excel ???????_??????_A??E_TSV???????????????",
            log_prefix,
            EXCLUDE_RULES_MATRIX_VBA_FILENAME,
        )

    _log_exclude_rules_sheet_debug(
        "OPENPYXL_VBA_FALLBACK",
        log_prefix,
        "openpyxl ????????? VBA ??? TSV ???????????? Excel ???????????????????",
        details=f"path={wb_path}",
    )
    return False


def _exclude_rules_e_sidecar_path() -> str:
    path = os.path.join(json_data_dir, EXCLUDE_RULES_E_SIDECAR_FILENAME)
    legacy = os.path.join(log_dir, EXCLUDE_RULES_E_SIDECAR_FILENAME)
    if os.path.isfile(legacy) and not os.path.isfile(path):
        try:
            shutil.move(legacy, path)
        except OSError:
            pass
    return path


def _exclude_rules_e_vba_tsv_path() -> str:
    return os.path.join(log_dir, EXCLUDE_RULES_E_VBA_TSV_FILENAME)


def _exclude_rules_matrix_vba_path() -> str:
    return os.path.join(log_dir, EXCLUDE_RULES_MATRIX_VBA_FILENAME)


def _serialize_cell_for_matrix_tsv(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    return str(val)


def _write_exclude_rules_matrix_vba_tsv(
    wb_path: str, ws, log_prefix: str
) -> bool:
    """VBA ?: ????? 1 ??? max_row ? A?E ? Base64(UTF-8) ?? TSV ??????"""
    max_r = max(1, int(ws.max_row or 1))
    lines = [
        "v1",
        "workbook\t" + os.path.abspath(wb_path),
        "sheet\t" + EXCLUDE_RULES_SHEET_NAME,
        "ncols\t5",
        "---",
    ]
    for r in range(1, max_r + 1):
        parts: list[str] = [str(r)]
        for c in range(1, 6):
            s = _serialize_cell_for_matrix_tsv(ws.cell(row=r, column=c).value)
            parts.append(base64.b64encode(s.encode("utf-8")).decode("ascii"))
        lines.append("\t".join(parts))
    path = _exclude_rules_matrix_vba_path()
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines) + "\n")
        _log_exclude_rules_sheet_debug(
            "MATRIX_TSV_WRITTEN",
            log_prefix,
            "????? A?E ? VBA ??? TSV ?????????openpyxl ???????",
            details=f"path={path} rows={max_r}",
        )
        return True
    except OSError as ex:
        logging.warning("%s: ?? VBA ? TSV ??????: %s", log_prefix, ex)
        return False


def _build_exclude_rules_list_from_openpyxl_ws(
    ws, c_proc: int, c_mach: int, c_flag: int, c_e: int
) -> list[dict]:
    """openpyxl ????????? _load_exclude_rules_from_workbook ???????????"""
    rules: list[dict] = []
    max_r = int(ws.max_row or 1)
    for r in range(2, max_r + 1):
        pv = ws.cell(row=r, column=c_proc).value
        proc = (
            ""
            if pv is None or (isinstance(pv, float) and pd.isna(pv))
            else str(pv).strip()
        )
        if not proc:
            continue
        mv = ws.cell(row=r, column=c_mach).value
        mach = (
            ""
            if mv is None or (isinstance(mv, float) and pd.isna(mv))
            else str(mv).strip()
        )
        cv = ws.cell(row=r, column=c_flag).value
        ev = ws.cell(row=r, column=c_e).value
        parsed = _parse_exclude_rule_json_cell(ev)
        rules.append(
            {"proc": proc, "mach": mach, "c_val": cv, "parsed": parsed}
        )
    return rules


def _set_exclude_rules_snapshot_from_ws(
    wb_path: str, ws, c_proc: int, c_mach: int, c_flag: int, c_e: int
) -> None:
    global _exclude_rules_rules_snapshot, _exclude_rules_snapshot_wb
    _exclude_rules_rules_snapshot = _build_exclude_rules_list_from_openpyxl_ws(
        ws, c_proc, c_mach, c_flag, c_e
    )
    _exclude_rules_snapshot_wb = os.path.normcase(os.path.abspath(wb_path))


def _clear_exclude_rules_e_apply_files() -> None:
    for p in (
        os.path.join(json_data_dir, EXCLUDE_RULES_E_SIDECAR_FILENAME),
        os.path.join(log_dir, EXCLUDE_RULES_E_SIDECAR_FILENAME),
    ):
        if os.path.isfile(p):
            try:
                os.remove(p)
            except OSError:
                pass
    for rel in (EXCLUDE_RULES_E_VBA_TSV_FILENAME, EXCLUDE_RULES_MATRIX_VBA_FILENAME):
        p = os.path.join(log_dir, rel)
        if os.path.isfile(p):
            try:
                os.remove(p)
            except OSError:
                pass


def _write_exclude_rules_e_vba_tsv_from_cells(
    wb_path: str, c_e: int, cells: dict[str, str], log_prefix: str
) -> None:
    """VBA ?: ???? Base64(UTF-8) ?????? TSV?"""
    lines = [
        "v1",
        "workbook\t" + os.path.abspath(wb_path),
        "sheet\t" + EXCLUDE_RULES_SHEET_NAME,
        "column_e\t" + str(int(c_e)),
        "---",
    ]
    for rk in sorted(cells.keys(), key=lambda x: int(x)):
        s = cells[rk]
        b64 = base64.b64encode(s.encode("utf-8")).decode("ascii")
        lines.append(rk + "\t" + b64)
    path_tsv = _exclude_rules_e_vba_tsv_path()
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path_tsv, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines) + "\n")
        _log_exclude_rules_sheet_debug(
            "E_VBA_TSV_WRITTEN",
            log_prefix,
            "E ?? VBA ??? TSV ?????????????????????????",
            details=f"path={path_tsv} cells={len(cells)}",
        )
    except OSError as ex:
        logging.warning("%s: E ? VBA ? TSV ??????: %s", log_prefix, ex)


def _write_exclude_rules_e_apply_artifacts(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> None:
    """
    E ?????? JSON ??????? VBA ? TSV ????????????????
    Python ?????? E ??? JSON ???????? E ???? TSV?
    """
    cells: dict[str, str] = {}
    max_r = int(ws.max_row or 1)
    for r in range(2, max_r + 1):
        ev = ws.cell(row=r, column=c_e).value
        if _cell_is_blank_for_rule(ev):
            continue
        s = str(ev).strip() if ev is not None else ""
        if not s:
            continue
        cells[str(r)] = s
    if not cells:
        _clear_exclude_rules_e_apply_files()
        return
    payload = {
        "version": 1,
        "workbook": os.path.abspath(wb_path),
        "sheet": EXCLUDE_RULES_SHEET_NAME,
        "column_e": c_e,
        "cells": cells,
    }
    path_sc = _exclude_rules_e_sidecar_path()
    try:
        os.makedirs(json_data_dir, exist_ok=True)
        with open(path_sc, "w", encoding="utf-8", newline="\n") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except OSError as ex:
        logging.warning("%s: E ? JSON ??????: %s", log_prefix, ex)
    _write_exclude_rules_e_vba_tsv_from_cells(wb_path, c_e, cells, log_prefix)
    _log_exclude_rules_sheet_debug(
        "E_APPLY_FILES_WRITTEN",
        log_prefix,
        "E ?? JSON ? VBA ? TSV ????????????? E ??????????????",
        details=f"cells={len(cells)}",
    )


def _try_apply_pending_exclude_rules_e_column(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> int:
    """
    ???????????????? JSON ?? E ???????
    ???????????????????????????????????
    """
    path_sc = _exclude_rules_e_sidecar_path()
    if not os.path.isfile(path_sc):
        return 0
    try:
        with open(path_sc, "r", encoding="utf-8") as f:
            payload = json.load(f)
    except Exception:
        return 0
    if int(payload.get("version") or 0) != 1:
        return 0
    target = os.path.normcase(os.path.abspath(wb_path))
    if os.path.normcase(str(payload.get("workbook") or "")) != target:
        return 0
    if str(payload.get("sheet") or "") != EXCLUDE_RULES_SHEET_NAME:
        return 0
    cells = payload.get("cells")
    if not isinstance(cells, dict):
        return 0
    n = 0
    for rk, val in cells.items():
        try:
            ri = int(rk)
        except (TypeError, ValueError):
            continue
        if ri < 2:
            continue
        if isinstance(val, dict):
            sval = json.dumps(val, ensure_ascii=False)
        else:
            sval = "" if val is None else str(val).strip()
        if not sval:
            continue
        ws.cell(row=ri, column=c_e, value=sval)
        n += 1
    try:
        os.remove(path_sc)
    except OSError:
        pass
    if n:
        _log_exclude_rules_sheet_debug(
            "E_SIDECAR_APPLIED",
            log_prefix,
            f"?????? E ?????????? {n} ?????????",
            details=path_sc,
        )
        logging.info(
            "%s: %s ??????????????????????????????????",
            log_prefix,
            path_sc,
        )
    return n


def _read_exclude_rules_d_cells_data_only_for_rows(
    wb_path: str, rows: list[int], c_d: int
) -> dict[int, object]:
    """
    D ????????openpyxl ??????? '=...' ???????
    data_only=True ???????????Excel ?????????????????????
    """
    out: dict[int, object] = {}
    if not rows or not os.path.isfile(wb_path):
        return out
    if _workbook_should_skip_openpyxl_io(wb_path):
        return out
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wbro = None
    try:
        wbro = load_workbook(
            wb_path,
            read_only=True,
            data_only=True,
            keep_vba=keep_vba,
        )
    except Exception:
        return out
    try:
        if EXCLUDE_RULES_SHEET_NAME not in wbro.sheetnames:
            return out
        wsro = wbro[EXCLUDE_RULES_SHEET_NAME]
        for r in rows:
            if r < 2:
                continue
            try:
                out[r] = wsro.cell(row=r, column=c_d).value
            except Exception:
                pass
    finally:
        if wbro is not None:
            try:
                wbro.close()
            except Exception:
                pass
    return out


def run_exclude_rules_sheet_maintenance(
    wb_path: str, pairs: list[tuple[str, str]], log_prefix: str
) -> None:
    """
    ???_????????????D?E ? AI ????????????? xlwings ? A?E ???Save?``EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1`` ??? openpyxl save ?????

    xlwings ??????????? ``log/exclude_rules_matrix_vba.tsv`` ???????
    ``??_??????_A??E_TSV????`` ? A?E ??????
    ???????? E ???? ``exclude_rules_e_column_vba.tsv`` ?????????? TSV ???????????
    ?????? TSV/JSON ???????

    ``json/exclude_rules_e_column_pending.json`` ? Python ?????? E ?????
    ????????? 1 ?????? VBA???_??????_????????
    """
    if not wb_path:
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_PATH",
            log_prefix,
            "TASK_INPUT_WORKBOOK ??????????????????",
        )
        return
    if not os.path.exists(wb_path):
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_FILE",
            log_prefix,
            "???????????",
            details=f"path={wb_path}",
        )
        return

    _log_exclude_rules_sheet_debug(
        "START",
        log_prefix,
        "?????????",
        details=f"path={wb_path} pairs={len(pairs)}",
    )
    global _exclude_rules_effective_read_path
    _exclude_rules_effective_read_path = None

    if _workbook_should_skip_openpyxl_io(wb_path):
        _log_exclude_rules_sheet_debug(
            "SKIP_OPENPYXL_INCOMPATIBLE_BOOK",
            log_prefix,
            f"?????{OPENPYXL_INCOMPATIBLE_SHEET_MARKER}?????????openpyxl ?????????????????",
            details=f"path={wb_path}",
        )
        logging.warning(
            "%s: ?%s???????%s?? openpyxl ????????????Excel?xlwings ???????????",
            log_prefix,
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
            EXCLUDE_RULES_SHEET_NAME,
        )
        return

    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = None
    try:
        wb = load_workbook(wb_path, keep_vba=keep_vba, read_only=False, data_only=False)
    except Exception as e1:
        if keep_vba:
            _log_exclude_rules_sheet_debug(
                "OPEN_RETRY",
                log_prefix,
                "keep_vba=True ???????? keep_vba=False ?????????????????????",
                exc=e1,
            )
            try:
                wb = load_workbook(wb_path, keep_vba=False, read_only=False, data_only=False)
            except Exception as e2:
                _log_exclude_rules_sheet_debug(
                    "OPEN_FAIL",
                    log_prefix,
                    "?????????????????????????",
                    details=f"path={wb_path}",
                    exc=e2,
                )
                return
        else:
            _log_exclude_rules_sheet_debug(
                "OPEN_FAIL",
                log_prefix,
                "?????????????????????????",
                details=f"path={wb_path}",
                exc=e1,
            )
            return

    _log_exclude_rules_sheet_debug(
        "OPEN_OK",
        log_prefix,
        "??????????",
        details=f"keep_vba={keep_vba} sheets={len(wb.sheetnames)}",
    )

    try:
        if EXCLUDE_RULES_SHEET_NAME not in wb.sheetnames:
            _log_exclude_rules_sheet_debug(
                "SKIP_NO_SHEET",
                log_prefix,
                "??????????VBA ????_??????_????????????????1/2 ???????????????",
                details=f"path={wb_path}",
            )
            logging.error(
                "%s: ?%s????????Python ?????????????",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
            )
            return

        ws = wb[EXCLUDE_RULES_SHEET_NAME]
        hm_before = _exclude_rules_sheet_header_map(ws)
        c_proc, c_mach, c_flag, c_d, c_e = _ensure_exclude_rules_sheet_headers_and_columns(
            ws, log_prefix
        )
        hm_after = _exclude_rules_sheet_header_map(ws)
        if tuple(hm_before.get(x) for x in (
            EXCLUDE_RULE_COL_PROCESS,
            EXCLUDE_RULE_COL_MACHINE,
            EXCLUDE_RULE_COL_FLAG,
            EXCLUDE_RULE_COL_LOGIC_JA,
            EXCLUDE_RULE_COL_LOGIC_JSON,
        )) != tuple(hm_after.get(x) for x in (
            EXCLUDE_RULE_COL_PROCESS,
            EXCLUDE_RULE_COL_MACHINE,
            EXCLUDE_RULE_COL_FLAG,
            EXCLUDE_RULE_COL_LOGIC_JA,
            EXCLUDE_RULE_COL_LOGIC_JSON,
        )):
            _log_exclude_rules_sheet_debug(
                "HEADER_FIX",
                log_prefix,
                "1????????????????????????????????",
                details=f"cols=({c_proc},{c_mach},{c_flag},{c_d},{c_e})",
            )

        # ?????????????????? E ????????????????????????????
        _try_apply_pending_exclude_rules_e_column(wb_path, ws, c_e, log_prefix)

        existing_keys: set[tuple[str, str]] = set()
        max_r = max(2, int(ws.max_row or 2))
        for r in range(2, max_r + 1):
            pv = ws.cell(row=r, column=c_proc).value
            mv = ws.cell(row=r, column=c_mach).value
            p = str(pv).strip() if pv is not None and not (isinstance(pv, float) and pd.isna(pv)) else ""
            m = str(mv).strip() if mv is not None and not (isinstance(mv, float) and pd.isna(mv)) else ""
            if not p:
                continue
            existing_keys.add(
                (_normalize_process_name_for_rule_match(p), _normalize_equipment_match_key(m))
            )

        added = 0
        for p, m in pairs:
            key = (_normalize_process_name_for_rule_match(p), _normalize_equipment_match_key(m))
            if key in existing_keys:
                continue
            ws.append([p, m, None, None, None])
            existing_keys.add(key)
            added += 1
        if added:
            _log_exclude_rules_sheet_debug(
                "SYNC_ROWS",
                log_prefix,
                f"??+????? {added} ????????",
            )
            logging.info(
                "%s: ?%s????+????????? %s ????????",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                added,
            )

        # ?????????1?????????????????????????????????????
        if added == 0 and not existing_keys:
            ws.append(["??", "", "yes", "", ""])
            existing_keys.add(
                (_normalize_process_name_for_rule_match("??"), _normalize_equipment_match_key(""))
            )
            _log_exclude_rules_sheet_debug(
                "EXAMPLE_ROW",
                log_prefix,
                "???????????????=yes??1????",
            )
            logging.info(
                "%s: ?%s??????????????????=yes??1????????",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
            )

        # ????? AI ??????????????????????????????????
        n_kept, n_removed_empty = _compact_exclude_rules_data_rows(
            ws, c_proc, c_mach, c_flag, c_d, c_e, log_prefix
        )
        if n_removed_empty:
            _log_exclude_rules_sheet_debug(
                "DATA_COMPACT",
                log_prefix,
                "??????????????????????????AI ??????",
                details=f"rows={n_kept} removed_empty={n_removed_empty}",
            )

        max_r = int(ws.max_row or 1)
        pending_rows: list[int] = []
        for r in range(2, max_r + 1):
            dv = ws.cell(row=r, column=c_d).value
            ev = ws.cell(row=r, column=c_e).value
            # C ??????????D ?????? E ???? D?E ???
            if _cell_is_blank_for_rule(dv):
                continue
            if not _cell_is_blank_for_rule(ev):
                continue
            pending_rows.append(r)

        # D ????????????? '=...' ??????data_only ?????????????
        formula_rows = [
            r
            for r in pending_rows
            if isinstance(ws.cell(row=r, column=c_d).value, str)
            and str(ws.cell(row=r, column=c_d).value).strip().startswith("=")
        ]
        d_cached = (
            _read_exclude_rules_d_cells_data_only_for_rows(wb_path, formula_rows, c_d)
            if formula_rows
            else {}
        )
        pending_texts: list[str] = []
        filtered_rows: list[int] = []
        for r in pending_rows:
            dv = ws.cell(row=r, column=c_d).value
            blob = (
                ""
                if dv is None or (isinstance(dv, float) and pd.isna(dv))
                else str(dv).strip()
            )
            if blob.startswith("="):
                alt = d_cached.get(r)
                if alt is not None and not (isinstance(alt, float) and pd.isna(alt)):
                    blob = str(alt).strip()
                else:
                    logging.warning(
                        "%s: ?%s?%s ??? D ??????????????????????Excel ???????? D ???????????",
                        log_prefix,
                        EXCLUDE_RULES_SHEET_NAME,
                        r,
                    )
                    continue
            if _cell_is_blank_for_rule(blob):
                continue
            filtered_rows.append(r)
            pending_texts.append(blob)
        pending_rows = filtered_rows

        ai_filled = 0
        ai_e_cell_addrs: list[str] = []
        if pending_texts:
            parsed_list = _ai_compile_exclude_rule_logics_batch(pending_texts)
            for r, parsed in zip(pending_rows, parsed_list):
                if not parsed:
                    logging.warning(
                        "%s: ?%s?%s ??? D ?? JSON ??????????API??????????",
                        log_prefix,
                        EXCLUDE_RULES_SHEET_NAME,
                        r,
                    )
                    continue
                jstr = json.dumps(parsed, ensure_ascii=False)
                ws.cell(row=r, column=c_e, value=jstr)
                cell_addr = f"{get_column_letter(c_e)}{r}"
                ai_e_cell_addrs.append(cell_addr)
                preview = jstr if len(jstr) <= 160 else (jstr[:160] + "�")
                logging.info(
                    "%s: ?%s????????%s??? %s ? JSON ?????: %s",
                    log_prefix,
                    EXCLUDE_RULES_SHEET_NAME,
                    EXCLUDE_RULE_COL_LOGIC_JSON,
                    cell_addr,
                    preview,
                )
                ai_filled += 1
        if ai_filled:
            _log_exclude_rules_sheet_debug(
                "AI_E_FILLED",
                log_prefix,
                f"D?E ? AI ??? {ai_filled} ????",
                details="cells=" + ",".join(ai_e_cell_addrs),
            )
            logging.info(
                "%s: ?%s?? D?E ? AI ??? %s ????: %s??",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                ai_filled,
                ",".join(ai_e_cell_addrs),
            )

        _er_test = os.environ.get("EXCLUDE_RULES_TEST_E1234", "").strip().lower()
        if _er_test in ("1", "yes", "true"):
            try:
                _er_row = int(os.environ.get("EXCLUDE_RULES_TEST_E1234_ROW", "9") or "9")
            except ValueError:
                _er_row = 9
            if _er_row < 2:
                _er_row = 9
            ws.cell(row=_er_row, column=c_e, value="1234")
            _e_addr = f"{get_column_letter(c_e)}{_er_row}"
            _log_exclude_rules_sheet_debug(
                "TEST_E1234",
                log_prefix,
                f'E? {_e_addr} ????? "1234" ?????',
                details=f"row={_er_row}",
            )
            logging.warning(
                '%s: ?????%s ? "1234" ??????EXCLUDE_RULES_TEST_E1234??',
                log_prefix,
                _e_addr,
            )

        _set_exclude_rules_snapshot_from_ws(
            wb_path, ws, c_proc, c_mach, c_flag, c_e
        )
        _write_exclude_rules_e_apply_artifacts(wb_path, ws, c_e, log_prefix)
        persisted = _persist_exclude_rules_workbook(wb, wb_path, ws, log_prefix)
        if not persisted:
            logging.warning(
                "%s: ?????? openpyxl ??????????"
                " log ??? TSV ???????_??????_A??E_TSV??????"
                "??? E ??????_??????_E?_TSV???????????????",
                log_prefix,
            )
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "FATAL",
            log_prefix,
            "??????????????????????",
            exc=ex,
        )
        logging.exception("%s: ??_????????????", log_prefix)
    finally:
        if wb is not None:
            wb.close()
            _log_exclude_rules_sheet_debug("CLOSED", log_prefix, "?????????????")


def _resolve_exclude_rules_workbook_path_for_read(wb_path: str) -> str:
    """???????????????????????????????????????"""
    p = _exclude_rules_effective_read_path
    if p and os.path.exists(p):
        return p
    return wb_path


def _load_exclude_rules_from_workbook(wb_path: str) -> list[dict]:
    """???????????????????????"""
    if not wb_path:
        return []
    global _exclude_rules_rules_snapshot, _exclude_rules_snapshot_wb
    ap_arg = os.path.normcase(os.path.abspath(wb_path))
    if (
        _exclude_rules_rules_snapshot is not None
        and _exclude_rules_snapshot_wb == ap_arg
    ):
        snap = list(_exclude_rules_rules_snapshot)
        _exclude_rules_rules_snapshot = None
        _exclude_rules_snapshot_wb = None
        return snap
    path = _resolve_exclude_rules_workbook_path_for_read(wb_path)
    if not os.path.exists(path):
        return []
    if _workbook_should_skip_openpyxl_io(path):
        logging.warning(
            "???????: ?????%s?????? pandas(openpyxl) ???%s??????????????????????",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
            EXCLUDE_RULES_SHEET_NAME,
        )
        return []
    try:
        df = pd.read_excel(path, sheet_name=EXCLUDE_RULES_SHEET_NAME)
    except Exception:
        return []
    df.columns = df.columns.str.strip()
    need = [EXCLUDE_RULE_COL_PROCESS, EXCLUDE_RULE_COL_MACHINE]
    for c in need:
        if c not in df.columns:
            return []
    rules = []
    for _, row in df.iterrows():
        proc = str(row.get(EXCLUDE_RULE_COL_PROCESS, "") or "").strip()
        if not proc:
            continue
        mach = str(row.get(EXCLUDE_RULE_COL_MACHINE, "") or "").strip()
        c_val = row.get(EXCLUDE_RULE_COL_FLAG)
        e_raw = row.get(EXCLUDE_RULE_COL_LOGIC_JSON)
        parsed = _parse_exclude_rule_json_cell(e_raw)
        rules.append(
            {
                "proc": proc,
                "mach": mach,
                "c_val": c_val,
                "parsed": parsed,
            }
        )
    return rules


def apply_exclude_rules_config_to_plan_df(
    df: pd.DataFrame, wb_path: str, log_prefix: str
) -> pd.DataFrame:
    """???????????????????C=yes ??? E ? JSON ????"""
    if df is None or df.empty:
        return df
    if TASK_COL_MACHINE not in df.columns or PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns:
        return df
    rules = _load_exclude_rules_from_workbook(wb_path)
    if not rules:
        return df
    df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT].astype(object)
    n = 0
    for i in df.index:
        try:
            row = df.loc[i]
        except Exception:
            continue
        tp = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        tm = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        if not tp:
            continue
        for ru in rules:
            if not _task_row_matches_exclude_rule_target(tp, tm, ru["proc"], ru["mach"]):
                continue
            if _exclude_rule_c_column_is_yes(ru["c_val"]):
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
                n += 1
                break
            if ru.get("parsed") and evaluate_exclude_rule_json_for_row(ru["parsed"], row):
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
                n += 1
                break
    if n:
        logging.info("%s: ???%s????????=yes ? %s ?????????", log_prefix, EXCLUDE_RULES_SHEET_NAME, n)
    return df


def _sort_stage1_plan_df_by_dispatch_trial_order_asc(plan_df: "pd.DataFrame") -> "pd.DataFrame":
    """
    ??1????: ????????????????? DataFrame ????
    ???????????????????????
    """
    col = RESULT_TASK_COL_DISPATCH_TRIAL_ORDER
    if plan_df is None or getattr(plan_df, "empty", True) or col not in plan_df.columns:
        return plan_df
    dto_positions = [i for i, c in enumerate(plan_df.columns) if c == col]
    if not dto_positions:
        return plan_df
    loc = dto_positions[0]
    n = len(plan_df)
    if n <= 1:
        return plan_df
    keys = []
    for i in range(n):
        dto = parse_optional_int(plan_df.iat[i, loc])
        if dto is not None and dto >= 1:
            keys.append((0, int(dto), i))
        else:
            keys.append((1, 10**9, i))
    order = sorted(range(n), key=lambda j: keys[j])
    if order == list(range(n)):
        return plan_df
    return plan_df.iloc[order].reset_index(drop=True)


# =============================================================================
# ??1?????task_extract_stage1.py ? run_stage1_extract?
#   ????DATA ?? ? ???????? ? ??????? ? plan_input_tasks.xlsx ??
# =============================================================================
def run_stage1_extract():
    """
    ??1: ????DATA ?????????????? output/plan_input_tasks.xlsx ????
    ????NO????????????????????????????????? yes ????????
    ??????????_??????????+??????????????AI????????????? VBA??
    """
    if not TASKS_INPUT_WORKBOOK:
        logging.error("TASK_INPUT_WORKBOOK ???????")
        return False
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        logging.error(f"TASK_INPUT_WORKBOOK ???????: {TASKS_INPUT_WORKBOOK}")
        return False
    reset_gemini_usage_tracker()
    df_src = load_tasks_df()
    try:
        _pm_pairs = _collect_process_machine_pairs_for_exclude_rules(df_src)
        run_exclude_rules_sheet_maintenance(TASKS_INPUT_WORKBOOK, _pm_pairs, "??1")
    except Exception:
        logging.exception("??1: ??_????????????????")
    records = []
    for _, row in df_src.iterrows():
        if row_has_completion_keyword(row):
            continue
        task_id = planning_task_id_str_from_scalar(row.get(TASK_COL_TASK_ID))
        machine = str(row.get(TASK_COL_MACHINE, "")).strip()
        machine_name = str(row.get(TASK_COL_MACHINE_NAME, "")).strip()
        qty_total = parse_float_safe(row.get(TASK_COL_QTY), 0.0)
        done_qty = calc_done_qty_equivalent_from_row(row)
        qty = max(0.0, qty_total - done_qty)
        if qty <= 0 or not machine or not task_id:
            continue
        rec = {c: row.get(c) for c in SOURCE_BASE_COLUMNS}
        rec[TASK_COL_TASK_ID] = task_id
        _pn_stage1 = row.get(TASK_COL_PRODUCT, None)
        _qty_total_s1 = parse_float_safe(row.get(TASK_COL_QTY), 0.0)
        _roll_len = infer_unit_m_from_product_name(
            _pn_stage1, fallback_unit=_qty_total_s1 if _qty_total_s1 > 0 else qty
        )
        try:
            _roll_len = float(_roll_len)
        except (TypeError, ValueError):
            _roll_len = _qty_total_s1 if _qty_total_s1 > 0 else qty
        if _roll_len <= 0:
            _roll_len = _qty_total_s1 if _qty_total_s1 > 0 else max(qty, 1e-9)
        rec[PLAN_COL_ROLL_UNIT_LENGTH] = _roll_len
        # ??? + ??? ?�??�??????????????????????
        if machine_name:
            rec[PLAN_COL_PROCESS_FACTOR] = f"{machine}+{machine_name}"
        else:
            rec[PLAN_COL_PROCESS_FACTOR] = f"{machine}+"
        rec[PLAN_COL_SPEED_OVERRIDE] = ""
        rec[PLAN_COL_RAW_INPUT_DATE_OVERRIDE] = ""
        rec[PLAN_COL_PREFERRED_OP] = ""
        rec[PLAN_COL_SPECIAL_REMARK] = ""
        rec[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = ""
        rec[PLAN_COL_AI_PARSE] = ""
        records.append(rec)
    if not records:
        logging.warning("??1: ??????????????")
    order = plan_input_sheet_column_order()
    out_df = pd.DataFrame(records)
    if out_df.empty:
        out_df = pd.DataFrame(columns=order)
    else:
        out_df = out_df.reindex(columns=order).fillna("")
    if PLAN_COL_EXCLUDE_FROM_ASSIGNMENT in out_df.columns:
        out_df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = out_df[
            PLAN_COL_EXCLUDE_FROM_ASSIGNMENT
        ].astype(object)
    if RESULT_TASK_COL_DISPATCH_TRIAL_ORDER in out_df.columns:
        out_df[RESULT_TASK_COL_DISPATCH_TRIAL_ORDER] = ""
        out_df[RESULT_TASK_COL_DISPATCH_TRIAL_ORDER] = out_df[
            RESULT_TASK_COL_DISPATCH_TRIAL_ORDER
        ].astype(object)
    try:
        (
            _skills_d_stage1,
            _members_stage1,
            equipment_list_stage1,
            req_map,
            need_rules,
            _surplus_map_stage1,
            need_combo_col_index_stage1,
        ) = load_skills_and_needs()
    except PlanningValidationError:
        logging.error("??1???: ??? skills ???????????????????")
        raise
    except Exception as e:
        logging.info("??1: ??? need ??????? need ??????? (%s)", e)
        req_map, need_rules = {}, []
        equipment_list_stage1 = []
        need_combo_col_index_stage1 = {}
    out_df = _merge_plan_sheet_user_overrides(out_df)
    _refresh_plan_reference_columns(out_df, req_map, need_rules)
    try:
        _apply_auto_exclude_bunkatsu_duplicate_machine(out_df, log_prefix="??1")
    except Exception as ex:
        logging.exception("??1: ??????????????????????: %s", ex)
    try:
        out_df = apply_exclude_rules_config_to_plan_df(out_df, TASKS_INPUT_WORKBOOK, "??1")
    except Exception as ex:
        logging.warning("??1: ?????????????????????: %s", ex)
    try:
        _ext_dt_s1 = _extract_data_extraction_datetime()
        _run_d_s1 = _ext_dt_s1.date() if _ext_dt_s1 is not None else datetime.now().date()
        fill_plan_dispatch_trial_order_column_stage1(
            out_df,
            _run_d_s1,
            req_map,
            need_rules,
            need_combo_col_index_stage1,
            equipment_list_stage1,
        )
    except Exception as ex:
        logging.warning("??1: ???????????????????????: %s", ex)
    out_df = _sort_stage1_plan_df_by_dispatch_trial_order_asc(out_df)
    out_path = os.path.join(output_dir, STAGE1_OUTPUT_FILENAME)
    out_df.to_excel(out_path, sheet_name="?????", index=False)
    _apply_excel_date_columns_date_only_display(out_path, "?????")
    _apply_plan_input_visual_format(out_path, "?????")
    logging.info(f"??1??: '{out_path}' ???????????? '{PLAN_INPUT_SHEET_NAME}' ???????????")
    _try_write_main_sheet_gemini_usage_summary("??1")
    return True


# ?????????????2026?3????
TARGET_YEAR = 2026
TARGET_MONTH = 3
DEFAULT_START_TIME = time(8, 45)
DEFAULT_END_TIME = time(17, 0)
DEFAULT_BREAKS = [
    (time(12, 0), time(12, 50)),
    (time(14, 45), time(15, 0))
]
# ????????: ASSIGN_END_OF_DAY_DEFER_MINUTES ??????team_end_limit ??????????????
# ?? remaining_units??????? ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS ???????????????None??
# ?????????ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS ????????????????????????
# ???????????_eod_reject_capacity_units_below_threshold??
# ASSIGN_END_OF_DAY_DEFER_MINUTES ?? 45????0 ????????????????
# ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS ?? 5?????????: 999999?????????????????????????
# ??: ????????? _defer_team_start_past_prebreak_and_end_of_day ???????????
# ??????????? _contiguous_work_minutes_until_next_break_or_limit ????
# ?? ASSIGN_DEFER_MIN_REMAINING_ROLLS / ASSIGN_PRE_BREAK_DEFER_GAP_MINUTES ???????
ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS = max(
    0,
    int(os.environ.get("ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS", "5").strip() or 0),
)
ASSIGN_END_OF_DAY_DEFER_MINUTES = max(
    0,
    int(os.environ.get("ASSIGN_END_OF_DAY_DEFER_MINUTES", "45").strip() or 0),
)


def _eod_minutes_window_covers_start(
    team_start: datetime, team_end_limit: datetime
) -> bool:
    """ASSIGN_END_OF_DAY_DEFER_MINUTES ??????????????????????"""
    gap = ASSIGN_END_OF_DAY_DEFER_MINUTES
    if gap <= 0:
        return False
    if team_start >= team_end_limit:
        return False
    return (team_end_limit - team_start) <= timedelta(minutes=gap)


def _eod_reject_capacity_units_below_threshold(
    units_fit_until_close: int, team_start: datetime, team_end_limit: datetime
) -> bool:
    """
    ?????? ASSIGN_END_OF_DAY_DEFER_MINUTES ????????????
    ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS ?????????????????????????? True?????????????????
    """
    th = ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS
    if th <= 0:
        return False
    if not _eod_minutes_window_covers_start(team_start, team_end_limit):
        return False
    return int(units_fit_until_close) < int(th)


# =========================================================
# 1. ???????? (?????)
#    ???????????????????????????????????
# =========================================================
def merge_time_intervals(intervals):
    """???????????????????????????"""
    if not intervals:
        return []
    intervals.sort(key=lambda x: x[0])
    merged = [intervals[0]]
    for current_start, current_end in intervals[1:]:
        last_start, last_end = merged[-1]
        if current_start <= last_end:
            merged[-1] = (last_start, max(last_end, current_end))
        else:
            merged.append((current_start, current_end))
    return merged


def _contiguous_work_minutes_until_next_break_or_limit(
    start_dt: datetime,
    breaks_dt: list,
    end_limit_dt: datetime,
) -> int:
    """
    start_dt ??????????????????????????????????
    ????????? 0???????????breaks_dt ? merge ?????
    """
    if start_dt >= end_limit_dt:
        return 0
    for bs, be in breaks_dt:
        if bs <= start_dt < be:
            return 0
    next_stop = end_limit_dt
    for bs, be in breaks_dt:
        if be <= start_dt:
            continue
        if start_dt < bs:
            next_stop = min(next_stop, bs)
    return max(0, int((next_stop - start_dt).total_seconds() / 60))


def _break_end_to_skip_if_contiguous_under(
    start_dt: datetime,
    breaks_dt: list,
    end_limit_dt: datetime,
    min_contiguous_mins: int,
) -> datetime | None:
    """
    ????????????????????? min_contiguous_mins ?????
    ?????????????????????? 1 ???????????????????
    ???????????????? None?
    """
    if min_contiguous_mins <= 0:
        return None
    if start_dt >= end_limit_dt:
        return None
    c = _contiguous_work_minutes_until_next_break_or_limit(
        start_dt, breaks_dt, end_limit_dt
    )
    if c >= min_contiguous_mins:
        return None
    next_stop = end_limit_dt
    for bs, be in breaks_dt:
        if be <= start_dt:
            continue
        if start_dt < bs:
            next_stop = min(next_stop, bs)
    if next_stop >= end_limit_dt:
        return None
    for bs, be in breaks_dt:
        if bs == next_stop:
            return be
    return None


def _defer_team_start_past_prebreak_and_end_of_day(
    task: dict,
    team: tuple,
    team_start: datetime,
    team_end_limit: datetime,
    team_breaks: list,
    refloor_fn,
    min_contiguous_work_mins: int | None = None,
) -> datetime | None:
    """
    - ASSIGN_END_OF_DAY_DEFER_MINUTES > 0 ?? (team_end_limit - ????) ?????????
      remaining_units ????? ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS ?????????????None??
    - ????????????? **???????????**??`refloor_fn` ??????avail ???????
      ????????????EOD ?????????? None?
    - min_contiguous_work_mins ?????????? **???????????**???????
      ?????????????????? refloor ??????
    """
    _tid = str(task.get("task_id", "") or "").strip()
    _team_txt = ", ".join(str(x) for x in team) if team else "�"

    def _trace_block(msg: str, *a) -> None:
        if not _trace_schedule_task_enabled(_tid):
            return
        _log_dispatch_trace_schedule(
            _tid,
            "[?????? task=%s] ??????: " + msg,
            _tid,
            *a,
        )

    ts = refloor_fn(team_start)
    for _ in range(64):
        if ts >= team_end_limit:
            _trace_block(
                "????(????) machine=%s team=%s rem=%.4f trial_start=%s end_limit=%s",
                task.get("machine"),
                _team_txt,
                float(task.get("remaining_units") or 0),
                ts,
                team_end_limit,
            )
            return None

        break_end = None
        for bs, be in team_breaks:
            if bs <= ts < be:
                break_end = be
                break
        if break_end is not None:
            _trace_block(
                "?????????????? machine=%s team=%s rem=%.4f break_end=%s trial_was=%s",
                task.get("machine"),
                _team_txt,
                float(task.get("remaining_units") or 0),
                break_end,
                ts,
            )
            ts = refloor_fn(break_end)
            continue

        if min_contiguous_work_mins is not None and min_contiguous_work_mins > 0:
            slip_end = _break_end_to_skip_if_contiguous_under(
                ts, team_breaks, team_end_limit, min_contiguous_work_mins
            )
            if slip_end is not None:
                _trace_block(
                    "??????????????????????? machine=%s team=%s rem=%.4f need_contig_min=%s trial_was=%s break_end=%s",
                    task.get("machine"),
                    _team_txt,
                    float(task.get("remaining_units") or 0),
                    min_contiguous_work_mins,
                    ts,
                    slip_end,
                )
                ts = refloor_fn(slip_end)
                continue

        gap_end = ASSIGN_END_OF_DAY_DEFER_MINUTES
        rem_ceil = math.ceil(float(task.get("remaining_units") or 0))
        if (
            gap_end > 0
            and (team_end_limit - ts) <= timedelta(minutes=gap_end)
            and rem_ceil <= ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS
        ):
            _trace_block(
                "????(??????????) machine=%s team=%s rem_ceil=%s max_rem=%s trial_start=%s end_limit=%s gap_end_min=%s",
                task.get("machine"),
                _team_txt,
                rem_ceil,
                ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS,
                ts,
                team_end_limit,
                gap_end,
            )
            return None

        return ts

    _trace_block(
        "????(?????????) machine=%s team=%s rem=%.4f trial_start=%s",
        task.get("machine"),
        _team_txt,
        float(task.get("remaining_units") or 0),
        ts,
    )
    return None


def _expand_timeline_events_for_equipment_grid(timeline_events: list) -> list:
    """
    ????????????????????????????
    1 ???????????????e["date"] ???????????????????????
    start_dt?end_dt ????? DEFAULT_START_TIME?DEFAULT_END_TIME ???????????????
    """
    expanded: list = []
    for e in timeline_events:
        sd = e.get("start_dt")
        ed = e.get("end_dt")
        if not isinstance(sd, datetime) or not isinstance(ed, datetime):
            expanded.append(e)
            continue
        if ed <= sd:
            expanded.append(e)
            continue
        segments: list = []
        cal = sd.date()
        last_d = ed.date()
        while cal <= last_d:
            w0 = datetime.combine(cal, DEFAULT_START_TIME)
            w1 = datetime.combine(cal, DEFAULT_END_TIME)
            a = max(sd, w0)
            b = min(ed, w1)
            if a < b:
                ne = dict(e)
                ne["date"] = cal
                ne["start_dt"] = a
                ne["end_dt"] = b
                segments.append(ne)
            cal += timedelta(days=1)
        if segments:
            expanded.extend(segments)
        else:
            expanded.append(e)
    return expanded


def get_actual_work_minutes(start_dt, end_dt, breaks_dt):
    """
    start_dt ?? end_dt ????????????????
    breaks_dt � (????, ????) ???datetime ??? time??????????????????
    """
    current = start_dt
    actual_mins = 0
    while current < end_dt:
        next_event = end_dt
        in_break = False
        b_end_time = None
        for b_s, b_e in breaks_dt:
            if b_s <= current < b_e:
                in_break = True
                b_end_time = b_e
                break
            elif current < b_s < next_event:
                next_event = b_s
        
        if in_break:
            current = b_end_time
        else:
            actual_mins += int((next_event - current).total_seconds() / 60)
            current = next_event
    return actual_mins


def calculate_end_time(start_dt, duration_minutes, breaks_dt, end_limit_dt):
    """
    start_dt ???? duration_minutes ?????? datetime ??????????????
    end_limit_dt ???????????????: (????, ?????????, ??????)
    """
    current = start_dt
    remaining_work = duration_minutes
    actual_work_time = 0 

    while current < end_limit_dt and remaining_work > 0:
        next_event = end_limit_dt
        in_break = False
        break_end = None
        for b_start, b_end in breaks_dt:
            if b_start <= current < b_end:
                in_break = True
                break_end = b_end
                break
            elif current < b_start < next_event:
                next_event = b_start
                
        if in_break:
            current = break_end
            continue
            
        block_mins = int((next_event - current).total_seconds() / 60)
        if remaining_work <= block_mins:
            actual_work_time += remaining_work
            current += timedelta(minutes=remaining_work)
            remaining_work = 0
        else:
            actual_work_time += block_mins
            remaining_work -= block_mins
            current = next_event

    end_dt = min(current, end_limit_dt)
    return end_dt, actual_work_time, remaining_work

def match_need_sheet_condition(condition_raw: str, task_id: str) -> bool:
    """
    need ??????NO????????
    ??*??? ? ??????
    prefix:ABC / ???:ABC ? ??NO ??????????
    regex:... / ????:... ? ??????????
    ???????????????????????????NO?JR?�?? JR ?????????JR???
    """
    cond = (condition_raw or "").strip()
    tid = str(task_id).strip()
    if not cond or cond in ("*", "??", "??", "any", "ANY"):
        return True
    low = cond.lower()
    cn = cond.replace("?", ":")
    if low.startswith("prefix:") or low.startswith("???:"):
        pref = cn.split(":", 1)[1].strip() if ":" in cn else ""
        return bool(pref) and tid.startswith(pref)
    if low.startswith("regex:") or low.startswith("????:"):
        pat = cn.split(":", 1)[1].strip() if ":" in cn else ""
        if not pat:
            return False
        try:
            return re.search(pat, tid) is not None
        except re.error:
            logging.warning(f"need ??NO????????????: {pat}")
            return False
    if "??" in cond and "JR" in cond.upper():
        return tid.upper().startswith("JR")
    return tid.startswith(cond)


def parse_need_sheet_special_rules(needs_df, label_col, equipment_list, cond_col):
    """????1?99 ????????????????1?99???????????????????"""
    rules = []
    for _, row in needs_df.iterrows():
        lab = str(row.get(label_col, "") or "").strip()
        m = re.match(r"????\s*(\d+)", lab)
        if not m:
            continue
        order = int(m.group(1))
        if order < 1 or order > 99:
            continue
        cond = ""
        if cond_col is not None:
            v = row.get(cond_col)
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                cond = str(v).strip()
        overrides = {}
        for eq in equipment_list:
            v = row.get(eq)
            n = parse_optional_int(v)
            if n is not None and 1 <= n <= 99:
                overrides[str(eq).strip()] = n
        if not overrides:
            continue
        rules.append({"order": order, "condition": cond, "overrides": overrides})
    rules.sort(key=lambda r: r["order"])
    return rules


def resolve_need_required_op(process: str, machine_name: str, task_id: str, req_map: dict, need_rules: list) -> int:
    """
    need ???????? + ???????OP??????????1?99? order ??????????

    req_map ?
      - f\"{process}+{machine_name}\"??????
      - machine_name??????????????
      - process??????????????
    ?????? base ???????
    need_rules ? overrides ??????????
    """
    p = str(process).strip()
    m = str(machine_name).strip()

    combo_key = f"{p}+{m}" if p and m else None

    base = None
    if combo_key and combo_key in req_map:
        base = req_map.get(combo_key)
    if base is None and m:
        base = req_map.get(m)
    if base is None and p:
        base = req_map.get(p)
    if base is None:
        base = 1

    for rule in need_rules:
        if not match_need_sheet_condition(rule["condition"], task_id):
            continue

        if combo_key and combo_key in rule["overrides"]:
            return int(rule["overrides"][combo_key])
        if m and m in rule["overrides"]:
            return int(rule["overrides"][m])
        if p and p in rule["overrides"]:
            return int(rule["overrides"][p])

    return int(base)


def resolve_need_required_op_explain(
    process: str, machine_name: str, task_id: str, req_map: dict, need_rules: list
) -> tuple[int, str]:
    """
    resolve_need_required_op ???????????????????????????
    """
    p = str(process).strip()
    m = str(machine_name).strip()
    combo_key = f"{p}+{m}" if p and m else None
    base = None
    base_src = ""
    if combo_key and combo_key in req_map:
        base = req_map.get(combo_key)
        base_src = f"req_map[{combo_key!r}]={base}"
    elif m and m in req_map:
        base = req_map[m]
        base_src = f"req_map[????? {m!r}]={base}????????"
    elif p and p in req_map:
        base = req_map[p]
        base_src = f"req_map[????? {p!r}]={base}???????????"
    else:
        base = 1
        base_src = "req_map???????1"
    for rule in need_rules:
        if not match_need_sheet_condition(rule["condition"], task_id):
            continue
        order = rule.get("order", "?")
        if combo_key and combo_key in rule["overrides"]:
            v = int(rule["overrides"][combo_key])
            return v, f"need????{order} [{combo_key!r}]={v}"
        if m and m in rule["overrides"]:
            v = int(rule["overrides"][m])
            return v, f"need????{order} [???{m!r}]={v}"
        if p and p in rule["overrides"]:
            v = int(rule["overrides"][p])
            return v, f"need????{order} [???{p!r}]={v}"
    return int(base), base_src


def _need_row_label_hints_surplus_add(label_a0: str) -> bool:
    """need ??? A?: ????????????????????????????????????"""
    s = unicodedata.normalize("NFKC", str(label_a0 or "").strip())
    if not s or s.startswith("????"):
        return False
    if "??" in s and "??" in s:
        return False
    if "??" in s and ("??" in s or "??" in s or "??" in s):
        return True
    if "??" in s or "??" in s:
        return True
    if "??" in s and ("??" in s or "?" in s or "??" in s):
        return True
    return False


def _find_need_surplus_add_row_index(
    needs_raw, base_row: int, col0: int, pm_cols: list
) -> int | None:
    """????????????????????????????????"""
    r = base_row + 1
    if r >= needs_raw.shape[0]:
        return None
    v0 = needs_raw.iat[r, col0]
    s0 = "" if pd.isna(v0) else str(v0).strip()
    if s0.startswith("????"):
        return None
    if _need_row_label_hints_surplus_add(s0):
        return r
    nz = 0
    for col_idx, _, _ in pm_cols:
        if parse_optional_int(needs_raw.iat[r, col_idx]) is not None:
            nz += 1
    if nz > 0 and not unicodedata.normalize("NFKC", s0).startswith("??"):
        return r
    return None


def resolve_need_surplus_extra_max(
    process: str,
    machine_name: str,
    task_id: str,
    surplus_map: dict,
    need_rules: list,
) -> int:
    """
    need ????????????????�???????????????????
    ????????????????0 ???????????????????
    need_rules ???????????????????? task_id ???????
    """
    _ = (task_id, need_rules)
    if not surplus_map:
        return 0
    p = str(process).strip()
    m = str(machine_name).strip()
    combo_key = f"{p}+{m}" if p and m else None
    v = None
    if combo_key and combo_key in surplus_map:
        v = surplus_map[combo_key]
    elif m and m in surplus_map:
        v = surplus_map[m]
    elif p and p in surplus_map:
        v = surplus_map[p]
    if v is None:
        return 0
    try:
        n = int(v)
    except (TypeError, ValueError):
        return 0
    return max(0, min(n, 50))


def resolve_need_surplus_extra_max_explain(
    process: str,
    machine_name: str,
    task_id: str,
    surplus_map: dict,
    need_rules: list,
) -> tuple[int, str]:
    """resolve_need_surplus_extra_max ???????????????"""
    val = resolve_need_surplus_extra_max(
        process, machine_name, task_id, surplus_map, need_rules
    )
    _ = need_rules
    if not surplus_map:
        return val, "surplus_map?????????????"
    p = str(process).strip()
    m = str(machine_name).strip()
    combo_key = f"{p}+{m}" if p and m else None
    if combo_key and combo_key in surplus_map:
        raw = surplus_map[combo_key]
        return val, f"surplus_map[{combo_key!r}]={raw}"
    if m and m in surplus_map:
        raw = surplus_map[m]
        return val, f"surplus_map[????? {m!r}]={raw}????????"
    if p and p in surplus_map:
        raw = surplus_map[p]
        return val, f"surplus_map[????? {p!r}]={raw}????????"
    return val, "surplus??????0"


def _surplus_team_time_factor(
    rq_base: int, team_len: int, extra_max_allowed: int
) -> float:
    """
    ???????????????????????????1.0???????
    ????extra_max_allowed??????????????? SURPLUS_TEAM_MAX_SPEEDUP_RATIO ????????????
    """
    rq = max(1, int(rq_base))
    tl = int(team_len)
    em = max(0, int(extra_max_allowed))
    extra = max(0, tl - rq)
    if extra <= 0 or em <= 0:
        return 1.0
    frac = min(1.0, extra / float(em))
    return 1.0 - SURPLUS_TEAM_MAX_SPEEDUP_RATIO * frac


def _team_assign_trace_tuple_label() -> str:
    if TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF:
        return "(-??, ??, -???, ?????)"
    if TEAM_ASSIGN_START_SLACK_WAIT_MINUTES <= 0:
        return "(??, -???, ?????)"
    return (
        f"??????{TEAM_ASSIGN_START_SLACK_WAIT_MINUTES}????"
        "(0,-??,??,-???,???)????(1,??,-??,-???,???)"
    )


def _team_assignment_sort_tuple(
    team: tuple,
    team_start: datetime,
    units_today: int,
    team_prio_sum: int,
    t_min: datetime | None = None,
) -> tuple:
    """
    ??????????????????????????
    - TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF: (-??, ??, -???, ?????)
    - ?????? TEAM_ASSIGN_START_SLACK_WAIT_MINUTES>0 ?? t_min ??:
        ???????????? ? (0, -??, ??, -???, ???) � ?????????
        ????? ? (1, ??, -??, -???, ???) � ?????
    - ????: (??, -???, ?????)
    """
    n = len(team)
    if TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF:
        return (-n, team_start, -units_today, team_prio_sum)
    sm = TEAM_ASSIGN_START_SLACK_WAIT_MINUTES
    if sm <= 0 or t_min is None:
        return (team_start, -units_today, team_prio_sum)
    sl = timedelta(minutes=sm)
    if team_start - t_min <= sl:
        return (0, -n, team_start, -units_today, team_prio_sum)
    return (1, team_start, -n, -units_today, team_prio_sum)


# skills ??: OP / AS + ?????????? OP1, AS 3????????????????????
_SKILL_OP_AS_CELL_RE = re.compile(r"^(OP|AS)(\d*)$", re.IGNORECASE)


def parse_op_as_skill_cell(cell_val):
    """
    master.xlsm?skills????1???????
    - ?OP?????AS?????????????????????????: OP, OP1, AS3, AS 12
    - ??????????????????????????????????????? 1?
    - OP/AS ???????????????
    """
    if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
        return None, 10**9
    s = str(cell_val).strip()
    if not s:
        return None, 10**9
    compact = re.sub(r"\s+", "", s).upper()
    m = _SKILL_OP_AS_CELL_RE.match(compact)
    if not m:
        return None, 10**9
    role = m.group(1).upper()
    tail = m.group(2) or ""
    if tail == "":
        pr = 1
    else:
        try:
            pr = int(tail)
        except ValueError:
            return None, 10**9
    if pr < 0:
        pr = 0
    return role, pr


def _validate_skills_op_as_priority_numbers_unique(
    skills_dict: dict, column_keys: list
) -> None:
    """
    master?skills???????+???????????OP/AS ???????**??**?
    ??????????????????????? PlanningValidationError?
    ?OP1 ? AS1 ??????????????????????????
    """
    errors: list[str] = []
    for combo in column_keys:
        ck = str(combo or "").strip()
        if not ck:
            continue
        pr_to_entries: dict[int, list[str]] = defaultdict(list)
        for mem, row in (skills_dict or {}).items():
            mnm = str(mem or "").strip()
            if not mnm or not isinstance(row, dict):
                continue
            raw = row.get(ck)
            if raw is None or (isinstance(raw, float) and pd.isna(raw)):
                continue
            sval = str(raw).strip()
            if not sval or sval.lower() in ("nan", "none", "null"):
                continue
            role, pr = parse_op_as_skill_cell(sval)
            if role not in ("OP", "AS"):
                continue
            pr_to_entries[int(pr)].append(f"{mnm}({role})")
        for pr, entries in sorted(pr_to_entries.items()):
            if len(entries) > 1:
                errors.append(f'??{ck}?: ??? {pr} ??? ? ' + "?".join(entries))
    if errors:
        cap = 50
        tail = errors[:cap]
        msg = (
            "????skills??????? OP/AS ???????????????"
            " ???????1????1??????????\n"
            + "\n".join(tail)
        )
        if len(errors) > cap:
            msg += f"\n�? {len(errors) - cap} ?"
        raise PlanningValidationError(msg)


def build_member_assignment_priority_reference(
    skills_dict: dict,
    members: list | None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    ??????: ??? skills ?????+????????????????????
    (??????, ???????) ????????????????????
    ?????????????????????????????????????????????
    """
    mem_list = list(members) if members else list((skills_dict or {}).keys())
    mem_list = [str(m).strip() for m in mem_list if m and str(m).strip()]

    surplus_on = bool(TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF)
    slack_m = TEAM_ASSIGN_START_SLACK_WAIT_MINUTES
    if surplus_on:
        team_rule = (
            "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF=??: "
            "(-??, ??, -???, ?????) ???????????????"
        )
    elif slack_m > 0:
        team_rule = (
            f"??: ??????????????????????"
            f"?????{slack_m}????????????????0,-??,??,-???,?????"
            f"???????????????1,??,-??,-???,?????"
            f"???? TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=0 ?????"
        )
    else:
        team_rule = (
            "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=0: "
            "(??, -???, ?????) ??????????"
        )

    legend_rows = [
        {
            "??": "???????",
            "??": "?????+????????????? OP/AS?+????????????????"
            " ??????????????????? 1?parse_op_as_skill_cell ?????"
            " ??????????????????????????????????",
        },
        {
            "??": "?????",
            "??": "???????????????????? AS/OP ?????????????"
            " ??????????NO????????????????OP???????",
        },
        {
            "??": "????????",
            "??": team_rule,
        },
        {
            "??": "???????????",
            "??": "??OP_???????????????? OP ??????????????",
        },
        {
            "??": "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF",
            "??": "1/????????????" if surplus_on else "0/??????",
        },
        {
            "??": "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES",
            "??": str(slack_m),
        },
    ]
    df_legend = pd.DataFrame(legend_rows)

    combo_keys: set[str] = set()
    for m in mem_list:
        row = (skills_dict or {}).get(m) or {}
        for k in row:
            ks = str(k).strip()
            if "+" in ks:
                combo_keys.add(ks)
    sorted_combos = sorted(combo_keys)

    out: list[dict] = []
    for combo in sorted_combos:
        parts = combo.split("+", 1)
        proc = parts[0].strip()
        mach = parts[1].strip() if len(parts) > 1 else ""
        ranked: list[tuple[int, str, str, str]] = []
        for m in sorted(mem_list):
            cell = ((skills_dict or {}).get(m) or {}).get(combo)
            if cell is None or (isinstance(cell, float) and pd.isna(cell)):
                cell_s = ""
            else:
                cell_s = str(cell).strip()
            role, pr = parse_op_as_skill_cell(cell_s if cell_s else None)
            if role in ("OP", "AS"):
                ranked.append((pr, m, role, cell_s))
        ranked.sort(key=lambda x: (x[0], x[1]))
        if not ranked:
            out.append(
                {
                    "???": proc,
                    "???": mach,
                    "??????": combo,
                    "????": "",
                    "????": "????",
                    "???": "",
                    "????_??????": "",
                    "skills???": "",
                    "??": "???? OP/AS ?????????????????",
                }
            )
            continue
        for i, (pr, m, role, cell_s) in enumerate(ranked, start=1):
            out.append(
                {
                    "???": proc,
                    "???": mach,
                    "??????": combo,
                    "????": i,
                    "????": m,
                    "???": role,
                    "????_??????": pr,
                    "skills???": cell_s,
                    "??": "",
                }
            )

    df_tbl = pd.DataFrame(out)
    return df_legend, df_tbl


def _normalize_person_name_for_match(s):
    """??????????????NFKC???/???????????????????????"""
    if s is None:
        return ""
    t = unicodedata.normalize("NFKC", str(s).strip())
    if "??" in t:
        t = t.replace("??", "??")
    t = re.sub(r"[\s?]+", "", t)
    t = re.sub(r"(??|?|?)$", "", t)
    return t


def _split_person_sei_mei(s) -> tuple[str, str]:
    """
    ???????????????????????????????????
    ???????? (??, '')????????
    ??? ?????? ??????????
    """
    if s is None:
        return "", ""
    t = unicodedata.normalize("NFKC", str(s).strip())
    if not t or t.lower() in ("nan", "none", "null"):
        return "", ""
    t = re.sub(r"(??|?|?)$", "", t)
    for i, ch in enumerate(t):
        if ch in " \u3000":
            sei = t[:i].strip()
            rest = t[i + 1 :]
            mei = re.sub(r"[\s?]+", "", rest.strip())
            return sei, mei
    return t.strip(), ""


def _normalize_sei_for_match(sei: str) -> str:
    """?????????????????????NFKC???/??????????"""
    if not sei:
        return ""
    t = unicodedata.normalize("NFKC", str(sei).strip())
    if "??" in t:
        t = t.replace("??", "??")
    t = re.sub(r"[\s?]+", "", t)
    return t


def _normalize_mei_for_match(mei: str) -> str:
    """????????????????NFKC???????????????????"""
    if not mei:
        return ""
    t = unicodedata.normalize("NFKC", str(mei).strip())
    t = re.sub(r"[\s?]+", "", t)
    return t


def _has_duplicate_surname_among_members(member_names) -> bool:
    """skills ?????????????????2???????"""
    cnt = Counter()
    for name in member_names or []:
        if name is None or (isinstance(name, float) and pd.isna(name)):
            continue
        s = str(name).strip()
        if not s:
            continue
        sei, _mei = _split_person_sei_mei(s)
        key = _normalize_sei_for_match(sei)
        if key:
            cnt[key] += 1
    return any(n >= 2 for n in cnt.values())


def _mei_matches_with_fuzzy_allowed(r_mei_n: str, m_mei_n: str) -> bool:
    """???????????????????????????"""
    if not r_mei_n and not m_mei_n:
        return True
    if not r_mei_n or not m_mei_n:
        return False
    if r_mei_n == m_mei_n:
        return True
    return r_mei_n in m_mei_n or m_mei_n in r_mei_n


def _resolve_preferred_name_to_capable_member(raw, capable_candidates, roster_member_names=None):
    """
    ??????????????? OP/AS ???????skills ??????????????
    capable_candidates: ????? OP ??? AS ?????????????????
    roster_member_names: skills ???????????? capable_candidates??????????????

    ???????:
    - ????????????????????????/?????????????
    - roster ?????2????????????????????????????????????????
    - ??????????????????????
    - ????????????????????????????????????None??
    """
    if not raw or not capable_candidates:
        return None
    r0 = unicodedata.normalize("NFKC", str(raw).strip())
    r = _normalize_person_name_for_match(r0)
    if not r:
        return None
    for m in capable_candidates:
        if _normalize_person_name_for_match(m) == r:
            return m
        if unicodedata.normalize("NFKC", str(m).strip()) == r0.strip():
            return m

    roster = (
        list(roster_member_names)
        if roster_member_names is not None
        else list(capable_candidates)
    )
    allow_mei_fuzzy = not _has_duplicate_surname_among_members(roster)

    r_sei, r_mei = _split_person_sei_mei(raw)
    r_sei_n = _normalize_sei_for_match(r_sei)
    r_mei_n = _normalize_mei_for_match(r_mei)
    if not r_sei_n:
        return None

    matches = []
    for m in capable_candidates:
        m_sei, m_mei = _split_person_sei_mei(m)
        m_sei_n = _normalize_sei_for_match(m_sei)
        m_mei_n = _normalize_mei_for_match(m_mei)
        if r_sei_n != m_sei_n:
            continue
        if allow_mei_fuzzy:
            if r_mei_n:
                if _mei_matches_with_fuzzy_allowed(r_mei_n, m_mei_n):
                    matches.append(m)
            else:
                matches.append(m)
        else:
            if r_mei_n == m_mei_n:
                matches.append(m)

    if len(matches) == 1:
        return matches[0]
    return None


def _resolve_preferred_op_to_member(raw, op_candidates, roster_member_names=None):
    """?????? OP ???????? API????? `_resolve_preferred_name_to_capable_member`?"""
    return _resolve_preferred_name_to_capable_member(
        raw, op_candidates, roster_member_names
    )


def _task_process_matches_global_contains(machine_val: str, contains: str) -> bool:
    """???????? machine???????NFKC???????"""
    m = unicodedata.normalize("NFKC", str(machine_val or "").strip()).casefold()
    c = unicodedata.normalize("NFKC", str(contains or "").strip()).casefold()
    if not c:
        return False
    return c in m


def _coerce_global_day_process_operator_rules(raw_val) -> list:
    """Gemini ? global_day_process_operator_rules ??????????????"""
    out: list[dict] = []
    if not isinstance(raw_val, list):
        return out
    seen_sig = set()
    for item in raw_val:
        if not isinstance(item, dict):
            continue
        d = parse_optional_date(item.get("date"))
        if d is None:
            continue
        pc = item.get("process_contains")
        if pc is None or (isinstance(pc, float) and pd.isna(pc)):
            continue
        pc_s = unicodedata.normalize("NFKC", str(pc).strip())
        if not pc_s:
            continue
        names = item.get("operator_names")
        if not isinstance(names, list):
            continue
        op_names: list[str] = []
        for n in names:
            if n is None or (isinstance(n, float) and pd.isna(n)):
                continue
            s = str(n).strip()
            if s and s.lower() not in ("nan", "none", "null"):
                op_names.append(s)
        if not op_names:
            continue
        sig = (d.isoformat(), pc_s.casefold(), tuple(op_names))
        if sig in seen_sig:
            continue
        seen_sig.add(sig)
        out.append(
            {
                "date": d.isoformat(),
                "process_contains": pc_s,
                "operator_names": op_names,
            }
        )
    return out


def _active_global_day_process_must_include(
    gpo: dict,
    task: dict,
    current_date: date,
    capable_members: list,
    roster_members: list,
) -> tuple[list[str], list[str]]:
    """
    ???????????????�??�???????????????????
    **?????????**?????skills ????????????????
    """
    rules = gpo.get("global_day_process_operator_rules") or []
    if not isinstance(rules, list):
        return [], []
    machine = task.get("machine")
    warns: list[str] = []
    acc: list[str] = []
    seen_m: set[str] = set()
    tid = str(task.get("task_id") or "").strip()
    for rule in rules:
        if not isinstance(rule, dict):
            continue
        rd = parse_optional_date(rule.get("date"))
        if rd is None or rd != current_date:
            continue
        pc = rule.get("process_contains") or ""
        pcn = unicodedata.normalize("NFKC", str(pc).strip())
        if not pcn or not _task_process_matches_global_contains(machine, pcn):
            continue
        for raw_name in rule.get("operator_names") or []:
            mem = _resolve_preferred_name_to_capable_member(
                raw_name, capable_members, roster_members
            )
            if mem:
                if mem not in seen_m:
                    seen_m.add(mem)
                    acc.append(mem)
            else:
                warns.append(
                    "????????(??�??)??: "
                    f"??NO={tid} ??={current_date} ??={machine!r} ? "
                    f"???{raw_name}?????????????????????"
                )
    return acc, warns


def _merge_global_day_process_and_pref_anchor(
    must_include: list, pref_mem, capable_members: list
) -> list[str]:
    """?????????OP???1???capable ?????????"""
    fixed: list[str] = []
    seen: set[str] = set()
    for m in must_include or []:
        if m in capable_members and m not in seen:
            seen.add(m)
            fixed.append(m)
    if (
        pref_mem
        and pref_mem in capable_members
        and pref_mem not in seen
    ):
        fixed.append(pref_mem)
    return fixed


# =========================================================
# 2. ??????????(?????) ? AI??
#    master.xlsm ? skills / need / ??????????????
#    ?????????????? Gemini ???????
# =========================================================
def load_skills_and_needs():
    """
    ??????(MASTER_FILE)?????? need ???????????

    ????7?????? need ?????????+????????????????????
    ``need_combo_col_index``????????????

    ??? need ??Excel???
      ??????????????????????A????????????
      ????: ??????????????????????????�????????? 0?
      ??: ????1?99
    ???????????OP?????+??????????

    skills ????? OP/AS ?????????? OP1, AS3??????????????????????
    ????? OP/AS ???? 1?
    ????????�??????????????????????????? PlanningValidationError??
    """
    try:
        # ?????? pd.read_excel ?????? I/O ??????ExcelFile ?1?????????? parse ???
        with pd.ExcelFile(MASTER_FILE) as _master_xls:
            # skills ????:
            #   1??: ???
            #   2??: ???
            #   A3??: ?????
            #   ????: OP ??? AS ????????????? OP1, AS3????????????????????
            #             ????? OP/AS ???? 1?????????????
            # ????????????1???????????????????
            skills_raw = pd.read_excel(
                _master_xls, sheet_name="skills", header=None
            )
            skills_dict = {}
            equipment_list = []
            members = []

            use_two_header = False
            if skills_raw.shape[0] >= 3 and skills_raw.shape[1] >= 2:
                non_empty_pm = 0
                for c in range(1, skills_raw.shape[1]):
                    p = skills_raw.iat[0, c]
                    m = skills_raw.iat[1, c]
                    if pd.isna(p) or pd.isna(m):
                        continue
                    p_s = str(p).strip()
                    m_s = str(m).strip()
                    if p_s and m_s and p_s.lower() != "nan" and m_s.lower() != "nan":
                        non_empty_pm += 1
                use_two_header = non_empty_pm > 0

            if use_two_header:
                pm_cols = []
                seen_combo = set()
                for c in range(1, skills_raw.shape[1]):
                    p = skills_raw.iat[0, c]
                    m = skills_raw.iat[1, c]
                    if pd.isna(p) or pd.isna(m):
                        continue
                    p_s = str(p).strip()
                    m_s = str(m).strip()
                    if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
                        continue
                    combo = f"{p_s}+{m_s}"
                    pm_cols.append((c, p_s, m_s, combo))
                    if combo not in seen_combo:
                        seen_combo.add(combo)
                        equipment_list.append(combo)

                for r in range(2, skills_raw.shape[0]):
                    m_name_raw = skills_raw.iat[r, 0]
                    if pd.isna(m_name_raw):
                        continue
                    m_name = str(m_name_raw).strip()
                    if not m_name or m_name.lower() in ("nan", "none", "null"):
                        continue
                    row_skills = {}
                    for c, p_s, m_s, combo in pm_cols:
                        v = skills_raw.iat[r, c] if c < skills_raw.shape[1] else None
                        sval = "" if pd.isna(v) else str(v).strip()
                        if not sval or sval.lower() in ("nan", "none", "null"):
                            continue
                        row_skills[combo] = sval
                        if m_s not in row_skills:
                            row_skills[m_s] = sval
                        if p_s not in row_skills:
                            row_skills[p_s] = sval
                    skills_dict[m_name] = row_skills
                members = list(skills_dict.keys())
                logging.info(
                    "skills???: 2?????????????????+??=%s?, ????=%s???",
                    len(pm_cols),
                    len(members),
                )
            else:
                skills_df = pd.read_excel(_master_xls, sheet_name="skills")
                skills_df.columns = skills_df.columns.str.strip()
                skill_cols = [
                    str(c).strip()
                    for c in skills_df.columns
                    if not str(c).startswith("Unnamed")
                ]

                member_col = None
                for c in skill_cols:
                    if c in ("????", "???", "??", "???"):
                        member_col = c
                        break
                if member_col is None and skill_cols:
                    member_col = skill_cols[0]
                    logging.warning(
                        "skills???: ????????????????????? '%s' ??????????????",
                        member_col,
                    )

                seen_eq = set()
                for c in skill_cols:
                    if c == member_col:
                        continue
                    cid = str(c).strip()
                    if not cid or cid.lower() in ("nan", "none", "null"):
                        continue
                    if cid not in seen_eq:
                        seen_eq.add(cid)
                        equipment_list.append(cid)

                for _, row in skills_df.iterrows():
                    m_name = str(row.get(member_col, "")).strip() if member_col else ""
                    if not m_name or m_name.lower() == "nan":
                        continue
                    row_skills = {}
                    for c in skill_cols:
                        if c == member_col:
                            continue
                        sval = str(row.get(c, "")).strip()
                        if not sval or sval.lower() in ("nan", "none", "null"):
                            continue
                        row_skills[c] = sval
                        if "+" in c:
                            p, m = c.split("+", 1)
                            p = p.strip()
                            m = m.strip()
                            if m and m not in row_skills:
                                row_skills[m] = sval
                            if p and p not in row_skills:
                                row_skills[p] = sval
                    skills_dict[m_name] = row_skills
                members = list(skills_dict.keys())
                logging.info(
                    "skills???: 1????????????????????????=%s???",
                    len(members),
                )

            if not members:
                logging.error("skills?????????????????????")
            else:
                _validate_skills_op_as_priority_numbers_unique(
                    skills_dict, equipment_list
                )

            # need ? header=None ???????????�????�?????
            needs_raw = pd.read_excel(
                _master_xls, sheet_name="need", header=None
            )

        col0 = 0
        process_header_row = None
        machine_header_row = None
        base_row = None

        for r in range(needs_raw.shape[0]):
            v0 = needs_raw.iat[r, col0]
            if pd.isna(v0):
                continue
            s0 = str(v0).strip()
            if process_header_row is None and s0 == "???":
                process_header_row = r
            elif machine_header_row is None and s0 == "???":
                machine_header_row = r
            if base_row is None and "????" in s0 and not s0.startswith("????"):
                base_row = r
            if process_header_row is not None and machine_header_row is not None and base_row is not None:
                break

        if process_header_row is None or machine_header_row is None or base_row is None:
            raise ValueError("need ?????????????/???/????????????????")

        # ???NO????????????? 1???
        cond_col_idx = 1
        for r in range(needs_raw.shape[0]):
            c1 = needs_raw.iat[r, 1] if needs_raw.shape[1] > 1 else None
            c2 = needs_raw.iat[r, 2] if needs_raw.shape[1] > 2 else None
            if pd.isna(c1) or pd.isna(c2):
                continue
            if str(c1).strip() == NEED_COL_CONDITION and str(c2).strip() == NEED_COL_NOTE:
                cond_col_idx = 1
                break

        # ???�??? ????????? Excel????????
        pm_cols = []
        for col_idx in range(needs_raw.shape[1]):
            if col_idx < 3:
                continue
            p = needs_raw.iat[process_header_row, col_idx]
            m = needs_raw.iat[machine_header_row, col_idx]
            if pd.isna(p) or pd.isna(m):
                continue
            p_s = str(p).strip()
            m_s = str(m).strip()
            if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
                continue
            pm_cols.append((col_idx, p_s, m_s))

        req_map = {}
        # ???+?????? ? need ????????????????????????????
        need_combo_col_index: dict[str, int] = {}
        # need_rules: [{'order': int, 'condition': str, 'overrides': {combo_key/machine/process: int}}]
        need_rules = []

        # ??????
        for col_idx, p_s, m_s in pm_cols:
            n = parse_optional_int(needs_raw.iat[base_row, col_idx])
            if n is None or n < 1:
                n = 1
            combo_key = f"{p_s}+{m_s}"
            need_combo_col_index[combo_key] = col_idx
            req_map[combo_key] = n
            # ???????????? or ???????????????
            if p_s not in req_map:
                req_map[p_s] = n
            if m_s not in req_map:
                req_map[m_s] = n

        surplus_map: dict[str, int] = {}
        surplus_row = _find_need_surplus_add_row_index(
            needs_raw, base_row, col0, pm_cols
        )
        if surplus_row is not None:
            for col_idx, p_s, m_s in pm_cols:
                raw_ex = parse_optional_int(needs_raw.iat[surplus_row, col_idx])
                ex = int(raw_ex) if raw_ex is not None and raw_ex >= 0 else 0
                ex = max(0, min(ex, 50))
                combo_key = f"{p_s}+{m_s}"
                surplus_map[combo_key] = ex
                if p_s not in surplus_map:
                    surplus_map[p_s] = ex
                if m_s not in surplus_map:
                    surplus_map[m_s] = ex
            logging.info(
                "need ???: ????????????Excel??%s?????????????????",
                surplus_row + 1,
            )
        else:
            logging.info(
                "need ???: ???????????????????????????????????"
            )

        if TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW:
            logging.info(
                "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW ???: ???????????????? 0 ???????????????????"
            )

        logging.info(
            "need?????: %s ? need ?????????skills ??? ExcelFile ???????need ???????????????AI json ???????",
            os.path.abspath(MASTER_FILE),
        )
        for _ci, _ps, _ms in pm_cols:
            _ck = f"{_ps}+{_ms}"
            _bn = req_map.get(_ck)
            _sx = surplus_map.get(_ck, 0) if surplus_map else 0
            logging.info(
                "need???? combo=%r ??????=%s ?????????=%s",
                _ck,
                _bn,
                _sx,
            )

        # ????
        for r in range(needs_raw.shape[0]):
            v0 = needs_raw.iat[r, col0]
            if pd.isna(v0):
                continue
            lab = str(v0).strip()
            m = re.match(r"????\s*(\d+)", lab)
            if not m:
                continue
            order = int(m.group(1))
            if order < 1 or order > 99:
                continue

            cond_raw = needs_raw.iat[r, cond_col_idx] if needs_raw.shape[1] > cond_col_idx else None
            cond = "" if pd.isna(cond_raw) else str(cond_raw).strip()

            overrides = {}
            for col_idx, p_s, m_s in pm_cols:
                v = needs_raw.iat[r, col_idx]
                n = parse_optional_int(v)
                if n is not None and 1 <= n <= 99:
                    combo_key = f"{p_s}+{m_s}"
                    overrides[combo_key] = n
                    # ????????
                    overrides[p_s] = n
                    overrides[m_s] = n

            if overrides:
                need_rules.append({"order": order, "condition": cond, "overrides": overrides})

        need_rules.sort(key=lambda rr: rr["order"])
        logging.info(f"need ???????: {len(need_rules)} ?????+???????")

        logging.info(f"?{MASTER_FILE}???????????(need)?????????")
        return (
            skills_dict,
            members,
            equipment_list,
            req_map,
            need_rules,
            surplus_map,
            need_combo_col_index,
        )

    except PlanningValidationError:
        raise
    except Exception as e:
        logging.error(f"???????({MASTER_FILE})????/need???????: {e}")
        return {}, [], [], {}, [], [], {}


def load_team_combination_presets_from_master() -> dict[
    str, list[tuple[int, int | None, tuple[str, ...], int | None]]
]:
    """
    master.xlsm??????????????+???????
    [(??????, ???????None, ???????, ????ID???None), ...] ????
    ????????????????????????
    ???????????? need ???????????????????????????
    ????????????????????????????????? team_start ???????
    ??????????????????????????????
    A ??????ID????????????? ID ? None?
    """
    if not TEAM_ASSIGN_USE_MASTER_COMBO_SHEET:
        return {}
    path = MASTER_FILE
    if not os.path.isfile(path):
        return {}
    try:
        df = pd.read_excel(path, sheet_name=MASTER_SHEET_TEAM_COMBINATIONS, header=0)
    except Exception as e:
        logging.info("????????????????????: %s", e)
        return {}
    if df is None or df.empty:
        return {}

    def norm_cell(x) -> str:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip()

    colmap = {norm_cell(c): c for c in df.columns if norm_cell(c)}
    id_c = colmap.get("????ID") or colmap.get("??????")
    proc_c = colmap.get("???")
    mach_c = colmap.get("???")
    combo_c = colmap.get("??+??")
    prio_c = colmap.get("??????")
    req_c = colmap.get("????")

    def mem_col_order(c) -> int:
        m = re.search(r"????\s*(\d+)", norm_cell(c))
        return int(m.group(1)) if m else 9999

    mem_keys = sorted(
        [c for c in df.columns if norm_cell(str(c)).startswith("????")],
        key=mem_col_order,
    )
    buckets: dict[
        str,
        list[tuple[int, int, int | None, tuple[str, ...], int | None]],
    ] = defaultdict(list)
    for row_i, (_, row) in enumerate(df.iterrows()):
        proc = norm_cell(row.get(proc_c)) if proc_c else ""
        mach = norm_cell(row.get(mach_c)) if mach_c else ""
        combo_cell = norm_cell(row.get(combo_c)) if combo_c else ""
        if proc and mach:
            key = f"{proc}+{mach}"
        elif combo_cell:
            key = combo_cell
        else:
            continue
        pr = parse_optional_int(row.get(prio_c)) if prio_c else None
        if pr is None:
            pr = 10**9
        sheet_req: int | None = None
        if req_c:
            sheet_req = parse_optional_int(row.get(req_c))
            if sheet_req is not None and sheet_req < 1:
                sheet_req = None
        sheet_combo_id: int | None = None
        if id_c:
            sheet_combo_id = parse_optional_int(row.get(id_c))
            if sheet_combo_id is not None and sheet_combo_id < 1:
                sheet_combo_id = None
        team: list[str] = []
        for mc in mem_keys:
            s = norm_cell(row.get(mc))
            if not s or s.lower() in ("nan", "none", "null"):
                continue
            team.append(s)
        if not team:
            continue
        buckets[key].append(
            (pr, row_i, sheet_req, tuple(team), sheet_combo_id)
        )

    out: dict[
        str, list[tuple[int, int | None, tuple[str, ...], int | None]]
    ] = {}
    for key, lst in buckets.items():
        lst.sort(key=lambda x: (x[0], x[1]))
        out[key] = [(t[0], t[2], t[3], t[4]) for t in lst]
    return out


def _lookup_combo_sheet_row_id_for_preset_team(
    preset_rows: list | None,
    team: tuple,
) -> int | None:
    """
    ?????????????NFKC?trim??????????????????????????
    ????????ID?A????????????????? combo_sheet_row_id ???????
    ???????????????????????????????????
    """
    if not preset_rows or not team:
        return None

    def _mem_key(members) -> frozenset:
        out: set[str] = set()
        for m in members:
            s = str(m).strip()
            if not s:
                continue
            out.add(unicodedata.normalize("NFKC", s))
        return frozenset(out)

    target = _mem_key(team)
    if not target:
        return None
    best_id: int | None = None
    best_prio: int | None = None
    for pr, _sheet_rs, preset_team, combo_row_id in preset_rows:
        if combo_row_id is None:
            continue
        if _mem_key(preset_team) != target:
            continue
        try:
            prio_val = int(pr)
        except (TypeError, ValueError):
            prio_val = 10**9
        if best_prio is None or prio_val < best_prio:
            best_prio = prio_val
            try:
                best_id = int(combo_row_id)
            except (TypeError, ValueError):
                best_id = None
    return best_id


def generate_default_calendar_dates(year, month):
    cal = calendar.Calendar()
    return [d for d in cal.itermonthdates(year, month) if d.year == year and d.month == month and d.weekday() < 5]

def parse_time_str(time_str, default_time):
    if time_str is None or pd.isna(time_str) or not str(time_str).strip() or str(time_str).strip().lower() == 'null':
        return default_time
    try:
        if isinstance(time_str, time): return time_str
        if isinstance(time_str, datetime): return time_str.time()
        time_str = str(time_str).strip()
        if len(time_str.split(':')) == 3:
            return datetime.strptime(time_str, "%H:%M:%S").time()
        return datetime.strptime(time_str, "%H:%M").time()
    except:
        return default_time


def _excel_scalar_to_time_optional(v) -> time | None:
    """master ?????????datetime / time / ????? time ??????? None?"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, time):
        return v
    if isinstance(v, datetime):
        return v.time()
    return parse_time_str(v, None)


def _pick_master_main_sheet_name(sheetnames: list[str]) -> str | None:
    """
    master.xlsm ??????????????????VBA MasterGetMainWorksheet ??????
    ???????????????????????????????????????
    ?????????????????????
    """
    for prefer in ("???", "???_", "Main"):
        if prefer in sheetnames:
            return prefer
    cand = [sn for sn in sheetnames if "???" in sn and "?????" not in sn]
    if not cand:
        return None
    return min(cand, key=len)


def _read_master_main_factory_operating_times(master_path: str) -> tuple[time | None, time | None]:
    """
    master.xlsm ??????? A12???????B12??????????
    ????????????>=?????? (None, None)?
    """
    p = (master_path or "").strip()
    if not p or not os.path.isfile(p):
        return None, None
    if _workbook_should_skip_openpyxl_io(p):
        return None, None
    try:
        wb = load_workbook(p, data_only=True, read_only=False)
    except Exception as e:
        logging.warning("??????: master ? openpyxl ??????????????????????: %s", e)
        return None, None
    try:
        sn = _pick_master_main_sheet_name(list(wb.sheetnames))
        if sn is None:
            return None, None
        ws = wb[sn]
        st = _excel_scalar_to_time_optional(ws.cell(row=12, column=1).value)
        et = _excel_scalar_to_time_optional(ws.cell(row=12, column=2).value)
        if st is None or et is None:
            return None, None
        if st >= et:
            logging.warning(
                "??????: master ??? A12/B12 ???>=?? (%s >= %s) ????????????",
                st,
                et,
            )
            return None, None
        return st, et
    finally:
        try:
            wb.close()
        except Exception:
            pass


def _read_master_main_regular_shift_times(master_path: str) -> tuple[time | None, time | None]:
    """
    master.xlsm ??????? A15???????B15??????????
    ????????????>=?????? (None, None)?
    """
    p = (master_path or "").strip()
    if not p or not os.path.isfile(p):
        return None, None
    if _workbook_should_skip_openpyxl_io(p):
        return None, None
    try:
        wb = load_workbook(p, data_only=True, read_only=False)
    except Exception as e:
        logging.warning(
            "????: master ? openpyxl ???????????????????????????: %s",
            e,
        )
        return None, None
    try:
        sn = _pick_master_main_sheet_name(list(wb.sheetnames))
        if sn is None:
            return None, None
        ws = wb[sn]
        st = _excel_scalar_to_time_optional(ws.cell(row=15, column=1).value)
        et = _excel_scalar_to_time_optional(ws.cell(row=15, column=2).value)
        if st is None or et is None:
            return None, None
        if st >= et:
            logging.warning(
                "????: master ??? A15/B15 ???>=?? (%s >= %s) ???????????????",
                st,
                et,
            )
            return None, None
        return st, et
    finally:
        try:
            wb.close()
        except Exception:
            pass


@contextmanager
def _override_default_factory_hours_from_master(master_path: str):
    """??2???? DEFAULT_START_TIME / DEFAULT_END_TIME ? master ??? A12/B12 ?????"""
    global DEFAULT_START_TIME, DEFAULT_END_TIME
    orig_s, orig_e = DEFAULT_START_TIME, DEFAULT_END_TIME
    ns, ne = _read_master_main_factory_operating_times(master_path)
    try:
        if ns is not None and ne is not None:
            DEFAULT_START_TIME = ns
            DEFAULT_END_TIME = ne
            logging.info(
                "?????: master.xlsm ??? A12/B12 ??? ? %s ? %s???_* ????????????",
                DEFAULT_START_TIME.strftime("%H:%M"),
                DEFAULT_END_TIME.strftime("%H:%M"),
            )
        yield
    finally:
        DEFAULT_START_TIME, DEFAULT_END_TIME = orig_s, orig_e


def infer_mid_break_from_reason(reason_text, start_t, end_t, break1_start=None, break1_end=None):
    """
    ?????????????????????
    AI?????????????????????????
    master.xlsm ????????????: ??=??????????????=???????????????.txt ?????
    ????????????????????1_??/???????? DEFAULT_BREAKS[0]???????
    """
    if reason_text is None:
        return None, None
    txt = str(reason_text).strip()
    if not txt or txt.lower() in ("nan", "none", "null", "??"):
        return None, None

    b1_s = break1_start if break1_start is not None else DEFAULT_BREAKS[0][0]
    b1_e = break1_end if break1_end is not None else DEFAULT_BREAKS[0][1]

    noon_end = time(12, 0)
    afternoon_start = time(13, 0)
    # ????????????????????????????????????????????????????
    if txt == "??":
        # ??????????1?????????????????????????????????????
        if start_t and start_t < b1_e:
            return start_t, b1_e
        return None, None
    if txt == "??":
        if end_t and b1_s < end_t:
            return b1_s, end_t
        return None, None

    # 1) ??????????: 11:00-14:00 / 11:00?14:00?
    m = re.search(r"(\d{1,2}[:?]\d{2})\s*[~?\-??]\s*(\d{1,2}[:?]\d{2})", txt)
    if m:
        s = parse_time_str(m.group(1).replace("?", ":"), None)
        e = parse_time_str(m.group(2).replace("?", ":"), None)
        if s and e and s < e:
            return s, e

    # 2) ????????/??/??? + ?????????????
    # ?????????????????????????????????????????????????
    leave_keywords = (
        "???", "??", "??", "??", "??", "??", "??", "???", "???",
        "??", "??", "??",
    )
    has_leave_hint = any(k in txt for k in leave_keywords)
    if not has_leave_hint:
        return None, None

    if ("??" in txt) or ("1?" in txt and "??" not in txt):
        return start_t, end_t
    if ("???" in txt) or ("??" in txt):
        return start_t, noon_end
    if ("??" in txt):
        return afternoon_start, end_t

    return None, None


# ??_?????(???) ??????VBA ??????????????????? ????1_??????????
_AFTERNOON_OFF_DISPLAY_END = DEFAULT_BREAKS[0][0]


def _reason_is_afternoon_off(reason: str) -> bool:
    """??????????????????????????"""
    r = str(reason or "")
    return ("??" in r and ("?" in r or "??" in r)) or ("??" in r)


def _reason_is_morning_off(reason: str) -> bool:
    """????????????????????????????????????????????????"""
    return "??" in str(reason or "")


def _calendar_display_clock_out_for_calendar_sheet(entry: dict, day_date: date):
    """
    ??? breaks_dt ???????????????end_dt ? 17:00 ?????????????????????
    ?????????????????????????????????????????????????1_???????end_dt ??????????
    """
    if not entry.get("is_working"):
        return None
    end_dt = entry.get("end_dt")
    if end_dt is None:
        return None
    reason = str(entry.get("reason") or "")
    afternoon_off = _reason_is_afternoon_off(reason)
    if not afternoon_off:
        return None
    breaks_dt = entry.get("breaks_dt") or []
    for b_s, b_e in breaks_dt:
        if b_s is None or b_e is None:
            continue
        bs = b_s.time() if isinstance(b_s, datetime) else b_s
        if isinstance(bs, datetime):
            bs = bs.time()
        if bs < DEFAULT_BREAKS[0][0]:
            continue
        if isinstance(b_e, datetime):
            be_cmp = b_e
        elif isinstance(b_e, time):
            be_cmp = datetime.combine(day_date, b_e)
        else:
            continue
        if be_cmp >= end_dt - timedelta(seconds=1):
            return datetime.combine(day_date, _AFTERNOON_OFF_DISPLAY_END)
    return None


def _member_schedule_break_cell_label(grid_mid_dt, breaks_dt, shift_end_dt, reason):
    """
    ??_* ???????10???????????????
    ??????????????????????????????????????????
    ????????????????????????????????????
    """
    reason = str(reason or "")
    afternoon_off = _reason_is_afternoon_off(reason)
    morning_off = _reason_is_morning_off(reason)
    for b_s, b_e in breaks_dt:
        if b_s is None or b_e is None:
            continue
        if not (b_s <= grid_mid_dt < b_e):
            continue
        if isinstance(b_e, datetime) and isinstance(shift_end_dt, datetime):
            bs = b_s.time() if isinstance(b_s, datetime) else b_s
            if isinstance(bs, datetime):
                bs = bs.time()
            if afternoon_off and bs >= DEFAULT_BREAKS[0][0] and b_e >= shift_end_dt - timedelta(seconds=2):
                return "??"
            if morning_off and bs < DEFAULT_BREAKS[0][0]:
                be_t = b_e.time() if isinstance(b_e, datetime) else b_e
                if be_t <= time(13, 0):
                    return "??"
        return "??"
    return None


def _member_schedule_off_shift_label(
    day_date: date,
    grid_mid_dt: datetime,
    d_start_dt: datetime,
    d_end_dt: datetime,
    reason: str,
) -> str:
    """
    ??_* ?????????????10???
    ?????????????????????????????????????????????????
    """
    r = str(reason or "")
    day_start = datetime.combine(day_date, DEFAULT_START_TIME)
    day_end = datetime.combine(day_date, DEFAULT_END_TIME)
    if grid_mid_dt < d_start_dt:
        if _reason_is_morning_off(r) and grid_mid_dt >= day_start:
            return "??"
        return "???"
    if grid_mid_dt >= d_end_dt:
        if _reason_is_afternoon_off(r) and grid_mid_dt < day_end:
            return "??"
        return "???"
    return "???"


def _member_schedule_full_day_off_label(entry) -> str:
    """
    ??????is_working=False????????????
    ????????????? *?????????????????????
    """
    if not entry:
        return "?"
    r = str(entry.get("reason") or "").strip()
    if r == "??" or r.startswith("?? "):
        return "??"
    return "?"


def _attendance_remark_text(row) -> str:
    """
    ??1????????????????????
    ??AI?????????????????reason ???? load_attendance ??????????????
    """
    if row is None:
        return ""
    try:
        v = row.get(ATT_COL_REMARK)
    except Exception:
        return ""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    return s


def _attendance_leave_type_text(row) -> str:
    """??1??????????????????? ??/?? ???"""
    if row is None:
        return ""
    try:
        v = row.get(ATT_COL_LEAVE_TYPE)
    except Exception:
        return ""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    return s


def _ai_json_bool(v, default: bool = False) -> bool:
    """???? AI ?????bool / ?? / ???????????"""
    if v is None:
        return default
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v != 0
    if isinstance(v, float):
        if pd.isna(v):
            return default
        return v != 0.0
    s = str(v).strip().lower()
    if s in ("true", "1", "yes", "y", "??", "?", "on"):
        return True
    if s in ("false", "0", "no", "n", "???", "?", "off", ""):
        return False
    return default


def _parse_attendance_overtime_end_optional(v) -> time | None:
    """??????????????????????? None?_excel_scalar_to_time_optional ??????"""
    return _excel_scalar_to_time_optional(v)


def load_attendance_and_analyze(members):
    attendance_data = {}
    # ???????? master ????????????????????????????_????API?generate_plan ??????
    ai_log = {
        "???????????": "??2????????????AI??????????????????JSON??_*_AI_API???",
        "????_AI_API": "??",
        "????_AI_??": "??????????",
    }
    
    # 1. ???????????????
    all_records = []
    try:
        xls = pd.ExcelFile(MASTER_FILE)
        for sheet_name in xls.sheet_names:
            if "?????" in sheet_name or sheet_name.lower() in ['skills', 'need', 'tasks']:
                continue 
                
            m_name = sheet_name.strip()
            if m_name not in members:
                continue 
                
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
            df_sheet.columns = df_sheet.columns.str.strip()
            df_sheet['????'] = m_name 
            all_records.append(df_sheet)
            
        if all_records:
            df = pd.concat(all_records, ignore_index=True)
            df['??'] = pd.to_datetime(df['??'], errors='coerce').dt.date
            df = df.dropna(subset=['??'])
            logging.info(f"?{MASTER_FILE}??????????????????????")
            _cols = {str(c).strip() for c in df.columns}
            if ATT_COL_REMARK in _cols and ATT_COL_LEAVE_TYPE in _cols:
                logging.info(
                    "???: AI ????%s????????????%s????????????????? reason ???????",
                    ATT_COL_REMARK,
                    ATT_COL_LEAVE_TYPE,
                )
            elif ATT_COL_REMARK not in _cols:
                logging.warning(
                    "???????%s??????????????? AI ????????????",
                    ATT_COL_REMARK,
                )
            if ATT_COL_OT_END in _cols:
                logging.info(
                    "???: ???%s???????????????????????????",
                    ATT_COL_OT_END,
                )
        else:
            raise FileNotFoundError("??????????????????????")
            
    except Exception as e:
        logging.warning(f"????????????: {e} ?????????????????")
        default_dates = generate_default_calendar_dates(TARGET_YEAR, TARGET_MONTH)
        records = []
        for d in default_dates:
            for m in members: records.append({'??': d, '????': m, '??': '??'})
        df = pd.DataFrame(records)

    # 2. AI ?????????????????????????? AI ???????????????????
    remarks_to_analyze = []
    for _, row in df.iterrows():
        m = str(row.get('????', '')).strip()
        if m not in members:
            continue
        rem = _attendance_remark_text(row)
        lt = _attendance_leave_type_text(row)
        d_str = row['??'].strftime("%Y-%m-%d") if pd.notna(row['??']) else ""
        if rem:
            remarks_to_analyze.append(f"{d_str}_{m} ???: {rem}")
        elif lt and lt not in ("??", ""):
            remarks_to_analyze.append(f"{d_str}_{m} ???????????: {lt}")

    if remarks_to_analyze:
        remarks_blob = "\n".join(remarks_to_analyze)
        cache_key = hashlib.sha256(
            (remarks_blob + "\n" + ATTENDANCE_REMARK_AI_SCHEMA_ID).encode("utf-8")
        ).hexdigest()
        ai_cache = load_ai_cache()

        # ????????????????????API??????
        cached_data = get_cached_ai_result(ai_cache, cache_key)
        if cached_data is not None:
            ai_parsed = cached_data
            ai_log["????_AI_API"] = "??(???????)"
            ai_log["????_AI_??"] = "????????"
        elif not API_KEY:
            ai_parsed = {}
            ai_log["????_AI_API"] = "??"
            ai_log["????_AI_??"] = "GEMINI_API_KEY??????????AI?????"
            logging.info("GEMINI_API_KEY ????????AI????????????")
        else:
            logging.info("? AI?????????????...")
            ai_log["????_AI_API"] = "??"
            
            prompt = f"""
            ?????????????????????????????????????????JSON????????????
            ????????(``` ?)??????????JSON??????????????

            ?JSON???????????????????
            {{
              "YYYY-MM-DD_?????": {{
                "????": "HH:MM", 
                "????": "HH:MM", 
                "?????": "HH:MM",
                "?????": "HH:MM",
                "????": 1.0,     
                "is_holiday": false,
                "?????": false
              }}
            }}
            ?????????????????????????????????
            ?????/????: ???????????????????????????????????????? null
            ??????/??: ??????????????????????????????????????? null
            ????????:
              - ????????????=> ????? "08:45", ????? "12:00"
              - ???????=> ????? "13:00", ????? "17:00"
            ?is_holiday: ???????????????????? **???????** ?????????? true????????????????? false???????????
            ??????: ?????? **??????????OP/AS ?????????????** ????????? true?????????????????
              ?: ????????????????????????????????????????????????????????????????????????????????
              ??????????????????? false
            ?????: 0.0?1.0???
            
            ?????????
            {chr(10).join(remarks_to_analyze)}
            """
            try:
                client = genai.Client(api_key=API_KEY)
                res = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                match = re.search(r'\{.*\}', res.text, re.DOTALL)
                if match:
                    ai_parsed = json.loads(match.group(0))
                    put_cached_ai_result(ai_cache, cache_key, ai_parsed)
                    save_ai_cache(ai_cache)
                    ai_log["????_AI_??"] = "????"
                else:
                    ai_parsed = {}
                    ai_log["????_AI_??"] = "JSON?????"
            except Exception as e:
                err_text = str(e)
                is_quota_or_rate = ("429" in err_text) or ("RESOURCE_EXHAUSTED" in err_text)
                retry_sec = extract_retry_seconds(err_text) if is_quota_or_rate else None

                if is_quota_or_rate and retry_sec is not None:
                    wait_sec = min(max(retry_sec, 1.0), 90.0)
                    logging.warning(f"AI?? 429/RESOURCE_EXHAUSTED?{wait_sec:.1f}?????1??????????")
                    time_module.sleep(wait_sec)
                    try:
                        res = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                        match = re.search(r'\{.*\}', res.text, re.DOTALL)
                        if match:
                            ai_parsed = json.loads(match.group(0))
                            put_cached_ai_result(ai_cache, cache_key, ai_parsed)
                            save_ai_cache(ai_cache)
                            ai_log["????_AI_??"] = "????????"
                        else:
                            ai_parsed = {}
                            ai_log["????_AI_??"] = "????JSON?????"
                    except Exception as e2:
                        ai_parsed = {}
                        logging.warning(f"AI??????: {e2}")
                        ai_log["????_AI_??"] = f"429??????: {e2}"
                else:
                    ai_parsed = {}
                    logging.warning(f"AI?????: {e}")
                    ai_log["????_AI_??"] = str(e)
    else:
        ai_parsed = {}

    # 3. ????????????
    for _, row in df.iterrows():
        if pd.isna(row['??']): continue
        curr_date = row['??']
        m = str(row.get('????', '')).strip()
        if m not in members: continue

        if curr_date not in attendance_data:
            attendance_data[curr_date] = {}

        original_reason = _attendance_remark_text(row)
        leave_type = _attendance_leave_type_text(row)

        key = f"{curr_date.strftime('%Y-%m-%d')}_{m}"
        ai_info = ai_parsed.get(key, {})

        is_empty_shift = pd.isna(row.get('????')) and pd.isna(row.get('????')) and not ai_info
        is_holiday = _ai_json_bool(ai_info.get("is_holiday"), False) or is_empty_shift
        exclude_from_line = _ai_json_bool(ai_info.get("?????"), False)

        ai_eff = ai_info.get("????")
        excel_eff = row.get('????')
        
        if ai_eff is not None:
            eff_val = ai_eff
        elif excel_eff is not None and not pd.isna(excel_eff):
            eff_val = excel_eff
        else:
            eff_val = 1.0
            
        try:
            efficiency = float(eff_val)
        except (ValueError, TypeError):
            efficiency = 1.0

        if original_reason:
            if (
                leave_type
                and leave_type not in ("??", "")
                and leave_type not in original_reason
            ):
                reason = f"{leave_type} {original_reason}"
            else:
                reason = original_reason
        elif leave_type and leave_type not in ("??", ""):
            reason = leave_type
        else:
            reason = '??' if not is_empty_shift else '?????'

        # ???????????????????????AI???/?????????????????????????????
        excel_s = row.get("????")
        excel_e = row.get("????")
        if not pd.isna(excel_s) and not pd.isna(excel_e):
            start_t = parse_time_str(excel_s, DEFAULT_START_TIME)
            end_t = parse_time_str(excel_e, DEFAULT_END_TIME)
        else:
            start_t = parse_time_str(ai_info.get("????") or excel_s, DEFAULT_START_TIME)
            end_t = parse_time_str(ai_info.get("????") or excel_e, DEFAULT_END_TIME)
        base_end_t = end_t

        b1_s = parse_time_str(row.get('????1_??'), DEFAULT_BREAKS[0][0])
        b1_e = parse_time_str(row.get('????1_??'), DEFAULT_BREAKS[0][1])
        b2_s = parse_time_str(row.get('????2_??'), DEFAULT_BREAKS[1][0])
        b2_e = parse_time_str(row.get('????2_??'), DEFAULT_BREAKS[1][1])

        # ???: AI??????????
        mid_break_s = parse_time_str(ai_info.get("?????"), None)
        mid_break_e = parse_time_str(ai_info.get("?????"), None)
        # AI??????????????????????????????
        if not (mid_break_s and mid_break_e):
            fb_s, fb_e = infer_mid_break_from_reason(reason, start_t, end_t, b1_s, b1_e)
            if fb_s and fb_e:
                mid_break_s, mid_break_e = fb_s, fb_e

        ot_applied_flag = False
        if not is_holiday:
            ot_end = _parse_attendance_overtime_end_optional(row.get(ATT_COL_OT_END))
            if ot_end is not None:
                end_t = ot_end
                ot_applied_flag = True

        def combine_dt(t): return datetime.combine(curr_date, t) if t else None
        
        start_dt = combine_dt(start_t)
        end_dt = combine_dt(end_t)
        if (not is_holiday) and start_dt and end_dt and end_dt <= start_dt:
            logging.warning(
                "?? %s %s: ?????????????????????????????????????????",
                curr_date,
                m,
            )
            end_t = base_end_t
            end_dt = combine_dt(end_t)
        breaks_dt = []
        
        # ????????
        if b1_s and b1_e: breaks_dt.append((combine_dt(b1_s), combine_dt(b1_e)))
        if b2_s and b2_e: breaks_dt.append((combine_dt(b2_s), combine_dt(b2_e)))
        
        # ???: ?????????????????????????????????
        if mid_break_s and mid_break_e: breaks_dt.append((combine_dt(mid_break_s), combine_dt(mid_break_e)))
        
        is_working = not is_holiday
        attendance_data[curr_date][m] = {
            "is_working": is_working,
            "eligible_for_assignment": is_working and (not exclude_from_line),
            "start_dt": start_dt,
            "end_dt": end_dt,
            "breaks_dt": merge_time_intervals(breaks_dt),
            "efficiency": efficiency,
            "reason": reason,
        }

    return attendance_data, ai_log


# ---------------------------------------------------------------------------
# ?????: ?????????? / ??: EC????????????
# ---------------------------------------------------------------------------
ROLL_PIPELINE_EC_PROCESS = "EC"
ROLL_PIPELINE_EC_MACHINE = "EC????"
ROLL_PIPELINE_INSP_PROCESS = "??"
ROLL_PIPELINE_INSP_MACHINE = "???????"
# �B-3: ??? B-2 ???????????????????????? EC ??????????????? B-2 ?????
ROLL_PIPELINE_REWIND_PROCESS = "???"
ROLL_PIPELINE_REWIND_MACHINE = "EC????"
ROLL_PIPELINE_INITIAL_BUFFER_ROLLS = 2
# ??????? min ????????? EC ??????? need??????????????ec_done=0 ??????????????
ROLL_PIPELINE_INSP_UNCAPPED_ROOM = 1.0e18


# ???????????????????????????????????????????????????
# False ?????2????????????????????????????????????
STAGE2_EXTEND_ATTENDANCE_CALENDAR = False
SCHEDULE_EXTEND_MAX_EXTRA_DAYS = 366

# ???????????????????????**????NO??** due_basis ? +1 ??
# ????????????????????**?????????**???????????????????????
# ????????????????????????????(?????????)??????????????????????????
# ?? **False**??????????????????????????????????????????????? True?
STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING = False
# ????? +1 ???????????????NO??????????6 ?????????????????????????????????????
STAGE2_RETRY_SHIFT_DUE_MAX_ROUNDS = 5

# True ??????????????????????NO???????????? 1 ?????
# ?????????????????????**???????????????**???
# ????????1????????????????????????NO???? ???=1???
# ?? False??????NO?????????? STAGE2_SERIAL_DISPATCH_BY_TASK_ID=1 ??????
STAGE2_SERIAL_DISPATCH_BY_TASK_ID = (
    os.environ.get("STAGE2_SERIAL_DISPATCH_BY_TASK_ID", "0")
    .strip()
    .lower()
    in ("1", "true", "yes", "on", "??")
)

# True: ?????????????????1?????????1????????
# ?????????????????? 13:00 ???same_day_raw_start_limit ? 13:00??
# ??????? max ?????????????????????????????
# ?????????????????????? avail ??????????????????????
# ???: ???? STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST=0
STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST = os.environ.get(
    "STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST", "1"
).strip().lower() not in ("0", "false", "no", "off", "???", "??")

# True????: start_date_req<=?? ????????????????????????????????
# ???????????????????????????????????????????????
STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT = os.environ.get(
    "STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT", "1"
).strip().lower() not in ("0", "false", "no", "off", "???", "??")

# True????: ?????????????????????????????????????????
# ??????????????Excel ???? I/O ???????
# False: ????? avail_dt / machine_avail_dt ???
DISPATCH_INTERVAL_MIRROR_ENFORCE = os.environ.get(
    "DISPATCH_INTERVAL_MIRROR_ENFORCE", "1"
).strip().lower() not in ("0", "false", "no", "off", "???", "??")


def _clone_attendance_day_shifted(source_day: dict, old_date: date, new_date: date) -> dict:
    """???????????? new_date ???????????????"""
    delta_days = (new_date - old_date).days
    if delta_days == 0:
        return {m: dict(st) for m, st in source_day.items()}
    delta = timedelta(days=delta_days)
    out: dict = {}
    for m, st in source_day.items():
        new_st = dict(st)
        sd = st.get("start_dt")
        ed = st.get("end_dt")
        new_st["start_dt"] = sd + delta if sd else None
        new_st["end_dt"] = ed + delta if ed else None
        nb = []
        for pair in st.get("breaks_dt") or []:
            if len(pair) >= 2:
                a, b = pair[0], pair[1]
                if a is not None and b is not None:
                    nb.append((a + delta, b + delta))
        new_st["breaks_dt"] = merge_time_intervals(nb)
        out[m] = new_st
    return out


def _pick_extension_template_date(attendance_data: dict, plan_dates: list):
    """??????????1???????????????????????????????????"""
    for i in range(len(plan_dates) - 1, -1, -1):
        d = plan_dates[i]
        day = attendance_data.get(d)
        if not day:
            continue
        if any(
            v.get("eligible_for_assignment", v.get("is_working", False))
            for v in day.values()
        ):
            return d
    return plan_dates[-1] if plan_dates else None


def _extend_attendance_one_calendar_day(
    attendance_data: dict,
    plan_dates: list,
) -> bool:
    """??????1??? plan_dates ????????????????? attendance ???????? False?"""
    if not plan_dates:
        return False
    last_d = plan_dates[-1]
    next_d = last_d + timedelta(days=1)
    tmpl_d = _pick_extension_template_date(attendance_data, plan_dates)
    if tmpl_d is None:
        return False
    template = attendance_data.get(tmpl_d)
    if not template:
        return False
    attendance_data[next_d] = _clone_attendance_day_shifted(template, tmpl_d, next_d)
    plan_dates.append(next_d)
    logging.info(
        "?????????????: %s ????????=%s??????=%s?",
        next_d,
        tmpl_d,
        len(attendance_data[next_d]),
    )
    return True


def _iter_plan_dates_extending(
    plan_dates: list,
    attendance_data: dict,
    task_queue: list,
):
    """
    plan_dates ??????? yield???????????????????1??????????
    plan_dates / attendance_data ?????????????
    """
    si = 0
    ext_used = 0
    while True:
        while si < len(plan_dates):
            yield plan_dates[si]
            si += 1
        pending = any(float(t.get("remaining_units") or 0) > 1e-12 for t in task_queue)
        if not pending:
            return
        if ext_used >= SCHEDULE_EXTEND_MAX_EXTRA_DAYS:
            logging.warning(
                "?????????????????????%s ?????????????????????????????",
                SCHEDULE_EXTEND_MAX_EXTRA_DAYS,
            )
            return
        if not _extend_attendance_one_calendar_day(attendance_data, plan_dates):
            logging.warning(
                "???1???????????????????????????????????????"
            )
            return
        ext_used += 1


def _parse_process_content_tokens(val) -> list[str]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return []
    s = unicodedata.normalize("NFKC", str(val).strip())
    if not s or s.lower() in ("nan", "none", "null"):
        return []
    return [p.strip() for p in s.split(",") if p.strip()]


def _collect_process_content_order_by_task_id(tasks_df) -> dict[str, list[str]]:
    """??NO ? ?????????????????????????????????"""
    out: dict[str, list[str]] = {}
    if tasks_df is None or tasks_df.empty:
        return out
    for _, row in tasks_df.iterrows():
        tid = planning_task_id_str_from_plan_row(row)
        if not tid:
            continue
        parts = _parse_process_content_tokens(row.get(TASK_COL_PROCESS_CONTENT))
        if not parts:
            continue
        if tid not in out:
            out[tid] = parts
    return out


def _process_name_matches_kakou_content_tokens(
    process_name: str, content_tokens: list[str]
) -> bool:
    """
    ?????????????????????????????????????????????
    ????????????????????????????????????? True?
    """
    if not content_tokens:
        return True
    proc = _normalize_process_name_for_rule_match(process_name)
    if not proc:
        return False
    for tok in content_tokens:
        if _normalize_process_name_for_rule_match(tok) == proc:
            return True
    return False


def _process_sequence_rank_for_machine(proc, order_list: list[str]):
    if not order_list:
        return None
    pn = _normalize_process_name_for_rule_match(proc)
    for i, token in enumerate(order_list):
        if _normalize_process_name_for_rule_match(token) == pn:
            return i
    return None


def _task_rank_int_or_none(task) -> int | None:
    r = task.get("process_sequence_rank")
    if r is None:
        return None
    try:
        return int(r)
    except (TypeError, ValueError):
        return None


def _plan_sheet_priority_sort_value(t: dict) -> int:
    """???????????????????????????? 999?"""
    p = t.get("priority", 999)
    try:
        return int(p)
    except (TypeError, ValueError):
        return 999


def _task_blocked_by_same_request_dependency(task, task_queue) -> bool:
    """
    ????NO????????????????????? �A-1?�A-2??
    - ?????????? rank ?????? rank ??????�A-1??
    - ????? rank ????????????????? same_request_line_seq ????�A-2??
    �B-2 / �B-3: ``roll_pipeline_inspection`` ??? ``roll_pipeline_rewind`` ??
    ``roll_pipeline_ec`` ????? �A-1 ???????
    ``_roll_pipeline_inspection_assign_room`` > 0 ????????????????
    ?????? ``_trial_order_flow_eligible_tasks`` ? EC ????????????
    EC ????????????????????????????????????
    """
    tid = str(task.get("task_id", "") or "").strip()
    if not tid:
        return False
    try:
        my_seq = int(task.get("same_request_line_seq", 0))
    except (TypeError, ValueError):
        my_seq = 0
    my_r = _task_rank_int_or_none(task)

    for t2 in task_queue:
        if str(t2.get("task_id", "") or "").strip() != tid:
            continue
        if float(t2.get("remaining_units") or 0) <= 1e-9:
            continue
        r2 = _task_rank_int_or_none(t2)
        try:
            s2 = int(t2.get("same_request_line_seq", 0))
        except (TypeError, ValueError):
            s2 = 0

        if my_r is not None and r2 is not None:
            precedes = r2 < my_r
        elif my_r is None and r2 is None:
            precedes = s2 < my_seq
        else:
            precedes = s2 < my_seq

        if precedes:
            if (
                (
                    task.get("roll_pipeline_inspection")
                    or task.get("roll_pipeline_rewind")
                )
                and t2.get("roll_pipeline_ec")
                and _roll_pipeline_inspection_assign_room(task_queue, tid) > 1e-12
            ):
                continue
            return True
    return False


def _task_not_yet_schedulable_due_to_dependency_or_b2_room(
    task: dict, task_queue: list
) -> bool:
    """
    ?????????????????????????�A ????????????? �B-2/�B-3 ??????
    `_min_pending_dispatch_trial_order_for_date` ? `_equipment_line_lower_dispatch_trial_still_pending`
    ??????????????????????????????????????????????
    """
    if _task_blocked_by_same_request_dependency(task, task_queue):
        return True
    if (task.get("roll_pipeline_inspection") or task.get("roll_pipeline_rewind")) and (
        _roll_pipeline_inspection_assign_room(
            task_queue, str(task.get("task_id", "") or "").strip()
        )
        <= 1e-12
    ):
        return True
    return False


def _row_matches_roll_pipeline_ec(proc, mach) -> bool:
    return (
        _normalize_process_name_for_rule_match(proc)
        == _normalize_process_name_for_rule_match(ROLL_PIPELINE_EC_PROCESS)
        and _normalize_equipment_match_key(mach)
        == _normalize_equipment_match_key(ROLL_PIPELINE_EC_MACHINE)
    )


def _row_matches_roll_pipeline_inspection(proc, mach) -> bool:
    return (
        _normalize_process_name_for_rule_match(proc)
        == _normalize_process_name_for_rule_match(ROLL_PIPELINE_INSP_PROCESS)
        and _normalize_equipment_match_key(mach)
        == _normalize_equipment_match_key(ROLL_PIPELINE_INSP_MACHINE)
    )


def _row_matches_roll_pipeline_rewind(proc, mach) -> bool:
    return (
        _normalize_process_name_for_rule_match(proc)
        == _normalize_process_name_for_rule_match(ROLL_PIPELINE_REWIND_PROCESS)
        and _normalize_equipment_match_key(mach)
        == _normalize_equipment_match_key(ROLL_PIPELINE_REWIND_MACHINE)
    )


def _pipeline_ec_roll_done_units(task_queue, tid: str) -> float:
    tid = str(tid or "").strip()
    s = 0.0
    for t in task_queue:
        if str(t.get("task_id", "") or "").strip() != tid:
            continue
        if not t.get("roll_pipeline_ec"):
            continue
        init = float(t.get("initial_remaining_units") or 0)
        rem = float(t.get("remaining_units") or 0)
        s += max(0.0, init - rem)
    return s


def _pipeline_inspection_roll_done_units(task_queue, tid: str) -> float:
    """????????????????????????"""
    tid = str(tid or "").strip()
    s = 0.0
    for t in task_queue:
        if str(t.get("task_id", "") or "").strip() != tid:
            continue
        if not t.get("roll_pipeline_inspection"):
            continue
        init = float(t.get("initial_remaining_units") or 0)
        rem = float(t.get("remaining_units") or 0)
        s += max(0.0, init - rem)
    return s


def _pipeline_b2_follower_roll_done_units(task_queue, tid: str) -> float:
    """�B-2 ????�B-3 ????????????????????????????"""
    tid = str(tid or "").strip()
    s = 0.0
    for t in task_queue:
        if str(t.get("task_id", "") or "").strip() != tid:
            continue
        if not (t.get("roll_pipeline_inspection") or t.get("roll_pipeline_rewind")):
            continue
        init = float(t.get("initial_remaining_units") or 0)
        rem = float(t.get("remaining_units") or 0)
        s += max(0.0, init - rem)
    return s


def _task_queue_has_roll_pipeline_ec_for_tid(task_queue, task_id: str) -> bool:
    """????NO? EC???????????????????????????"""
    tid = str(task_id or "").strip()
    if not tid:
        return False
    for t in task_queue:
        if str(t.get("task_id", "") or "").strip() != tid:
            continue
        if t.get("roll_pipeline_ec"):
            return True
    return False


def _pipeline_ec_fully_done_for_tid(task_queue, task_id: str) -> bool:
    """????NO? EC ????????????????????????"""
    tid = str(task_id or "").strip()
    if not tid:
        return False
    found = False
    for t in task_queue:
        if str(t.get("task_id", "") or "").strip() != tid:
            continue
        if not t.get("roll_pipeline_ec"):
            continue
        found = True
        if float(t.get("remaining_units") or 0) > 1e-9:
            return False
    return found


def _roll_pipeline_inspection_assign_room(task_queue, task_id: str) -> float:
    tid = str(task_id or "").strip()
    if not _task_queue_has_roll_pipeline_ec_for_tid(task_queue, tid):
        return float(ROLL_PIPELINE_INSP_UNCAPPED_ROOM)
    ec_done = _pipeline_ec_roll_done_units(task_queue, task_id)
    insp_done = _pipeline_b2_follower_roll_done_units(task_queue, task_id)
    # EC ?????????EC ?????????????????????? max_insp ? ec_done ?
    # ???????????????????????? EC ??????????????
    # max_insp - insp_done ? 0 ?????????? eligible ????????????????
    # ?????: ec_fully_done ?? insp_done==max_insp==ec_done ? room=0 ? ????????????
    if _pipeline_ec_fully_done_for_tid(task_queue, task_id):
        return float(ROLL_PIPELINE_INSP_UNCAPPED_ROOM)
    # EC ???: ?????? B ??????????? ec_done ????????B=2 ???????????
    max_insp = max(0.0, ec_done - float(ROLL_PIPELINE_INITIAL_BUFFER_ROLLS) + 1.0)
    _room = max(0.0, max_insp - insp_done)
    return _room


def _roll_pipeline_inspection_task_row_for_tid(
    task_queue: list, task_id: str
) -> dict | None:
    """????NO? �B-2 ?????? �B-3 ?????1???????? None?"""
    tid = str(task_id or "").strip()
    if not tid:
        return None
    for t in task_queue:
        if str(t.get("task_id") or "").strip() != tid:
            continue
        if t.get("roll_pipeline_inspection") or t.get("roll_pipeline_rewind"):
            return t
    return None


def _pipeline_b2_ec_roll_end_datetimes_sorted(
    task_queue: list, task_id: str
) -> list[datetime]:
    """????? EC ????????????????????assigned_history ? end_dt??"""
    tid = str(task_id or "").strip()
    ends: list[datetime] = []
    if not tid:
        return ends
    for t in task_queue:
        if str(t.get("task_id") or "").strip() != tid:
            continue
        if not t.get("roll_pipeline_ec"):
            continue
        for h in t.get("assigned_history") or []:
            ed = h.get("end_dt")
            if isinstance(ed, datetime):
                ends.append(ed)
    ends.sort()
    return ends


def _roll_pipeline_b2_inspection_ec_completion_floor_dt(
    task_queue: list, task_id: str
) -> datetime | None:
    """
    ???????????????????
    ??????????? K?????? B?=ROLL_PIPELINE_INITIAL_BUFFER_ROLLS??????
    EC ?????????? (K+B) ???????????????? end_dt???????????
    ??????: ?????? EC_RollEndCount - KENSA_RollEndCount >= B ??????????????
    ?????????????????
    """
    tid = str(task_id or "").strip()
    if not tid or not _task_queue_has_roll_pipeline_ec_for_tid(task_queue, tid):
        return None
    insp_done = int(
        math.floor(float(_pipeline_b2_follower_roll_done_units(task_queue, tid)))
    )
    need_n = insp_done + int(ROLL_PIPELINE_INITIAL_BUFFER_ROLLS)
    ends = _pipeline_b2_ec_roll_end_datetimes_sorted(task_queue, tid)
    if need_n < 1 or len(ends) < need_n:
        return None
    return ends[need_n - 1]


def _pipeline_b2_team_history_names(team_cell) -> set[str]:
    """assigned_history ? team ?????????,??????????????????NFKC??"""
    if team_cell is None:
        return set()
    s = str(team_cell).strip()
    if not s:
        return set()
    out: set[str] = set()
    for part in re.split(r"[,?]", s):
        t = part.strip()
        if t:
            out.add(unicodedata.normalize("NFKC", t))
    return out


def _pipeline_b2_assigned_member_names_nfkc_for_side(
    task_queue: list, task_id: str, *, ec_side: bool
) -> set[str]:
    """????? EC ???????? assigned_history ????????NFKC ????"""
    tid = str(task_id or "").strip()
    if not tid:
        return set()
    names: set[str] = set()
    for t in task_queue:
        if str(t.get("task_id") or "").strip() != tid:
            continue
        if ec_side:
            if not t.get("roll_pipeline_ec"):
                continue
        else:
            if not (
                t.get("roll_pipeline_inspection") or t.get("roll_pipeline_rewind")
            ):
                continue
        for h in t.get("assigned_history") or []:
            names |= _pipeline_b2_team_history_names(h.get("team"))
    return names


def _b2_ec_insp_pair_in_queue(task_queue: list, task_id: str) -> bool:
    """????NO? �B-2/�B-3 ? EC ???????????????????????????"""
    tid = str(task_id or "").strip()
    if not tid:
        return False
    return bool(
        _task_queue_has_roll_pipeline_ec_for_tid(task_queue, tid)
        and _roll_pipeline_inspection_task_row_for_tid(task_queue, tid) is not None
    )


def _filter_capable_members_b2_disjoint_teams(
    task: dict, task_queue: list, capable_members: list
) -> list:
    """
    �B-2 / �B-3 ???????EC ?????????????????????????????
    ???????? EC ????????
    ??????: ????????????`PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS` ??????
    """
    if not capable_members:
        return capable_members
    tid = str(task.get("task_id") or "").strip()
    if not tid or not _b2_ec_insp_pair_in_queue(task_queue, tid):
        return capable_members
    is_ec = bool(task.get("roll_pipeline_ec"))
    is_follower = bool(
        task.get("roll_pipeline_inspection") or task.get("roll_pipeline_rewind")
    )
    if not is_ec and not is_follower:
        return capable_members
    if not PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS:
        return capable_members
    if is_ec:
        excl = _pipeline_b2_assigned_member_names_nfkc_for_side(
            task_queue, tid, ec_side=False
        )
    else:
        excl = _pipeline_b2_assigned_member_names_nfkc_for_side(
            task_queue, tid, ec_side=True
        )
    if not excl:
        return capable_members
    filtered = [
        m
        for m in capable_members
        if unicodedata.normalize("NFKC", str(m).strip()) not in excl
    ]
    removed = [m for m in capable_members if m not in filtered]
    if removed and _trace_schedule_task_enabled(tid):
        if is_ec:
            _side = "EC"
        elif task.get("roll_pipeline_rewind"):
            _side = "???"
        else:
            _side = "??"
        _log_dispatch_trace_schedule(
            tid,
            "[?????? task=%s] ??????: B-2????? side=%s machine=%s "
            "????=%s ???=%s(%s)",
            tid,
            _side,
            task.get("machine"),
            ",".join(str(x) for x in removed),
            len(filtered),
            ",".join(str(x) for x in filtered) if filtered else "??",
        )
    return filtered


def _exclusive_b1_inspection_holder_for_machine(task_queue, occupant_key: str):
    """
    ??????????????????????�B-2 ???????? �B-3 ???? **???????** ?????????????
    ????? dict ?1???????? None??

    ???????????????????????????????????? **???** ?????????
    ??_???????????????????????????????????????????????????
    """
    m = str(occupant_key or "").strip()
    if not m:
        return None
    holders: list = []
    for t in task_queue:
        _tm = t.get("machine")
        _eqt = str(t.get("equipment_line_key") or _tm or "").strip() or (_tm or "")
        lk = _machine_occupancy_key_resolve(t, _eqt)
        if lk != m:
            continue
        if not (t.get("roll_pipeline_inspection") or t.get("roll_pipeline_rewind")):
            continue
        rem = float(t.get("remaining_units") or 0)
        if rem <= 1e-9:
            continue
        init = float(t.get("initial_remaining_units") or 0)
        started = (init - rem) > 1e-9 or bool(t.get("assigned_history"))
        if not started:
            continue
        holders.append(t)
    if not holders:
        return None
    return min(
        holders,
        key=lambda t: (
            int(t.get("dispatch_trial_order") or 10**9),
            str(t.get("task_id") or ""),
            int(t.get("same_request_line_seq") or 0),
        ),
    )


def _need_sheet_pm_column_rank(
    process,
    machine_name,
    need_combo_col_index: dict | None,
) -> int:
    """need ????????????+???????????????????"""
    if not need_combo_col_index:
        return 10**9
    p = str(process or "").strip()
    m = str(machine_name or "").strip()
    if not p or not m:
        return 10**9
    ck = f"{p}+{m}"
    v = need_combo_col_index.get(ck)
    return int(v) if v is not None else 10**9


def _generate_plan_task_queue_sort_key(
    task: dict,
    _req_map: dict,
    _need_rules: list,
    need_combo_col_index: dict | None = None,
) -> tuple:
    """
    generate_plan ??????????????? task_queue.sort ????

    1. ?????in_progress???
    2. ???? due_basis_date?????????????????
    3. �B-1 ? �B-2/�B-3 ? ? ????b_tier?
    4. �B-2/�B-3 ???? EC ???????????????b2_queue_sub?
    5. need ????????????+????????
    6. ??NO???????_task_id_same_machine_due_tiebreak_key?

    _req_map / _need_rules ?????????????
    """
    insp = bool(task.get("roll_pipeline_inspection"))
    rw = bool(task.get("roll_pipeline_rewind"))
    ip = bool(task.get("in_progress"))
    ec = bool(task.get("roll_pipeline_ec"))
    if insp and ip:
        b_tier = 0  # �B-1
    elif ec or (insp and not ip) or (rw and not ip):
        b_tier = 1  # �B-2 / �B-3 ?
    else:
        b_tier = 2
    if b_tier == 1:
        if ec:
            b2_queue_sub = 0
        elif (insp and not ip) or (rw and not ip):
            b2_queue_sub = 1
        else:
            b2_queue_sub = 2
    else:
        b2_queue_sub = 0
    need_rank = _need_sheet_pm_column_rank(
        task.get("machine"), task.get("machine_name"), need_combo_col_index
    )
    return (
        0 if ip else 1,
        task["due_basis_date"] or date.max,
        b_tier,
        b2_queue_sub,
        need_rank,
        _task_id_same_machine_due_tiebreak_key(task.get("task_id")),
    )


def _reorder_task_queue_b2_ec_inspection_consecutive(task_queue: list) -> None:
    """
    �B-2 / �B-3: ?? task_id ? `roll_pipeline_ec` ?????????????
    ?`roll_pipeline_inspection` ??? `roll_pipeline_rewind`???????????
    """
    if len(task_queue) < 2:
        return
    moved_tids: list[str] = []
    n_rounds = 0
    max_rounds = max(len(task_queue) * 4, 8)
    while n_rounds < max_rounds:
        n_rounds += 1
        by_tid: dict = {}
        for t in task_queue:
            tid = str(t.get("task_id") or "").strip()
            if not tid:
                continue
            if t.get("roll_pipeline_ec"):
                by_tid.setdefault(tid, {})["ec"] = t
            if (t.get("roll_pipeline_inspection") and not t.get("in_progress")) or (
                t.get("roll_pipeline_rewind") and not t.get("in_progress")
            ):
                by_tid.setdefault(tid, {}).setdefault("followers", []).append(t)
        blocks = []
        for tid, d in by_tid.items():
            ec_task = d.get("ec")
            followers = d.get("followers") or []
            if ec_task is None or not followers:
                continue
            followers = sorted(
                followers,
                key=lambda x: (
                    int(x.get("same_request_line_seq") or 0),
                    task_queue.index(x),
                ),
            )
            blocks.append((tid, ec_task, followers))
        if not blocks:
            break
        blocks.sort(key=lambda b: task_queue.index(b[1]))
        moved = False
        for tid, ec_task, followers in blocks:
            chain = [ec_task] + followers
            try:
                indices = [task_queue.index(x) for x in chain]
            except ValueError:
                continue
            if all(indices[i] == indices[0] + i for i in range(len(indices))):
                continue
            insert_at = min(indices)
            for idx in sorted(indices, reverse=True):
                task_queue.pop(idx)
            for j, item in enumerate(chain):
                task_queue.insert(insert_at + j, item)
            moved_tids.append(tid)
            moved = True
            break
        if not moved:
            break
    if moved_tids:
        logging.info(
            "�B-2/�B-3 ?????: EC ?????????????????????NO: %s",
            ",".join(moved_tids),
        )


def _assign_sequential_dispatch_trial_order(task_queue: list) -> None:
    """
    `task_queue` ?????????? `dispatch_trial_order` ? 1..n ??????
    `_reorder_task_queue_b2_ec_inspection_consecutive` ??????????????????????
    EC ??????????????????????
    """
    for i, t in enumerate(task_queue, start=1):
        t["dispatch_trial_order"] = i


def _task_queue_all_have_sheet_dispatch_trial_order(task_queue: list) -> bool:
    """???????????????????????????????????"""
    if not task_queue:
        return False
    for t in task_queue:
        v = t.get("dispatch_trial_order_from_sheet")
        if v is None:
            return False
        try:
            iv = int(v)
        except (TypeError, ValueError):
            return False
        if iv < 1:
            return False
    return True


def _apply_dispatch_trial_order_for_generate_plan(
    task_queue: list,
    req_map: dict,
    need_rules: list,
    need_combo_col_index: dict | None,
) -> None:
    """
    ??????????????????????????????�B-2/3 ??????????????
    ??????????????????need ??????????EC ???? 1..n ????
    """
    if _task_queue_all_have_sheet_dispatch_trial_order(task_queue):
        task_queue.sort(
            key=lambda t: (
                int(t.get("dispatch_trial_order_from_sheet") or 10**9),
                int(t.get("planning_sheet_row_seq") or 10**9),
            )
        )
        for t in task_queue:
            t["dispatch_trial_order"] = int(t.get("dispatch_trial_order_from_sheet") or 10**9)
        logging.info(
            "??????: ?%s????????????????? %s ???",
            RESULT_TASK_COL_DISPATCH_TRIAL_ORDER,
            len(task_queue),
        )
        return
    task_queue.sort(
        key=lambda x: _generate_plan_task_queue_sort_key(
            x, req_map, need_rules, need_combo_col_index
        )
    )
    _reorder_task_queue_b2_ec_inspection_consecutive(task_queue)
    _assign_sequential_dispatch_trial_order(task_queue)
    logging.info(
        "??????: ???????????????? 1..%s ????????",
        len(task_queue),
    )


def fill_plan_dispatch_trial_order_column_stage1(
    plan_df: "pd.DataFrame",
    run_date: date,
    req_map: dict,
    need_rules: list,
    need_combo_col_index: dict | None,
    equipment_list: list,
) -> None:
    """
    ??1?? DataFrame ?????????????2 ????????????�B-2/3 ???????????
    ?????????????
    """
    if plan_df is None or getattr(plan_df, "empty", True):
        return
    if RESULT_TASK_COL_DISPATCH_TRIAL_ORDER not in plan_df.columns:
        return
    col = RESULT_TASK_COL_DISPATCH_TRIAL_ORDER
    global_priority_raw = load_main_sheet_global_priority_override_text()
    members_for_gpo: list = []
    try:
        with pd.ExcelFile(MASTER_FILE) as _xf:
            _skills = pd.read_excel(_xf, sheet_name="skills", header=None)
        for r in range(2, _skills.shape[0]):
            cell = _skills.iat[r, 0]
            if pd.isna(cell):
                continue
            name = str(cell).strip()
            if name and name.lower() not in ("nan", "none", "null"):
                members_for_gpo.append(name)
    except Exception:
        members_for_gpo = []
    gpo = analyze_global_priority_override_comment(
        global_priority_raw, members_for_gpo, run_date.year, ai_sheet_sink={}
    )
    tq = build_task_queue_from_planning_df(
        plan_df,
        run_date,
        req_map,
        None,
        gpo,
        equipment_list,
    )
    _apply_dispatch_trial_order_for_generate_plan(
        tq, req_map, need_rules, need_combo_col_index
    )
    try:
        col_idx = plan_df.columns.get_loc(col)
    except Exception:
        return
    for t in tq:
        iloc = t.get("planning_df_iloc")
        if iloc is None:
            continue
        if not isinstance(iloc, int) or iloc < 0 or iloc >= len(plan_df):
            continue
        dto = t.get("dispatch_trial_order")
        if dto is None:
            continue
        try:
            # Excel ????????????????????????????????????????????
            plan_df.iat[iloc, col_idx] = int(dto)
        except (TypeError, ValueError):
            plan_df.iat[iloc, col_idx] = float("nan")


def _equipment_schedule_unified_sub_string_map(timeline_for_eq_grid: list) -> dict:
    """
    ????????????????NO ??????????????????????????
    ?????????????? `sub` ????????????????? ", " ?????
    ????????????????????? `sub` ?????????????
    """
    acc: dict = defaultdict(set)
    for e in timeline_for_eq_grid or []:
        if not _is_machining_timeline_event(e):
            continue
        tid = str(e.get("task_id") or "").strip()
        m = str(e.get("machine") or "").strip()
        d0 = e.get("date")
        if not tid or not m or d0 is None:
            continue
        for s in str(e.get("sub") or "").split(","):
            t = s.strip()
            if t:
                acc[(d0, m, tid)].add(t)
    return {k: ", ".join(sorted(v)) for k, v in acc.items() if v}


def _build_equipment_schedule_dataframe(
    sorted_dates: list,
    equipment_list: list,
    attendance_data: dict,
    timeline_events: list,
    *,
    first_eq_schedule_cell_by_task_id: dict | None = None,
) -> "pd.DataFrame":
    """
    ??_???????????? DataFrame?10 ????????????
    first_eq_schedule_cell_by_task_id ???????????????????????????????
    """
    timeline_for_eq_grid = _expand_timeline_events_for_equipment_grid(timeline_events)
    _eq_sched_unify_sub = _equipment_schedule_unified_sub_string_map(timeline_for_eq_grid)
    events_by_date = defaultdict(list)
    for e in timeline_for_eq_grid:
        events_by_date[e["date"]].append(e)

    all_eq_rows = []
    eq_empty_cols = {}
    for eq in equipment_list:
        eq_empty_cols[eq] = ""
        eq_empty_cols[f"{eq}??"] = ""

    for d in sorted_dates:
        d_start = datetime.combine(d, DEFAULT_START_TIME)
        d_end = datetime.combine(d, DEFAULT_END_TIME)
        events_today = events_by_date[d]
        machine_to_events = defaultdict(list)
        for ev in events_today:
            machine_to_events[ev["machine"]].append(ev)
        for _eq_k, _evs in machine_to_events.items():
            _evs.sort(
                key=lambda e: (e.get("start_dt") or datetime.min, str(e.get("task_id") or ""))
            )

        is_anyone_working = any(
            daily_status["is_working"] for daily_status in attendance_data[d].values()
        )
        if not events_today and not is_anyone_working:
            continue

        all_eq_rows.append({"???": f"? {d.strftime('%Y/%m/%d (%a)')} ?", **eq_empty_cols})

        def _eq_cell_display_sub(ev, day_d) -> str:
            tid0 = str(ev.get("task_id") or "").strip()
            m0 = str(ev.get("machine") or "").strip()
            if tid0 and m0:
                u0 = _eq_sched_unify_sub.get((day_d, m0, tid0))
                if u0 is not None:
                    return u0
            return str(ev.get("sub") or "").strip()

        curr_grid = d_start
        while curr_grid < d_end:
            next_grid = curr_grid + timedelta(minutes=10)
            if next_grid > d_end:
                next_grid = d_end

            mid_t = curr_grid + (next_grid - curr_grid) / 2
            row_data = {
                "???": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"
            }

            for eq in equipment_list:
                eq_text = ""
                progress_text = ""
                active_ev = None
                for ev in machine_to_events.get(eq, ()):
                    if ev["start_dt"] <= mid_t < ev["end_dt"]:
                        active_ev = ev
                        break

                if active_ev:
                    _use_prog = (
                        _is_machining_timeline_event(active_ev)
                        and all(
                            k in active_ev
                            for k in (
                                "eff_time_per_unit",
                                "units_done",
                                "total_units",
                                "already_done_units",
                            )
                        )
                        and float(active_ev.get("eff_time_per_unit") or 0) > 0
                    )
                    if any(b_s <= mid_t < b_e for b_s, b_e in active_ev["breaks"]):
                        eq_text = "??"
                    elif not _use_prog:
                        _ek_disp = _timeline_event_kind(active_ev)
                        _tag = {
                            TIMELINE_EVENT_MACHINE_DAILY_STARTUP: "??????",
                            TIMELINE_EVENT_CHANGEOVER_CLEANUP: "???????",
                            TIMELINE_EVENT_CHANGEOVER_PREP: "?????",
                        }.get(
                            _ek_disp,
                            "??????",
                        )
                        _sub_n = _eq_cell_display_sub(active_ev, d)
                        _sub_text = f" ?:{_sub_n}" if _sub_n else ""
                        _tid_d = str(active_ev.get("task_id") or "").strip()
                        eq_text = (
                            f"[{_tid_d}] ?:{active_ev.get('op', '')}{_sub_text} ({_tag})"
                        )
                        progress_text = ""
                    else:
                        elapsed = get_actual_work_minutes(
                            active_ev["start_dt"],
                            min(next_grid, active_ev["end_dt"]),
                            active_ev["breaks"],
                        )
                        block_done_now = min(
                            int(elapsed / active_ev["eff_time_per_unit"]),
                            active_ev["units_done"],
                        )

                        cumulative_done = active_ev["already_done_units"] + block_done_now
                        total_u = active_ev["total_units"]

                        _sub_s = _eq_cell_display_sub(active_ev, d)
                        sub_text = f" ?:{_sub_s}" if _sub_s else ""
                        eq_text = f"[{active_ev['task_id']}] ?:{active_ev['op']}{sub_text}"
                        progress_text = f"{cumulative_done}/{total_u}R"
                        _tid_sched = str(active_ev.get("task_id") or "").strip()
                        if (
                            first_eq_schedule_cell_by_task_id is not None
                            and _tid_sched
                            and _tid_sched not in first_eq_schedule_cell_by_task_id
                        ):
                            _row_ex = len(all_eq_rows) + 2
                            _ci = 2 + 2 * equipment_list.index(eq)
                            first_eq_schedule_cell_by_task_id[_tid_sched] = (
                                f"{get_column_letter(_ci)}{_row_ex}"
                            )

                row_data[eq] = eq_text
                row_data[f"{eq}??"] = progress_text

            all_eq_rows.append(row_data)
            curr_grid = next_grid
        all_eq_rows.append({"???": "", **eq_empty_cols})

    df_eq = pd.DataFrame(all_eq_rows)
    _eq_hdr = _equipment_schedule_header_labels(equipment_list)
    _eq_rename = {}
    for _eq, _lab in zip(equipment_list, _eq_hdr):
        if _eq in df_eq.columns:
            _eq_rename[_eq] = _lab
        _pqc = f"{_eq}??"
        if _pqc in df_eq.columns:
            _eq_rename[_pqc] = f"{_lab}??"
    if _eq_rename:
        df_eq = df_eq.rename(columns=_eq_rename)
    return df_eq


def _machine_display_key_for_equipment(eq: str) -> str:
    """skills ??????+???????????????????????????????"""
    s = str(eq).strip()
    if "+" in s:
        mpart = s.split("+", 1)[1].strip()
        return mpart if mpart else s
    return s


def _build_equipment_schedule_by_machine_name_dataframe(
    sorted_dates: list,
    equipment_list: list,
    attendance_data: dict,
    timeline_events: list,
) -> "pd.DataFrame":
    """
    ????????????? 10 ?????????NO???????????????
    ?????????????+??????????????????????????1???????
    """
    timeline_for_eq_grid = _expand_timeline_events_for_equipment_grid(timeline_events)
    events_by_date = defaultdict(list)
    for e in timeline_for_eq_grid:
        events_by_date[e["date"]].append(e)

    # ?????????????????1?????? equipment_list ???????????
    occ_key_to_header: dict[str, str] = {}
    machine_cols: list[str] = []
    eq_to_mcol: dict[str, str] = {}
    for eq in equipment_list:
        occ_key = _equipment_line_key_to_physical_occupancy_key(eq)
        if not occ_key:
            occ_key = _normalize_equipment_match_key(str(eq).strip())
        disp = _machine_display_key_for_equipment(eq).strip() or str(eq).strip()
        if occ_key not in occ_key_to_header:
            occ_key_to_header[occ_key] = disp
            machine_cols.append(disp)
        eq_to_mcol[eq] = occ_key_to_header[occ_key]

    empty_tail = {mcol: "" for mcol in machine_cols}
    all_rows = []

    for d in sorted_dates:
        d_start = datetime.combine(d, DEFAULT_START_TIME)
        d_end = datetime.combine(d, DEFAULT_END_TIME)
        events_today = events_by_date[d]
        machine_to_events = defaultdict(list)
        for ev in events_today:
            machine_to_events[ev["machine"]].append(ev)
        for _eq_k, _evs in machine_to_events.items():
            _evs.sort(
                key=lambda e: (e.get("start_dt") or datetime.min, str(e.get("task_id") or ""))
            )

        is_anyone_working = any(
            daily_status["is_working"] for daily_status in attendance_data[d].values()
        )
        if not events_today and not is_anyone_working:
            continue

        all_rows.append({"???": f"? {d.strftime('%Y/%m/%d (%a)')} ?", **empty_tail})

        curr_grid = d_start
        while curr_grid < d_end:
            next_grid = curr_grid + timedelta(minutes=10)
            if next_grid > d_end:
                next_grid = d_end
            mid_t = curr_grid + (next_grid - curr_grid) / 2
            row_data = {
                "???": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"
            }
            for mcol in machine_cols:
                row_data[mcol] = ""
            tids_by_mcol: dict[str, set[str]] = defaultdict(set)
            for eq, evs in machine_to_events.items():
                mcol = eq_to_mcol.get(eq)
                if not mcol:
                    continue
                active_ev = None
                for ev in evs:
                    if ev["start_dt"] <= mid_t < ev["end_dt"]:
                        active_ev = ev
                        break
                if not active_ev:
                    continue
                if any(b_s <= mid_t < b_e for b_s, b_e in active_ev["breaks"]):
                    tids_by_mcol[mcol].add("????")
                else:
                    tid = str(active_ev.get("task_id") or "").strip()
                    if tid:
                        tids_by_mcol[mcol].add(tid)
            for mcol in machine_cols:
                parts = sorted(tids_by_mcol.get(mcol, ()))
                row_data[mcol] = "?".join(parts) if parts else ""
            all_rows.append(row_data)
            curr_grid = next_grid
        all_rows.append({"???": "", **empty_tail})

    return pd.DataFrame(all_rows)


def _build_block_table_dataframe(
    sorted_dates: list,
    equipment_list: list,
    members: list,
    attendance_data: dict,
    timeline_events: list,
) -> "pd.DataFrame":
    """
    ??????????NO??????????? 10 ??????????????????
    """
    timeline_for_eq_grid = _expand_timeline_events_for_equipment_grid(timeline_events)
    events_by_date = defaultdict(list)
    for e in timeline_for_eq_grid:
        events_by_date[e["date"]].append(e)

    _eq_hdr = _equipment_schedule_header_labels(equipment_list)
    eq_disp_to_key: dict[str, str] = {}
    for eq, lab in zip(equipment_list, _eq_hdr):
        eq_disp_to_key[f"??:{lab}"] = eq

    mem_cols = [f"?:{m}" for m in members]
    eq_cols = [f"??:{lab}" for lab in _eq_hdr]
    all_cols = ["???"] + eq_cols + mem_cols
    rows_out = []

    for d in sorted_dates:
        d_start = datetime.combine(d, DEFAULT_START_TIME)
        d_end = datetime.combine(d, DEFAULT_END_TIME)
        events_today = events_by_date[d]
        machine_to_events = defaultdict(list)
        for ev in events_today:
            machine_to_events[ev["machine"]].append(ev)
        for _evs in machine_to_events.values():
            _evs.sort(
                key=lambda e: (e.get("start_dt") or datetime.min, str(e.get("task_id") or ""))
            )

        is_anyone_working = any(
            daily_status["is_working"] for daily_status in attendance_data[d].values()
        )
        if not events_today and not is_anyone_working:
            continue

        banner = {"???": f"? {d.strftime('%Y/%m/%d (%a)')} ?"}
        banner.update({c: "" for c in all_cols if c != "???"})
        rows_out.append(banner)

        curr_grid = d_start
        while curr_grid < d_end:
            next_grid = curr_grid + timedelta(minutes=10)
            if next_grid > d_end:
                next_grid = d_end
            mid_t = curr_grid + (next_grid - curr_grid) / 2
            row_data: dict = {
                "???": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"
            }
            for c in eq_cols + mem_cols:
                row_data[c] = ""

            for col_eq, lab in zip(eq_cols, _eq_hdr):
                eq_key = eq_disp_to_key.get(col_eq)
                if not eq_key:
                    continue
                active_ev = None
                for ev in machine_to_events.get(eq_key, ()):
                    if ev["start_dt"] <= mid_t < ev["end_dt"]:
                        active_ev = ev
                        break
                if not active_ev:
                    continue
                if any(b_s <= mid_t < b_e for b_s, b_e in active_ev["breaks"]):
                    row_data[col_eq] = "??"
                else:
                    tid = str(active_ev.get("task_id") or "").strip()
                    row_data[col_eq] = tid if tid else "??"

            busy_member_task: dict[str, set[str]] = defaultdict(set)
            for ev in events_today:
                st = ev.get("start_dt")
                ed = ev.get("end_dt")
                if not isinstance(st, datetime) or not isinstance(ed, datetime):
                    continue
                if not (st <= mid_t < ed):
                    continue
                tid = str(ev.get("task_id") or "").strip()
                op = str(ev.get("op") or "").strip()
                if op:
                    if any(
                        b_s <= mid_t < b_e for b_s, b_e in ev.get("breaks") or ()
                    ):
                        busy_member_task[op].add("??" if tid else "??")
                    elif tid:
                        busy_member_task[op].add(tid)
                for s in str(ev.get("sub") or "").split(","):
                    s = s.strip()
                    if not s:
                        continue
                    if any(
                        b_s <= mid_t < b_e for b_s, b_e in ev.get("breaks") or ()
                    ):
                        busy_member_task[s].add("??")
                    elif tid:
                        busy_member_task[s].add(tid)

            for m in members:
                col_m = f"?:{m}"
                parts = sorted(busy_member_task.get(m, ()))
                row_data[col_m] = "?".join(parts) if parts else ""

            rows_out.append(row_data)
            curr_grid = next_grid

        tail = {"???": ""}
        tail.update({c: "" for c in all_cols if c != "???"})
        rows_out.append(tail)

    return pd.DataFrame(rows_out, columns=all_cols)


def _day_schedule_task_sort_key(
    task: dict,
    _task_queue: list | None = None,
    need_combo_col_index: dict | None = None,
):
    """
    ???????????STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST=0 ????????
    ????? _generate_plan_task_queue_sort_key ?????????????? due_basis_date?�B ??b2_queue_sub?need ?????NO??
    ??? �B-1 ????????????? rank?dispatch_trial_order?�B-2 ?? EC ?????????????
    ??????????????? _equipment_line_lower_dispatch_trial_still_pending ??????????
    STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT=1 ???? _task_blocked_by_global_dispatch_trial_order ?
    ??????????????????????????????
    """
    raw_r = task.get("process_sequence_rank")
    if raw_r is None:
        r = 10**9
    else:
        r = int(raw_r)
    try:
        line_seq = int(task.get("same_request_line_seq", 0))
    except (TypeError, ValueError):
        line_seq = 0
    try:
        dto = int(task.get("dispatch_trial_order") or 10**9)
    except (TypeError, ValueError):
        dto = 10**9
    insp = bool(task.get("roll_pipeline_inspection"))
    rw = bool(task.get("roll_pipeline_rewind"))
    ip = bool(task.get("in_progress"))
    ec = bool(task.get("roll_pipeline_ec"))
    if insp and ip:
        b_tier = 0
    elif ec or (insp and not ip) or (rw and not ip):
        b_tier = 1
    else:
        b_tier = 2
    if b_tier == 1:
        if ec:
            b2_queue_sub = 0
        elif (insp and not ip) or (rw and not ip):
            b2_queue_sub = 1
        else:
            b2_queue_sub = 2
    else:
        b2_queue_sub = 0
    if ec:
        b2_roll_pipeline_stage = 0
    elif (insp and not ip) or (rw and not ip):
        b2_roll_pipeline_stage = 1
    else:
        b2_roll_pipeline_stage = 2
    dbk = task.get("due_basis_date")
    if not isinstance(dbk, date):
        dbk = date.max
    need_rank = _need_sheet_pm_column_rank(
        task.get("machine"), task.get("machine_name"), need_combo_col_index
    )
    tb = _task_id_same_machine_due_tiebreak_key(task.get("task_id"))
    b1_trial_early = (0, dto) if (insp and ip) else (1, 0)
    return (
        (
            0 if ip else 1,
            dbk,
            b_tier,
            b2_queue_sub,
            need_rank,
            tb,
            b1_trial_early,
            r,
            line_seq,
            dto,
            b2_roll_pipeline_stage,
            _plan_sheet_priority_sort_value(task),
        )
        + _result_task_sheet_sort_key(task)
    )


def _equipment_line_lower_dispatch_trial_still_pending(
    task_queue: list,
    machine_occ_key: str,
    my_dispatch_order: int,
    current_date: date,
    *,
    daily_status: dict | None = None,
    members: list | None = None,
    machine_avail_dt: dict | None = None,
    machine_day_start: datetime | None = None,
    machine_handoff: dict | None = None,
    skills_dict: dict | None = None,
    abolish_all_scheduling_limits: bool = False,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
    assign_probe_ctx: dict | None = None,
) -> bool:
    """
    ???????machine ??????????????????????????????
    machine_avail_dt ???????????????????????????????????
    ????????????? _task_blocked_by_global_dispatch_trial_order ????????

    ???????????????????????tasks_today ????
    start_date_req <= current_date ???????????????????
    ????????????????????????????????????????????

    ??????????? **??????????????????**?????????????????
    ????? eligible ?????????????????????????????????????

    ??????????? **?????????????????????**?????????????????
    ?????????????????????????????????????????
    """
    line = (machine_occ_key or "").strip()
    if not line:
        return False
    try:
        my_o = int(my_dispatch_order)
    except (TypeError, ValueError):
        my_o = 10**9
    for t in task_queue:
        if float(t.get("remaining_units") or 0) <= 1e-12:
            continue
        _sdr = t.get("start_date_req")
        if not isinstance(_sdr, date) or _sdr > current_date:
            continue
        _tm = t.get("machine")
        _eqt = str(t.get("equipment_line_key") or _tm or "").strip() or (_tm or "")
        t_occ = _machine_occupancy_key_resolve(t, _eqt)
        if t_occ != line:
            continue
        try:
            o = int(t.get("dispatch_trial_order") or 10**9)
        except (TypeError, ValueError):
            o = 10**9
        if o < my_o:
            if _task_not_yet_schedulable_due_to_dependency_or_b2_room(t, task_queue):
                continue
            if _task_fully_machine_calendar_blocked_on_date(
                t, current_date, daily_status, members
            ):
                continue
            if _task_no_machining_window_left_from_avail_floor(
                t,
                current_date,
                daily_status,
                members,
                machine_avail_dt,
                machine_day_start,
                machine_handoff=machine_handoff,
                skills_dict=skills_dict,
                abolish_all_scheduling_limits=abolish_all_scheduling_limits,
                dispatch_interval_mirror=dispatch_interval_mirror,
            ):
                continue
            if assign_probe_ctx is not None and _trial_order_assign_probe_fails(
                t, current_date, daily_status, assign_probe_ctx
            ):
                continue
            return True
    return False


def _min_pending_dispatch_trial_order_for_date(
    task_queue: list,
    current_date: date,
    *,
    daily_status: dict | None = None,
    members: list | None = None,
    machine_avail_dt: dict | None = None,
    machine_day_start: datetime | None = None,
    machine_handoff: dict | None = None,
    skills_dict: dict | None = None,
    abolish_all_scheduling_limits: bool = False,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
) -> int | None:
    """
    start_date_req <= current_date ?????????????????????
    _equipment_line_lower_dispatch_trial_still_pending ??????????????????
    ????????????????

    **????????????**?STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT????
    ???????????????????????????????????????
    �A-1/�A-2 ???????????????????????????????????
    ????????????
    - `_task_not_yet_schedulable_due_to_dependency_or_b2_room` ? True ??
    - ?daily_status?members ???????????????????????????
    - ?machine_avail_dt ??????????????????????????????

    1 ????????????????????`_effective_min_dispatch_trial_order_from_pool` ??????????
    """
    pool = _tasks_in_min_pending_dispatch_pool(
        task_queue,
        current_date,
        daily_status=daily_status,
        members=members,
        machine_avail_dt=machine_avail_dt,
        machine_day_start=machine_day_start,
        machine_handoff=machine_handoff,
        skills_dict=skills_dict,
        abolish_all_scheduling_limits=abolish_all_scheduling_limits,
        dispatch_interval_mirror=dispatch_interval_mirror,
    )
    orders: list[int] = []
    for t in pool:
        try:
            orders.append(int(t.get("dispatch_trial_order") or 10**9))
        except (TypeError, ValueError):
            orders.append(10**9)
    return min(orders) if orders else None


def _task_blocked_by_global_dispatch_trial_order(
    task: dict,
    task_queue: list,
    current_date: date,
    *,
    daily_status: dict | None = None,
    members: list | None = None,
    machine_avail_dt: dict | None = None,
    machine_day_start: datetime | None = None,
    machine_handoff: dict | None = None,
    skills_dict: dict | None = None,
    abolish_all_scheduling_limits: bool = False,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
    min_dispatch_effective: int | None = None,
) -> bool:
    """
    ?????????????????????????????????????????
    min_dispatch_effective: ??????????????????????????????????????
    """
    if not STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT:
        return False
    if min_dispatch_effective is not None:
        m = min_dispatch_effective
    else:
        m = _min_pending_dispatch_trial_order_for_date(
            task_queue,
            current_date,
            daily_status=daily_status,
            members=members,
            machine_avail_dt=machine_avail_dt,
            machine_day_start=machine_day_start,
            machine_handoff=machine_handoff,
            skills_dict=skills_dict,
            abolish_all_scheduling_limits=abolish_all_scheduling_limits,
            dispatch_interval_mirror=dispatch_interval_mirror,
        )
    if m is None:
        return False
    try:
        my_o = int(task.get("dispatch_trial_order") or 10**9)
    except (TypeError, ValueError):
        my_o = 10**9
    return my_o > m


def _purge_attendance_days_not_in_set(attendance_data: dict, keep_dates: frozenset) -> None:
    """??????????????????????????????????"""
    for dk in list(attendance_data.keys()):
        if dk not in keep_dates:
            del attendance_data[dk]


def _partial_task_id_due_shift_outcome(
    task_queue: list, task_id: str, calendar_last: date
) -> tuple[bool, bool]:
    """
    ??????NO??????+1?????????
    ???: (shift_ok, calendar_shortfall)
    - shift_ok: ?????due_basis_date???????????????? +1 ????????????
    - calendar_shortfall: ????????????????? +1 ??????????????
    ????????????? (False, False)????????????
    """
    tid = (task_id or "").strip()
    if not tid:
        return False, False
    rows = [t for t in task_queue if str(t.get("task_id", "") or "").strip() == tid]
    if not rows:
        return False, False
    basis_rows = [t for t in rows if t.get("due_basis_date") is not None]
    if not basis_rows:
        return False, False
    for t in basis_rows:
        db = t["due_basis_date"]
        if db + timedelta(days=1) > calendar_last:
            return False, True
    return True, False


def _shift_task_due_calendar_fields_one_day(task: dict, run_date: date) -> None:
    """
    ????????: **????????due_basis_date???**? +1 ????
    ??_??????? ``due_basis_date_result_sheet`` ???????+1 ?????????
    ???????????????????????
    due_urgent ????? due_basis_date ???????
    """
    if task.get("due_basis_date") is not None:
        task["due_basis_date"] = task["due_basis_date"] + timedelta(days=1)
    db = task.get("due_basis_date")
    if db is not None:
        task["due_urgent"] = db <= run_date


def _seed_avail_from_timeline_for_date(
    timeline_events: list,
    current_date: date,
    machine_avail_dt: dict,
    avail_dt: dict,
    machine_day_start: datetime,
) -> None:
    """??????? timeline ??????????????????????????????"""
    for e in timeline_events:
        if e.get("date") != current_date:
            continue
        end_dt = e.get("end_dt")
        if end_dt is None or not hasattr(end_dt, "replace"):
            continue
        occ = str(e.get("machine_occupancy_key") or "").strip()
        if not occ:
            mraw = str(e.get("machine") or "").strip()
            occ = (
                _normalize_equipment_match_key(mraw.split("+", 1)[1])
                if "+" in mraw
                else _normalize_equipment_match_key(mraw)
            )
        if occ:
            prev = machine_avail_dt.get(occ, machine_day_start)
            if end_dt > prev:
                machine_avail_dt[occ] = end_dt
        op = str(e.get("op") or "").strip()
        if op and op in avail_dt:
            prev_m = avail_dt[op]
            if end_dt > prev_m:
                avail_dt[op] = end_dt
        sub_raw = e.get("sub") or ""
        for sn in str(sub_raw).split(","):
            sm = sn.strip()
            if sm and sm in avail_dt:
                prev_s = avail_dt[sm]
                if end_dt > prev_s:
                    avail_dt[sm] = end_dt


def _merge_machine_calendar_intervals(
    intervals: list[tuple[datetime, datetime]],
) -> list[tuple[datetime, datetime]]:
    if not intervals:
        return []
    iv = sorted(intervals, key=lambda x: (x[0], x[1]))
    out = [iv[0]]
    for s, e in iv[1:]:
        ps, pe = out[-1]
        if s <= pe:
            out[-1] = (ps, max(pe, e))
        else:
            out.append((s, e))
    return out


def _bump_dt_past_machine_calendar_blocks(
    t: datetime,
    blocks: list[tuple[datetime, datetime]],
) -> datetime:
    """???????? [start,end) ? t ??????????????"""
    if not blocks:
        return t
    changed = True
    while changed:
        changed = False
        for s, e in blocks:
            if s <= t < e:
                t = e
                changed = True
                break
    return t


def _machine_cal_parse_slot_datetime(cell) -> datetime | None:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return None
    try:
        dt = pd.to_datetime(cell, errors="coerce")
    except Exception:
        return None
    if dt is None or (isinstance(dt, float) and pd.isna(dt)):
        return None
    if isinstance(dt, pd.Timestamp):
        dt = dt.to_pydatetime()
    if getattr(dt, "tzinfo", None) is not None:
        dt = dt.replace(tzinfo=None)
    return dt


def _machine_cal_cell_is_occupied(cell) -> bool:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return False
    if isinstance(cell, str):
        return bool(cell.strip())
    if isinstance(cell, bool):
        return cell
    # Excel ? 0 ????????????????? 0 ????????? True ??????????????
    if isinstance(cell, (int, float)):
        try:
            return float(cell) != 0.0
        except (TypeError, ValueError):
            return True
    return True


def _clip_machine_calendar_slot_to_factory_window(
    day_d: date, slot_start: datetime, slot_end: datetime
) -> tuple[datetime, datetime] | None:
    """
    ???????1???? [slot_start, slot_end) ??????????????
    ?????????? None???????????2?? master ??? A12/B12 ?
    DEFAULT_START_TIME / DEFAULT_END_TIME ???????generate_plan ?????????????
    """
    w0 = datetime.combine(day_d, DEFAULT_START_TIME)
    w1 = datetime.combine(day_d, DEFAULT_END_TIME)
    s2 = max(slot_start, w0)
    e2 = min(slot_end, w1)
    if s2 < e2:
        return (s2, e2)
    return None


def _machine_calendar_planning_window_end_dt(
    current_date: date,
    daily_status: dict,
    members: list,
) -> datetime:
    """
    ??????????????????????????DEFAULT_END_TIME???
    ????????????????????????????????????????
    ????????????????????
    """
    w_factory = datetime.combine(current_date, DEFAULT_END_TIME)
    ends: list[datetime] = []
    for m in members:
        if m not in daily_status:
            continue
        st = daily_status[m]
        if not st.get("eligible_for_assignment", st.get("is_working", False)):
            continue
        et = st.get("end_dt")
        if et is not None and hasattr(et, "replace"):
            ends.append(et)
    if not ends:
        return w_factory
    return min(w_factory, min(ends))


def _clip_machine_busy_blocks_to_planning_window(
    blocks: list[tuple[datetime, datetime]],
    w0: datetime,
    w1: datetime,
) -> list[tuple[datetime, datetime]]:
    """??????? [w0, w1) ???????????????"""
    out: list[tuple[datetime, datetime]] = []
    for s, e in blocks or []:
        s2 = max(s, w0)
        e2 = min(e, w1)
        if s2 < e2:
            out.append((s2, e2))
    if not out:
        return []
    return _merge_machine_calendar_intervals(out)


def _machine_cal_resolve_column_to_equipment_key(
    p_raw,
    m_raw,
    eq_lookup: dict,
    elist_set: set,
) -> str | None:
    p_s = (
        str(p_raw).strip()
        if p_raw is not None and not (isinstance(p_raw, float) and pd.isna(p_raw))
        else ""
    )
    m_s = (
        str(m_raw).strip()
        if m_raw is not None and not (isinstance(m_raw, float) and pd.isna(m_raw))
        else ""
    )
    if p_s and m_s:
        combo = f"{p_s}+{m_s}"
    elif p_s:
        combo = p_s
    else:
        return None
    if combo in elist_set:
        return combo
    nk = _normalize_equipment_match_key(combo)
    return eq_lookup.get(nk)


def load_machine_calendar_occupancy_blocks(
    master_path: str,
    equipment_list: list,
) -> dict[date, dict[str, list[tuple[datetime, datetime]]]]:
    """
    master.xlsm???????????????????????? 1 ?????????????
    ??: ?? -> equipment_list ??? -> ???? [start, end) ????????????
    """
    if not master_path or not os.path.isfile(master_path):
        return {}
    try:
        xls = pd.ExcelFile(master_path)
        if SHEET_MACHINE_CALENDAR not in xls.sheet_names:
            return {}
        raw = pd.read_excel(master_path, sheet_name=SHEET_MACHINE_CALENDAR, header=None)
    except Exception as e:
        logging.warning("???????: ?????????????? (%s)", e)
        return {}
    if raw.shape[0] < 3 or raw.shape[1] < 3:
        return {}

    ncols = raw.shape[1]
    non_empty_pm = 0
    for c in range(2, ncols):
        p = raw.iat[0, c]
        m = raw.iat[1, c]
        if pd.isna(p) or pd.isna(m):
            continue
        p_s = str(p).strip()
        m_s = str(m).strip()
        if p_s and m_s and p_s.lower() != "nan" and m_s.lower() != "nan":
            non_empty_pm += 1
    use_two_header = non_empty_pm > 0

    eq_lookup = _equipment_lookup_normalized_to_canonical(equipment_list)
    elist_set = set(str(x).strip() for x in equipment_list if str(x).strip())
    col_to_eq: dict[int, str] = {}
    for c in range(2, ncols):
        p = raw.iat[0, c]
        m = raw.iat[1, c] if use_two_header else None
        if use_two_header:
            if pd.isna(p) or pd.isna(m):
                continue
            p_s = str(p).strip()
            m_s = str(m).strip()
            if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
                continue
        else:
            if pd.isna(p):
                continue
            p_s = str(p).strip()
            if not p_s or p_s.lower() == "nan":
                continue
            m_s = ""
        canon = _machine_cal_resolve_column_to_equipment_key(
            p_s, m_s, eq_lookup, elist_set
        )
        if canon:
            col_to_eq[c] = canon

    if not col_to_eq:
        return {}

    acc: dict[date, dict[str, list[tuple[datetime, datetime]]]] = defaultdict(
        lambda: defaultdict(list)
    )
    for r in range(2, raw.shape[0]):
        slot0 = _machine_cal_parse_slot_datetime(raw.iat[r, 0])
        if slot0 is None:
            continue
        try:
            day_d = slot0.date()
        except Exception:
            continue
        for c, eq_key in col_to_eq.items():
            if c >= raw.shape[1]:
                continue
            cell = raw.iat[r, c]
            if not _machine_cal_cell_is_occupied(cell):
                continue
            slot_start = slot0
            slot_end = slot_start + timedelta(hours=1)
            _clipped_mc = _clip_machine_calendar_slot_to_factory_window(
                day_d, slot_start, slot_end
            )
            if _clipped_mc is None:
                continue
            slot_start, slot_end = _clipped_mc
            acc[day_d][eq_key].append((slot_start, slot_end))

    out: dict[date, dict[str, list[tuple[datetime, datetime]]]] = {}
    for d, eqmap in acc.items():
        merged_eq = {
            eq: _merge_machine_calendar_intervals(iv)
            for eq, iv in eqmap.items()
            if iv
        }
        phys_accum: dict[str, list] = defaultdict(list)
        for eq, iv in merged_eq.items():
            pk = _equipment_line_key_to_physical_occupancy_key(str(eq).strip())
            if pk:
                phys_accum[pk].extend(iv)
        merged_all = dict(merged_eq)
        for pk, iv in phys_accum.items():
            merged_all[pk] = _merge_machine_calendar_intervals(iv)
        out[d] = merged_all
    return out


def _apply_machine_calendar_floor_for_date(
    current_date: date,
    machine_avail_dt: dict,
    equipment_list: list,
    machine_day_start: datetime,
    *,
    machine_calendar_plan_end: datetime | None = None,
) -> None:
    """?????????????????????????????????????"""
    day_blocks = _MACHINE_CALENDAR_BLOCKS_BY_DATE.get(current_date)
    if not day_blocks:
        return
    candidates: set[str] = set()
    for k in machine_avail_dt.keys():
        sk = str(k).strip() if k is not None else ""
        if sk:
            candidates.add(sk)
    for el in equipment_list:
        ek = str(el).strip() if el is not None else ""
        if not ek:
            continue
        pk = _equipment_line_key_to_physical_occupancy_key(ek)
        if pk:
            candidates.add(pk)
    w0 = machine_day_start
    w1 = machine_calendar_plan_end
    if w1 is None:
        w1 = datetime.combine(current_date, DEFAULT_END_TIME)
    for eq_s in candidates:
        blocks = day_blocks.get(eq_s) or _machine_calendar_blocks_for_occ_key(
            day_blocks, eq_s
        )
        if not blocks:
            continue
        blocks_c = _clip_machine_busy_blocks_to_planning_window(blocks, w0, w1)
        if not blocks_c:
            continue
        t0 = machine_avail_dt.get(eq_s, machine_day_start)
        t1 = _bump_dt_past_machine_calendar_blocks(t0, blocks_c)
        if t1 > t0:
            machine_avail_dt[eq_s] = t1


def _machine_calendar_blocks_for_occ_key(
    day_blocks: dict[str, list[tuple[datetime, datetime]]],
    occ: str,
) -> list[tuple[datetime, datetime]] | None:
    """day_blocks ?????????????????????????????"""
    o = str(occ or "").strip()
    if not o or not day_blocks:
        return None
    if o in day_blocks:
        return day_blocks[o]
    nk = _normalize_equipment_match_key(o)
    for k, iv in day_blocks.items():
        if _normalize_equipment_match_key(str(k)) == nk:
            return iv
    return None


def _machine_calendar_occ_blocks_full_plan_window(
    occ_key: str,
    current_date: date,
    daily_status: dict,
    members: list,
) -> bool:
    """
    ???????????????? [??, min(??,????????) ) ??????
    ???????? 1 ????????????? True?
    """
    day_blocks = _MACHINE_CALENDAR_BLOCKS_BY_DATE.get(current_date)
    if not day_blocks:
        return False
    blocks = _machine_calendar_blocks_for_occ_key(day_blocks, occ_key)
    if not blocks:
        return False
    w0 = datetime.combine(current_date, DEFAULT_START_TIME)
    w1 = _machine_calendar_planning_window_end_dt(current_date, daily_status, members)
    blocks_c = _clip_machine_busy_blocks_to_planning_window(blocks, w0, w1)
    if not blocks_c:
        return False
    t1 = _bump_dt_past_machine_calendar_blocks(w0, blocks_c)
    return t1 >= w1


def _task_fully_machine_calendar_blocked_on_date(
    t: dict,
    current_date: date,
    daily_status: dict | None,
    members: list | None,
) -> bool:
    """
    ??????????????????????????????????????
    ??????????????????????????????????????????
    """
    if daily_status is None or members is None:
        return False
    _tm = t.get("machine")
    _eqt = str(t.get("equipment_line_key") or _tm or "").strip() or (_tm or "")
    occ = (_machine_occupancy_key_resolve(t, _eqt) or "").strip()
    if not occ:
        return False
    return _machine_calendar_occ_blocks_full_plan_window(
        occ, current_date, daily_status, members
    )


def _task_no_machining_window_left_from_avail_floor(
    t: dict,
    current_date: date,
    daily_status: dict | None,
    members: list | None,
    machine_avail_dt: dict | None,
    machine_day_start: datetime | None,
    *,
    machine_handoff: dict | None = None,
    skills_dict: dict | None = None,
    abolish_all_scheduling_limits: bool = False,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
) -> bool:
    """
    machine_avail_dt???????????????????????????
    ?????????????????????????????????
    `machine_handoff` ??????? `_resolve_machine_changeover_floor_segments` ???
    `_assign_one_roll_trial_order_flow` ??? **????????** ?????
    ??? machine_avail ??????????????????????? min_dto ?????????
    ???????????????????? **????? 1 ?????????**
    ????????? True???????????????
    ?????????????????????
    """
    if (
        daily_status is None
        or members is None
        or machine_avail_dt is None
        or machine_day_start is None
    ):
        return False
    w1 = _machine_calendar_planning_window_end_dt(current_date, daily_status, members)
    _tm = t.get("machine")
    _eqt = str(t.get("equipment_line_key") or _tm or "").strip() or (_tm or "")
    occ = (_machine_occupancy_key_resolve(t, _eqt) or "").strip()
    if not occ:
        return False
    use_co = machine_handoff is not None and skills_dict is not None
    if use_co:
        machine_name = str(t.get("machine_name", "") or "").strip()
        machine_proc = str(_tm or "").strip()
        eq_line = str(
            t.get("equipment_line_key") or _tm or ""
        ).strip() or str(_tm or "")
        machine_occ_key = _machine_occupancy_key_resolve(t, eq_line)
        t_floor, _segs, abort = _resolve_machine_changeover_floor_segments(
            abolish_all_scheduling_limits=bool(abolish_all_scheduling_limits),
            machine_occ_key=machine_occ_key,
            task_id=str(t.get("task_id") or "").strip(),
            eq_line=eq_line,
            machine_name=machine_name,
            machine_proc=machine_proc,
            machine_avail_dt=machine_avail_dt,
            machine_day_floor=machine_day_start,
            current_date=current_date,
            machine_handoff=machine_handoff,
            daily_status=daily_status,
            skills_dict=skills_dict,
            dispatch_interval_mirror=dispatch_interval_mirror,
        )
        if abort:
            return True
    else:
        t_floor = machine_avail_dt.get(occ)
        if t_floor is None:
            nk = _normalize_equipment_match_key(occ)
            for k, v in machine_avail_dt.items():
                if _normalize_equipment_match_key(str(k)) == nk:
                    t_floor = v
                    break
        if t_floor is None:
            t_floor = machine_day_start
    if t_floor >= w1:
        return True
    rem = w1 - t_floor
    if rem <= timedelta(0):
        return True
    btp = parse_float_safe(t.get("base_time_per_unit"), 0.0)
    if btp <= 0:
        return False
    t_eff = parse_float_safe(t.get("task_eff_factor"), 1.0)
    if t_eff <= 0:
        t_eff = 1.0
    # eff_time_per_unit ? base / avg_eff / t_eff � ?????avg_eff ???????????
    _avg_eff_floor = 0.5
    approx_need_mins = max(1.0, float(btp) / t_eff / _avg_eff_floor)
    return rem < timedelta(minutes=approx_need_mins)


def _bump_machine_avail_after_roll_for_calendar(
    current_date: date,
    eq_line: str,
    machine_avail_dt: dict,
    *,
    machine_calendar_plan_end: datetime | None = None,
    machine_day_floor: datetime | None = None,
) -> None:
    """???????: ????????????????????????????"""
    day_blocks = _MACHINE_CALENDAR_BLOCKS_BY_DATE.get(current_date)
    if not day_blocks:
        return
    eq_s = str(eq_line).strip() if eq_line is not None else ""
    if not eq_s:
        return
    blocks = day_blocks.get(eq_s)
    if not blocks:
        return
    t0 = machine_avail_dt.get(eq_s)
    if t0 is None:
        return
    w0 = (
        machine_day_floor
        if machine_day_floor is not None
        else datetime.combine(current_date, DEFAULT_START_TIME)
    )
    w1 = (
        machine_calendar_plan_end
        if machine_calendar_plan_end is not None
        else datetime.combine(current_date, DEFAULT_END_TIME)
    )
    blocks_c = _clip_machine_busy_blocks_to_planning_window(blocks, w0, w1)
    if not blocks_c:
        return
    t1 = _bump_dt_past_machine_calendar_blocks(t0, blocks_c)
    if t1 > t0:
        machine_avail_dt[eq_s] = t1


def _parse_nonneg_minutes_cell(v) -> int:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0
    try:
        n = int(round(float(v)))
    except (TypeError, ValueError):
        return 0
    return max(0, n)


def _df_pick_column(df, *candidates: str) -> str | None:
    cols = [str(c).strip() for c in df.columns]
    low_map = {str(c).strip().lower(): str(c).strip() for c in df.columns}
    for cand in candidates:
        c0 = str(cand).strip()
        if c0 in df.columns:
            return c0
        cl = c0.lower()
        if cl in low_map:
            return low_map[cl]
    return None


def load_machine_changeover_settings(
    master_path: str,
) -> tuple[dict[str, tuple[int, int]], dict[str, int]]:
    """
    master.xlsm ??????:
      - ???_?????????� ?????????????????1 ??????2 ????????
      - ???_??_???????� ???????????

    ??NO???????????????????????????????????????????
    ????????????????NO????????????????
    ??????????????????????????????????

    ??: (????????+??????????? -> (???, ????),
          ??????????? -> ?????)
    """
    changeover: dict[str, tuple[int, int]] = {}
    startup: dict[str, int] = {}
    if not master_path or not os.path.isfile(master_path):
        return changeover, startup
    try:
        xls = pd.ExcelFile(master_path)
    except Exception as e:
        logging.warning("????/????: ????????? (%s)", e)
        return changeover, startup

    if SHEET_MACHINE_CHANGEOVER in xls.sheet_names:
        try:
            df = pd.read_excel(
                master_path, sheet_name=SHEET_MACHINE_CHANGEOVER, header=0
            )
            df.columns = [str(c).strip() for c in df.columns]
            c_proc = _df_pick_column(df, "???", "??")
            c_mac = _df_pick_column(df, "???", "??")
            c_prep = _df_pick_column(
                df,
                "????_?",
                "???",
                "?????_?",
                "???????_?",
            )
            c_clean = _df_pick_column(
                df,
                "?????_?",
                "????",
                "??????_?",
                "????????_?",
            )
            if c_proc and c_mac and c_prep and c_clean:
                n_ent = 0
                for _, row in df.iterrows():
                    p = row.get(c_proc)
                    m = row.get(c_mac)
                    if p is None or m is None:
                        continue
                    if isinstance(p, float) and pd.isna(p):
                        continue
                    if isinstance(m, float) and pd.isna(m):
                        continue
                    p_s = str(p).strip()
                    m_s = str(m).strip()
                    if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
                        continue
                    combo = f"{p_s}+{m_s}"
                    prep = _parse_nonneg_minutes_cell(row.get(c_prep))
                    clean = _parse_nonneg_minutes_cell(row.get(c_clean))
                    if prep == 0 and clean == 0:
                        continue
                    changeover[combo] = (prep, clean)
                    nk = _normalize_equipment_match_key(combo)
                    if nk:
                        changeover[nk] = (prep, clean)
                    n_ent += 1
                if n_ent:
                    logging.info(
                        "????%s?: ??+?? %s ????/???????????????",
                        SHEET_MACHINE_CHANGEOVER,
                        n_ent,
                    )
        except Exception as e:
            logging.warning(
                "????%s?????????: %s", SHEET_MACHINE_CHANGEOVER, e
            )

    if SHEET_MACHINE_DAILY_STARTUP in xls.sheet_names:
        try:
            df2 = pd.read_excel(
                master_path, sheet_name=SHEET_MACHINE_DAILY_STARTUP, header=0
            )
            df2.columns = [str(c).strip() for c in df2.columns]
            c_mn = _df_pick_column(df2, "???", "??")
            c_su = _df_pick_column(
                df2, "??????_?", "????_?", "?????_?"
            )
            if c_mn and c_su:
                for _, row in df2.iterrows():
                    mn = row.get(c_mn)
                    if mn is None or (isinstance(mn, float) and pd.isna(mn)):
                        continue
                    mn_s = str(mn).strip()
                    if not mn_s or mn_s.lower() == "nan":
                        continue
                    su = _parse_nonneg_minutes_cell(row.get(c_su))
                    if su <= 0:
                        continue
                    startup[mn_s] = su
                    nk = _normalize_equipment_match_key(mn_s)
                    if nk:
                        startup[nk] = su
                if startup:
                    logging.info(
                        "????%s?: ?? %s ????????????????????",
                        SHEET_MACHINE_DAILY_STARTUP,
                        len({k for k in startup if "+" not in str(k)}),
                    )
        except Exception as e:
            logging.warning(
                "????%s?????????: %s", SHEET_MACHINE_DAILY_STARTUP, e
            )

    return changeover, startup


def _lookup_changeover_minutes_for_eq(
    eq_line: str,
    by_eq: dict[str, tuple[int, int]] | None,
) -> tuple[int, int]:
    mp = by_eq if by_eq is not None else _STAGE2_MACHINE_CHANGEOVER_BY_EQ
    k0 = str(eq_line or "").strip()
    if not k0:
        return (0, 0)
    if k0 in mp:
        return mp[k0]
    nk = _normalize_equipment_match_key(k0)
    if nk in mp:
        return mp[nk]
    for k, v in mp.items():
        if _normalize_equipment_match_key(str(k)) == nk:
            return v
    return (0, 0)


def _lookup_daily_startup_minutes(
    machine_name: str,
    by_m: dict[str, int] | None,
) -> int:
    st = by_m if by_m is not None else _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE
    mn = str(machine_name or "").strip()
    if not mn:
        return 0
    if mn in st:
        return st[mn]
    nk = _normalize_equipment_match_key(mn)
    if nk in st:
        return st[nk]
    for k, v in st.items():
        if _normalize_equipment_match_key(str(k)) == nk:
            return v
    return 0


def _timeline_event_kind(ev: dict) -> str:
    k = str(ev.get("event_kind") or "").strip()
    return k if k else TIMELINE_EVENT_MACHINING


def _is_machining_timeline_event(ev: dict) -> bool:
    return _timeline_event_kind(ev) == TIMELINE_EVENT_MACHINING


def _pick_skilled_op_for_changeover_interval(
    machine_proc: str,
    machine_name: str,
    skills_dict: dict,
    daily_status: dict,
) -> str | None:
    """
    ?? eligible ????????+??? OP ?????????????????1??
    ???????????????????avail_dt ??????
    """
    cands: list[tuple[int, str]] = []
    proc = (machine_proc or "").strip()
    mnm = (machine_name or "").strip()
    for mem, st in daily_status.items():
        if not st.get("eligible_for_assignment", st.get("is_working", False)):
            continue
        srow = skills_dict.get(mem, {})
        if proc and mnm:
            v = srow.get(f"{proc}+{mnm}", "")
        elif mnm:
            v = srow.get(mnm, "")
        elif proc:
            v = srow.get(proc, "")
        else:
            v = ""
        role, prio = parse_op_as_skill_cell(v)
        if role == "OP":
            cands.append((prio, mem))
    if not cands:
        return None
    return min(cands)[1]


def _machine_effective_floor_timedelta_only(
    machine_occ_key: str,
    task_id: str,
    eq_line: str,
    machine_name: str,
    machine_avail_dt: dict,
    machine_handoff: dict,
    machine_day_floor: datetime,
    abolish_limits: bool,
    *,
    changeover_by_eq: dict[str, tuple[int, int]] | None = None,
    daily_startup_by_machine: dict[str, int] | None = None,
    current_date: date | None = None,
) -> datetime:
    """??? OP ?????????????????????????????????????????? max??"""
    if abolish_limits:
        return machine_day_floor
    mf = machine_avail_dt.get(machine_occ_key, machine_day_floor)
    mto = machine_handoff.get("machining_today_occ") or machine_handoff.get(
        "started_today", set()
    )
    if machine_occ_key not in mto:
        su = _lookup_daily_startup_minutes(machine_name, daily_startup_by_machine)
        if su:
            reg_ts = _STAGE2_REGULAR_SHIFT_START
            if reg_ts is not None:
                cd = current_date if current_date is not None else machine_day_floor.date()
                reg_end = datetime.combine(cd, reg_ts) + timedelta(minutes=su)
                mf = max(mf, reg_end)
            else:
                mf = mf + timedelta(minutes=su)
    prev_tid = (machine_handoff.get("last_tid") or {}).get(machine_occ_key)
    cur_tid = str(task_id or "").strip()
    prev_eq = (machine_handoff.get("last_eq") or {}).get(machine_occ_key, "")
    if prev_tid and cur_tid and prev_tid != cur_tid:
        _, cu = _lookup_changeover_minutes_for_eq(prev_eq, changeover_by_eq)
        if cu:
            mf = mf + timedelta(minutes=cu)
    prep, _ = _lookup_changeover_minutes_for_eq(eq_line, changeover_by_eq)
    _pt = str(prev_tid or "").strip()
    _ct = str(cur_tid or "").strip()
    if prep > 0 and (not _pt or _pt != _ct):
        mf = mf + timedelta(minutes=prep)
    return mf


def _changeover_plan_segments_and_machining_lower_bound(
    *,
    prev_machining_end_dt: datetime,
    machine_day_floor: datetime,
    current_date: date,
    machine_occ_key: str,
    task_id: str,
    eq_line: str,
    machine_name: str,
    machine_proc: str,
    machine_handoff: dict,
    daily_status: dict,
    skills_dict: dict,
    abolish_limits: bool,
) -> tuple[datetime | None, list[dict]]:
    """
    ???????? prev_machining_end_dt ??????????????????????????????
    ?????(????????, ??????????????) ????
    ????? master ??? A15??????????? [??, ??+N?) ??????? forward ?????
    A15 ???????????????????? OP ?????????? forward ???
    ????????????????NO??????????????????????
    ?????????? op ????????????????????????????? op ? forward ?????????
    ????? dict ? start_dt, end_dt, op, event_kind, machine, machine_occupancy_key ????
    """
    if abolish_limits:
        return prev_machining_end_dt, []
    mach_occ = str(machine_occ_key or "").strip()
    reg_ts = _STAGE2_REGULAR_SHIFT_START
    machining_today_occ = machine_handoff.get("machining_today_occ") or machine_handoff.get(
        "started_today", set()
    )
    last_tid = (machine_handoff.get("last_tid") or {}).get(mach_occ, "")
    last_eq = (machine_handoff.get("last_eq") or {}).get(mach_occ, "")
    last_d = (machine_handoff.get("last_machining_date") or {}).get(mach_occ)
    last_lead = (machine_handoff.get("last_lead_op") or {}).get(mach_occ, "")
    cur_tid = str(task_id or "").strip()
    su = _lookup_daily_startup_minutes(machine_name, None)
    prep, _cu_line = _lookup_changeover_minutes_for_eq(eq_line, None)
    _prep_unused, cu_prev = (
        _lookup_changeover_minutes_for_eq(last_eq, None) if last_eq else (0, 0)
    )

    rep = _pick_skilled_op_for_changeover_interval(
        machine_proc, machine_name, skills_dict, daily_status
    )
    st_r = daily_status.get(rep) if rep else None
    br_r = merge_time_intervals(list(st_r.get("breaks_dt") or [])) if st_r else []
    end_r = st_r["end_dt"] if st_r else None
    start_r = st_r["start_dt"] if st_r else None

    segments: list[dict] = []
    t = prev_machining_end_dt

    if mach_occ not in machining_today_occ and su > 0:
        if reg_ts is not None:
            reg_start_dt = datetime.combine(current_date, reg_ts)
            reg_end_dt = reg_start_dt + timedelta(minutes=su)
            segments.append(
                {
                    "start_dt": reg_start_dt,
                    "end_dt": reg_end_dt,
                    "op": "",
                    "event_kind": TIMELINE_EVENT_MACHINE_DAILY_STARTUP,
                    "machine": eq_line,
                    "machine_occupancy_key": mach_occ,
                }
            )
            t = max(t, reg_end_dt)
        else:
            if rep is None or not st_r or end_r is None or start_r is None:
                return None, []
            t0 = max(t, machine_day_floor, start_r)
            ce, act, rem = calculate_end_time(t0, su, br_r, end_r)
            if rem > 0 or act < su:
                return None, []
            segments.append(
                {
                    "start_dt": t0,
                    "end_dt": ce,
                    "op": "",
                    "event_kind": TIMELINE_EVENT_MACHINE_DAILY_STARTUP,
                    "machine": eq_line,
                    "machine_occupancy_key": mach_occ,
                }
            )
            t = ce

    need_cleanup = (
        bool(last_tid)
        and bool(cur_tid)
        and last_tid != cur_tid
        and last_d == current_date
        and cu_prev > 0
        and mach_occ in machining_today_occ
    )
    if need_cleanup:
        cop = (str(last_lead).strip() or rep)
        if not cop:
            return None, []
        st_c = daily_status.get(cop) or st_r
        if not st_c:
            return None, []
        br_c = merge_time_intervals(list(st_c.get("breaks_dt") or []))
        end_c = st_c["end_dt"]
        cs = prev_machining_end_dt
        ce, act, rem = calculate_end_time(cs, cu_prev, br_c, end_c)
        if rem > 0 or act < cu_prev:
            return None, []
        segments.append(
            {
                "start_dt": cs,
                "end_dt": ce,
                "op": cop,
                "event_kind": TIMELINE_EVENT_CHANGEOVER_CLEANUP,
                "machine": last_eq or eq_line,
                "machine_occupancy_key": mach_occ,
            }
        )
        t = max(t, ce)

    _lt_s = str(last_tid or "").strip()
    need_prep = prep > 0 and (not _lt_s or _lt_s != cur_tid)
    if need_prep:
        if rep is None or not st_r or end_r is None:
            return None, []
        pe, act, rem = calculate_end_time(t, prep, br_r, end_r)
        if rem > 0 or act < prep:
            return None, []
        segments.append(
            {
                "start_dt": t,
                "end_dt": pe,
                "op": rep,
                "event_kind": TIMELINE_EVENT_CHANGEOVER_PREP,
                "machine": eq_line,
                "machine_occupancy_key": mach_occ,
            }
        )
        t = pe

    return t, segments


def _machine_effective_floor_for_assign(
    machine_occ_key: str,
    task_id: str,
    eq_line: str,
    machine_name: str,
    machine_avail_dt: dict,
    machine_handoff: dict,
    machine_day_floor: datetime,
    abolish_limits: bool,
    *,
    changeover_by_eq: dict[str, tuple[int, int]] | None = None,
    daily_startup_by_machine: dict[str, int] | None = None,
    current_date: date | None = None,
    daily_status: dict | None = None,
    skills_dict: dict | None = None,
    machine_proc: str | None = None,
) -> datetime:
    """
    ????????????????????????????
    daily_status?skills_dict?current_date ???????skills ?? OP ??????????
    ???????????? forward ????????????????????????????????
    """
    if abolish_limits:
        return machine_day_floor
    prev_mach = machine_avail_dt.get(machine_occ_key, machine_day_floor)
    if (
        current_date is not None
        and daily_status is not None
        and skills_dict is not None
        and machine_proc is not None
    ):
        lb, _segs = _changeover_plan_segments_and_machining_lower_bound(
            prev_machining_end_dt=prev_mach,
            machine_day_floor=machine_day_floor,
            current_date=current_date,
            machine_occ_key=machine_occ_key,
            task_id=task_id,
            eq_line=eq_line,
            machine_name=machine_name,
            machine_proc=str(machine_proc or "").strip(),
            machine_handoff=machine_handoff,
            daily_status=daily_status,
            skills_dict=skills_dict,
            abolish_limits=False,
        )
        if lb is not None:
            return lb
    return _machine_effective_floor_timedelta_only(
        machine_occ_key,
        task_id,
        eq_line,
        machine_name,
        machine_avail_dt,
        machine_handoff,
        machine_day_floor,
        False,
        changeover_by_eq=changeover_by_eq,
        daily_startup_by_machine=daily_startup_by_machine,
        current_date=current_date,
    )


def _resolve_machine_changeover_floor_segments(
    *,
    abolish_all_scheduling_limits: bool,
    machine_occ_key: str,
    task_id: str,
    eq_line: str,
    machine_name: str,
    machine_proc: str,
    machine_avail_dt: dict,
    machine_day_floor: datetime,
    current_date: date,
    machine_handoff: dict,
    daily_status: dict,
    skills_dict: dict,
    dispatch_interval_mirror: DispatchIntervalMirror | None,
) -> tuple[datetime, list[dict], bool]:
    """
    ?????????????????????????????
    ??? (floor_dt, segments, abort)?abort ? True ??????????????????????
    """
    if abolish_all_scheduling_limits:
        prev = machine_avail_dt.get(machine_occ_key, machine_day_floor)
        return prev, [], False
    prev_mach = machine_avail_dt.get(machine_occ_key, machine_day_floor)
    co_lb, co_segs = _changeover_plan_segments_and_machining_lower_bound(
        prev_machining_end_dt=prev_mach,
        machine_day_floor=machine_day_floor,
        current_date=current_date,
        machine_occ_key=machine_occ_key,
        task_id=task_id,
        eq_line=eq_line,
        machine_name=machine_name,
        machine_proc=str(machine_proc or "").strip(),
        machine_handoff=machine_handoff,
        daily_status=daily_status,
        skills_dict=skills_dict,
        abolish_limits=False,
    )
    if co_lb is None:
        if (
            _pick_skilled_op_for_changeover_interval(
                str(machine_proc or "").strip(),
                str(machine_name or "").strip(),
                skills_dict,
                daily_status,
            )
            is None
        ):
            mf = _machine_effective_floor_timedelta_only(
                machine_occ_key,
                task_id,
                eq_line,
                machine_name,
                machine_avail_dt,
                machine_handoff,
                machine_day_floor,
                False,
                current_date=current_date,
            )
            return mf, [], False
        return machine_day_floor, [], True
    if dispatch_interval_mirror is not None and co_segs:
        _last_sub_m = machine_handoff.get("last_machining_sub") or {}
        for seg in co_segs:
            sop = str(seg.get("op") or "").strip()
            sok = str(seg.get("machine_occupancy_key") or machine_occ_key).strip()
            st_seg = seg.get("start_dt")
            ed_seg = seg.get("end_dt")
            ek_chk = str(seg.get("event_kind") or "").strip()
            if not isinstance(st_seg, datetime) or not isinstance(ed_seg, datetime):
                continue
            if ek_chk == TIMELINE_EVENT_CHANGEOVER_CLEANUP and sok:
                _sc = str(_last_sub_m.get(sok, "") or "").strip()
                _team_chk: list[str] = []
                if sop:
                    _team_chk.append(sop)
                for _p in _sc.split(","):
                    _t = _p.strip()
                    if _t and _t not in _team_chk:
                        _team_chk.append(_t)
                if dispatch_interval_mirror.would_block_roll(
                    sok, tuple(_team_chk), st_seg, ed_seg
                ):
                    return machine_day_floor, [], True
                continue
            if (
                sop
                and dispatch_interval_mirror.would_block_member(sop, st_seg, ed_seg)
            ):
                return machine_day_floor, [], True
            if (
                sok
                and dispatch_interval_mirror.would_block_equipment(
                    sok, st_seg, ed_seg
                )
            ):
                return machine_day_floor, [], True
    return co_lb, co_segs, False


def _changeover_timeline_op_sub_for_event(
    *,
    event_kind: str,
    op_from_segment: str,
    machine_occ_key: str,
    machining_lead_op: str,
    machining_sub_str: str,
    machine_handoff: dict,
    daily_status: dict,
) -> tuple[str, str]:
    """?????????????????????????????????? handoff ???????"""
    ek = str(event_kind or "").strip()
    op_s = str(op_from_segment or "").strip()
    _lead = str(machining_lead_op or "").strip()
    _sub_new = str(machining_sub_str or "").strip()
    if ek == TIMELINE_EVENT_MACHINE_DAILY_STARTUP:
        return "", ""
    if ek == TIMELINE_EVENT_CHANGEOVER_PREP:
        op_u = _lead if _lead in daily_status else op_s
        return op_u, _sub_new
    if ek == TIMELINE_EVENT_CHANGEOVER_CLEANUP:
        mocc = str(machine_occ_key or "").strip()
        sub_prev = str(
            (machine_handoff.get("last_machining_sub") or {}).get(mocc, "") or ""
        ).strip()
        return op_s, sub_prev
    return op_s, ""


def _append_changeover_segments_to_timeline(
    timeline_events: list,
    dispatch_interval_mirror: DispatchIntervalMirror | None,
    avail_dt: dict,
    daily_status: dict,
    *,
    current_date: date,
    task_id: str,
    machine_occ_key: str,
    segments: list[dict],
    machining_lead_op: str | None = None,
    machining_sub_str: str | None = None,
    machine_handoff: dict | None = None,
) -> None:
    """??????????????????????????? avail ????"""
    _mh = machine_handoff or {}
    _lead_m = str(machining_lead_op or "").strip()
    _sub_roll = str(machining_sub_str or "").strip()
    for seg in segments or []:
        op_seg = str(seg.get("op") or "").strip()
        st = seg.get("start_dt")
        ed = seg.get("end_dt")
        if not isinstance(st, datetime) or not isinstance(ed, datetime):
            continue
        m_line = str(seg.get("machine") or "").strip()
        m_occ = str(seg.get("machine_occupancy_key") or machine_occ_key).strip()
        ek = str(seg.get("event_kind") or "").strip() or TIMELINE_EVENT_CHANGEOVER_PREP
        op, sub = _changeover_timeline_op_sub_for_event(
            event_kind=ek,
            op_from_segment=op_seg,
            machine_occ_key=m_occ,
            machining_lead_op=_lead_m,
            machining_sub_str=_sub_roll,
            machine_handoff=_mh,
            daily_status=daily_status,
        )
        br_acc: list = []
        for nm in (op, *[_p.strip() for _p in sub.split(",") if _p.strip()]):
            if nm and nm in daily_status:
                br_acc.extend(daily_status[nm].get("breaks_dt") or [])
        br_seg = merge_time_intervals(br_acc)
        tid_ev = (
            ""
            if ek == TIMELINE_EVENT_MACHINE_DAILY_STARTUP
            else str(task_id or "").strip()
        )
        ev = {
            "date": current_date,
            "task_id": tid_ev,
            "machine": m_line,
            "machine_occupancy_key": m_occ,
            "op": op,
            "sub": sub,
            "start_dt": st,
            "end_dt": ed,
            "breaks": br_seg,
            "units_done": 0,
            "event_kind": ek,
        }
        timeline_events.append(ev)
        if dispatch_interval_mirror is not None:
            dispatch_interval_mirror.register_from_event(ev)
        for nm in (op, *[_p.strip() for _p in sub.split(",") if _p.strip()]):
            if not nm:
                continue
            prev_a = avail_dt.get(nm, st)
            if isinstance(prev_a, datetime):
                avail_dt[nm] = max(prev_a, ed)
            else:
                avail_dt[nm] = ed


def _collect_task_ids_missed_deadline_after_day(task_queue: list, current_date: date) -> set:
    """
    ?????????????????????????????????NO?
    ???????????????= ???????????
    """
    out = set()
    eps = 1e-9
    for t in task_queue:
        if float(t.get("remaining_units") or 0) <= eps:
            continue
        db = t.get("due_basis_date")
        if db is None:
            continue
        sdr = t.get("start_date_req")
        if isinstance(sdr, date) and sdr > current_date:
            continue
        if current_date >= db:
            tid = str(t.get("task_id", "") or "").strip()
            if tid:
                out.add(tid)
    return out


def _normalize_timeline_task_id(ev: dict) -> str:
    return str(ev.get("task_id", "") or "").strip()


def _machine_handoff_state_from_timeline(
    timeline_events: list,
    current_date: date,
) -> dict:
    """
    ?????????? machine_occupancy_key ????
    ??? current_date ??? **?? (machining)** ???????????????
    ??????? event_kind ? last_tid / ???????????
    """
    best: dict[str, tuple[datetime, str, str, date, str, str]] = {}
    for e in timeline_events:
        if not _is_machining_timeline_event(e):
            continue
        ed = e.get("date")
        if not isinstance(ed, date):
            continue
        if ed > current_date:
            continue
        occ = str(e.get("machine_occupancy_key") or "").strip()
        if not occ:
            mraw = str(e.get("machine") or "").strip()
            occ = (
                _normalize_equipment_match_key(mraw.split("+", 1)[1])
                if "+" in mraw
                else _normalize_equipment_match_key(mraw)
            )
        if not occ:
            continue
        end_dt = e.get("end_dt")
        if end_dt is None or not hasattr(end_dt, "replace"):
            continue
        eq_line = str(e.get("machine") or "").strip()
        tid = _normalize_timeline_task_id(e)
        lead_op = str(e.get("op") or "").strip()
        sub_csv = str(e.get("sub") or "").strip()
        prev = best.get(occ)
        if prev is None or end_dt > prev[0]:
            best[occ] = (end_dt, tid, eq_line, ed, lead_op, sub_csv)
    last_tid = {k: v[1] for k, v in best.items()}
    last_eq = {k: v[2] for k, v in best.items()}
    last_machining_dt = {k: v[0] for k, v in best.items()}
    last_machining_date = {k: v[3] for k, v in best.items()}
    last_lead_op = {k: v[4] for k, v in best.items()}
    last_machining_sub = {k: v[5] for k, v in best.items()}
    machining_today_occ: set[str] = set()
    for e in timeline_events:
        if not _is_machining_timeline_event(e):
            continue
        if e.get("date") != current_date:
            continue
        occ = str(e.get("machine_occupancy_key") or "").strip()
        if not occ:
            mraw = str(e.get("machine") or "").strip()
            occ = (
                _normalize_equipment_match_key(mraw.split("+", 1)[1])
                if "+" in mraw
                else _normalize_equipment_match_key(mraw)
            )
        if occ:
            machining_today_occ.add(occ)
    started_today = set(machining_today_occ)
    return {
        "last_tid": last_tid,
        "last_eq": last_eq,
        "last_machining_dt": last_machining_dt,
        "last_machining_date": last_machining_date,
        "last_lead_op": last_lead_op,
        "last_machining_sub": last_machining_sub,
        "machining_today_occ": machining_today_occ,
        "started_today": started_today,
    }


def _trial_order_flow_day_start_floor(
    task: dict,
    current_date: date,
    macro_run_date: date,
    macro_now_dt: datetime,
    task_queue: list | None = None,
) -> datetime:
    """??????????????????????????? 13:00 ???????"""
    floor = datetime.combine(current_date, DEFAULT_START_TIME)
    # �B-2 ?? / �B-3 ???? EC ??????????????
    # ??????=??13:00??????????????????????????????
    # EC???????_roll_pipeline_b2_inspection_ec_completion_floor_dt????????
    _tid_floor = str(task.get("task_id", "") or "").strip()
    is_b2_follower_delayed = bool(
        (task.get("roll_pipeline_inspection") or task.get("roll_pipeline_rewind"))
        and _tid_floor
        and task_queue is not None
        and _task_queue_has_roll_pipeline_ec_for_tid(task_queue, _tid_floor)
    )
    rid = task.get("raw_input_date")
    if not is_b2_follower_delayed and isinstance(rid, date) and rid == current_date:
        floor = max(floor, datetime.combine(current_date, time(13, 0)))
    sdl = task.get("same_day_raw_start_limit")
    s_req = task.get("start_date_req")
    if (
        (not is_b2_follower_delayed)
        and sdl
        and isinstance(s_req, date)
        and current_date == s_req
        and isinstance(sdl, time)
    ):
        floor = max(floor, datetime.combine(current_date, sdl))
    est = task.get("earliest_start_time")
    if (
        (not is_b2_follower_delayed)
        and isinstance(s_req, date)
        and current_date == s_req
        and est
    ):
        if isinstance(est, time):
            floor = max(floor, datetime.combine(current_date, est))
    if current_date == macro_run_date and floor < macro_now_dt:
        floor = macro_now_dt
    return floor


def _trial_order_flow_eligible_tasks(
    tasks_today: list,
    task_queue: list,
    current_date: date,
    *,
    daily_status: dict | None = None,
    members: list | None = None,
    machine_avail_dt: dict | None = None,
    machine_day_start: datetime | None = None,
    machine_handoff: dict | None = None,
    skills_dict: dict | None = None,
    abolish_all_scheduling_limits: bool = False,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
    min_dispatch_effective: int | None = None,
    assign_probe_ctx: dict | None = None,
) -> list:
    out = []
    for task in tasks_today:
        if float(task.get("remaining_units") or 0) <= 1e-12:
            continue
        if _task_blocked_by_same_request_dependency(task, task_queue):
            continue
        if _task_blocked_by_global_dispatch_trial_order(
            task,
            task_queue,
            current_date,
            daily_status=daily_status,
            members=members,
            machine_avail_dt=machine_avail_dt,
            machine_day_start=machine_day_start,
            machine_handoff=machine_handoff,
            skills_dict=skills_dict,
            abolish_all_scheduling_limits=abolish_all_scheduling_limits,
            dispatch_interval_mirror=dispatch_interval_mirror,
            min_dispatch_effective=min_dispatch_effective,
        ):
            continue
        # min_dto ???????????????????????????????? my_o>m ?????
        # ???=min ???????????????????????????????????????
        if daily_status is not None and members is not None:
            if _task_fully_machine_calendar_blocked_on_date(
                task, current_date, daily_status, members
            ):
                continue
            if _task_no_machining_window_left_from_avail_floor(
                task,
                current_date,
                daily_status,
                members,
                machine_avail_dt,
                machine_day_start,
                machine_handoff=machine_handoff,
                skills_dict=skills_dict,
                abolish_all_scheduling_limits=abolish_all_scheduling_limits,
                dispatch_interval_mirror=dispatch_interval_mirror,
            ):
                continue
        if (
            task.get("roll_pipeline_inspection") or task.get("roll_pipeline_rewind")
        ) and (
            _roll_pipeline_inspection_assign_room(
                task_queue, str(task.get("task_id", "") or "").strip()
            )
            <= 1e-12
        ):
            continue
        machine = task["machine"]
        eq_line = str(
            task.get("equipment_line_key") or machine or ""
        ).strip() or machine
        _mocc_trial = _machine_occupancy_key_resolve(task, eq_line)
        if PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE:
            _b1_holder = _exclusive_b1_inspection_holder_for_machine(
                task_queue,
                _mocc_trial,
            )
            if _b1_holder is not None and _b1_holder is not task:
                continue
        try:
            _my_dispatch_ord = int(task.get("dispatch_trial_order") or 10**9)
        except (TypeError, ValueError):
            _my_dispatch_ord = 10**9
        if _equipment_line_lower_dispatch_trial_still_pending(
            task_queue,
            _mocc_trial,
            _my_dispatch_ord,
            current_date,
            daily_status=daily_status,
            members=members,
            machine_avail_dt=machine_avail_dt,
            machine_day_start=machine_day_start,
            machine_handoff=machine_handoff,
            skills_dict=skills_dict,
            abolish_all_scheduling_limits=abolish_all_scheduling_limits,
            dispatch_interval_mirror=dispatch_interval_mirror,
            assign_probe_ctx=assign_probe_ctx,
        ):
            continue
        out.append(task)
    return out


def _combo_preset_team_size_bounds(
    preset_team: tuple,
    sheet_req_n: int | None,
    max_team_size_need: int,
) -> tuple[int, int] | None:
    """
    ???????????1?????? (lo, hi)?need ?????????????????
    - ???????????????????????????
    - hi ? need ??????????????????? need ?????????????
    """
    nmem = len(preset_team)
    if nmem < 1:
        return None
    if sheet_req_n is not None and sheet_req_n >= 1:
        if nmem != sheet_req_n:
            return None
        lo = sheet_req_n
    else:
        lo = nmem
    hi = max(max_team_size_need, nmem)
    if not (lo <= nmem <= hi):
        return None
    return lo, hi


def _plan_sheet_required_op_optional(task: dict) -> int | None:
    """????????????????????????? None?"""
    ro = task.get("required_op")
    if ro is None or (isinstance(ro, float) and pd.isna(ro)):
        return None
    try:
        n = int(ro)
    except (TypeError, ValueError):
        return None
    return n if n >= 1 else None


def _append_legacy_dispatch_candidate_for_team(
    task: dict,
    team: tuple,
    avail_dt: dict,
    machine_avail_dt: dict,
    daily_status: dict,
    current_date: date,
    macro_run_date: date,
    macro_now_dt: datetime,
    skill_role_priority,
    eq_line: str,
    rq_base: int,
    extra_max: int,
    global_priority_override: dict,
    team_candidates: list,
    *,
    combo_sheet_row_id: int | None = None,
    combo_preset_team: tuple[str, ...] | None = None,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
    machine_handoff: dict | None = None,
    machine_day_floor: datetime | None = None,
    machine_floor_cached: datetime | None = None,
) -> bool:
    """????????????: ??????????? team_candidates ? 1 ????? True?"""
    _machine_occ_key = _machine_occupancy_key_resolve(task, eq_line)
    _gpo = global_priority_override or {}
    _floor_default = datetime.combine(current_date, DEFAULT_START_TIME)
    _mdf = machine_day_floor if machine_day_floor is not None else _floor_default
    _mh_legacy = machine_handoff or {
        "last_tid": {},
        "last_eq": {},
        "started_today": set(),
        "machining_today_occ": set(),
        "last_machining_dt": {},
        "last_machining_date": {},
        "last_lead_op": {},
        "last_machining_sub": {},
    }
    op_list = [m for m in team if skill_role_priority(m)[0] == "OP"]
    if not op_list:
        return False
    team_start = max(avail_dt[m] for m in team)
    if not _gpo.get("abolish_all_scheduling_limits"):
        if machine_floor_cached is not None:
            machine_free_dt = machine_floor_cached
        else:
            machine_free_dt = _machine_effective_floor_for_assign(
                _machine_occ_key,
                str(task.get("task_id") or "").strip(),
                eq_line,
                str(task.get("machine_name") or "").strip(),
                machine_avail_dt,
                _mh_legacy,
                _mdf,
                False,
                current_date=current_date,
            )
        if team_start < machine_free_dt:
            team_start = machine_free_dt
        if task.get("same_day_raw_start_limit") and current_date == task["start_date_req"]:
            min_start_dt = datetime.combine(
                current_date, task["same_day_raw_start_limit"]
            )
            if team_start < min_start_dt:
                team_start = min_start_dt
        if current_date == task["start_date_req"] and task.get("earliest_start_time"):
            min_user_t = datetime.combine(
                current_date, task["earliest_start_time"]
            )
            if team_start < min_user_t:
                team_start = min_user_t
        if current_date == macro_run_date and team_start < macro_now_dt:
            team_start = macro_now_dt
    team_end_limit = min(daily_status[m]["end_dt"] for m in team)
    if team_start >= team_end_limit:
        return False
    team_breaks = []
    for m in team:
        team_breaks.extend(daily_status[m]["breaks_dt"])
    team_breaks = merge_time_intervals(team_breaks)

    avg_eff = sum(daily_status[m]["efficiency"] for m in team) / len(team)
    if avg_eff <= 0:
        avg_eff = 0.01
    t_eff = parse_float_safe(task.get("task_eff_factor"), 1.0)
    if t_eff <= 0:
        t_eff = 1.0
    eff_time_per_unit = (
        task["base_time_per_unit"]
        / avg_eff
        / t_eff
        * _surplus_team_time_factor(rq_base, len(team), extra_max)
    )
    _defer_min_contig = max(1, int(math.ceil(float(eff_time_per_unit))))

    def _refloor_legacy_roll(ts: datetime) -> datetime:
        ts = max(ts, max(avail_dt[m] for m in team))
        if not _gpo.get("abolish_all_scheduling_limits"):
            if machine_floor_cached is not None:
                mf = machine_floor_cached
            else:
                mf = _machine_effective_floor_for_assign(
                    _machine_occ_key,
                    str(task.get("task_id") or "").strip(),
                    eq_line,
                    str(task.get("machine_name") or "").strip(),
                    machine_avail_dt,
                    _mh_legacy,
                    _mdf,
                    False,
                    current_date=current_date,
                )
            if ts < mf:
                ts = mf
            if task.get("same_day_raw_start_limit") and current_date == task["start_date_req"]:
                min_start_dt = datetime.combine(
                    current_date, task["same_day_raw_start_limit"]
                )
                if ts < min_start_dt:
                    ts = min_start_dt
            if current_date == task["start_date_req"] and task.get("earliest_start_time"):
                min_user_t = datetime.combine(
                    current_date, task["earliest_start_time"]
                )
                if ts < min_user_t:
                    ts = min_user_t
            if current_date == macro_run_date and ts < macro_now_dt:
                ts = macro_now_dt
        return ts

    team_start_adj = _defer_team_start_past_prebreak_and_end_of_day(
        task,
        team,
        team_start,
        team_end_limit,
        team_breaks,
        _refloor_legacy_roll,
        min_contiguous_work_mins=_defer_min_contig,
    )
    if team_start_adj is None:
        return False
    team_start = team_start_adj
    if team_start >= team_end_limit:
        return False

    _, avail_mins, _ = calculate_end_time(team_start, 9999, team_breaks, team_end_limit)
    units_can_do = int(avail_mins / eff_time_per_unit)
    if units_can_do == 0:
        return False
    units_today = min(units_can_do, math.ceil(task["remaining_units"]))
    if _eod_reject_capacity_units_below_threshold(
        units_today, team_start, team_end_limit
    ):
        return False
    work_mins_needed = int(units_today * eff_time_per_unit)
    if (
        _contiguous_work_minutes_until_next_break_or_limit(
            team_start, team_breaks, team_end_limit
        )
        < work_mins_needed
    ):
        return False
    actual_end_dt, _, _ = calculate_end_time(
        team_start, work_mins_needed, team_breaks, team_end_limit
    )
    if dispatch_interval_mirror is not None and dispatch_interval_mirror.would_block_roll(
        _machine_occ_key, team, team_start, actual_end_dt
    ):
        return False
    team_prio_sum = sum(skill_role_priority(m)[1] for m in team)
    team_candidates.append(
        {
            "team": team,
            "team_start": team_start,
            "actual_end_dt": actual_end_dt,
            "units_today": units_today,
            "team_breaks": team_breaks,
            "avg_eff": avg_eff,
            "prio_sum": team_prio_sum,
            "op_list": op_list,
            "eff_time_per_unit": eff_time_per_unit,
            "combo_sheet_row_id": combo_sheet_row_id,
            "combo_preset_team": combo_preset_team,
        }
    )
    return True


def _assign_one_roll_trial_order_flow(
    task: dict,
    current_date: date,
    daily_status: dict,
    avail_dt: dict,
    machine_avail_dt: dict,
    task_queue: list,
    skills_dict: dict,
    members: list,
    req_map: dict,
    need_rules: list,
    surplus_map: dict,
    global_priority_override: dict,
    macro_run_date: date,
    macro_now_dt: datetime,
    preferred_team: tuple | None,
    _need_headcount_logged_orders: set,
    team_combo_presets: dict | None = None,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
    machine_handoff: dict | None = None,
) -> dict | None:
    """
    1??????????????????????????? team_start ??????
    preferred_team ?????????????????????????????
    ????????????????????????????
    ???: team(tuple), start_dt, end_dt, breaks, eff, op, eff_time_per_unit, extra_max, rq_base, need_src_line, extra_src_line, machine, machine_name, eq_line, req_num, max_team_size
    """
    machine = task["machine"]
    machine_name = str(task.get("machine_name", "") or "").strip()
    machine_proc = str(machine or "").strip()
    eq_line = str(task.get("equipment_line_key") or machine or "").strip() or machine
    machine_occ_key = _machine_occupancy_key_resolve(task, eq_line)
    _gpo = global_priority_override or {}
    _mh = machine_handoff or {
        "last_tid": {},
        "last_eq": {},
        "started_today": set(),
        "machining_today_occ": set(),
        "last_machining_dt": {},
        "last_machining_date": {},
        "last_lead_op": {},
        "last_machining_sub": {},
    }

    plan_ro = _plan_sheet_required_op_optional(task)
    need_src_line = ""
    if TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY:
        req_num, need_src_line = resolve_need_required_op_explain(
            machine,
            machine_name,
            task["task_id"],
            req_map,
            need_rules,
        )
        if plan_ro is not None and plan_ro != req_num:
            need_src_line = (need_src_line + "?") if need_src_line else ""
            need_src_line += f"?????????{plan_ro}?????need??={req_num}?"
    else:
        if plan_ro is not None:
            req_num = plan_ro
            need_src_line = f"????????OP(??)?={req_num}"
        else:
            req_num, need_src_line = resolve_need_required_op_explain(
                machine,
                machine_name,
                task["task_id"],
                req_map,
                need_rules,
            )
    if _gpo.get("ignore_need_minimum"):
        req_num = 1
        need_src_line = (
            (need_src_line + " ? ") if need_src_line else ""
        ) + "?????ignore_need_minimum?req=1"

    skill_meta_cache: dict = {}

    def skill_role_priority(mem):
        if _gpo.get("ignore_skill_requirements"):
            return ("OP", 100)
        if mem not in skill_meta_cache:
            srow = skills_dict.get(mem, {})
            v = ""
            if machine_proc and machine_name:
                v = srow.get(f"{machine_proc}+{machine_name}", "")
            elif machine_name:
                v = srow.get(machine_name, "")
            elif machine_proc:
                v = srow.get(machine_proc, "")
            skill_meta_cache[mem] = parse_op_as_skill_cell(v)
        return skill_meta_cache[mem]

    capable_members = [m for m in avail_dt if skill_role_priority(m)[0] in ("OP", "AS")]
    capable_members.sort(key=lambda mm: (skill_role_priority(mm)[1], mm))
    capable_members = _filter_capable_members_b2_disjoint_teams(
        task, task_queue, capable_members
    )

    pref_raw = str(task.get("preferred_operator_raw") or "").strip()
    op_today = [m for m in capable_members if skill_role_priority(m)[0] == "OP"]
    pref_mem = (
        _resolve_preferred_op_to_member(pref_raw, op_today, members)
        if pref_raw
        else None
    )

    _gdp_must, _gdp_warns = _active_global_day_process_must_include(
        _gpo, task, current_date, capable_members, members
    )
    for _gw in _gdp_warns:
        logging.warning(_gw)
    fixed_team_anchor = _merge_global_day_process_and_pref_anchor(
        _gdp_must, pref_mem, capable_members
    )
    if _gdp_must:
        logging.info(
            "????????(??�??): task=%s date=%s ??=%r ?????=%s",
            task.get("task_id"),
            current_date,
            machine,
            ",".join(_gdp_must),
        )
    if fixed_team_anchor:
        _nfix = len(fixed_team_anchor)
        if _nfix > req_num:
            need_src_line = (need_src_line + " ? ") if need_src_line else ""
            need_src_line += f"?????(??�??)?????{_nfix}?"
        req_num = max(req_num, _nfix)

    extra_max_sheet, extra_src_line = resolve_need_surplus_extra_max_explain(
        machine,
        machine_name,
        task["task_id"],
        surplus_map,
        need_rules,
    )
    if TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW:
        extra_max_sheet = 0
        extra_src_line = (
            (extra_src_line + " ? ") if extra_src_line else ""
        ) + "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW?0"
    extra_max = (
        extra_max_sheet if TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS else 0
    )
    if (
        extra_max_sheet > 0
        and not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
        and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
    ):
        extra_src_line = (
            (extra_src_line + " ? ") if extra_src_line else ""
        ) + "???????????????????????�???????"
    max_team_size = min(req_num + extra_max, len(capable_members))
    if max_team_size < req_num:
        max_team_size = req_num
    rq_base = max(1, int(req_num))
    combo_key_assign = (
        f"{machine_proc}+{machine_name}"
        if machine_proc and machine_name
        else ""
    )
    preset_rows_assign = (
        (team_combo_presets or {}).get(combo_key_assign)
        if (team_combo_presets and combo_key_assign)
        else None
    )

    _dto_head = task.get("dispatch_trial_order")
    if _dto_head is not None and _dto_head not in _need_headcount_logged_orders:
        _need_headcount_logged_orders.add(_dto_head)
        logging.info(
            "need??(????????) order=%s task=%s ??/??=%s/%s "
            "req_num=%s [%s] extra_max=%s [%s] max_team??=%s capable=%s?",
            _dto_head,
            task["task_id"],
            machine,
            machine_name,
            req_num,
            need_src_line,
            extra_max,
            extra_src_line,
            max_team_size,
            len(capable_members),
        )

    day_floor = _trial_order_flow_day_start_floor(
        task, current_date, macro_run_date, macro_now_dt, task_queue
    )
    machine_day_floor = datetime.combine(current_date, DEFAULT_START_TIME)
    b2_insp_ec_floor: datetime | None = None
    _tid_assign = str(task.get("task_id") or "").strip()
    _trace_assign_enabled = _trace_schedule_task_enabled(_tid_assign)
    def _trace_assign(msg: str, *args) -> None:
        if not _trace_assign_enabled:
            return
        _log_dispatch_trace_schedule(
            _tid_assign,
            "[?????? task=%s] " + msg,
            _tid_assign,
            *args,
        )
    if (
        (
            task.get("roll_pipeline_inspection")
            or task.get("roll_pipeline_rewind")
        )
        and _task_queue_has_roll_pipeline_ec_for_tid(task_queue, _tid_assign)
    ):
        b2_insp_ec_floor = _roll_pipeline_b2_inspection_ec_completion_floor_dt(
            task_queue, _tid_assign
        )

    _prev_mach_before_co = machine_avail_dt.get(
        machine_occ_key, machine_day_floor
    )
    _mach_floor_eff, _co_segs, _co_abort = _resolve_machine_changeover_floor_segments(
        abolish_all_scheduling_limits=bool(_gpo.get("abolish_all_scheduling_limits")),
        machine_occ_key=machine_occ_key,
        task_id=str(task.get("task_id") or "").strip(),
        eq_line=eq_line,
        machine_name=machine_name,
        machine_proc=machine_proc,
        machine_avail_dt=machine_avail_dt,
        machine_day_floor=machine_day_floor,
        current_date=current_date,
        machine_handoff=_mh,
        daily_status=daily_status,
        skills_dict=skills_dict,
        dispatch_interval_mirror=dispatch_interval_mirror,
    )
    if _co_abort:
        return None

    def _one_roll_from_team(
        team: tuple,
        min_n: int | None = None,
        max_n: int | None = None,
    ) -> dict | None:
        lo = req_num if min_n is None else min_n
        hi = max_team_size if max_n is None else max_n
        if len(team) < lo or len(team) > hi:
            _trace_assign(
                "????: ?????? team=%s size=%s req=%s max=%s",
                ",".join(str(x) for x in team),
                len(team),
                lo,
                hi,
            )
            return None
        op_list = [m for m in team if skill_role_priority(m)[0] == "OP"]
        if not op_list:
            _trace_assign(
                "????: OP?? team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        if not all(m in daily_status for m in team):
            _trace_assign(
                "????: ???????? team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        team_start = max(avail_dt[m] for m in team)
        if not _gpo.get("abolish_all_scheduling_limits"):
            machine_free_dt = _mach_floor_eff
            if team_start < machine_free_dt:
                team_start = machine_free_dt
            if team_start < day_floor:
                team_start = day_floor
        if b2_insp_ec_floor is not None and team_start < b2_insp_ec_floor:
            team_start = b2_insp_ec_floor
        team_end_limit = min(daily_status[m]["end_dt"] for m in team)
        if team_start >= team_end_limit:
            _trace_assign(
                "????: ??>=?? team=%s start=%s end_limit=%s",
                ",".join(str(x) for x in team),
                team_start,
                team_end_limit,
            )
            return None
        team_breaks = []
        for m in team:
            team_breaks.extend(daily_status[m]["breaks_dt"])
        team_breaks = merge_time_intervals(team_breaks)

        avg_eff = sum(daily_status[m]["efficiency"] for m in team) / len(team)
        if avg_eff <= 0:
            avg_eff = 0.01
        t_eff = parse_float_safe(task.get("task_eff_factor"), 1.0)
        if t_eff <= 0:
            t_eff = 1.0
        eff_time_per_unit = (
            task["base_time_per_unit"]
            / avg_eff
            / t_eff
            * _surplus_team_time_factor(rq_base, len(team), extra_max)
        )
        _defer_min_contig = max(1, int(math.ceil(float(eff_time_per_unit))))

        def _refloor_trial_roll(ts: datetime) -> datetime:
            ts = max(ts, max(avail_dt[m] for m in team))
            if not _gpo.get("abolish_all_scheduling_limits"):
                mf = _mach_floor_eff
                if ts < mf:
                    ts = mf
                if ts < day_floor:
                    ts = day_floor
            if b2_insp_ec_floor is not None and ts < b2_insp_ec_floor:
                ts = b2_insp_ec_floor
            return ts

        team_start_d = _defer_team_start_past_prebreak_and_end_of_day(
            task,
            team,
            team_start,
            team_end_limit,
            team_breaks,
            _refloor_trial_roll,
            min_contiguous_work_mins=_defer_min_contig,
        )
        if team_start_d is None:
            _trace_assign(
                "????: ?????????(??)????? team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        team_start = team_start_d
        if team_start >= team_end_limit:
            _trace_assign(
                "????: ????????>=?? team=%s start=%s end_limit=%s",
                ",".join(str(x) for x in team),
                team_start,
                team_end_limit,
            )
            return None

        _, avail_mins, _ = calculate_end_time(
            team_start, 9999, team_breaks, team_end_limit
        )
        _trial_units_cap = int(avail_mins / eff_time_per_unit)
        if _trial_units_cap < 1:
            _trace_assign(
                "????: ???? team=%s start=%s avail_mins=%s need_mins=%.2f",
                ",".join(str(x) for x in team),
                team_start,
                avail_mins,
                eff_time_per_unit,
            )
            return None
        if _eod_reject_capacity_units_below_threshold(
            _trial_units_cap, team_start, team_end_limit
        ):
            _trace_assign(
                "????: ?????????????????? team=%s cap=%s th=%s start=%s",
                ",".join(str(x) for x in team),
                _trial_units_cap,
                ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS,
                team_start,
            )
            return None
        work_mins_needed = int(eff_time_per_unit)
        _contig = _contiguous_work_minutes_until_next_break_or_limit(
            team_start, team_breaks, team_end_limit
        )
        if _contig < work_mins_needed:
            _trace_assign(
                "????: ?????????????? team=%s contiguous_min=%s need_mins=%s start=%s",
                ",".join(str(x) for x in team),
                _contig,
                work_mins_needed,
                team_start,
            )
            return None
        actual_end_dt, _, _ = calculate_end_time(
            team_start, work_mins_needed, team_breaks, team_end_limit
        )
        if dispatch_interval_mirror is not None and dispatch_interval_mirror.would_block_roll(
            machine_occ_key, team, team_start, actual_end_dt
        ):
            _trace_assign(
                "???????: team=%s start=%s end=%s eq=%s",
                ",".join(str(x) for x in team),
                team_start,
                actual_end_dt,
                eq_line,
            )
            return None
        if pref_mem and pref_mem in op_list:
            lead_op = pref_mem
        else:
            lead_op = min(op_list, key=lambda mm: (skill_role_priority(mm)[1], mm))
        team_prio_sum = sum(skill_role_priority(m)[1] for m in team)
        return {
            "team": team,
            "team_start": team_start,
            "actual_end_dt": actual_end_dt,
            "team_breaks": team_breaks,
            "avg_eff": avg_eff,
            "prio_sum": team_prio_sum,
            "op_list": op_list,
            "eff_time_per_unit": eff_time_per_unit,
            "lead_op": lead_op,
            "changeover_segments": _co_segs,
        }

    # ????: ???????????????????????????????
    _hist = task.get("assigned_history") or []
    _last_hist_date = _hist[-1].get("date") if _hist else None
    _same_day_last_roll = _last_hist_date == current_date.strftime("%m/%d")
    if preferred_team and _same_day_last_roll:
        pt = tuple(preferred_team)
        _pref_pt_ok = (not fixed_team_anchor) or all(
            m in pt for m in fixed_team_anchor
        )
        if _pref_pt_ok and all(m in capable_members and m in avail_dt for m in pt):
            got = _one_roll_from_team(pt)
            if got is not None:
                _cid_pt = _lookup_combo_sheet_row_id_for_preset_team(
                    preset_rows_assign, pt
                )
                return {
                    **got,
                    "extra_max": extra_max,
                    "rq_base": rq_base,
                    "need_src_line": need_src_line,
                    "extra_src_line": extra_src_line,
                    "machine": machine,
                    "machine_name": machine_name,
                    "eq_line": eq_line,
                    "req_num": req_num,
                    "max_team_size": max_team_size,
                    "combo_sheet_row_id": _cid_pt,
                    "combo_preset_team": pt if _cid_pt is not None else None,
                    "changeover_segments": _co_segs,
                }

    team_candidates: list[dict] = []
    # ??????????????????? return??????????????
    # team_start / ?????????????????????????????????
    if preset_rows_assign:
        for _prio, sheet_rs, preset_team, combo_row_id in preset_rows_assign:
            bounds = _combo_preset_team_size_bounds(
                tuple(preset_team), sheet_rs, max_team_size
            )
            if bounds is None:
                continue
            lo_pt, hi_pt = bounds
            if fixed_team_anchor and not all(m in preset_team for m in fixed_team_anchor):
                continue
            if pref_mem is not None and pref_mem not in preset_team:
                continue
            if not all(m in capable_members for m in preset_team):
                continue
            if sum(1 for m in preset_team if skill_role_priority(m)[0] == "OP") < 1:
                continue
            got = _one_roll_from_team(
                tuple(preset_team), min_n=lo_pt, max_n=hi_pt
            )
            if got is not None:
                team_candidates.append(
                    {
                        **got,
                        "combo_sheet_row_id": combo_row_id,
                        "combo_preset_team": tuple(preset_team),
                    }
                )
    for tsize in range(req_num, max_team_size + 1):
        if fixed_team_anchor:
            _ft = list(fixed_team_anchor)
            others = [m for m in capable_members if m not in _ft]
            need_extra = tsize - len(_ft)
            if need_extra < 0:
                teams_iter = []
            elif need_extra == 0:
                teams_iter = [tuple(_ft)]
            elif len(others) >= need_extra:
                teams_iter = [
                    tuple(_ft + list(rest))
                    for rest in itertools.combinations(others, need_extra)
                ]
            else:
                teams_iter = []
        elif (
            pref_mem is not None
            and pref_mem in capable_members
            and skill_role_priority(pref_mem)[0] == "OP"
        ):
            if tsize == 1:
                _trace_assign(
                    "????: ??OP??=%s ??? 1???????????",
                    pref_mem,
                )
            others = [m for m in capable_members if m != pref_mem]
            if tsize == 1:
                teams_iter = [(pref_mem,)]
            elif len(others) >= tsize - 1:
                teams_iter = [
                    tuple([pref_mem] + list(rest))
                    for rest in itertools.combinations(others, tsize - 1)
                ]
            else:
                teams_iter = itertools.combinations(capable_members, tsize)
        else:
            teams_iter = itertools.combinations(capable_members, tsize)

        for team in teams_iter:
            got = _one_roll_from_team(team)
            if got is not None:
                team_candidates.append(
                    {
                        **got,
                        "combo_sheet_row_id": None,
                        "combo_preset_team": None,
                    }
                )

    if not team_candidates:
        _mem_max_end: datetime | None = None
        for _m in capable_members:
            if _m not in daily_status:
                continue
            _ed = daily_status[_m].get("end_dt")
            if isinstance(_ed, datetime):
                _mem_max_end = (
                    _ed if _mem_max_end is None else max(_mem_max_end, _ed)
                )
        if (
            len(capable_members) >= req_num
            and _mem_max_end is not None
            and isinstance(_mach_floor_eff, datetime)
            and _mach_floor_eff >= _mem_max_end
        ):
            logging.warning(
                "??2: ??NO=%s ??=%s ??/??=%s/%s ???????0??"
                "?????(OP/AS)? %s ???????????????=%s ?"
                "??????????(%s)?????????????????"
                "master??????????????????????????????"
                "?????????????????????????????????????"
                "?????? 3.2.1 ??????????????????"
                "??: changeover????????=%s ????=%s",
                task.get("task_id"),
                current_date,
                machine,
                machine_name,
                len(capable_members),
                _mach_floor_eff.strftime("%Y-%m-%d %H:%M"),
                _mem_max_end.strftime("%H:%M"),
                _prev_mach_before_co.strftime("%Y-%m-%d %H:%M"),
                machine_occ_key,
            )
        return None
    t_min = min(c["team_start"] for c in team_candidates)

    def _team_cand_key(c):
        return _team_assignment_sort_tuple(
            c["team"],
            c["team_start"],
            1,
            c["prio_sum"],
            t_min,
        )

    best_c = min(team_candidates, key=_team_cand_key)
    if best_c.get("combo_sheet_row_id") is None and preset_rows_assign:
        _lcid = _lookup_combo_sheet_row_id_for_preset_team(
            preset_rows_assign, tuple(best_c["team"])
        )
        if _lcid is not None:
            best_c = {
                **best_c,
                "combo_sheet_row_id": _lcid,
                "combo_preset_team": tuple(best_c["team"]),
            }
    return {
        **best_c,
        "extra_max": extra_max,
        "rq_base": rq_base,
        "need_src_line": need_src_line,
        "extra_src_line": extra_src_line,
        "machine": machine,
        "machine_name": machine_name,
        "eq_line": eq_line,
        "req_num": req_num,
        "max_team_size": max_team_size,
        "combo_sheet_row_id": best_c.get("combo_sheet_row_id"),
        "combo_preset_team": best_c.get("combo_preset_team"),
        "changeover_segments": _co_segs,
    }


def _trial_order_assign_probe_fails(
    task: dict,
    current_date: date,
    daily_status: dict,
    ctx: dict,
) -> bool:
    """
    ??? avail_dt / machine_avail_dt / machine_handoff ??????????
    `_assign_one_roll_trial_order_flow` ? None ????? True?
    ??????????????????????????????????????????????
    ??????need ????? set ??????
    """
    try:
        r = _assign_one_roll_trial_order_flow(
            task,
            current_date,
            daily_status,
            ctx["avail_dt"],
            ctx["machine_avail_dt"],
            ctx["task_queue"],
            ctx["skills_dict"],
            ctx["members"],
            ctx["req_map"],
            ctx["need_rules"],
            ctx["surplus_map"],
            ctx["global_priority_override"],
            ctx["macro_run_date"],
            ctx["macro_now_dt"],
            None,
            set(),
            team_combo_presets=ctx.get("team_combo_presets"),
            dispatch_interval_mirror=ctx.get("dispatch_interval_mirror"),
            machine_handoff=ctx["machine_handoff"],
        )
    except Exception as ex:
        logging.warning(
            "trial_order_assign_probe ??????????????: task=%s err=%s",
            task.get("task_id"),
            ex,
        )
        return False
    return r is None


def _tasks_in_min_pending_dispatch_pool(
    task_queue: list,
    current_date: date,
    *,
    daily_status: dict | None = None,
    members: list | None = None,
    machine_avail_dt: dict | None = None,
    machine_day_start: datetime | None = None,
    machine_handoff: dict | None = None,
    skills_dict: dict | None = None,
    abolish_all_scheduling_limits: bool = False,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
) -> list:
    """`_min_pending_dispatch_trial_order_for_date` ???????????????????????"""
    out: list = []
    for t in task_queue:
        if float(t.get("remaining_units") or 0) <= 1e-12:
            continue
        sdr = t.get("start_date_req")
        if not isinstance(sdr, date) or sdr > current_date:
            continue
        if _task_not_yet_schedulable_due_to_dependency_or_b2_room(t, task_queue):
            continue
        if _task_fully_machine_calendar_blocked_on_date(
            t, current_date, daily_status, members
        ):
            continue
        if _task_no_machining_window_left_from_avail_floor(
            t,
            current_date,
            daily_status,
            members,
            machine_avail_dt,
            machine_day_start,
            machine_handoff=machine_handoff,
            skills_dict=skills_dict,
            abolish_all_scheduling_limits=abolish_all_scheduling_limits,
            dispatch_interval_mirror=dispatch_interval_mirror,
        ):
            continue
        out.append(t)
    return out


def _effective_min_dispatch_trial_order_from_pool(
    pool: list,
    current_date: date,
    daily_status: dict,
    assign_probe_ctx: dict,
) -> int | None:
    """
    pool ??? dto ????**?? dto ???????? 1 ???** 1 ?????????????
    ?? dto ???????????????
    ?? dto ???????????????????????????? dto ??????????????
    ?????????? pool ??? dto ????
    """
    if not pool:
        return None
    dtos = sorted(
        {
            int(t.get("dispatch_trial_order") or 10**9)
            for t in pool
        }
    )
    if not assign_probe_ctx:
        return min(dtos)
    for d in dtos:
        at_d = [
            t
            for t in pool
            if int(t.get("dispatch_trial_order") or 10**9) == d
        ]
        if any(
            not _trial_order_assign_probe_fails(
                t, current_date, daily_status, assign_probe_ctx
            )
            for t in at_d
        ):
            return d
    return None


def _trial_order_first_schedule_pass(
    current_date: date,
    tasks_today: list,
    task_queue: list,
    daily_status: dict,
    machine_avail_dt: dict,
    avail_dt: dict,
    timeline_events: list,
    skills_dict: dict,
    members: list,
    req_map: dict,
    need_rules: list,
    surplus_map: dict,
    global_priority_override: dict,
    macro_run_date: date,
    macro_now_dt: datetime,
    _need_headcount_logged_orders: set,
    team_combo_presets: dict | None = None,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
) -> bool:
    """
    ???????????????????1 ?????
    **?????�B-2 / �B-3?**: **????1**? **?????????**???????????**???**???EC????????????????
    **`_drain_rolls_for_task`** ??**????2**? �B-2 ???�B-3 ???????**????? EC ?????????**???????
    EC ?????? `_trial_order_flow_eligible_tasks` ?????????????? EC ???????
    ???????? EC ????`_run_b2_inspection_rewind_pass` ??????????????????
    EC ???? **??????** ??????? EC ??????????????????1???????
    ?????????????? **?????????**?????1? **????? EC** ?????????
    ??? `start_ge_end_initial`????????????????????�B-2/�B-3 ????????
    ?????EC???????????????1????? **?????**??????
    ???? **??? EC ????**?**????????1** ????? **?????**????
    **??1?????**???????????rest ??????????????OP?????
    ??????? start_ge_end_initial ?????????
    ??????????????????? `_roll_pipeline_inspection_assign_room` ???
    `_roll_pipeline_b2_inspection_ec_completion_floor_dt`?EC ????????????????
    ?????????????????????**????????????????**????????
    ???????????????????????
    """
    _mc_w0 = datetime.combine(current_date, DEFAULT_START_TIME)
    _mh_init = _machine_handoff_state_from_timeline(timeline_events, current_date)
    _gpo = global_priority_override or {}
    _assign_probe_ctx: dict | None = None
    if STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT:
        _assign_probe_ctx = {
            "avail_dt": avail_dt,
            "machine_avail_dt": machine_avail_dt,
            "task_queue": task_queue,
            "skills_dict": skills_dict,
            "members": members,
            "req_map": req_map,
            "need_rules": need_rules,
            "surplus_map": surplus_map,
            "global_priority_override": global_priority_override,
            "macro_run_date": macro_run_date,
            "macro_now_dt": macro_now_dt,
            "machine_handoff": _mh_init,
            "team_combo_presets": team_combo_presets,
            "dispatch_interval_mirror": dispatch_interval_mirror,
        }
    _min_dispatch_eff: int | None = None
    if STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT and _assign_probe_ctx:
        _pool_min = _tasks_in_min_pending_dispatch_pool(
            task_queue,
            current_date,
            daily_status=daily_status,
            members=members,
            machine_avail_dt=machine_avail_dt,
            machine_day_start=_mc_w0,
            machine_handoff=_mh_init,
            skills_dict=skills_dict,
            abolish_all_scheduling_limits=bool(
                _gpo.get("abolish_all_scheduling_limits")
            ),
            dispatch_interval_mirror=dispatch_interval_mirror,
        )
        _min_dispatch_eff = _effective_min_dispatch_trial_order_from_pool(
            _pool_min, current_date, daily_status, _assign_probe_ctx
        )
    eligible = _trial_order_flow_eligible_tasks(
        tasks_today,
        task_queue,
        current_date,
        daily_status=daily_status,
        members=members,
        machine_avail_dt=machine_avail_dt,
        machine_day_start=_mc_w0,
        machine_handoff=_mh_init,
        skills_dict=skills_dict,
        abolish_all_scheduling_limits=bool(_gpo.get("abolish_all_scheduling_limits")),
        dispatch_interval_mirror=dispatch_interval_mirror,
        min_dispatch_effective=_min_dispatch_eff,
        assign_probe_ctx=_assign_probe_ctx,
    )
    if not eligible:
        return False
    eligible_sorted = sorted(
        eligible,
        key=lambda t: int(t.get("dispatch_trial_order") or 10**9),
    )
    _mc_plan_end = _machine_calendar_planning_window_end_dt(
        current_date, daily_status, members
    )
    machine_handoff = {
        "last_tid": dict(_mh_init["last_tid"]),
        "last_eq": dict(_mh_init["last_eq"]),
        "started_today": set(_mh_init["started_today"]),
        "machining_today_occ": set(_mh_init.get("machining_today_occ") or set()),
        "last_machining_dt": dict(_mh_init.get("last_machining_dt") or {}),
        "last_machining_date": dict(_mh_init.get("last_machining_date") or {}),
        "last_lead_op": dict(_mh_init.get("last_lead_op") or {}),
        "last_machining_sub": dict(_mh_init.get("last_machining_sub") or {}),
    }

    def _drain_rolls_for_task(
        task: dict, *, max_rolls: int | None = None
    ) -> bool:
        preferred_team: tuple | None = None
        made_local = False
        rolls_done = 0
        while float(task.get("remaining_units") or 0) > 1e-12:
            if max_rolls is not None and rolls_done >= max_rolls:
                break
            res = _assign_one_roll_trial_order_flow(
                task,
                current_date,
                daily_status,
                avail_dt,
                machine_avail_dt,
                task_queue,
                skills_dict,
                members,
                req_map,
                need_rules,
                surplus_map,
                global_priority_override,
                macro_run_date,
                macro_now_dt,
                preferred_team,
                _need_headcount_logged_orders,
                team_combo_presets,
                dispatch_interval_mirror=dispatch_interval_mirror,
                machine_handoff=machine_handoff,
            )
            if res is None:
                break
            done_units = 1
            if task.get("roll_pipeline_inspection") or task.get(
                "roll_pipeline_rewind"
            ):
                _rp_room = _roll_pipeline_inspection_assign_room(
                    task_queue, str(task.get("task_id", "") or "").strip()
                )
                if _rp_room <= 1e-12:
                    break
                done_units = min(
                    1, int(min(_rp_room, math.ceil(task["remaining_units"])))
                )
            if done_units <= 0:
                break
            best_team = tuple(res["team"])
            lead_op = res["lead_op"]
            sub_members = [m for m in best_team if m != lead_op]
            best_start = res["team_start"]
            best_end = res["actual_end_dt"]
            best_breaks = res["team_breaks"]
            best_eff = res["avg_eff"]
            rq_base = res["rq_base"]
            extra_max = res["extra_max"]
            eq_line = res["eq_line"]
            machine_occ_key = _machine_occupancy_key_resolve(task, eq_line)
            _te_disp = parse_float_safe(task.get("task_eff_factor"), 1.0)
            if _te_disp <= 0:
                _te_disp = 1.0

            total_u = (
                math.ceil(task["total_qty_m"] / task["unit_m"]) if task["unit_m"] else 0
            )
            rem_u_before = math.ceil(task["remaining_units"])
            already_done = total_u - rem_u_before
            try:
                tot_qty = parse_float_safe(task.get("total_qty_m"), 0.0)
                done_qty = parse_float_safe(task.get("done_qty_reported"), 0.0)
                if tot_qty > 0:
                    pct_macro = max(
                        0, min(100, int(round((done_qty / tot_qty) * 100)))
                    )
                else:
                    pct_macro = 0
            except Exception:
                pct_macro = 0

            _mach_sub_line = ", ".join(
                str(s).strip() for s in sub_members if s and str(s).strip()
            )
            _append_changeover_segments_to_timeline(
                timeline_events,
                dispatch_interval_mirror,
                avail_dt,
                daily_status,
                current_date=current_date,
                task_id=str(task.get("task_id") or ""),
                machine_occ_key=machine_occ_key,
                segments=list(res.get("changeover_segments") or []),
                machining_lead_op=str(lead_op or "").strip() or None,
                machining_sub_str=_mach_sub_line or None,
                machine_handoff=machine_handoff,
            )
            timeline_events.append(
                {
                    "date": current_date,
                    "task_id": task["task_id"],
                    "machine": eq_line,
                    "machine_occupancy_key": machine_occ_key,
                    "op": lead_op,
                    "sub": ", ".join(sub_members),
                    "start_dt": best_start,
                    "end_dt": best_end,
                    "breaks": best_breaks,
                    "units_done": done_units,
                    "already_done_units": already_done,
                    "total_units": total_u,
                    "pct_macro": pct_macro,
                    "eff_time_per_unit": task["base_time_per_unit"]
                    / best_eff
                    / _te_disp
                    * _surplus_team_time_factor(
                        rq_base, len(best_team), extra_max
                    ),
                    "unit_m": task["unit_m"],
                    "event_kind": TIMELINE_EVENT_MACHINING,
                }
            )
            if dispatch_interval_mirror is not None:
                dispatch_interval_mirror.register_from_event(timeline_events[-1])
            task["remaining_units"] -= float(done_units)
            op_main = (lead_op or "").strip()
            subs_part = ",".join(
                s.strip() for s in sub_members if s and str(s).strip()
            )
            team_s = f"{op_main}, {subs_part}" if subs_part else op_main
            req_num_run = int(res.get("req_num") or 0)
            extra_max_run = int(res.get("extra_max") or 0)
            need_surplus_assigned = (
                TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
                and extra_max_run > 0
                and len(best_team) > req_num_run
            )
            names_ordered: list[str] = []
            if op_main:
                names_ordered.append(op_main)
            for _m in sub_members:
                if _m and str(_m).strip():
                    names_ordered.append(str(_m).strip())
            surplus_member_names = (
                names_ordered[req_num_run:]
                if need_surplus_assigned
                and len(names_ordered) > req_num_run
                else []
            )
            task["assigned_history"].append(
                {
                    "date": current_date.strftime("%m/%d"),
                    "team": team_s,
                    "done_m": int(done_units * task["unit_m"]),
                    "start_dt": best_start,
                    "end_dt": best_end,
                    "need_surplus_assigned": need_surplus_assigned,
                    "combo_sheet_row_id": res.get("combo_sheet_row_id"),
                    "surplus_member_names": surplus_member_names,
                }
            )
            for m in best_team:
                avail_dt[m] = best_end
            if not _gpo.get("abolish_all_scheduling_limits"):
                machine_avail_dt[machine_occ_key] = best_end
                _bump_machine_avail_after_roll_for_calendar(
                    current_date,
                    machine_occ_key,
                    machine_avail_dt,
                    machine_calendar_plan_end=_mc_plan_end,
                    machine_day_floor=_mc_w0,
                )
            machine_handoff["last_tid"][machine_occ_key] = str(
                task.get("task_id") or ""
            ).strip()
            machine_handoff["last_eq"][machine_occ_key] = eq_line
            machine_handoff["started_today"].add(machine_occ_key)
            machine_handoff["machining_today_occ"].add(machine_occ_key)
            machine_handoff["last_machining_dt"][machine_occ_key] = best_end
            machine_handoff["last_machining_date"][machine_occ_key] = current_date
            machine_handoff["last_lead_op"][machine_occ_key] = lead_op
            machine_handoff.setdefault("last_machining_sub", {})
            machine_handoff["last_machining_sub"][machine_occ_key] = _mach_sub_line
            if _trace_schedule_task_enabled(task.get("task_id")):
                _log_dispatch_trace_schedule(
                    task.get("task_id"),
                    "[?????? task=%s] ????? ??? day=%s machine=%s machine_name=%s "
                    "start=%s end=%s ????=%s req_num=%s ?????extra_max=%s "
                    "??????(???)=%s team=%s",
                    task.get("task_id"),
                    current_date,
                    eq_line,
                    str(task.get("machine_name") or "").strip(),
                    best_start,
                    best_end,
                    len(best_team),
                    req_num_run,
                    extra_max_run,
                    need_surplus_assigned,
                    team_s,
                )
            preferred_team = best_team
            made_local = True
            rolls_done += 1
        return made_local

    def _is_b2_follower_phase2_row(t: dict) -> bool:
        _tid = str(t.get("task_id") or "").strip()
        return bool(
            (
                t.get("roll_pipeline_inspection")
                or t.get("roll_pipeline_rewind")
            )
            and _task_queue_has_roll_pipeline_ec_for_tid(task_queue, _tid)
        )

    phase1_tasks = [t for t in eligible_sorted if not _is_b2_follower_phase2_row(t)]
    phase2_tasks = [t for t in eligible_sorted if _is_b2_follower_phase2_row(t)]

    phase2_tids: set[str] = {
        str(t.get("task_id") or "").strip()
        for t in phase2_tasks
        if str(t.get("task_id") or "").strip()
    }
    phase2_mocc: set[str] = set()
    for t in phase2_tasks:
        _tm = t.get("machine")
        _eqt = str(t.get("equipment_line_key") or _tm or "").strip() or (_tm or "")
        _pk = (_machine_occupancy_key_resolve(t, _eqt) or "").strip()
        if _pk:
            phase2_mocc.add(_pk)

    phase1_interleave: list = []
    phase1_rest: list = []
    for t in phase1_tasks:
        _tid1 = str(t.get("task_id") or "").strip()
        _tm = t.get("machine")
        _eqt = str(t.get("equipment_line_key") or _tm or "").strip() or (_tm or "")
        _mk = (_machine_occupancy_key_resolve(t, _eqt) or "").strip()
        _same_tid_ec = bool(t.get("roll_pipeline_ec") and _tid1 and _tid1 in phase2_tids)
        _share_m = bool(_mk and _mk in phase2_mocc)
        if _same_tid_ec or _share_m:
            phase1_interleave.append(t)
        else:
            phase1_rest.append(t)

    def _b2_merged_sort_key(t: dict) -> tuple:
        # ???????????????????? EC ???????????????
        # ?????????????�B-2 ?????? EC ?????????????
        _fol = bool(
            t.get("roll_pipeline_inspection") or t.get("roll_pipeline_rewind")
        )
        return (
            int(t.get("dispatch_trial_order") or 10**9),
            0 if _fol else 1,
            str(t.get("task_id") or ""),
            int(t.get("same_request_line_seq") or 0),
        )

    pass_made = False
    if phase2_tasks:
        merged_b2 = sorted(
            phase1_interleave + phase2_tasks,
            key=_b2_merged_sort_key,
        )
        _merged_row_ids = {id(x) for x in merged_b2}

        def _b2_rr_key(t: dict) -> tuple:
            if id(t) in _merged_row_ids:
                return _b2_merged_sort_key(t)
            return (
                int(t.get("dispatch_trial_order") or 10**9),
                2,
                str(t.get("task_id") or ""),
                int(t.get("same_request_line_seq") or 0),
            )

        all_rr = sorted(merged_b2 + phase1_rest, key=_b2_rr_key)
        while True:
            round_made = False
            for task in all_rr:
                if float(task.get("remaining_units") or 0) <= 1e-12:
                    continue
                if _drain_rolls_for_task(task, max_rolls=1):
                    round_made = True
            if not round_made:
                break
            pass_made = True
    else:
        for task in phase1_tasks:
            if _drain_rolls_for_task(task):
                pass_made = True
    return pass_made


def _run_b2_inspection_rewind_pass(
    sorted_dates: list,
    attendance_data: dict,
    task_queue: list,
    timeline_events: list,
    skills_dict: dict,
    members: list,
    equipment_list: list,
    req_map: dict,
    need_rules: list,
    surplus_map: dict,
    global_priority_override: dict,
    macro_run_date: date,
    macro_now_dt: datetime,
    _need_headcount_logged_orders: set,
    team_combo_presets: dict | None = None,
    dispatch_interval_mirror: DispatchIntervalMirror | None = None,
) -> bool:
    """
    �B-2 / �B-3: EC ??????????????????????????????????????
    timeline_events ????????????????????????????????
    """
    target_tids: set[str] = set()
    for t in task_queue:
        if float(t.get("remaining_units") or 0) <= 1e-12:
            continue
        if not (t.get("roll_pipeline_inspection") or t.get("roll_pipeline_rewind")):
            continue
        tid = str(t.get("task_id", "") or "").strip()
        if not tid:
            continue
        if not _task_queue_has_roll_pipeline_ec_for_tid(task_queue, tid):
            continue
        if not _pipeline_ec_fully_done_for_tid(task_queue, tid):
            continue
        target_tids.add(tid)
    if not target_tids:
        return False

    _gpo = global_priority_override or {}
    _any_progress = False
    _machine_day_start_cache: dict[date, datetime] = {}
    for current_date in sorted_dates:
        daily_status = attendance_data.get(current_date)
        if not daily_status:
            continue
        machine_avail_dt: dict = {}
        avail_dt: dict = {}
        for m in members:
            if m not in daily_status:
                continue
            st = daily_status[m]
            if st.get("eligible_for_assignment", st.get("is_working", False)):
                avail_dt[m] = st["start_dt"]
        if not avail_dt:
            continue

        _machine_day_start = _machine_day_start_cache.get(current_date)
        if _machine_day_start is None:
            _machine_day_start = datetime.combine(current_date, DEFAULT_START_TIME)
            _machine_day_start_cache[current_date] = _machine_day_start
        _seed_avail_from_timeline_for_date(
            timeline_events,
            current_date,
            machine_avail_dt,
            avail_dt,
            _machine_day_start,
        )
        if _gpo.get("abolish_all_scheduling_limits"):
            machine_avail_dt.clear()
        _mc_plan_end_b2 = _machine_calendar_planning_window_end_dt(
            current_date, daily_status, members
        )
        _apply_machine_calendar_floor_for_date(
            current_date,
            machine_avail_dt,
            equipment_list,
            _machine_day_start,
            machine_calendar_plan_end=_mc_plan_end_b2,
        )

        tasks_today = [
            t
            for t in task_queue
            if float(t.get("remaining_units") or 0) > 1e-12
            and (
                t.get("roll_pipeline_inspection")
                or t.get("roll_pipeline_rewind")
            )
            and str(t.get("task_id", "") or "").strip() in target_tids
            and t.get("start_date_req") <= current_date
        ]
        if not tasks_today:
            continue

        _sched_max_passes = max(96, max(1, len(tasks_today)) * 15)
        _sched_pi = 0
        while _sched_pi < _sched_max_passes:
            _sched_pi += 1
            _made = _trial_order_first_schedule_pass(
                current_date,
                tasks_today,
                task_queue,
                daily_status,
                machine_avail_dt,
                avail_dt,
                timeline_events,
                skills_dict,
                members,
                req_map,
                need_rules,
                surplus_map,
                global_priority_override,
                macro_run_date,
                macro_now_dt,
                _need_headcount_logged_orders,
                team_combo_presets,
                dispatch_interval_mirror=dispatch_interval_mirror,
            )
            if not _made:
                break
            _any_progress = True
    return _any_progress


def _timeline_event_team_names_set(ev: dict) -> set:
    names: set = set()
    op = str(ev.get("op") or "").strip()
    if op:
        names.add(op)
    sub = str(ev.get("sub") or "").strip()
    if sub:
        for s in sub.split(","):
            t = s.strip()
            if t:
                names.add(t)
    return names


def _task_dict_for_timeline_event(ev: dict, task_queue: list) -> dict | None:
    tid = str(ev.get("task_id") or "").strip()
    if not tid:
        return None
    eq = str(ev.get("machine") or "").strip()
    for t in task_queue:
        if str(t.get("task_id") or "").strip() != tid:
            continue
        t_eq = str(t.get("equipment_line_key") or t.get("machine") or "").strip()
        if t_eq == eq:
            return t
    for t in task_queue:
        if str(t.get("task_id") or "").strip() == tid:
            return t
    return None


def _member_overlaps_busy(
    busy_map: dict, member: str, st: datetime, ed: datetime
) -> bool:
    for bs, be in busy_map.get(member, ()):
        if st < be and bs < ed:
            return True
    return False


def append_surplus_staff_after_main_dispatch(
    timeline_events: list,
    attendance_data: dict,
    skills_dict: dict,
    members: list,
    task_queue: list,
    req_map: dict,
    need_rules: list,
    surplus_map: dict,
    global_priority_override: dict | None,
) -> int:
    """
    need??????????????????????????????????????????????
    ?????????????????????????????????????????
    eligible ?? OP/AS ??????????????
    ???????????????????event_kind ?????????????????????????????
    """
    gpo = global_priority_override or {}
    if not surplus_map or TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW:
        return 0

    busy: dict[str, list[tuple[datetime, datetime]]] = defaultdict(list)
    for e in timeline_events:
        st = e.get("start_dt")
        ed = e.get("end_dt")
        if not isinstance(st, datetime) or not isinstance(ed, datetime):
            continue
        for name in _timeline_event_team_names_set(e):
            busy[name].append((st, ed))

    appended_total = 0
    sorted_evs = sorted(
        (
            e
            for e in timeline_events
            if isinstance(e.get("start_dt"), datetime)
            and isinstance(e.get("end_dt"), datetime)
        ),
        key=lambda x: (x.get("date"), x.get("start_dt") or datetime.min),
    )

    for ev in sorted_evs:
        d = ev.get("date")
        if d is None or d not in attendance_data:
            continue
        if not _is_machining_timeline_event(ev):
            continue
        daily_status = attendance_data[d]
        task = _task_dict_for_timeline_event(ev, task_queue)
        if task is None:
            continue
        machine = task.get("machine")
        machine_name = str(task.get("machine_name") or "").strip()
        tid = str(task.get("task_id") or "").strip()

        if TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY:
            req_num = resolve_need_required_op(
                str(machine or "").strip(),
                machine_name,
                tid,
                req_map,
                need_rules,
            )
        else:
            ro = task.get("required_op")
            if ro is not None:
                try:
                    riv = int(ro)
                    if riv >= 1:
                        req_num = riv
                    else:
                        req_num = resolve_need_required_op(
                            str(machine or "").strip(),
                            machine_name,
                            tid,
                            req_map,
                            need_rules,
                        )
                except (TypeError, ValueError):
                    req_num = resolve_need_required_op(
                        str(machine or "").strip(),
                        machine_name,
                        tid,
                        req_map,
                        need_rules,
                    )
            else:
                req_num = resolve_need_required_op(
                    str(machine or "").strip(),
                    machine_name,
                    tid,
                    req_map,
                    need_rules,
                )
        if gpo.get("ignore_need_minimum"):
            req_num = 1

        extra_max_sheet = resolve_need_surplus_extra_max(
            str(machine or "").strip(),
            machine_name,
            tid,
            surplus_map,
            need_rules,
        )
        if extra_max_sheet <= 0:
            continue

        names = _timeline_event_team_names_set(ev)
        team_size = len(names)
        cap_add = req_num + extra_max_sheet - team_size
        if cap_add <= 0:
            continue

        skill_meta_cache: dict = {}

        def skill_role_priority(mem):
            if gpo.get("ignore_skill_requirements"):
                return ("OP", 100)
            if mem not in skill_meta_cache:
                srow = skills_dict.get(mem, {})
                machine_proc = str(machine or "").strip()
                v = ""
                if machine_proc and machine_name:
                    v = srow.get(f"{machine_proc}+{machine_name}", "")
                elif machine_name:
                    v = srow.get(machine_name, "")
                elif machine_proc:
                    v = srow.get(machine_proc, "")
                skill_meta_cache[mem] = parse_op_as_skill_cell(v)
            return skill_meta_cache[mem]

        capable = []
        for mem in members:
            if mem not in daily_status:
                continue
            st_ent = daily_status[mem]
            if not st_ent.get(
                "eligible_for_assignment", st_ent.get("is_working", False)
            ):
                continue
            if skill_role_priority(mem)[0] not in ("OP", "AS"):
                continue
            capable.append(mem)
        capable.sort(key=lambda mm: (skill_role_priority(mm)[1], mm))

        st = ev["start_dt"]
        ed = ev["end_dt"]
        candidates = [
            m
            for m in capable
            if m not in names and not _member_overlaps_busy(busy, m, st, ed)
        ]
        candidates.sort(
            key=lambda mm: (
                0 if skill_role_priority(mm)[0] == "AS" else 1,
                skill_role_priority(mm)[1],
                mm,
            )
        )

        chosen = candidates[:cap_add]
        if not chosen:
            continue

        team_size_before = team_size
        final_team_size = team_size_before + len(chosen)
        highlight_surplus = final_team_size > req_num

        old_sub = str(ev.get("sub") or "").strip()
        parts = [s.strip() for s in old_sub.split(",") if s.strip()]
        parts.extend(chosen)
        ev["sub"] = ", ".join(parts)
        for m in chosen:
            busy[m].append((st, ed))
        appended_total += len(chosen)

        op_sync = str(ev.get("op") or "").strip()
        subs_sync = ",".join(
            s.strip()
            for s in str(ev.get("sub") or "").split(",")
            if s.strip()
        )
        team_sync = f"{op_sync}, {subs_sync}" if subs_sync else op_sync

        _hist = task.get("assigned_history")
        if _hist:
            for h in _hist:
                if (
                    h.get("start_dt") == st
                    and h.get("end_dt") == ed
                ):
                    if highlight_surplus:
                        h["need_surplus_assigned"] = True
                    h["team"] = team_sync
                    prev_pd = h.get("post_dispatch_surplus_names") or []
                    h["post_dispatch_surplus_names"] = prev_pd + [
                        str(x) for x in chosen
                    ]
                    break

        if _trace_schedule_task_enabled(tid):
            _log_dispatch_trace_schedule(
                tid,
                "[?????? task=%s] ????(??????) day=%s machine=%s machine_name=%s "
                "start=%s end=%s ????=%s ?????=%s ?????=%s req_num=%s "
                "need???(???)=%s ???(??????)=%s ??????=%s",
                tid,
                d,
                str(machine or "").strip(),
                machine_name,
                st,
                ed,
                len(chosen),
                team_size_before,
                final_team_size,
                req_num,
                extra_max_sheet,
                highlight_surplus,
                ",".join(chosen),
            )

    return appended_total


# =========================================================
# 3. ??????? (????????????)
#    ??2????plan_simulation_stage2 ???????????
#    ????????? ? ?????? ? ????????OP?? ? ????????
# =========================================================
def generate_plan():
    """
    ??2????????????????Excel ???????

    ??: ???? TASK_INPUT_WORKBOOK??????????????????????
    ??: ``output_dir`` ??? ``production_plan_multi_day_*.xlsx`` / ``member_schedule_*.xlsx``???1???????? log/execution_log.txt?
    """
    master_abs = os.path.abspath(os.path.join(os.getcwd(), MASTER_FILE))
    with _override_default_factory_hours_from_master(master_abs):
        _generate_plan_impl()


def _generate_plan_impl():
    # ???????????? A3 ???????????0????? return ???
    # execution_log ????? skills ??????????????
    global TRACE_SCHEDULE_TASK_IDS, DEBUG_DISPATCH_ONLY_TASK_IDS
    _wb_trace = (os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK)
    _ids_from_sheet = _read_trace_schedule_task_ids_from_config_sheet(_wb_trace)
    TRACE_SCHEDULE_TASK_IDS = frozenset(
        str(x).strip() for x in _ids_from_sheet if str(x).strip()
    )
    if _ids_from_sheet:
        _preview = _ids_from_sheet[:25]
        _suffix = " �" if len(_ids_from_sheet) > 25 else ""
        logging.info(
            "??????%s?A3 ??: ???????NO? %s ??????%s%s?",
            APP_CONFIG_SHEET_NAME,
            len(_ids_from_sheet),
            ", ".join(_preview),
            _suffix,
        )
    else:
        logging.info(
            "??????%s?A3 ??: ???????NO??????????????",
            APP_CONFIG_SHEET_NAME,
        )
    if TRACE_SCHEDULE_TASK_IDS:
        logging.info(
            "??????: ?? task_id = %s?????? A3 ???",
            ", ".join(sorted(TRACE_SCHEDULE_TASK_IDS)),
        )
    else:
        logging.info(
            "??????: ?????[?????? �] ????????"
        )
    _ids_debug_dispatch_raw = _read_debug_dispatch_task_ids_from_config_sheet(_wb_trace)
    _dbg_norm: list[str] = []
    for _dx in _ids_debug_dispatch_raw:
        _dt = planning_task_id_str_from_scalar(_dx)
        if _dt:
            _dbg_norm.append(_dt)
    DEBUG_DISPATCH_ONLY_TASK_IDS = frozenset(_dbg_norm)
    if DEBUG_DISPATCH_ONLY_TASK_IDS:
        logging.warning(
            "??????: ?%s?B3?????????? %s ????NO??????: %s",
            APP_CONFIG_SHEET_NAME,
            len(DEBUG_DISPATCH_ONLY_TASK_IDS),
            ", ".join(sorted(DEBUG_DISPATCH_ONLY_TASK_IDS)),
        )
        _show_stage2_debug_dispatch_mode_dialog(sorted(DEBUG_DISPATCH_ONLY_TASK_IDS))
    if TRACE_TEAM_ASSIGN_TASK_ID:
        logging.info(
            "???? TRACE_TEAM_ASSIGN_TASK_ID=%r ? ???????????",
            TRACE_TEAM_ASSIGN_TASK_ID,
        )

    _reset_dispatch_trace_per_task_logfiles()

    (
        skills_dict,
        members,
        equipment_list,
        req_map,
        need_rules,
        surplus_map,
        need_combo_col_index,
    ) = load_skills_and_needs()
    team_combo_presets = load_team_combination_presets_from_master()
    if team_combo_presets:
        _nrules = sum(len(v) for v in team_combo_presets.values())
        logging.info(
            "??????: ??+???? %s ?????? %s ???????????????????",
            len(team_combo_presets),
            _nrules,
        )
    elif TEAM_ASSIGN_USE_MASTER_COMBO_SHEET:
        logging.info(
            "??????: ????????????????????????????????????"
        )
    if not members:
        master_abs = os.path.abspath(MASTER_FILE)
        logging.error(
            "??2???????: ?????0???????? skills ??????????????"
            " ????: %s ?????: %s??????????? master.xlsm ????"
            "planning_core ????????????????? python\\ ????????????"
            " ?????? production_plan / member_schedule ?????????",
            master_abs,
            os.getcwd(),
        )
        return
    global _MACHINE_CALENDAR_BLOCKS_BY_DATE
    global _STAGE2_MACHINE_CHANGEOVER_BY_EQ, _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE
    global _STAGE2_REGULAR_SHIFT_START
    try:
        _MACHINE_CALENDAR_BLOCKS_BY_DATE = load_machine_calendar_occupancy_blocks(
            os.path.abspath(os.path.join(os.getcwd(), MASTER_FILE)),
            equipment_list,
        )
    except Exception as e:
        logging.warning(
            "???????: ??????????????????? (%s)", e
        )
        _MACHINE_CALENDAR_BLOCKS_BY_DATE = {}
    try:
        (
            _STAGE2_MACHINE_CHANGEOVER_BY_EQ,
            _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE,
        ) = load_machine_changeover_settings(
            os.path.abspath(os.path.join(os.getcwd(), MASTER_FILE))
        )
    except Exception as e:
        logging.warning(
            "????/???????????: ???????????? (%s)", e
        )
        _STAGE2_MACHINE_CHANGEOVER_BY_EQ = {}
        _STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE = {}
    try:
        _rs_a15, _ = _read_master_main_regular_shift_times(
            os.path.abspath(os.path.join(os.getcwd(), MASTER_FILE))
        )
        _STAGE2_REGULAR_SHIFT_START = _rs_a15
        if _rs_a15 is not None:
            logging.info(
                "??????: ???? master ??? A15=%s ????[??, ??+?) ????????A15 ????????? forward?",
                _rs_a15.strftime("%H:%M"),
            )
    except Exception as e:
        logging.warning("????(A15) ????: ?????????? forward ???????? (%s)", e)
        _STAGE2_REGULAR_SHIFT_START = None
    if _MACHINE_CALENDAR_BLOCKS_BY_DATE:
        _n_iv = sum(
            len(ivs)
            for _dm in _MACHINE_CALENDAR_BLOCKS_BY_DATE.values()
            for ivs in _dm.values()
        )
        logging.info(
            "???????: %s ???????????? %s ??????????",
            len(_MACHINE_CALENDAR_BLOCKS_BY_DATE),
            _n_iv,
        )
    reset_gemini_usage_tracker()
    _clear_stage2_blocking_message_file()
    if (
        not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
        and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
    ):
        logging.info(
            "need???????: ???????????????"
            "???????????????????????????????????????????"
            "???????????????: TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS=1?"
        )

    # ??2??????????????????????????????
    data_extract_dt = _extract_data_extraction_datetime()
    base_now_dt = data_extract_dt if data_extract_dt is not None else datetime.now()
    run_date = base_now_dt.date()
    data_extract_dt_str = (
        base_now_dt.strftime("%Y/%m/%d %H:%M:%S") if data_extract_dt is not None else "�"
    )
    logging.info(
        "??????: %s?%s?",
        base_now_dt.strftime("%Y/%m/%d %H:%M:%S"),
        "??????" if data_extract_dt is not None else "???????????",
    )

    attendance_data, ai_log_data = load_attendance_and_analyze(members)
    global_priority_raw = load_main_sheet_global_priority_override_text()
    global_priority_override = analyze_global_priority_override_comment(
        global_priority_raw, members, run_date.year,         ai_sheet_sink=ai_log_data
    )
    _factory_closure_dates: set[date] = set()
    for _iso in global_priority_override.get("factory_closure_dates") or []:
        _d = parse_optional_date(_iso)
        if _d is not None:
            _factory_closure_dates.add(_d)
    if _factory_closure_dates:
        apply_factory_closure_dates_to_attendance(
            attendance_data, members, _factory_closure_dates
        )
        logging.info(
            "?????????????: ????????? ? %s",
            ", ".join(str(x) for x in sorted(_factory_closure_dates)),
        )
    ai_log_data["???_?????_?????(??)"] = (
        ", ".join(str(x) for x in sorted(_factory_closure_dates))
        if _factory_closure_dates
        else "????"
    )
    _sn = str(global_priority_override.get("scheduler_notes_ja") or "").strip()
    if _sn:
        ai_log_data["???_?????_?????(AI)"] = _sn[:2000]

    sorted_dates = sorted(list(attendance_data.keys()))
    # ?????????????????????????????????
    sorted_dates = [d for d in sorted_dates if d >= run_date]
    if not sorted_dates:
        logging.error("??????????????????")
        _try_write_main_sheet_gemini_usage_summary("??2")
        return

    # ?????: ?????????_?????????1?????????????
    try:
        tasks_df = load_planning_tasks_df()
    except Exception as e:
        logging.error(f"?????????????????: {e}")
        _try_write_main_sheet_gemini_usage_summary("??2")
        return

    if DEBUG_DISPATCH_ONLY_TASK_IDS:
        _n_tasks_before = len(tasks_df)
        _dbg_mask = tasks_df.apply(
            lambda row: planning_task_id_str_from_plan_row(row)
            in DEBUG_DISPATCH_ONLY_TASK_IDS,
            axis=1,
        )
        tasks_df = tasks_df.loc[_dbg_mask].copy()
        _n_tasks_after = len(tasks_df)
        logging.warning(
            "??????: ?%s???? %s ? %s ?????????",
            PLAN_INPUT_SHEET_NAME,
            _n_tasks_before,
            _n_tasks_after,
        )
        if _n_tasks_after == 0:
            logging.error(
                "??????: B3?????NO???????????????2???????"
            )
            _try_write_main_sheet_gemini_usage_summary("??2")
            return

    if global_priority_raw.strip():
        snip = global_priority_raw[:2500]
        if len(global_priority_raw) > 2500:
            snip += "�"
        ai_log_data["???_???????(??)"] = snip
    else:
        ai_log_data["???_???????(??)"] = (
            "????????????????????????????????????"
        )
    ai_log_data["???_???????(AI)"] = json.dumps(
        global_priority_override, ensure_ascii=False
    )
    if global_priority_override.get("ignore_skill_requirements"):
        logging.warning(
            "????????: ????????????????%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("ignore_need_minimum"):
        logging.warning(
            "????????: ??????1????????need?????OP?????????%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("abolish_all_scheduling_limits"):
        logging.warning(
            "????????: ????????????????????????????????????%s",
            global_priority_override.get("interpretation_ja", ""),
        )

    # ??????????????????????????????
    macro_now_dt = base_now_dt
    macro_run_date = macro_now_dt.date()
    ai_task_by_tid = analyze_task_special_remarks(
        tasks_df, reference_year=run_date.year,         ai_sheet_sink=ai_log_data
    )
    task_queue = build_task_queue_from_planning_df(
        tasks_df,
        run_date,
        req_map,
        ai_task_by_tid,
        global_priority_override,
        equipment_list,
    )
    # ????????????????????????: 4/4, 4/5 ?????? 4/3 ??
    working_days = [
        d for d in sorted_dates
        if any(attendance_data[d][m]["is_working"] for m in attendance_data[d])
    ]
    if working_days:
        for t in task_queue:
            req_d = t.get("start_date_req")
            if not isinstance(req_d, date):
                continue
            if req_d in working_days:
                continue
            prev_work = None
            for wd in working_days:
                if wd <= req_d:
                    prev_work = wd
                else:
                    break
            if prev_work is not None:
                if str(t.get("task_id", "")).strip() == DEBUG_TASK_ID:
                    logging.info(
                        "DEBUG[task=%s] start_date_req ???????: %s -> %s",
                        DEBUG_TASK_ID,
                        req_d,
                        prev_work,
                    )
                t["start_date_req"] = prev_work
    conflict_rows = collect_planning_conflicts_by_excel_row(tasks_df, ai_task_by_tid)
    _try_write_plan_input_global_parse_and_conflicts_one_save(
        global_priority_override,
        data_extract_dt_str,
        len(tasks_df),
        conflict_rows,
    )

    if not task_queue:
        logging.warning(
            f"??????????????{PLAN_INPUT_SHEET_NAME}?????NO?????????????"
            "???????????????????????????????????"
        )

    # ?????: ????????????????????? �B ?????need ??????? EC ???? 1..n
    _apply_dispatch_trial_order_for_generate_plan(
        task_queue, req_map, need_rules, need_combo_col_index
    )
    if DEBUG_TASK_ID:
        dbg_items = [t for t in task_queue if str(t.get("task_id", "")).strip() == DEBUG_TASK_ID]
        if dbg_items:
            t0 = dbg_items[0]
            logging.info(
                "DEBUG[task=%s] queue??: start_date_req=%s due_basis=%s answer_due=%s specified_due=%s specified_due_ov=%s due_source=%s priority=%s in_progress=%s remark=%s",
                DEBUG_TASK_ID,
                t0.get("start_date_req"),
                t0.get("due_basis_date"),
                t0.get("answer_due_date"),
                t0.get("specified_due_date"),
                t0.get("specified_due_override"),
                t0.get("due_source"),
                t0.get("priority"),
                t0.get("in_progress"),
                t0.get("has_special_remark"),
            )
        else:
            logging.info("DEBUG[task=%s] task_queue??????????/??0/??NO?????????", DEBUG_TASK_ID)
    timeline_events = []

    # ---------------------------------------------------------
    # ??????????????
    # STAGE2_EXTEND_ATTENDANCE_CALENDAR ? True ?????????????????????????
    # STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING ? True ?????: ???????????????????
    # due_basis +1??????????????????????????????????????????
    # ?? False ?????? 1 ?????????? 1 ?????
    # ---------------------------------------------------------
    _master_attendance_date_set = frozenset(attendance_data.keys())
    _master_plan_dates_template = list(sorted_dates)
    _calendar_last_plan_day = _master_plan_dates_template[-1]

    for t in task_queue:
        t["remaining_units"] = float(t.get("initial_remaining_units") or 0)
        t["assigned_history"].clear()
    timeline_events.clear()

    _dispatch_interval_mirror: DispatchIntervalMirror | None = None
    if DISPATCH_INTERVAL_MIRROR_ENFORCE:
        _dispatch_interval_mirror = DispatchIntervalMirror()
        logging.info(
            "DISPATCH_INTERVAL_MIRROR_ENFORCE: ???????????????????"
            "????? ??_???? ?? DISPATCH_INTERVAL_MIRROR_ENFORCE=0??"
        )

    if STAGE2_SERIAL_DISPATCH_BY_TASK_ID:
        logging.info(
            "??NO????: ???STAGE2_SERIAL_DISPATCH_BY_TASK_ID??"
            " ???????????NO????????????????????????????????"
        )
    else:
        logging.info(
            "??NO????: ???start_date ???????????????????????????????????"
        )

    _due_shift_retry_count_by_request: dict[str, int] = {}
    _due_shift_exhausted_requests: set[str] = set()
    _due_shift_cap_warned_tids: set[str] = set()
    _outer_retry_round = 0
    while True:
        _dispatch_trace_begin_outer_round(_outer_retry_round)
        _need_headcount_logged_orders: set = set()
        if _outer_retry_round > 0:
            _purge_attendance_days_not_in_set(
                attendance_data, _master_attendance_date_set
            )
            sorted_dates[:] = list(_master_plan_dates_template)

        for t in task_queue:
            t.pop("_partial_retry_calendar_blocked", None)

        if STAGE2_SERIAL_DISPATCH_BY_TASK_ID:
            _serial_order_tids = _serial_dispatch_order_task_ids(task_queue)
        else:
            _serial_order_tids = []

        _plan_day_iter = (
            _iter_plan_dates_extending(sorted_dates, attendance_data, task_queue)
            if STAGE2_EXTEND_ATTENDANCE_CALENDAR
            else sorted_dates
        )
        _full_calendar_without_deadline_restart = True
        for current_date in _plan_day_iter:
            daily_status = attendance_data[current_date]
            # ?????????????????????????
            machine_avail_dt = {}
            
            avail_dt = {}
            for m in members:
                if m not in daily_status:
                    continue
                st = daily_status[m]
                if st.get("eligible_for_assignment", st.get("is_working", False)):
                    avail_dt[m] = st["start_dt"]

            _machine_day_start = datetime.combine(current_date, DEFAULT_START_TIME)
            _machine_calendar_plan_end = _machine_calendar_planning_window_end_dt(
                current_date, daily_status, members
            )
            if avail_dt:
                _seed_avail_from_timeline_for_date(
                    timeline_events,
                    current_date,
                    machine_avail_dt,
                    avail_dt,
                    _machine_day_start,
                )
                _apply_machine_calendar_floor_for_date(
                    current_date,
                    machine_avail_dt,
                    equipment_list,
                    _machine_day_start,
                    machine_calendar_plan_end=_machine_calendar_plan_end,
                )

            if not avail_dt:
                logging.info("DEBUG[day=%s] ??????0?????????", current_date)
                continue
    
            tasks_today = [t for t in task_queue if t['remaining_units'] > 0 and t['start_date_req'] <= current_date]
            if STAGE2_SERIAL_DISPATCH_BY_TASK_ID and _serial_order_tids:
                _tasks_today_before_serial = len(tasks_today)
                _active_serial_tid = None
                for _tid in _serial_order_tids:
                    if any(
                        float(x.get("remaining_units") or 0) > 1e-12
                        for x in task_queue
                        if str(x.get("task_id", "") or "").strip() == _tid
                    ):
                        _active_serial_tid = _tid
                        break
                if _active_serial_tid is not None:
                    tasks_today = [
                        t
                        for t in tasks_today
                        if str(t.get("task_id", "") or "").strip() == _active_serial_tid
                    ]
                _serial_pos = (
                    _serial_order_tids.index(_active_serial_tid) + 1
                    if _active_serial_tid in _serial_order_tids
                    else 0
                )
                _pending_rows = sum(1 for t in task_queue if t["remaining_units"] > 0)
                logging.info(
                    "??NO???? day=%s ???????NO=%s ???????=%s/%s "
                    "??????(???)=%s ???=%s ?????(??)=%s",
                    current_date,
                    _active_serial_tid if _active_serial_tid is not None else "�",
                    _serial_pos if _serial_pos else "�",
                    len(_serial_order_tids),
                    _tasks_today_before_serial,
                    len(tasks_today),
                    _pending_rows,
                )
            pending_total = sum(1 for t in task_queue if t["remaining_units"] > 0)
            if not tasks_today:
                earliest_wait = min(
                    [t["start_date_req"] for t in task_queue if t["remaining_units"] > 0],
                    default=None,
                )
                logging.info(
                    "DEBUG[day=%s] ???????0? pending_total=%s earliest_start_date_req=%s",
                    current_date,
                    pending_total,
                    earliest_wait,
                )
            elif DEBUG_TASK_ID:
                has_dbg_today = any(str(t.get("task_id", "")).strip() == DEBUG_TASK_ID for t in tasks_today)
                if current_date.isoformat() == "2026-04-03" or has_dbg_today:
                    logging.info(
                        "DEBUG[day=%s] avail_members=%s tasks_today=%s (task=%s ??=%s)",
                        current_date,
                        len(avail_dt),
                        len(tasks_today),
                        DEBUG_TASK_ID,
                        has_dbg_today,
                    )
            
            _sched_max_passes = max(96, max(1, len(tasks_today)) * 15)
            _sched_pi = 0
            while _sched_pi < _sched_max_passes:
                _sched_pi += 1
                _sched_made_progress = False
                if STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST:
                    _sched_made_progress = _trial_order_first_schedule_pass(
                        current_date,
                        tasks_today,
                        task_queue,
                        daily_status,
                        machine_avail_dt,
                        avail_dt,
                        timeline_events,
                        skills_dict,
                        members,
                        req_map,
                        need_rules,
                        surplus_map,
                        global_priority_override,
                        macro_run_date,
                        macro_now_dt,
                        _need_headcount_logged_orders,
                        team_combo_presets,
                        dispatch_interval_mirror=_dispatch_interval_mirror,
                    )
                if not STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST:
                    _mh_legacy_day = _machine_handoff_state_from_timeline(
                        timeline_events, current_date
                    )
                    machine_handoff_legacy = {
                        "last_tid": dict(_mh_legacy_day["last_tid"]),
                        "last_eq": dict(_mh_legacy_day["last_eq"]),
                        "started_today": set(_mh_legacy_day["started_today"]),
                        "machining_today_occ": set(
                            _mh_legacy_day.get("machining_today_occ") or set()
                        ),
                        "last_machining_dt": dict(
                            _mh_legacy_day.get("last_machining_dt") or {}
                        ),
                        "last_machining_date": dict(
                            _mh_legacy_day.get("last_machining_date") or {}
                        ),
                        "last_lead_op": dict(_mh_legacy_day.get("last_lead_op") or {}),
                        "last_machining_sub": dict(
                            _mh_legacy_day.get("last_machining_sub") or {}
                        ),
                    }
                    _assign_probe_ctx_legacy: dict | None = None
                    _min_dispatch_eff_legacy: int | None = None
                    if STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT:
                        _assign_probe_ctx_legacy = {
                            "avail_dt": avail_dt,
                            "machine_avail_dt": machine_avail_dt,
                            "task_queue": task_queue,
                            "skills_dict": skills_dict,
                            "members": members,
                            "req_map": req_map,
                            "need_rules": need_rules,
                            "surplus_map": surplus_map,
                            "global_priority_override": global_priority_override,
                            "macro_run_date": macro_run_date,
                            "macro_now_dt": macro_now_dt,
                            "machine_handoff": machine_handoff_legacy,
                            "team_combo_presets": team_combo_presets,
                            "dispatch_interval_mirror": _dispatch_interval_mirror,
                        }
                        _pool_legacy = _tasks_in_min_pending_dispatch_pool(
                            task_queue,
                            current_date,
                            daily_status=daily_status,
                            members=members,
                            machine_avail_dt=machine_avail_dt,
                            machine_day_start=_machine_day_start,
                            machine_handoff=machine_handoff_legacy,
                            skills_dict=skills_dict,
                            abolish_all_scheduling_limits=bool(
                                global_priority_override.get(
                                    "abolish_all_scheduling_limits"
                                )
                            ),
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                        )
                        _min_dispatch_eff_legacy = (
                            _effective_min_dispatch_trial_order_from_pool(
                                _pool_legacy,
                                current_date,
                                daily_status,
                                _assign_probe_ctx_legacy,
                            )
                        )
                    for task in sorted(
                        [t for t in tasks_today if float(t.get("remaining_units") or 0) > 1e-12],
                        key=lambda t: _day_schedule_task_sort_key(
                            t, task_queue, need_combo_col_index
                        ),
                    ):
                        if _task_blocked_by_same_request_dependency(task, task_queue):
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[?????? task=%s] ????: ????NO??????? day=%s machine=%s rem=%.4f",
                                    task.get("task_id"),
                                    current_date,
                                    task.get("machine"),
                                    float(task.get("remaining_units") or 0),
                                )
                            continue
                        if (
                            task.get("roll_pipeline_inspection")
                            or task.get("roll_pipeline_rewind")
                        ) and (
                            _roll_pipeline_inspection_assign_room(
                                task_queue, str(task.get("task_id", "")).strip()
                            )
                            <= 1e-12
                        ):
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _tid_tr = str(task.get("task_id", "") or "").strip()
                                _ec_d = _pipeline_ec_roll_done_units(task_queue, _tid_tr)
                                _in_d = _pipeline_b2_follower_roll_done_units(
                                    task_queue, _tid_tr
                                )
                                _log_dispatch_trace_schedule(
                                    _tid_tr,
                                    "[?????? task=%s] ????: �B-2/�B-3 ???????? day=%s machine=%s "
                                    "ec????R=%.4f ??????R=%.4f rem_follower=%.4f",
                                    _tid_tr,
                                    current_date,
                                    task.get("machine"),
                                    _ec_d,
                                    _in_d,
                                    float(task.get("remaining_units") or 0),
                                )
                            continue
                        machine = task["machine"]
                        eq_line = str(
                            task.get("equipment_line_key") or machine or ""
                        ).strip() or machine
                        machine_occ_key = _machine_occupancy_key_resolve(task, eq_line)
                        if PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE:
                            _b1_holder = _exclusive_b1_inspection_holder_for_machine(
                                task_queue,
                                machine_occ_key,
                            )
                            if _b1_holder is not None and _b1_holder is not task:
                                if _trace_schedule_task_enabled(task.get("task_id")):
                                    _log_dispatch_trace_schedule(
                                        task.get("task_id"),
                                        "[?????? task=%s] ????: ?????????? day=%s "
                                        "?????NO=%s ??????=%s",
                                        task.get("task_id"),
                                        current_date,
                                        _b1_holder.get("task_id"),
                                        _b1_holder.get("dispatch_trial_order"),
                                    )
                                continue
                        if DEBUG_TASK_ID and str(task.get("task_id", "")).strip() == DEBUG_TASK_ID:
                            logging.info(
                                "DEBUG[task=%s] day=%s ????: start_date_req=%s remaining_units=%s machine=%s",
                                DEBUG_TASK_ID,
                                current_date,
                                task.get("start_date_req"),
                                task.get("remaining_units"),
                                task.get("machine"),
                            )
                        if task.get("has_done_deadline_override"):
                            logging.info(
                                "DEBUG[?????] ??NO=%s ??=%s start_date_req=%s due_basis=%s ????(???)=%s ??=%s/%s",
                                task.get("task_id"),
                                current_date,
                                task.get("start_date_req"),
                                task.get("due_basis_date"),
                                task.get("specified_due_override"),
                                task.get("done_qty_reported"),
                                task.get("total_qty_m"),
                            )
                        try:
                            _my_dispatch_ord = int(
                                task.get("dispatch_trial_order") or 10**9
                            )
                        except (TypeError, ValueError):
                            _my_dispatch_ord = 10**9
                        if _task_blocked_by_global_dispatch_trial_order(
                            task,
                            task_queue,
                            current_date,
                            daily_status=daily_status,
                            members=members,
                            machine_avail_dt=machine_avail_dt,
                            machine_day_start=_machine_day_start,
                            machine_handoff=machine_handoff_legacy,
                            skills_dict=skills_dict,
                            abolish_all_scheduling_limits=bool(
                                global_priority_override.get(
                                    "abolish_all_scheduling_limits"
                                )
                            ),
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                            min_dispatch_effective=_min_dispatch_eff_legacy,
                        ):
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[?????? task=%s] ????: ???????????????? "
                                    "day=%s my_order=%s",
                                    task.get("task_id"),
                                    current_date,
                                    _my_dispatch_ord,
                                )
                            continue
                        if _task_fully_machine_calendar_blocked_on_date(
                            task, current_date, daily_status, members
                        ):
                            continue
                        if _task_no_machining_window_left_from_avail_floor(
                            task,
                            current_date,
                            daily_status,
                            members,
                            machine_avail_dt,
                            _machine_day_start,
                            machine_handoff=machine_handoff_legacy,
                            skills_dict=skills_dict,
                            abolish_all_scheduling_limits=bool(
                                global_priority_override.get(
                                    "abolish_all_scheduling_limits"
                                )
                            ),
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                        ):
                            continue
                        if _equipment_line_lower_dispatch_trial_still_pending(
                            task_queue,
                            machine_occ_key,
                            _my_dispatch_ord,
                            current_date,
                            daily_status=daily_status,
                            members=members,
                            machine_avail_dt=machine_avail_dt,
                            machine_day_start=_machine_day_start,
                            machine_handoff=machine_handoff_legacy,
                            skills_dict=skills_dict,
                            abolish_all_scheduling_limits=bool(
                                global_priority_override.get(
                                    "abolish_all_scheduling_limits"
                                )
                            ),
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                        ):
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[?????? task=%s] ????: ?????????????????? "
                                    "day=%s eq_line=%s my_order=%s",
                                    task.get("task_id"),
                                    current_date,
                                    eq_line,
                                    _my_dispatch_ord,
                                )
                            continue
                        machine_name = str(task.get("machine_name", "") or "").strip()
                        machine_proc = str(machine or "").strip()
                        plan_ro = _plan_sheet_required_op_optional(task)
                        need_src_line = ""
                        if TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY:
                            req_num, need_src_line = resolve_need_required_op_explain(
                                machine,
                                machine_name,
                                task["task_id"],
                                req_map,
                                need_rules,
                            )
                            if plan_ro is not None and plan_ro != req_num:
                                need_src_line = (
                                    (need_src_line + "?") if need_src_line else ""
                                )
                                need_src_line += (
                                    f"?????????{plan_ro}?????need??={req_num}?"
                                )
                        else:
                            if plan_ro is not None:
                                req_num = plan_ro
                                need_src_line = f"????????OP(??)?={req_num}"
                            else:
                                req_num, need_src_line = resolve_need_required_op_explain(
                                    machine,
                                    machine_name,
                                    task["task_id"],
                                    req_map,
                                    need_rules,
                                )
                        if global_priority_override.get("ignore_need_minimum"):
                            req_num = 1
                            need_src_line = (
                                (need_src_line + " ? ")
                                if need_src_line
                                else ""
                            ) + "?????ignore_need_minimum?req=1"
    
                        # ????�??????parse_op_as_skill_cell: ???????????????????
                        # skills ???????????????????????????+???????????
                        # ????????+??????????????????? OP ??????????????
                        skill_meta_cache = {}
                        _gpo = global_priority_override
    
                        def skill_role_priority(mem):
                            if _gpo.get("ignore_skill_requirements"):
                                return ("OP", 100)
                            if mem not in skill_meta_cache:
                                srow = skills_dict.get(mem, {})
                                v = ""
                                if machine_proc and machine_name:
                                    v = srow.get(f"{machine_proc}+{machine_name}", "")
                                elif machine_name:
                                    v = srow.get(machine_name, "")
                                elif machine_proc:
                                    v = srow.get(machine_proc, "")
                                skill_meta_cache[mem] = parse_op_as_skill_cell(v)
                            return skill_meta_cache[mem]
    
                        capable_members = [m for m in avail_dt.keys() if skill_role_priority(m)[0] in ("OP", "AS")]
                        capable_members.sort(key=lambda mm: (skill_role_priority(mm)[1], mm))
                        capable_members = _filter_capable_members_b2_disjoint_teams(
                            task, task_queue, capable_members
                        )
                        if task.get("has_done_deadline_override"):
                            machine_free_dbg = _machine_effective_floor_for_assign(
                                machine_occ_key,
                                str(task.get("task_id") or "").strip(),
                                eq_line,
                                str(task.get("machine_name") or "").strip(),
                                machine_avail_dt,
                                machine_handoff_legacy,
                                _machine_day_start,
                                bool(_gpo.get("abolish_all_scheduling_limits")),
                                current_date=current_date,
                                daily_status=daily_status,
                                skills_dict=skills_dict,
                                machine_proc=machine_proc,
                            )
                            logging.info(
                                "DEBUG[?????] ??NO=%s ??=%s req_num=%s capable_members=%s machine_free=%s",
                                task.get("task_id"),
                                eq_line,
                                req_num,
                                len(capable_members),
                                machine_free_dbg,
                            )
    
                        pref_raw = str(task.get("preferred_operator_raw") or "").strip()
                        op_today = [m for m in capable_members if skill_role_priority(m)[0] == "OP"]
                        pref_mem = (
                            _resolve_preferred_op_to_member(pref_raw, op_today, members)
                            if pref_raw
                            else None
                        )
                        if pref_raw and pref_mem is None and op_today:
                            logging.info(
                                "??OP??: ???OP??????????? task=%s raw=%r",
                                task.get("task_id"),
                                pref_raw,
                            )

                        _gdp_must, _gdp_warns = _active_global_day_process_must_include(
                            _gpo,
                            task,
                            current_date,
                            capable_members,
                            members,
                        )
                        for _gw in _gdp_warns:
                            logging.warning(_gw)
                        fixed_team_anchor = _merge_global_day_process_and_pref_anchor(
                            _gdp_must, pref_mem, capable_members
                        )
                        if _gdp_must:
                            logging.info(
                                "????????(??�??): task=%s date=%s ??=%r ?????=%s",
                                task.get("task_id"),
                                current_date,
                                machine,
                                ",".join(_gdp_must),
                            )
                        if fixed_team_anchor:
                            _nfix = len(fixed_team_anchor)
                            if _nfix > req_num:
                                need_src_line = (
                                    (need_src_line + " ? ")
                                    if need_src_line
                                    else ""
                                )
                                need_src_line += (
                                    f"?????(??�??)?????{_nfix}?"
                                )
                            req_num = max(req_num, _nfix)
    
                        extra_max_sheet, extra_src_line = resolve_need_surplus_extra_max_explain(
                            machine,
                            machine_name,
                            task["task_id"],
                            surplus_map,
                            need_rules,
                        )
                        if TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW:
                            extra_max_sheet = 0
                            extra_src_line = (
                                (extra_src_line + " ? ")
                                if extra_src_line
                                else ""
                            ) + "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW?0"
                        extra_max = (
                            extra_max_sheet
                            if TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
                            else 0
                        )
                        if (
                            extra_max_sheet > 0
                            and not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
                            and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
                        ):
                            extra_src_line = (
                                (extra_src_line + " ? ")
                                if extra_src_line
                                else ""
                            ) + "???????????????????????�???????"
                        max_team_size = min(req_num + extra_max, len(capable_members))
                        if max_team_size < req_num:
                            max_team_size = req_num
                        rq_base = max(1, int(req_num))
    
                        _dto_head = task.get("dispatch_trial_order")
                        if (
                            _dto_head is not None
                            and _dto_head not in _need_headcount_logged_orders
                        ):
                            _need_headcount_logged_orders.add(_dto_head)
                            logging.info(
                                "need??(???????) order=%s task=%s ??/??=%s/%s "
                                "req_num=%s [%s] extra_max=%s [%s] max_team??=%s capable=%s?",
                                _dto_head,
                                task["task_id"],
                                machine,
                                machine_name,
                                req_num,
                                need_src_line,
                                extra_max,
                                extra_src_line,
                                max_team_size,
                                len(capable_members),
                            )
    
                        trace_assign = bool(TRACE_TEAM_ASSIGN_TASK_ID) and (
                            str(task.get("task_id", "")).strip() == TRACE_TEAM_ASSIGN_TASK_ID
                        )
                        if trace_assign:
                            logging.info(
                                "TRACE??[%s] %s ??/??=%s / %s req_num=%s extra_max=%s ? max_team=%s "
                                "capable(n=%s)=%s ignore_need1=%s ignore_skill=%s abolish=%s ??OP??=%r?%s",
                                task["task_id"],
                                current_date,
                                machine,
                                machine_name,
                                req_num,
                                extra_max,
                                max_team_size,
                                len(capable_members),
                                capable_members,
                                global_priority_override.get("ignore_need_minimum"),
                                global_priority_override.get("ignore_skill_requirements"),
                                global_priority_override.get("abolish_all_scheduling_limits"),
                                pref_raw,
                                pref_mem,
                            )
    
                        team_candidates: list[dict] = []
                        combo_key = (
                            f"{machine_proc}+{machine_name}"
                            if machine_proc and machine_name
                            else ""
                        )
                        preset_rows = (
                            (team_combo_presets or {}).get(combo_key)
                            if (team_combo_presets and combo_key)
                            else None
                        )
                        (
                            _mach_floor_legacy,
                            _co_segs_legacy,
                            _abort_legacy,
                        ) = _resolve_machine_changeover_floor_segments(
                            abolish_all_scheduling_limits=bool(
                                _gpo.get("abolish_all_scheduling_limits")
                            ),
                            machine_occ_key=machine_occ_key,
                            task_id=str(task.get("task_id") or "").strip(),
                            eq_line=eq_line,
                            machine_name=str(task.get("machine_name") or "").strip(),
                            machine_proc=machine_proc,
                            machine_avail_dt=machine_avail_dt,
                            machine_day_floor=_machine_day_start,
                            current_date=current_date,
                            machine_handoff=machine_handoff_legacy,
                            daily_status=daily_status,
                            skills_dict=skills_dict,
                            dispatch_interval_mirror=_dispatch_interval_mirror,
                        )
                        if _abort_legacy:
                            continue
                        # ?????????????????????????????????????
                        if preset_rows:
                            for _prio, sheet_rs, preset_team, combo_row_id in preset_rows:
                                pteam = tuple(preset_team)
                                bounds = _combo_preset_team_size_bounds(
                                    pteam, sheet_rs, max_team_size
                                )
                                if bounds is None:
                                    continue
                                if fixed_team_anchor and not all(
                                    m in pteam for m in fixed_team_anchor
                                ):
                                    continue
                                if pref_mem is not None and pref_mem not in pteam:
                                    continue
                                if not all(m in capable_members for m in pteam):
                                    continue
                                _append_legacy_dispatch_candidate_for_team(
                                    task,
                                    pteam,
                                    avail_dt,
                                    machine_avail_dt,
                                    daily_status,
                                    current_date,
                                    macro_run_date,
                                    macro_now_dt,
                                    skill_role_priority,
                                    eq_line,
                                    rq_base,
                                    extra_max,
                                    global_priority_override,
                                    team_candidates,
                                    combo_sheet_row_id=combo_row_id,
                                    combo_preset_team=pteam,
                                    dispatch_interval_mirror=_dispatch_interval_mirror,
                                    machine_handoff=machine_handoff_legacy,
                                    machine_day_floor=_machine_day_start,
                                    machine_floor_cached=_mach_floor_legacy,
                                )
    
                        for tsize in range(req_num, max_team_size + 1):
                            if fixed_team_anchor:
                                _ft = list(fixed_team_anchor)
                                others = [m for m in capable_members if m not in _ft]
                                need_extra = tsize - len(_ft)
                                if need_extra < 0:
                                    teams_iter = []
                                elif need_extra == 0:
                                    teams_iter = [tuple(_ft)]
                                elif len(others) >= need_extra:
                                    teams_iter = [
                                        tuple(_ft + list(rest))
                                        for rest in itertools.combinations(
                                            others, need_extra
                                        )
                                    ]
                                else:
                                    teams_iter = []
                            elif (
                                pref_mem is not None
                                and pref_mem in capable_members
                                and skill_role_priority(pref_mem)[0] == "OP"
                            ):
                                others = [m for m in capable_members if m != pref_mem]
                                if tsize == 1:
                                    teams_iter = [(pref_mem,)]
                                elif len(others) >= tsize - 1:
                                    teams_iter = [
                                        tuple([pref_mem] + list(rest))
                                        for rest in itertools.combinations(others, tsize - 1)
                                    ]
                                else:
                                    logging.info(
                                        "??OP??: ?????????????????? task=%s size=%s raw=%r",
                                        task.get("task_id"),
                                        tsize,
                                        pref_raw,
                                    )
                                    teams_iter = itertools.combinations(capable_members, tsize)
                            else:
                                teams_iter = itertools.combinations(capable_members, tsize)
    
                            for team in teams_iter:
                                op_list = [m for m in team if skill_role_priority(m)[0] == "OP"]
                                if not op_list:
                                    continue
    
                                team_start = max(avail_dt[m] for m in team)
                                if not _gpo.get("abolish_all_scheduling_limits"):
                                    # ?????1???1???????????????/????????????
                                    machine_free_dt = _mach_floor_legacy
                                    if team_start < machine_free_dt:
                                        team_start = machine_free_dt
                                    # ???????????? 13:00 ???????????????
                                    if task.get("same_day_raw_start_limit") and current_date == task["start_date_req"]:
                                        min_start_dt = datetime.combine(
                                            current_date, task["same_day_raw_start_limit"]
                                        )
                                        if team_start < min_start_dt:
                                            team_start = min_start_dt
                                    if current_date == task["start_date_req"] and task.get("earliest_start_time"):
                                        min_user_t = datetime.combine(
                                            current_date, task["earliest_start_time"]
                                        )
                                        if team_start < min_user_t:
                                            team_start = min_user_t
                                    # ????????????????????????
                                    if current_date == macro_run_date and team_start < macro_now_dt:
                                        team_start = macro_now_dt
                                team_end_limit = min(daily_status[m]['end_dt'] for m in team)
    
                                if team_start >= team_end_limit:
                                    continue
    
                                team_breaks = []
                                for m in team:
                                    team_breaks.extend(daily_status[m]['breaks_dt'])
                                team_breaks = merge_time_intervals(team_breaks)

                                avg_eff = sum(daily_status[m]['efficiency'] for m in team) / len(team)
                                if avg_eff <= 0:
                                    avg_eff = 0.01
                                t_eff = parse_float_safe(task.get("task_eff_factor"), 1.0)
                                if t_eff <= 0:
                                    t_eff = 1.0
                                eff_time_per_unit = (
                                    task["base_time_per_unit"]
                                    / avg_eff
                                    / t_eff
                                    * _surplus_team_time_factor(rq_base, len(team), extra_max)
                                )
                                _defer_min_contig = max(1, int(math.ceil(float(eff_time_per_unit))))
    
                                def _refloor_legacy_inline(ts):
                                    ts = max(ts, max(avail_dt[m] for m in team))
                                    if not _gpo.get("abolish_all_scheduling_limits"):
                                        _mfd = _mach_floor_legacy
                                        if ts < _mfd:
                                            ts = _mfd
                                        if task.get(
                                            "same_day_raw_start_limit"
                                        ) and current_date == task["start_date_req"]:
                                            _msd = datetime.combine(
                                                current_date,
                                                task["same_day_raw_start_limit"],
                                            )
                                            if ts < _msd:
                                                ts = _msd
                                        if current_date == task[
                                            "start_date_req"
                                        ] and task.get("earliest_start_time"):
                                            _mut = datetime.combine(
                                                current_date,
                                                task["earliest_start_time"],
                                            )
                                            if ts < _mut:
                                                ts = _mut
                                        if (
                                            current_date == macro_run_date
                                            and ts < macro_now_dt
                                        ):
                                            ts = macro_now_dt
                                    return ts
    
                                _ts_adj = _defer_team_start_past_prebreak_and_end_of_day(
                                    task,
                                    tuple(team),
                                    team_start,
                                    team_end_limit,
                                    team_breaks,
                                    _refloor_legacy_inline,
                                    min_contiguous_work_mins=_defer_min_contig,
                                )
                                if _ts_adj is None:
                                    continue
                                team_start = _ts_adj
                                if team_start >= team_end_limit:
                                    continue
    
                                _, avail_mins, _ = calculate_end_time(team_start, 9999, team_breaks, team_end_limit)
    
                                units_can_do = int(avail_mins / eff_time_per_unit)
                                if units_can_do == 0:
                                    continue
    
                                units_today = min(units_can_do, math.ceil(task['remaining_units']))
                                if _eod_reject_capacity_units_below_threshold(
                                    units_today, team_start, team_end_limit
                                ):
                                    continue
                                work_mins_needed = int(units_today * eff_time_per_unit)
                                if (
                                    _contiguous_work_minutes_until_next_break_or_limit(
                                        team_start, team_breaks, team_end_limit
                                    )
                                    < work_mins_needed
                                ):
                                    continue
                                actual_end_dt, _, _ = calculate_end_time(team_start, work_mins_needed, team_breaks, team_end_limit)
    
                                team_prio_sum = sum(skill_role_priority(m)[1] for m in team)
                                if (
                                    _dispatch_interval_mirror is not None
                                    and _dispatch_interval_mirror.would_block_roll(
                                        machine_occ_key,
                                        team,
                                        team_start,
                                        actual_end_dt,
                                    )
                                ):
                                    continue
                                team_candidates.append(
                                    {
                                        "team": team,
                                        "team_start": team_start,
                                        "actual_end_dt": actual_end_dt,
                                        "units_today": units_today,
                                        "team_breaks": team_breaks,
                                        "avg_eff": avg_eff,
                                        "prio_sum": team_prio_sum,
                                        "op_list": op_list,
                                        "eff_time_per_unit": eff_time_per_unit,
                                        "combo_sheet_row_id": None,
                                        "combo_preset_team": None,
                                    }
                                )
    
                        best_team = None
                        best_info = {
                            "start_dt": datetime.max,
                            "units_today": 0,
                            "prio_sum": 10**9,
                        }
                        t_min = (
                            min(c["team_start"] for c in team_candidates)
                            if team_candidates
                            else None
                        )
    
                        def _team_cand_key(c):
                            return _team_assignment_sort_tuple(
                                c["team"],
                                c["team_start"],
                                c["units_today"],
                                c["prio_sum"],
                                t_min,
                            )
    
                        if team_candidates:
                            best_c = min(team_candidates, key=_team_cand_key)
                            if best_c.get("combo_sheet_row_id") is None and preset_rows:
                                _lcid_l = _lookup_combo_sheet_row_id_for_preset_team(
                                    preset_rows, tuple(best_c["team"])
                                )
                                if _lcid_l is not None:
                                    best_c = {
                                        **best_c,
                                        "combo_sheet_row_id": _lcid_l,
                                        "combo_preset_team": tuple(best_c["team"]),
                                    }
                            if pref_mem and pref_mem in best_c["op_list"]:
                                lead_op = pref_mem
                            else:
                                lead_op = min(
                                    best_c["op_list"],
                                    key=lambda mm: (skill_role_priority(mm)[1], mm),
                                )
                            best_team = best_c["team"]
                            best_info = {
                                "start_dt": best_c["team_start"],
                                "end_dt": best_c["actual_end_dt"],
                                "op": lead_op,
                                "units_today": best_c["units_today"],
                                "breaks": best_c["team_breaks"],
                                "eff": best_c["avg_eff"],
                                "prio_sum": best_c["prio_sum"],
                            }
    
                        if trace_assign:
                            _tk = _team_assign_trace_tuple_label()
                            tid = task["task_id"]
                            for tsize in range(req_num, max_team_size + 1):
                                sub = [c for c in team_candidates if len(c["team"]) == tsize]
                                if not sub:
                                    logging.info(
                                        "TRACE??[%s] %s tsize=%s ? ??????????????",
                                        tid,
                                        current_date,
                                        tsize,
                                    )
                                else:
                                    sm = min(sub, key=_team_cand_key)
                                    logging.info(
                                        "TRACE??[%s] %s tsize=%s ?????: members=%s "
                                        "start=%s units_today=%s prio_sum=%s eff_t/unit=%.6f "
                                        "?????=%s ???????=%s ???????????????",
                                        tid,
                                        current_date,
                                        tsize,
                                        sm["team"],
                                        sm["team_start"],
                                        sm["units_today"],
                                        sm["prio_sum"],
                                        sm["eff_time_per_unit"],
                                        _tk,
                                        t_min.isoformat(sep=" ") if t_min else "�",
                                    )
    
                        if trace_assign and best_team is not None:
                            logging.info(
                                "TRACE??[%s] %s ??? n=%s members=%s start=%s units_today=%s prio_sum=%s",
                                task["task_id"],
                                current_date,
                                len(best_team),
                                best_team,
                                best_info["start_dt"],
                                best_info["units_today"],
                                best_info["prio_sum"],
                            )
                            if len(best_team) == 1 and max_team_size > req_num:
                                if TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF:
                                    logging.info(
                                        "TRACE??[%s] %s 1????TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF?: "
                                        "?????????????????OP???0?????>=?????",
                                        task["task_id"],
                                        current_date,
                                    )
                                else:
                                    logging.info(
                                        "TRACE??[%s] %s 1???: ?????????????"
                                        "????????????1??????????"
                                        "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=%s????????????????????",
                                        task["task_id"],
                                        current_date,
                                        TEAM_ASSIGN_START_SLACK_WAIT_MINUTES,
                                    )
    
                        if best_team:
                            if len(best_team) > req_num:
                                logging.info(
                                    "??????>req_num task=%s day=%s order=%s ??/??=%s/%s "
                                    "??=%s? req_num=%s extra_max=%s max_team=%s [%s] [%s]",
                                    task["task_id"],
                                    current_date,
                                    task.get("dispatch_trial_order"),
                                    machine,
                                    machine_name,
                                    len(best_team),
                                    req_num,
                                    extra_max,
                                    max_team_size,
                                    need_src_line,
                                    extra_src_line,
                                )
                            sub_members = [m for m in best_team if m != best_info["op"]]
                            done_units = best_info["units_today"]
                            if task.get("roll_pipeline_inspection") or task.get(
                                "roll_pipeline_rewind"
                            ):
                                _rp_room = _roll_pipeline_inspection_assign_room(
                                    task_queue, str(task.get("task_id", "")).strip()
                                )
                                done_units = min(
                                    int(done_units),
                                    int(min(_rp_room, math.ceil(task["remaining_units"]))),
                                )
                            else:
                                done_units = int(done_units)
                            if done_units <= 0:
                                if _trace_schedule_task_enabled(task.get("task_id")):
                                    _rp_log = None
                                    if task.get(
                                        "roll_pipeline_inspection"
                                    ) or task.get("roll_pipeline_rewind"):
                                        _rp_log = _roll_pipeline_inspection_assign_room(
                                            task_queue,
                                            str(task.get("task_id", "") or "").strip(),
                                        )
                                    _log_dispatch_trace_schedule(
                                        task.get("task_id"),
                                        "[?????? task=%s] ????: ?????????????0 "
                                        "day=%s machine=%s best_units_today=%s rp_room=%s rem=%.4f",
                                        task.get("task_id"),
                                        current_date,
                                        machine,
                                        best_info.get("units_today"),
                                        _rp_log,
                                        float(task.get("remaining_units") or 0),
                                    )
                                continue
                            if done_units < best_info["units_today"]:
                                team_end_limit = min(
                                    daily_status[m]["end_dt"] for m in best_team
                                )
                                _teff = parse_float_safe(task.get("task_eff_factor"), 1.0)
                                if _teff <= 0:
                                    _teff = 1.0
                                _eff_t = (
                                    task["base_time_per_unit"]
                                    / best_info["eff"]
                                    / _teff
                                    * _surplus_team_time_factor(rq_base, len(best_team), extra_max)
                                )
                                _wm = int(done_units * _eff_t)
                                _end_dt, _, _ = calculate_end_time(
                                    best_info["start_dt"],
                                    _wm,
                                    best_info["breaks"],
                                    team_end_limit,
                                )
                                best_info = dict(best_info)
                                best_info["end_dt"] = _end_dt
                                best_info["units_today"] = done_units
    
                            total_u = math.ceil(task['total_qty_m'] / task['unit_m']) if task['unit_m'] else 0
                            rem_u_before = math.ceil(task['remaining_units'])
                            already_done = total_u - rem_u_before
                            
                            # ????????????????????????????????
                            try:
                                tot_qty = parse_float_safe(task.get('total_qty_m'), 0.0)
                                done_qty = parse_float_safe(task.get('done_qty_reported'), 0.0)
                                if tot_qty > 0:
                                    pct_macro = max(0, min(100, int(round((done_qty / tot_qty) * 100))))
                                else:
                                    pct_macro = 0
                            except Exception:
                                pct_macro = 0
                            
                            _te_disp = parse_float_safe(task.get("task_eff_factor"), 1.0)
                            if _te_disp <= 0:
                                _te_disp = 1.0
                            _legacy_mach_sub = ", ".join(
                                str(s).strip()
                                for s in sub_members
                                if s and str(s).strip()
                            )
                            _append_changeover_segments_to_timeline(
                                timeline_events,
                                _dispatch_interval_mirror,
                                avail_dt,
                                daily_status,
                                current_date=current_date,
                                task_id=str(task.get("task_id") or ""),
                                machine_occ_key=machine_occ_key,
                                segments=list(_co_segs_legacy or []),
                                machining_lead_op=str(
                                    best_info.get("op") or ""
                                ).strip()
                                or None,
                                machining_sub_str=_legacy_mach_sub or None,
                                machine_handoff=machine_handoff_legacy,
                            )
                            timeline_events.append({
                                "date": current_date, "task_id": task['task_id'], "machine": eq_line,
                                "machine_occupancy_key": machine_occ_key,
                                "op": best_info["op"], "sub": ", ".join(sub_members),
                                "start_dt": best_info["start_dt"], "end_dt": best_info["end_dt"],
                                "breaks": best_info["breaks"], "units_done": done_units,
                                "already_done_units": already_done,
                                "total_units": total_u,
                                "pct_macro": pct_macro,
                                "eff_time_per_unit": task["base_time_per_unit"]
                                / best_info["eff"]
                                / _te_disp
                                * _surplus_team_time_factor(rq_base, len(best_team), extra_max),
                                "unit_m": task['unit_m'],
                                "event_kind": TIMELINE_EVENT_MACHINING,
                            })
                            if _dispatch_interval_mirror is not None:
                                _dispatch_interval_mirror.register_from_event(
                                    timeline_events[-1]
                                )
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _rp_tr = None
                                if task.get("roll_pipeline_inspection") or task.get(
                                    "roll_pipeline_rewind"
                                ):
                                    _rp_tr = _roll_pipeline_inspection_assign_room(
                                        task_queue,
                                        str(task.get("task_id", "") or "").strip(),
                                    )
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[?????? task=%s] ???????? chunk day=%s machine=%s "
                                    "done_units=%s already_done=%s total_u=%s rem_after=%.4f "
                                    "start=%s end=%s eff_t/unit=%.4f rp_room(??)=%s",
                                    task.get("task_id"),
                                    current_date,
                                    eq_line,
                                    done_units,
                                    already_done,
                                    total_u,
                                    float(task.get("remaining_units") or 0)
                                    - float(done_units),
                                    best_info["start_dt"],
                                    best_info["end_dt"],
                                    float(
                                        task["base_time_per_unit"]
                                        / best_info["eff"]
                                        / _te_disp
                                        * _surplus_team_time_factor(
                                            rq_base, len(best_team), extra_max
                                        )
                                    ),
                                    _rp_tr,
                                )

                            task['remaining_units'] -= done_units
                            op_main = (best_info.get("op") or "").strip()
                            subs_part = ",".join(
                                s.strip() for s in sub_members if s and str(s).strip()
                            )
                            team_s = f"{op_main}, {subs_part}" if subs_part else op_main
                            need_surplus_assigned = (
                                TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
                                and extra_max > 0
                                and len(best_team) > req_num
                            )
                            _lo = (best_info.get("op") or "").strip()
                            _subs_legacy = [
                                str(s).strip()
                                for s in sub_members
                                if s and str(s).strip()
                            ]
                            _names_ord = ([] if not _lo else [_lo]) + _subs_legacy
                            _surplus_names = (
                                _names_ord[int(req_num) :]
                                if need_surplus_assigned
                                and len(_names_ord) > int(req_num)
                                else []
                            )
                            task["assigned_history"].append(
                                {
                                    "date": current_date.strftime("%m/%d"),
                                    "team": team_s,
                                    "done_m": int(done_units * task["unit_m"]),
                                    "start_dt": best_info["start_dt"],
                                    "end_dt": best_info["end_dt"],
                                    "need_surplus_assigned": need_surplus_assigned,
                                    "combo_sheet_row_id": best_c.get(
                                        "combo_sheet_row_id"
                                    ),
                                    "surplus_member_names": _surplus_names,
                                }
                            )

                            for m in best_team:
                                avail_dt[m] = best_info["end_dt"]
                            if not _gpo.get("abolish_all_scheduling_limits"):
                                machine_avail_dt[machine_occ_key] = best_info["end_dt"]
                                _bump_machine_avail_after_roll_for_calendar(
                                    current_date,
                                    machine_occ_key,
                                    machine_avail_dt,
                                    machine_calendar_plan_end=_machine_calendar_plan_end,
                                    machine_day_floor=_machine_day_start,
                                )
                            machine_handoff_legacy["last_tid"][machine_occ_key] = str(
                                task.get("task_id") or ""
                            ).strip()
                            machine_handoff_legacy["last_eq"][machine_occ_key] = eq_line
                            machine_handoff_legacy["started_today"].add(
                                machine_occ_key
                            )
                            machine_handoff_legacy["machining_today_occ"].add(
                                machine_occ_key
                            )
                            machine_handoff_legacy["last_machining_dt"][
                                machine_occ_key
                            ] = best_info["end_dt"]
                            machine_handoff_legacy["last_machining_date"][
                                machine_occ_key
                            ] = current_date
                            machine_handoff_legacy["last_lead_op"][
                                machine_occ_key
                            ] = best_info["op"]
                            machine_handoff_legacy.setdefault("last_machining_sub", {})
                            machine_handoff_legacy["last_machining_sub"][
                                machine_occ_key
                            ] = _legacy_mach_sub
                            if _trace_schedule_task_enabled(task.get("task_id")):
                                _log_dispatch_trace_schedule(
                                    task.get("task_id"),
                                    "[?????? task=%s] ????? ??? day=%s machine=%s machine_name=%s "
                                    "start=%s end=%s ????=%s req_num=%s ?????extra_max=%s "
                                    "??????(???)=%s team=%s",
                                    task.get("task_id"),
                                    current_date,
                                    eq_line,
                                    str(machine_name or "").strip(),
                                    best_info["start_dt"],
                                    best_info["end_dt"],
                                    len(best_team),
                                    int(req_num),
                                    int(extra_max),
                                    need_surplus_assigned,
                                    team_s,
                                )
                            _sched_made_progress = True
                        else:
                            if task.get("has_done_deadline_override"):
                                logging.info(
                                    "DEBUG[?????] ??NO=%s ??=%s ????????/???????????????remaining_units=%s",
                                    task.get("task_id"),
                                    current_date,
                                    task.get("remaining_units"),
                                )
    
                if not _sched_made_progress:
                    break

            if TRACE_SCHEDULE_TASK_IDS:
                for _tt in TRACE_SCHEDULE_TASK_IDS:
                    for _t in task_queue:
                        if str(_t.get("task_id", "")).strip() != _tt:
                            continue
                        _rem_tr = float(_t.get("remaining_units") or 0)
                        if _rem_tr <= 1e-9:
                            continue
                        _log_dispatch_trace_schedule(
                            _tt,
                            "[?????? task=%s] ???????? day=%s machine=%s "
                            "machine_name=%s rem=%.4f roll_b2_follower=%s ???=%s",
                            _tt,
                            current_date,
                            _t.get("machine"),
                            _t.get("machine_name"),
                            _rem_tr,
                            bool(
                                _t.get("roll_pipeline_inspection")
                                or _t.get("roll_pipeline_rewind")
                            ),
                            _t.get("dispatch_trial_order"),
                        )

            if STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING:
                missed_tids = _collect_task_ids_missed_deadline_after_day(
                    task_queue, current_date
                )
                if missed_tids:
                    blocked_tids = set()
                    shift_tid_list = []
                    for _ptid in sorted(missed_tids):
                        _do_shift, _cal_short = _partial_task_id_due_shift_outcome(
                            task_queue, _ptid, _calendar_last_plan_day
                        )
                        if _cal_short:
                            blocked_tids.add(_ptid)
                        if _do_shift:
                            shift_tid_list.append(_ptid)
                    for t in task_queue:
                        _tid = str(t.get("task_id", "") or "").strip()
                        if _tid in blocked_tids:
                            t["_partial_retry_calendar_blocked"] = True
                    if shift_tid_list:
                        allowed_shift_tids = [
                            tid
                            for tid in shift_tid_list
                            if _due_shift_retry_count_by_request.get(tid, 0)
                            < STAGE2_RETRY_SHIFT_DUE_MAX_ROUNDS
                        ]
                        for tid in shift_tid_list:
                            if tid not in allowed_shift_tids:
                                _due_shift_exhausted_requests.add(tid)
                        if allowed_shift_tids:
                            _outer_retry_round += 1
                            for tid in allowed_shift_tids:
                                _due_shift_retry_count_by_request[tid] = (
                                    _due_shift_retry_count_by_request.get(tid, 0) + 1
                                )
                            shift_set = set(allowed_shift_tids)
                            for t in task_queue:
                                if str(t.get("task_id", "") or "").strip() in shift_set:
                                    _shift_task_due_calendar_fields_one_day(t, run_date)
                            timeline_events[:] = [
                                e
                                for e in timeline_events
                                if _normalize_timeline_task_id(e) not in shift_set
                            ]
                            if _dispatch_interval_mirror is not None:
                                _dispatch_interval_mirror.rebuild_from_timeline(
                                    timeline_events
                                )
                            for t in task_queue:
                                if str(t.get("task_id", "") or "").strip() in shift_set:
                                    t["remaining_units"] = float(
                                        t.get("initial_remaining_units") or 0
                                    )
                                    t["assigned_history"].clear()
                            _apply_dispatch_trial_order_for_generate_plan(
                                task_queue,
                                req_map,
                                need_rules,
                                need_combo_col_index,
                            )
                            _trials_detail = ",".join(
                                f"{tid}:{_due_shift_retry_count_by_request[tid]}"
                                for tid in sorted(allowed_shift_tids)
                            )
                            logging.info(
                                "????????: ????+1????????????????=%s ??NO=%s ?????????=%s?",
                                current_date.isoformat(),
                                ",".join(sorted(allowed_shift_tids)),
                                _trials_detail,
                            )
                            _full_calendar_without_deadline_restart = False
                            break
                        else:
                            # ???????????????????????????????break ???????????????????????
                            _cap_tids = sorted(
                                tid
                                for tid in shift_tid_list
                                if tid not in allowed_shift_tids
                            )
                            _first_cap_warn = [
                                tid for tid in _cap_tids if tid not in _due_shift_cap_warned_tids
                            ]
                            for tid in _first_cap_warn:
                                _due_shift_cap_warned_tids.add(tid)
                            if _first_cap_warn:
                                logging.warning(
                                    "?????????: ????NO?????????? %s ??????????? +1 ?????"
                                    " ??????????????????????????????????: %s",
                                    STAGE2_RETRY_SHIFT_DUE_MAX_ROUNDS,
                                    ",".join(_cap_tids),
                                )

        if _full_calendar_without_deadline_restart:
            _rewind_made = _run_b2_inspection_rewind_pass(
                sorted_dates,
                attendance_data,
                task_queue,
                timeline_events,
                skills_dict,
                members,
                equipment_list,
                req_map,
                need_rules,
                surplus_map,
                global_priority_override,
                macro_run_date,
                macro_now_dt,
                _need_headcount_logged_orders,
                team_combo_presets,
                dispatch_interval_mirror=_dispatch_interval_mirror,
            )
            if _rewind_made:
                logging.info(
                    "�B-2/�B-3 ?????: EC ??????????????????????????timeline_events ??????????????"
                )
            break

    if TRACE_SCHEDULE_TASK_IDS:
        for _tt in TRACE_SCHEDULE_TASK_IDS:
            for _t in task_queue:
                if str(_t.get("task_id", "")).strip() != _tt:
                    continue
                _log_dispatch_trace_schedule(
                    _tt,
                    "[?????? task=%s] ??????????? machine=%s machine_name=%s "
                    "rem=%.4f initial=%.4f roll_b2_follower=%s",
                    _tt,
                    _t.get("machine"),
                    _t.get("machine_name"),
                    float(_t.get("remaining_units") or 0),
                    float(_t.get("initial_remaining_units") or 0),
                    bool(
                        _t.get("roll_pipeline_inspection")
                        or _t.get("roll_pipeline_rewind")
                    ),
                )
            _evs_tr = sorted(
                (
                    e
                    for e in timeline_events
                    if str(e.get("task_id", "")).strip() == _tt
                ),
                key=lambda e: (e.get("date"), e.get("start_dt") or datetime.min),
            )
            _last_ev_by_machine: dict = {}
            for _e in _evs_tr:
                _last_ev_by_machine[str(_e.get("machine") or "")] = _e
            for _mk, _ev in sorted(_last_ev_by_machine.items()):
                _ad = int(_ev.get("already_done_units") or 0)
                _ud = int(_ev.get("units_done") or 0)
                _log_dispatch_trace_schedule(
                    _tt,
                    "[?????? task=%s] ?????????(?????) machine=%s "
                    "already_done+units_done=%s+%s=%s total_units=%s end_dt=%s",
                    _tt,
                    _mk,
                    _ad,
                    _ud,
                    _ad + _ud,
                    _ev.get("total_units"),
                    _ev.get("end_dt"),
                )

    # ???????????????need ???????TEMP_?????????
    timeline_before_need_surplus = copy.deepcopy(timeline_events)

    # need?????????: ???????????�????????????????
    if (
        not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
        and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
        and surplus_map
        and timeline_events
    ):
        _n_sur = append_surplus_staff_after_main_dispatch(
            timeline_events,
            attendance_data,
            skills_dict,
            members,
            task_queue,
            req_map,
            need_rules,
            surplus_map,
            global_priority_override,
        )
        if _n_sur:
            logging.info(
                "need??: ??????????? %s ????????�????????????",
                _n_sur,
            )

    if _dispatch_interval_mirror is not None:
        _dispatch_interval_mirror.rebuild_from_timeline(timeline_events)

    # ???????????????????????????????????????????????
    for e in timeline_events:
        e["subs_list"] = [s.strip() for s in e["sub"].split(",")] if e.get("sub") else []

    events_by_date = defaultdict(list)
    for e in timeline_events:
        events_by_date[e["date"]].append(e)

    # =========================================================
    # 4. Excel?? (?????)
    # =========================================================
    _remove_prior_stage2_workbooks_and_prune_empty_dirs(output_dir)
    # ????????????????????????????????
    _stage2_out_stamp = base_now_dt.strftime("%Y%m%d_%H%M%S_%f")
    output_filename = os.path.join(
        output_dir, f"production_plan_multi_day_{_stage2_out_stamp}.xlsx"
    )
    # ???ID ? ??_???????????????????????? B12????_???????????
    first_eq_schedule_cell_by_task_id: dict[str, str] = {}
    df_eq_schedule = _build_equipment_schedule_dataframe(
        sorted_dates,
        equipment_list,
        attendance_data,
        timeline_events,
        first_eq_schedule_cell_by_task_id=first_eq_schedule_cell_by_task_id,
    )
    df_temp_equipment_schedule = _build_equipment_schedule_dataframe(
        sorted_dates,
        equipment_list,
        attendance_data,
        timeline_before_need_surplus,
        first_eq_schedule_cell_by_task_id=None,
    )
    df_block_table = _build_block_table_dataframe(
        sorted_dates,
        equipment_list,
        members,
        attendance_data,
        timeline_events,
    )
    df_equipment_by_machine_name = _build_equipment_schedule_by_machine_name_dataframe(
        sorted_dates,
        equipment_list,
        attendance_data,
        timeline_events,
    )

    # ??_??????: ??????????????????????????timeline_events ???
    plan_window_by_task_id: dict = {}
    for _ev in timeline_events:
        tid = _ev.get("task_id")
        if tid is None:
            continue
        sd = _ev.get("start_dt")
        ed = _ev.get("end_dt")
        if sd is None or ed is None:
            continue
        if tid not in plan_window_by_task_id:
            plan_window_by_task_id[tid] = [sd, ed]
        else:
            w = plan_window_by_task_id[tid]
            if sd < w[0]:
                w[0] = sd
            if ed > w[1]:
                w[1] = ed

    # ??_???????????????????????_??????????????
    # ??????????????????????????????????????????????????
    _result_sheet_answer_spec_by_line = {}
    _result_sheet_raw_input_by_line: dict = {}
    if tasks_df is not None and not getattr(tasks_df, "empty", True):
        for _, _r in tasks_df.iterrows():
            if _plan_row_exclude_from_assignment(_r):
                continue
            _tid = str(_planning_df_cell_scalar(_r, TASK_COL_TASK_ID) or "").strip()
            _mach = str(_planning_df_cell_scalar(_r, TASK_COL_MACHINE) or "").strip()
            if not _tid or not _mach:
                continue
            _ad = parse_optional_date(_planning_df_cell_scalar(_r, TASK_COL_ANSWER_DUE))
            _sd = parse_optional_date(_planning_df_cell_scalar(_r, TASK_COL_SPECIFIED_DUE))
            _rid = parse_optional_date(_planning_df_cell_scalar(_r, TASK_COL_RAW_INPUT_DATE))
            _rid_ov = parse_optional_date(
                _planning_df_cell_scalar(_r, PLAN_COL_RAW_INPUT_DATE_OVERRIDE)
            )
            if _rid_ov is not None:
                _rid = _rid_ov
            _result_sheet_answer_spec_by_line[(_tid, _mach)] = (_ad, _sd)
            _result_sheet_raw_input_by_line[(_tid, _mach)] = _rid

    task_results = []
    max_history_len = max([len(t['assigned_history']) for t in task_queue] + [0])
    
    # ???????????????????=???????=?????????=???
    # ????+1 ???????NO?????????????????????????????????
    sorted_tasks_for_result = sorted(task_queue, key=_result_task_sheet_sort_key)
    for t in sorted_tasks_for_result:
        rem_u = float(t.get("remaining_units") or 0)
        hist = bool(t.get("assigned_history"))
        if rem_u <= 1e-9:
            status = "???"
        elif hist and t.get("_partial_retry_calendar_blocked"):
            status = "???(?????????)"
        elif not hist and rem_u > 1e-9:
            status = "????"
        else:
            status = "???"
        _tid_res = str(t.get("task_id", "") or "").strip()
        if (
            _tid_res in _due_shift_exhausted_requests
            and rem_u > 1e-9
            and "???????" not in status
        ):
            status = f"{status}?????????"
        
        total_r = int(t['total_qty_m'] / t['unit_m']) if t['unit_m'] else 0
        rem_r = int(t['remaining_units'])
        
        _line_key = (str(t.get("task_id", "") or "").strip(), str(t.get("machine", "") or "").strip())
        _sheet_pair = _result_sheet_answer_spec_by_line.get(_line_key)
        if _sheet_pair is not None:
            _ans_d, _spec_d = _sheet_pair
            ans_s = _ans_d.strftime("%Y/%m/%d") if _ans_d else ""
            spec_s = _spec_d.strftime("%Y/%m/%d") if _spec_d else ""
        else:
            _ans_d = t.get("answer_due_date")
            _spec_d = t.get("specified_due_date")
            ans_s = _ans_d.strftime("%Y/%m/%d") if _ans_d else ""
            spec_s = _spec_d.strftime("%Y/%m/%d") if _spec_d else ""
        _basis_for_sheet = t.get("due_basis_date_result_sheet")
        if _basis_for_sheet is None:
            _basis_for_sheet = t.get("due_basis_date")
        basis_s = (
            _basis_for_sheet.strftime("%Y/%m/%d")
            if _basis_for_sheet is not None and hasattr(_basis_for_sheet, "strftime")
            else ""
        )
        if _line_key in _result_sheet_raw_input_by_line:
            _rid_d = _result_sheet_raw_input_by_line[_line_key]
            kenhan_s = _rid_d.strftime("%Y/%m/%d") if _rid_d else ""
        else:
            _rid_t = t.get("raw_input_date")
            kenhan_s = (
                _rid_t.strftime("%Y/%m/%d")
                if _rid_t is not None and hasattr(_rid_t, "strftime")
                else ""
            )
        start_req = t["start_date_req"]
        start_req_s = start_req.strftime("%Y/%m/%d") if hasattr(start_req, "strftime") else str(start_req)
        rov = t.get("required_op")
        # ??: A=????? ? ???ID/??/??/??? ? ??1..n ? ??? ? ???????_AI
        row_status = {"?????": status}
        _dto = t.get("dispatch_trial_order")
        row_core = {
            "???ID": t['task_id'],
            "???": t['machine'],
            "???": t.get("machine_name", ""),
            "???": t.get("priority", 999),
            RESULT_TASK_COL_DISPATCH_TRIAL_ORDER: _dto if _dto is not None else "",
        }
        row_history = {}
        for i in range(max_history_len):
            if i < len(t['assigned_history']):
                h = t['assigned_history'][i]
                row_history[f"??{i+1}"] = _format_result_task_history_cell(t, h)
            else:
                row_history[f"??{i+1}"] = ""

        try:
            tot_qty = parse_float_safe(t.get("total_qty_m"), 0.0)
            done_qty = parse_float_safe(t.get("done_qty_reported"), 0.0)
            pct_macro = max(0, min(100, int(round((done_qty / tot_qty) * 100)))) if tot_qty > 0 else 0
        except Exception:
            pct_macro = 0

        _pw = plan_window_by_task_id.get(t["task_id"])
        if _pw:
            _ps, _pe = _pw[0], _pw[1]
            plan_assign_start_s = (
                _ps.strftime("%Y/%m/%d %H:%M") if hasattr(_ps, "strftime") else ""
            )
            plan_assign_end_s = (
                _pe.strftime("%Y/%m/%d %H:%M") if hasattr(_pe, "strftime") else ""
            )
        else:
            plan_assign_start_s = ""
            plan_assign_end_s = ""

        _plan_end_ans_spec16 = _result_task_plan_end_within_answer_or_spec_16_label(
            _pw, _ans_d, _spec_d
        )

        row_tail = {
            "??OP(??)": rov if rov is not None else "",
            "?????": parse_float_safe(t.get("task_eff_factor"), 1.0),
            "????": "??" if t.get("in_progress") else "???",
            "??????": "??" if t.get("has_special_remark") else "???",
            "??OP??": (t.get("preferred_operator_raw") or "")[:120],
            "????": ans_s,
            "????": spec_s,
            "??????": basis_s,
            TASK_COL_RAW_INPUT_DATE: kenhan_s,
            "????": "??" if t.get("due_urgent") else "???",
            "?????": start_req_s,
            "??_????": plan_assign_start_s,
            "??_????": plan_assign_end_s,
            RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16: _plan_end_ans_spec16,
            "????": f"{total_r}R ({t['total_qty_m']}m)",
            "????": f"{rem_r}R ({int(t['remaining_units'] * t['unit_m'])}m)",
            "???(????)": f"{pct_macro}%",
        }
        row_ai_last = {"????_AI": (t.get("task_special_ai_note") or "")[:300]}
        row_data = {**row_status, **row_core, **row_history, **row_tail, **row_ai_last}
        task_results.append(row_data)
        
    cal_rows = []
    for d in sorted_dates:
        for m in members:
            if m in attendance_data[d]:
                data = attendance_data[d][m]
                if data['is_working']:
                    cal_end = _calendar_display_clock_out_for_calendar_sheet(data, d)
                    end_disp = cal_end if cal_end is not None else data['end_dt']
                    clock_out_s = end_disp.strftime("%H:%M")
                else:
                    clock_out_s = "?"
                cal_rows.append({
                    "??": d,
                    "????": m,
                    "??": data['start_dt'].strftime("%H:%M") if data['is_working'] else "?",
                    "??": clock_out_s,
                    "??": data['efficiency'],
                    "??": data['reason'],
                })

    utilization_data = []
    for d in sorted_dates:
        row_data = {"???": d.strftime("%Y/%m/%d (%a)")}
        # ?????????????????????????????�????????????????
        member_worked_mins = defaultdict(int)
        for ev in events_by_date[d]:
            mins = get_actual_work_minutes(ev["start_dt"], ev["end_dt"], ev["breaks"])
            member_worked_mins[ev["op"]] += mins
            for s in ev["subs_list"]:
                if s:
                    member_worked_mins[s] += mins
        for m in members:
            if m in attendance_data[d] and attendance_data[d][m]['is_working']:
                default_start = datetime.combine(d, DEFAULT_START_TIME)
                default_end = datetime.combine(d, DEFAULT_END_TIME)
                
                actual_start = attendance_data[d][m]['start_dt']
                actual_end = attendance_data[d][m]['end_dt']
                clip_start = max(actual_start, default_start)
                clip_end = min(actual_end, default_end)
                
                if clip_start >= clip_end:
                    total_avail_mins = 0
                else:
                    breaks_dt = attendance_data[d][m]['breaks_dt']
                    total_avail_mins = get_actual_work_minutes(clip_start, clip_end, breaks_dt)
                
                if total_avail_mins <= 0:
                    row_data[m] = "0.0%"
                    continue
                
                worked_mins = member_worked_mins.get(m, 0)
                ratio = (worked_mins / total_avail_mins) * 100
                row_data[m] = f"{ratio:.1f}% ({worked_mins}/{total_avail_mins}?)"
            else:
                row_data[m] = "?"
        utilization_data.append(row_data)
        
    df_utilization = pd.DataFrame(utilization_data)

    df_mprio_legend, df_mprio_tbl = build_member_assignment_priority_reference(
        skills_dict, members
    )
    if df_mprio_tbl.empty:
        df_mprio_tbl = pd.DataFrame(
            [
                {
                    "???": "",
                    "???": "",
                    "??????": "",
                    "????": "",
                    "????": "",
                    "???": "",
                    "????_??????": "",
                    "skills???": "",
                    "??": "??? skills ?????+???????????????????????????",
                }
            ]
        )

    _usage_txt = build_gemini_usage_summary_text()
    if _usage_txt:
        ai_log_data["Gemini_??????????"] = _usage_txt[:50000]

    _master_abs_for_result_fmt = os.path.abspath(os.path.join(os.getcwd(), MASTER_FILE))
    _reg_shift_start, _reg_shift_end = _read_master_main_regular_shift_times(
        _master_abs_for_result_fmt
    )
    if _reg_shift_start is not None and _reg_shift_end is not None:
        logging.info(
            "???: master ??? A15/B15 ? %s ? %s???????????????",
            _reg_shift_start.strftime("%H:%M"),
            _reg_shift_end.strftime("%H:%M"),
        )

    logging.info(
        "??2: ??????????? ? %s",
        os.path.basename(output_filename),
    )
    try:
        with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
            df_eq_schedule.to_excel(
                writer, sheet_name=RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME, index=False
            )
            df_temp_equipment_schedule.to_excel(
                writer, sheet_name=TEMP_EQUIPMENT_SCHEDULE_SHEET_NAME, index=False
            )
            df_block_table.to_excel(writer, sheet_name=BLOCK_TABLE_SHEET_NAME, index=False)
            df_equipment_by_machine_name.to_excel(
                writer, sheet_name=RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME, index=False
            )
            pd.DataFrame(cal_rows).to_excel(writer, sheet_name='??_?????(???)', index=False)
            df_utilization.to_excel(writer, sheet_name='??_?????????', index=False)
            df_tasks = pd.DataFrame(task_results)
            df_tasks, task_column_order, _, vis_map = apply_result_task_sheet_column_order(
                df_tasks, max_history_len
            )
            seen_tc: set[str] = set()
            task_column_order_dedup: list = []
            vis_list_dedup: list = []
            for c in task_column_order:
                if c in seen_tc:
                    continue
                seen_tc.add(c)
                task_column_order_dedup.append(c)
                vis_list_dedup.append(bool(vis_map.get(c, True)))
            pd.DataFrame(
                {
                    "??": task_column_order_dedup,
                    "??": vis_list_dedup,
                }
            ).to_excel(writer, sheet_name=COLUMN_CONFIG_SHEET_NAME, index=False)
            df_tasks.to_excel(writer, sheet_name=RESULT_TASK_SHEET_NAME, index=False)
            pd.DataFrame(list(ai_log_data.items()), columns=["??", "??"]).to_excel(writer, sheet_name='??_AI??', index=False)

            _mprio_sheet = RESULT_MEMBER_PRIORITY_SHEET_NAME
            df_mprio_legend.to_excel(writer, sheet_name=_mprio_sheet, index=False)
            _mprio_gap = len(df_mprio_legend) + 2
            df_mprio_tbl.to_excel(
                writer, sheet_name=_mprio_sheet, index=False, startrow=_mprio_gap
            )

            logging.info(
                "??2: ??????????????????????????????????"
            )
            _write_results_equipment_gantt_sheet(
                writer,
                timeline_events,
                equipment_list,
                sorted_dates,
                attendance_data,
                data_extract_dt_str,
                base_now_dt,
                regular_shift_times=(_reg_shift_start, _reg_shift_end),
            )

            for sheet_name, ws_out in writer.sheets.items():
                if sheet_name == RESULT_SHEET_GANTT_NAME:
                    continue
                _apply_output_font_to_result_sheet(ws_out)

            if _reg_shift_start is not None and _reg_shift_end is not None:
                for _eq_sched_sheet in (
                    RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME,
                    TEMP_EQUIPMENT_SCHEDULE_SHEET_NAME,
                    RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME,
                ):
                    if _eq_sched_sheet in writer.sheets:
                        _apply_equipment_schedule_outside_regular_fill(
                            writer.sheets[_eq_sched_sheet],
                            _reg_shift_start,
                            _reg_shift_end,
                        )

            if RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME in writer.sheets:
                _apply_equipment_by_machine_dispatched_request_fill(
                    writer.sheets[RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME]
                )

            for _prep_sheet in (
                RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME,
                TEMP_EQUIPMENT_SCHEDULE_SHEET_NAME,
            ):
                if _prep_sheet in writer.sheets:
                    _apply_equipment_schedule_prep_cleanup_fill(
                        writer.sheets[_prep_sheet]
                    )

            if RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME in writer.sheets:
                _apply_equipment_schedule_machine_calendar_fill(
                    writer.sheets[RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME],
                    equipment_list,
                    _MACHINE_CALENDAR_BLOCKS_BY_DATE,
                )

            ws_cfg = writer.sheets[COLUMN_CONFIG_SHEET_NAME]
            _add_column_config_sheet_helpers(ws_cfg, len(task_column_order_dedup))

            worksheet_tasks = writer.sheets[RESULT_TASK_SHEET_NAME]
            max_col = worksheet_tasks.max_column
            for row in worksheet_tasks.iter_rows(min_row=1, max_row=worksheet_tasks.max_row, max_col=max_col):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical="top")

            _apply_result_task_sheet_column_visibility(
                worksheet_tasks, list(df_tasks.columns), vis_map
            )

            _apply_result_task_history_rich_text(worksheet_tasks, list(df_tasks.columns))
            _apply_result_task_date_columns_blue_font(worksheet_tasks, list(df_tasks.columns))

            # ????????????????????????
            status_col_idx = None
            for col_idx, col_name in enumerate(df_tasks.columns, 1):
                if str(col_name) == "?????":
                    status_col_idx = col_idx
                    break
            if status_col_idx is not None:
                unscheduled_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                for r in range(2, worksheet_tasks.max_row + 1):
                    st_val = worksheet_tasks.cell(row=r, column=status_col_idx).value
                    st = str(st_val).strip() if st_val is not None else ""
                    if st in ("????", "???"):
                        for c in range(1, max_col + 1):
                            worksheet_tasks.cell(row=r, column=c).fill = unscheduled_fill

            _apply_result_task_history_need_surplus_highlight(
                worksheet_tasks, list(df_tasks.columns), sorted_tasks_for_result
            )

            _apply_result_task_task_id_content_mismatch_highlight(
                worksheet_tasks, list(df_tasks.columns), sorted_tasks_for_result
            )
            _apply_result_task_plan_end_answer_spec_16_no_highlight(
                worksheet_tasks, list(df_tasks.columns)
            )
            _apply_result_task_id_hyperlinks_to_equipment_schedule(
                worksheet_tasks,
                list(df_tasks.columns),
                sorted_tasks_for_result,
                first_eq_schedule_cell_by_task_id,
                RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME,
            )

    except OSError as e:
        logging.error(
            "??2: ??????????????????: %s?%s??"
            "output ?? production_plan_multi_day_*.xlsx ? Excel ?????????????????",
            output_filename,
            e,
        )
        raise

    try:
        _apply_excel_date_columns_date_only_display(
            output_filename, "??_?????(???)", frozenset({"??"})
        )
    except Exception as e:
        logging.warning(f"??_?????(???)????????: {e}")

    _stage2_try_copy_column_config_shapes_from_input(
        output_filename,
        (os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK),
    )

    logging.info(f"??: '{output_filename}' ????????")

    # =========================================================
    # 5. ???: ?????????????? (?????) ??
    # =========================================================
    member_output_filename = os.path.join(
        output_dir, f"member_schedule_{_stage2_out_stamp}.xlsx"
    )
    
    # ????????????1??????????????????????
    time_labels = []
    time_grids = []
    curr_dt = datetime.combine(run_date, DEFAULT_START_TIME)
    end_dt_grid = datetime.combine(run_date, DEFAULT_END_TIME)
    while curr_dt < end_dt_grid:
        next_dt = curr_dt + timedelta(minutes=10)
        if next_dt > end_dt_grid:
            next_dt = end_dt_grid
        time_labels.append(f"{curr_dt.strftime('%H:%M')}-{next_dt.strftime('%H:%M')}")
        time_grids.append((curr_dt.time(), next_dt.time()))
        curr_dt = next_dt
    
    logging.info(
        "??2: ????????????????? ? %s",
        os.path.basename(member_output_filename),
    )
    try:
        with pd.ExcelWriter(member_output_filename, engine="openpyxl") as member_writer:
            for m in members:
                # ?????????
                m_schedule = {t_label: {"???": t_label} for t_label in time_labels}
            
                # ????????????????????
                for d in sorted_dates:
                    d_str = d.strftime("%m/%d (%a)")
                
                    # ?????: ???????? *?????????????????
                    if m not in attendance_data[d] or not attendance_data[d][m]['is_working']:
                        off_label = _member_schedule_full_day_off_label(
                            attendance_data[d].get(m) if m in attendance_data[d] else None
                        )
                        for t_label in time_labels:
                            m_schedule[t_label][d_str] = off_label
                        continue
                
                    daily_info = attendance_data[d][m]
                    d_start_dt = daily_info['start_dt']
                    d_end_dt = daily_info['end_dt']
                    breaks_dt = daily_info['breaks_dt']
                
                    events_today = events_by_date[d]
                
                    for i, (t_start, t_end) in enumerate(time_grids):
                        t_label = time_labels[i]
                    
                        # ???????????
                        grid_start_dt = datetime.combine(d, t_start)
                        grid_end_dt = datetime.combine(d, t_end)
                        grid_mid_dt = grid_start_dt + (grid_end_dt - grid_start_dt) / 2
                    
                        text = ""
                        if grid_mid_dt < d_start_dt or grid_mid_dt >= d_end_dt:
                            text = _member_schedule_off_shift_label(
                                d, grid_mid_dt, d_start_dt, d_end_dt, daily_info.get("reason")
                            )
                        else:
                            br_txt = _member_schedule_break_cell_label(
                                grid_mid_dt, breaks_dt, d_end_dt, daily_info.get("reason")
                            )
                            if br_txt is not None:
                                text = br_txt
                        if text == "":
                            # ???????????subs_list ????????
                            active_ev = next((e for e in events_today if e['start_dt'] <= grid_mid_dt < e['end_dt'] and (e['op'] == m or m in e.get('subs_list', []))), None)
                            if active_ev:
                                role = "?" if active_ev['op'] == m else "?"
                                text = f"[{active_ev['task_id']}] {active_ev['machine']}({role})"
                            else:
                                text = "" # ????????????????
                    
                        m_schedule[t_label][d_str] = text
                    
                # ??????????????????
                df_m = pd.DataFrame(list(m_schedule.values()))
                cols = ["???"] + [d.strftime("%m/%d (%a)") for d in sorted_dates]
                df_m = df_m[[c for c in cols if c in df_m.columns]]
                df_m.to_excel(member_writer, sheet_name=m, index=False)
            
                # --- ??????????????????? VBA ?????? AutoFit? ---
                worksheet = member_writer.sheets[m]
                _apply_output_font_to_result_sheet(worksheet)
                header_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
                for cell in worksheet[1]:
                    cell.fill = header_fill
                
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, max_col=worksheet.max_column):
                    for cell in row:
                        cell.border = thin_border

    except OSError as e:
        logging.error(
            "??2: ?????????????????????: %s?%s??"
            "member_schedule_*.xlsx ? Excel ?????????????????",
            member_output_filename,
            e,
        )
        raise

    logging.info(f"??: ?????????? '{member_output_filename}' ????????")
    _try_write_main_sheet_gemini_usage_summary("??2")