"""planning_core å®Ÿè£…æœ¬ä½“ï¼ˆãƒ‘ãƒƒã‚±ãƒ¼ã‚¸å†…ï¼‰ã€‚``import planning_core`` ã�§ bootstrap ã�Œå…ˆã�«å®Ÿè¡Œã�•ã‚Œã‚‹ã�“ã�¨ã€‚"""
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

# AI å‚™è€ƒãƒ»é…�å�°ä¸�èƒ½ãƒ­ã‚¸ãƒƒã‚¯ Dâ†’E ã�® TTL ã‚­ãƒ£ãƒƒã‚·ãƒ¥ï¼ˆæ—§ output/ ã�‹ã‚‰ json/ ã�¸ç§»è¡Œï¼‰
_ai_remarks_cache_name = "ai_remarks_cache.json"
_ai_cache_legacy = os.path.join(output_dir, _ai_remarks_cache_name)
_ai_cache_new = os.path.join(json_data_dir, _ai_remarks_cache_name)
if os.path.isfile(_ai_cache_legacy) and not os.path.isfile(_ai_cache_new):
    try:
        shutil.move(_ai_cache_legacy, _ai_cache_new)
    except OSError:
        pass
ai_cache_path = _ai_cache_new
# ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�ã‚·ãƒ¼ãƒˆä½œæˆ�ãƒ»ä¿�å­˜ã�®æˆ�å�¦ãƒ‡ãƒ�ãƒƒã‚°ï¼ˆexecution_log ã�¨ä½µç”¨ï¼‰
exclude_rules_sheet_debug_log_path = os.path.join(log_dir, "exclude_rules_sheet_debug.txt")
# ä¿�å­˜å¤±æ•—æ™‚ã�« E åˆ—ï¼ˆãƒ­ã‚¸ãƒƒã‚¯å¼�ï¼‰ã� ã�‘ã‚’é€€é�¿ã�—ã€�æ¬¡å›ž run_exclude_rules_sheet_maintenance ã�§è‡ªå‹•é�©ç”¨ã�™ã‚‹ï¼ˆjson ãƒ•ã‚©ãƒ«ãƒ€ï¼‰
EXCLUDE_RULES_E_SIDECAR_FILENAME = "exclude_rules_e_column_pending.json"
# openpyxl ä¿�å­˜å¤±æ•—æ™‚ã�« VBA ã�Œ E åˆ—ã�¸æ›¸ã��è¾¼ã‚€ã�Ÿã‚�ã�® UTF-8 TSVï¼ˆBase64ï¼‰ã€‚
EXCLUDE_RULES_E_VBA_TSV_FILENAME = "exclude_rules_e_column_vba.tsv"
# openpyxl ä¿�å­˜å¤±æ•—æ™‚ã�« VBA ã�Œ Aã€œE ã‚’ä¸€æ‹¬å��æ˜ ã�™ã‚‹ UTF-8 TSVï¼ˆè¡Œã�”ã�¨ã�« 5 ã‚»ãƒ«åˆ† Base64ï¼‰ã€‚
EXCLUDE_RULES_MATRIX_VBA_FILENAME = "exclude_rules_matrix_vba.tsv"
# VBA ã�Œãƒ¡ã‚¤ãƒ³ P åˆ—ã�¸æ›¸ã��è¾¼ã‚€ã�Ÿã‚�ã�® UTF-8 ãƒ†ã‚­ã‚¹ãƒˆï¼ˆExcel é–‹ã�„ã�Ÿã�¾ã�¾ save ã�§ã��ã�ªã�„å•�é¡Œã�®å›žé�¿ï¼‰
GEMINI_USAGE_SUMMARY_FOR_MAIN_FILE = "gemini_usage_summary_for_main.txt"
# å…¨å®Ÿè¡Œã‚’é€šã�—ã�Ÿ Gemini åˆ©ç”¨ãƒ»æŽ¨å®šæ–™é‡‘ã�®ç´¯è¨ˆï¼ˆAPI å¿œç­”ã�”ã�¨ã�«æ›´æ–°ã€‚ä¿�å­˜å…ˆã�¯ API_Payment ãƒ•ã‚©ãƒ«ãƒ€ï¼‰
GEMINI_USAGE_CUMULATIVE_JSON_FILE = "gemini_usage_cumulative.json"
# æœŸé–“åˆ¥ãƒ�ã‚±ãƒƒãƒˆã‚’ãƒ•ãƒ©ãƒƒãƒˆåŒ–ã�—ã�Ÿ CSVï¼ˆExcel ã�®æŠ˜ã‚Œç·šãƒ»æ£’ã‚°ãƒ©ãƒ•ç”¨ï¼‰
GEMINI_USAGE_BUCKETS_CSV_FILE = "gemini_usage_buckets_for_chart.csv"
# ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆãƒ»Gemini æ—¥æ¬¡æŽ¨ç§»ï¼ˆxlwings: Qã€œRï¼�æ–™é‡‘ã�¾ã�Ÿã�¯å‘¼å‡ºã�—ã€�Sã€œTï¼�å�ˆè¨ˆãƒˆãƒ¼ã‚¯ãƒ³ï¼‰
GEMINI_USAGE_CHART_COL_DATE = 17  # Q
GEMINI_USAGE_CHART_COL_VALUE = 18  # R
GEMINI_USAGE_CHART_COL_TOK_DATE = 19  # Sï¼ˆã‚°ãƒ©ãƒ•ç”¨ã�«æ—¥ä»˜ã‚’è¤‡è£½ï¼‰
GEMINI_USAGE_CHART_COL_TOK_VALUE = 20  # Tï¼ˆtotal_tokens ç›¸å½“ï¼‰
GEMINI_USAGE_CHART_HEADER_ROW = 16
GEMINI_USAGE_CHART_ANCHOR_CELL = "T16"
GEMINI_USAGE_CHART_TOKENS_ANCHOR_CELL = "AA16"
GEMINI_USAGE_CHART_MAX_DAYS = 14
GEMINI_USAGE_CHART_CLEAR_ROWS = 36
# xlwings ã�§è²¼ã‚‹æŠ˜ã‚Œç·šã�®å��å‰�ï¼ˆå†�å®Ÿè¡Œæ™‚ã�«å‰Šé™¤ã�—ã�¦ã�‹ã‚‰ä½œã‚Šç›´ã�™ï¼‰
GEMINI_USAGE_XLW_CHART_NAME = "_GeminiApiDailyTrend"
GEMINI_USAGE_XLW_CHART_TOKENS_NAME = "_GeminiApiDailyTokens"
# ãƒ†ã‚¹ãƒˆ: EXCLUDE_RULES_TEST_E1234=1 ã�§ EXCLUDE_RULES_SHEET_NAMEï¼ˆã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�ï¼‰ã�® E åˆ—ã�« "1234" ã‚’æ›¸ã��ï¼ˆä¿�å­˜çµŒè·¯ã�®ç¢ºèª�ç”¨ï¼‰ã€‚
# TASK_INPUT_WORKBOOK ã�¯ã€ŒåŠ å·¥è¨ˆç”»DATAã€�ã‚·ãƒ¼ãƒˆä»˜ã��ãƒ–ãƒƒã‚¯ï¼ˆä¾‹: ç”Ÿç”£ç®¡ç�†_AIé…�å�°ãƒ†ã‚¹ãƒˆ.xlsmï¼‰ã‚’æŒ‡å®šã�™ã‚‹ã�“ã�¨ã€‚
# è¡Œã�¯ EXCLUDE_RULES_TEST_E1234_ROWï¼ˆæ—¢å®š 9ã€�2 æœªæº€ã�¯ 9 ã�«ä¸¸ã‚�ã‚‹ï¼‰ã€‚

# =========================================================
# ã€�è¨­å®šã€‘APIã‚­ãƒ¼ / åŸºæœ¬ãƒ«ãƒ¼ãƒ« / ãƒ•ã‚¡ã‚¤ãƒ«å��
# =========================================================
# Gemini API ã‚­ãƒ¼ã�¯ TASK_INPUT_WORKBOOK ç¢ºå®šå¾Œã€�ä¸‹è¨˜ã€Œè¨­å®šã€�B1 ã�® JSON ã�‹ã‚‰è§£æ±ºï¼ˆå¹³æ–‡ã�¾ã�Ÿã�¯ format_version 2 ã�®æš—å�·åŒ–ï¼‰ã€‚
# æœªè¨­å®šæ™‚ã�®ã�¿ç§»è¡Œç”¨ã�«ç’°å¢ƒå¤‰æ•° GEMINI_API_KEY ã‚’å�‚ç…§ã€‚

GEMINI_MODEL_FLASH = "gemini-2.5-flash"
# æŽ¨å®šæ–™é‡‘: USD / 1M tokensï¼ˆå…¥åŠ›, å‡ºåŠ›ï¼‰ã€‚å…¬å¼�ã�®æœ€æ–°å�˜ä¾¡ã�«å�ˆã‚�ã�›ã�¦æ›´æ–°ã�™ã‚‹ã�“ã�¨ã€‚
# ç’°å¢ƒå¤‰æ•° GEMINI_PRICE_USD_IN_PER_M / GEMINI_PRICE_USD_OUT_PER_M ã�§ä¸Šæ›¸ã��å�¯ï¼ˆFlash å�‘ã�‘ï¼‰ã€‚
_GEMINI_FLASH_IN_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_IN_PER_M", "0.075") or 0.075
)
_GEMINI_FLASH_OUT_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_OUT_PER_M", "0.30") or 0.30
)
GEMINI_JPY_PER_USD = float(os.environ.get("GEMINI_JPY_PER_USD", "150") or 150)

# ---------------------------------------------------------------------------
# ä»¥é™�ã�®å®šæ•°ãƒ–ãƒ­ãƒƒã‚¯ã�¯ã€ŒExcel åˆ—è¦‹å‡ºã�—ã€�ã�¨ 1:1 ã�§å¯¾å¿œã�•ã�›ã‚‹ã€‚
# åˆ—å��ã‚’å¤‰ã�ˆã‚‹å ´å�ˆã�¯ VBAãƒ»ãƒžã‚¯ãƒ­å�´ã‚·ãƒ¼ãƒˆã�¨å�Œæ™‚ã�«ç›´ã�™ã�“ã�¨ã€‚
# ---------------------------------------------------------------------------

MASTER_FILE = "master.xlsm" # skillsã�¨attendance(ã�Šã‚ˆã�³tasks)ã‚’çµ±å�ˆã�—ã�Ÿãƒ•ã‚¡ã‚¤ãƒ«
# VBAã€Œmaster_æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã‚’ä½œæˆ�ã€�ã‚·ãƒ¼ãƒˆï¼ˆ1 æ™‚é–“ã‚¹ãƒ­ãƒƒãƒˆå� æœ‰ã‚’æ®µéšŽ2ã�® machine_avail_dt ã�«å��æ˜ ï¼‰
SHEET_MACHINE_CALENDAR = "æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼"
# ``generate_plan`` é–‹å§‹æ™‚ã�«å†�è¨­å®šã€‚date -> è¨­å‚™ã‚­ãƒ¼ -> [ (start, end), ... ] å�Šé–‹åŒºé–“ [start, end)
_MACHINE_CALENDAR_BLOCKS_BY_DATE: dict[
    date, dict[str, list[tuple[datetime, datetime]]]
] = {}

# master.xlsm: ä¾�é ¼NO ã�Œå¤‰ã‚�ã‚‹å‰�å¾Œã�®å·¥ç¨‹Ã—æ©Ÿæ¢°ã�”ã�¨ã�®æº–å‚™ãƒ»å¾Œå§‹æœ«ï¼ˆåˆ†ï¼‰ï¼�æ©Ÿæ¢°ã�”ã�¨ã�®æ—¥æ¬¡å§‹æ¥­æº–å‚™ï¼ˆåˆ†ï¼‰
SHEET_MACHINE_CHANGEOVER = "è¨­å®š_ä¾�é ¼åˆ‡æ›¿å‰�å¾Œæ™‚é–“"
SHEET_MACHINE_DAILY_STARTUP = "è¨­å®š_æ©Ÿæ¢°_æ—¥æ¬¡å§‹æ¥­æº–å‚™"
# ``generate_plan`` é–‹å§‹æ™‚ã�«å†�è¨­å®šï¼ˆã‚·ãƒ¼ãƒˆç„¡ã�—ãƒ»ç©ºã�¯ç©ºè¾žæ›¸ï¼�å¾“æ�¥ã�©ã�Šã‚Šï¼‰
_STAGE2_MACHINE_CHANGEOVER_BY_EQ: dict[str, tuple[int, int]] = {}
_STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE: dict[str, int] = {}
# master ãƒ¡ã‚¤ãƒ³ A15ï¼ˆå®šå¸¸é–‹å§‹ï¼‰ã€‚æ—¥æ¬¡å§‹æ¥­æº–å‚™ã‚’å‹¤æ€  forward ã�§ã�¯ã�ªã�� [é–‹å§‹, é–‹å§‹+Nåˆ†) ã�®å£�æ™‚è¨ˆã�«è¼‰ã�›ã‚‹ã€‚
_STAGE2_REGULAR_SHIFT_START: time | None = None
# timeline_events ã�® event_kindï¼ˆçœ�ç•¥æ™‚ã�¯åŠ å·¥ã�¨ã�¿ã�ªã�™ï¼‰
TIMELINE_EVENT_MACHINING = "machining"
TIMELINE_EVENT_MACHINE_DAILY_STARTUP = "machine_daily_startup"
TIMELINE_EVENT_CHANGEOVER_CLEANUP = "changeover_cleanup"
TIMELINE_EVENT_CHANGEOVER_PREP = "changeover_prep"
# VBAã€Œmaster_çµ„ã�¿å�ˆã‚�ã�›è¡¨ã‚’æ›´æ–°ã€�ã�§ä½œã‚‹ã‚·ãƒ¼ãƒˆï¼ˆå·¥ç¨‹+æ©Ÿæ¢°ã‚­ãƒ¼ã�¨ãƒ¡ãƒ³ãƒ�ãƒ¼ç·¨æˆ�ï¼‰
MASTER_SHEET_TEAM_COMBINATIONS = "çµ„ã�¿å�ˆã‚�ã�›è¡¨"
# ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ¥å‹¤æ€ ã‚·ãƒ¼ãƒˆ: master.xlsm ã�§ã�¯ã€Œä¼‘æš‡åŒºåˆ†ã€�ã�¨ã€Œå‚™è€ƒã€�ã�Œåˆ¥åˆ—ã€‚
# å‹¤æ€ AIã�®å…¥åŠ›ã�¯å‚™è€ƒã�®ã�¿ã€‚ã�Ÿã� ã�— reasonï¼ˆè¡¨ç¤ºãƒ»ä¸­æŠœã�‘è£œæ­£ãƒ»å€‹äººã‚·ãƒ¼ãƒˆã�®ä¼‘æ†©/ä¼‘æš‡æ–‡è¨€ï¼‰ã�¯ã€�å‚™è€ƒã�Œç©ºã�®ã�¨ã��ä¼‘æš‡åŒºåˆ†ã‚’å¼•ã��ç¶™ã��ã€‚
# master ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ï¼�å‡ºå‹¤ç°¿.txt æº–æ‹ : å‰�ä¼‘=å�ˆå‰�å¹´ä¼‘ãƒ»ä¼‘æ†©æ™‚é–“1_çµ‚äº†ï½žå®šå¸¸çµ‚äº†ï¼ˆå�ˆå¾Œä¼‘æ†©14:45ï½ž15:00ï¼‰ï¼�å¾Œä¼‘=å®šå¸¸é–‹å§‹ï½žä¼‘æ†©æ™‚é–“1_é–‹å§‹ãƒ»å�ˆå¾Œå¹´ä¼‘ï¼�å›½=ä»–æ‹ ç‚¹å‹¤å‹™ã€‚
# å‚™è€ƒåˆ—ãƒ»ä¼‘æš‡åŒºåˆ†ã�¯å‹¤æ€  AI ã�§æ§‹é€ åŒ–ï¼ˆé…�å�°ä¸�å�‚åŠ ãƒ»is_holidayãƒ»ä¸­æŠœã�‘ç­‰ï¼‰ã€‚å‚™è€ƒã�Œç©ºã�§ã‚‚ä¼‘æš‡åŒºåˆ†ã�®ã�¿ã�®è¡Œã�¯ AI ã�«æ¸¡ã�™ã€‚
ATT_COL_LEAVE_TYPE = "ä¼‘æš‡åŒºåˆ†"
ATT_COL_REMARK = "å‚™è€ƒ"
# ãƒ¡ãƒ³ãƒ�ãƒ¼å‹¤æ€ ã‚·ãƒ¼ãƒˆï¼ˆmaster.xlsmï¼‰: å®šæ™‚ã�®ã€Œé€€å‹¤æ™‚é–“ã€�ã�¨åˆ†ã�‘ã�¦é€€å‹¤ä¸Šé™�ã‚’æŒ‡å®šï¼ˆä»»æ„�åˆ—ï¼‰
ATT_COL_OT_END = "æ®‹æ¥­çµ‚æ¥­"
# å‹¤æ€ å‚™è€ƒ AI ã�® JSON ã‚¹ã‚­ãƒ¼ãƒžã‚’å¤‰ã�ˆã�Ÿã‚‰æ›´æ–°ã�—ã€�ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã‚­ãƒ¼ã‚’ç„¡åŠ¹åŒ–ã�™ã‚‹
ATTENDANCE_REMARK_AI_SCHEMA_ID = "v2_haitai_fuka"
# need ã‚·ãƒ¼ãƒˆ: ã€ŒåŸºæœ¬å¿…è¦�äººæ•°ã€�è¡Œï¼ˆAåˆ—ã�«ã€Œå¿…è¦�äººæ•°ã€�ã‚’å�«ã‚€ï¼‰ï¼‹ ã��ã�®ç›´ä¸‹ã�®ã€Œé…�å�°æ™‚è¿½åŠ äººæ•°ï¼�ä½™åŠ›æ™‚è¿½åŠ äººæ•°ã€�ç­‰
# ï¼ˆExcel ä¸Šã�¯æ¦‚ã�­ 5 è¡Œç›®ä»˜è¿‘ã€‚ä½™å‰°æ™‚ã�«å¢—ã‚„ã�›ã‚‹äººæ•°ä¸Šé™�ãƒ»å·¥ç¨‹Ã—æ©Ÿæ¢°åˆ—ï¼‰
# ï¼‹ è¡Œã€Œç‰¹åˆ¥æŒ‡å®š1ã€�ï½žã€Œç‰¹åˆ¥æŒ‡å®š99ã€�ï¼ˆå¿…è¦�äººæ•°ã�®ä¸Šæ›¸ã��ãƒ»1ï½ž99ï¼‰
NEED_COL_CONDITION = "ä¾�é ¼NOæ�¡ä»¶"
NEED_COL_NOTE = "å‚™è€ƒ"
# needã€Œé…�å�°æ™‚è¿½åŠ äººæ•°ã€�ã‚’æº€æž ä½¿ã�£ã�¦ã‚‚ã€�å�˜ä½�ã�‚ã�Ÿã‚ŠåŠ å·¥æ™‚é–“ã�ŒçŸ­ã��ã�ªã‚‹ã�®ã�¯æœ€å¤§ã�§ã�“ã�®å‰²å�ˆï¼ˆä¾‹: 0.05 â‰’ 5%ï¼‰
SURPLUS_TEAM_MAX_SPEEDUP_RATIO = 0.05
# ã‚¿ã‚¹ã‚¯ã�¯ tasks.xlsx ã‚’ä½¿ã‚�ã�šã€�VBA ã�‹ã‚‰æ¸¡ã�™ TASK_INPUT_WORKBOOK ã�®ã€ŒåŠ å·¥è¨ˆç”»DATAã€�ã�®ã�¿
TASKS_INPUT_WORKBOOK = os.environ.get("TASK_INPUT_WORKBOOK", "").strip()
TASKS_SHEET_NAME = "åŠ å·¥è¨ˆç”»DATA"

# ã�“ã�®ã‚·ãƒ¼ãƒˆå��ã‚’å�«ã‚€ãƒ–ãƒƒã‚¯ã�¯ openpyxl ã�Œèª­ã�¿æ›¸ã��ã�«å¤±æ•—ã�™ã‚‹ã�“ã�¨ã�Œã�‚ã‚‹ã�Ÿã‚�ã€�load_workbook ã‚’è©¦è¡Œã�—ã�ªã�„
OPENPYXL_INCOMPATIBLE_SHEET_MARKER = "é…�å�°_é…�å�°ä¸�è¦�å·¥ç¨‹"


def _ooxml_workbook_sheet_names(wb_path: str) -> list[str] | None:
    """ZIP å†… xl/workbook.xml ã�‹ã‚‰ã‚·ãƒ¼ãƒˆå��ä¸€è¦§ã‚’å�–ã‚‹ï¼ˆopenpyxl ã‚’ä½¿ã‚�ã�ªã�„ï¼‰ã€‚"""
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
    """å½“è©²ãƒ‘ã‚¹ã�Œ OOXML ã�§ã‚·ãƒ¼ãƒˆã€Œé…�å�°_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�ã‚’å�«ã‚€ã�¨ã�� Trueï¼ˆopenpyxl åˆ©ç”¨ã‚’é�¿ã�‘ã‚‹ï¼‰ã€‚"""
    p = (wb_path or "").strip()
    if not p:
        return False
    names = _ooxml_workbook_sheet_names(p)
    if not names:
        return False
    return OPENPYXL_INCOMPATIBLE_SHEET_MARKER in names


# ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œè¨­å®šã€�B1: ç¤¾å†…å…±æœ‰ä¸Šã�® Gemini èª�è¨¼ JSON ã�®ãƒ‘ã‚¹
APP_CONFIG_SHEET_NAME = "è¨­å®š"
# æš—å�·åŒ–èª�è¨¼ JSONï¼ˆformat_version 2ï¼‰ã�®å¾©å�·ã�¯å¸¸ã�«ã�“ã�®å®šæ•°ã�®ã�¿ï¼ˆç¤¾å†…æ‰‹é †ã�®ãƒ‘ã‚¹ãƒ•ãƒ¬ãƒ¼ã‚ºã�¨ä¸€è‡´ã�•ã�›ã‚‹ã€‚ãƒ­ã‚°ãƒ»UI ã�«å‡ºã�•ã�ªã�„ï¼‰ã€‚
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
    """ã€Œè¨­å®šã€�ã‚·ãƒ¼ãƒˆ B1 ã�‹ã‚‰èª�è¨¼ JSON ãƒ•ã‚¡ã‚¤ãƒ«ãƒ‘ã‚¹ã‚’èª­ã‚€ã€‚ç„¡ã�‘ã‚Œã�° Noneã€‚"""
    if not wb_path or not os.path.isfile(wb_path):
        return None
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.debug(
            "Gemini: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� openpyxl ã�§ã€Œ%sã€�!B1 ã‚’èª­ã�¿ã�¾ã�›ã‚“ã€‚",
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
            "Gemini: ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œ%sã€�ã�®ã€Œ%sã€�!B1 ã‚’èª­ã‚�ã�¾ã�›ã‚“: %s",
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
    ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œè¨­å®šã€�ã‚·ãƒ¼ãƒˆã�®æŒ‡å®šåˆ—ï¼ˆ1=A, 2=Bï¼‰3 è¡Œç›®ä»¥é™�ã�‹ã‚‰ä¾�é ¼NOã‚’èª­ã‚€ã€‚
    ç©ºã‚»ãƒ«ã�¯ã‚¹ã‚­ãƒƒãƒ—ã€‚é€£ç¶š 30 ã‚»ãƒ«ç©ºã�§æ‰“ã�¡åˆ‡ã‚Šã€‚æœ€å¤§ 500 è¡Œã€‚ã‚«ãƒ³ãƒžåŒºåˆ‡ã‚Šã�§è¤‡æ•°å�¯ã€‚
    """
    out: list[str] = []
    if not wb_path or not os.path.isfile(wb_path):
        return out
    if _workbook_should_skip_openpyxl_io(wb_path):
        msg = (
            f"{log_label}: ãƒ–ãƒƒã‚¯ã�«ã€Œ{OPENPYXL_INCOMPATIBLE_SHEET_MARKER}ã€�ã�Œã�‚ã‚‹ã�Ÿã‚�"
            f"ã€Œ{APP_CONFIG_SHEET_NAME}ã€�!{column_letter_desc}3 ä»¥é™�ã�¯ openpyxl ã�§èª­ã‚�ã�¾ã�›ã‚“ã€‚"
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
            "%s: ã€Œ%sã€�!%s3 ä»¥é™�ã�®ä¾�é ¼NOã‚’èª­ã‚�ã�¾ã�›ã‚“ï¼ˆç„¡è¦–ï¼‰: %s",
            log_label,
            APP_CONFIG_SHEET_NAME,
            column_letter_desc,
            ex,
        )
        return []
    return out


def _read_trace_schedule_task_ids_from_config_sheet(wb_path: str) -> list[str]:
    """
    ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œè¨­å®šã€�ã‚·ãƒ¼ãƒˆ A åˆ—ã�® 3 è¡Œç›®ä»¥é™�ã‚’ã€�é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹å¯¾è±¡ã�®ä¾�é ¼NOã�¨ã�—ã�¦èª­ã‚€ã€‚
    ç©ºã‚»ãƒ«ã�¯ã‚¹ã‚­ãƒƒãƒ—ã€‚é€£ç¶š 30 ã‚»ãƒ«ç©ºã�ªã‚‰æ‰“ã�¡åˆ‡ã‚Šã€‚æœ€å¤§ 500 è¡Œã�¾ã�§èµ°æŸ»ã€‚
    """
    return _read_task_ids_from_config_sheet_column(
        wb_path,
        1,
        "é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹",
        "A",
        openpyxl_skip_hint="é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ã�¯ã€Œè¨­å®šã€�ã‚·ãƒ¼ãƒˆ A åˆ—ã‚’ openpyxl ã�§èª­ã‚�ã�ªã�„ã�Ÿã‚�ç„¡åŠ¹ã�§ã�™ã€‚",
    )


def _read_debug_dispatch_task_ids_from_config_sheet(wb_path: str) -> list[str]:
    """
    ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œè¨­å®šã€�ã‚·ãƒ¼ãƒˆ B åˆ—ã�® 3 è¡Œç›®ä»¥é™�ã‚’ã€�æ®µéšŽ2ãƒ‡ãƒ�ãƒƒã‚°é…�å�°ã�®å¯¾è±¡ä¾�é ¼NOã�¨ã�—ã�¦èª­ã‚€ã€‚
    1 ä»¶ã‚‚ç„¡ã�„å ´å�ˆã�¯æ®µéšŽ2ã�¯é€šå¸¸ãƒ¢ãƒ¼ãƒ‰ï¼ˆå…¨ä»¶é…�å�°ï¼‰ã€‚ç©ºã‚»ãƒ«ãƒ»æ‰“ã�¡åˆ‡ã‚Šç­‰ã�¯ A åˆ—ãƒˆãƒ¬ãƒ¼ã‚¹ã�¨å�Œã�˜ã€‚
    """
    return _read_task_ids_from_config_sheet_column(
        wb_path,
        2,
        "ãƒ‡ãƒ�ãƒƒã‚°é…�å�°",
        "B",
        openpyxl_skip_hint="ãƒ‡ãƒ�ãƒƒã‚°é…�å�°ã�¯ã€Œè¨­å®šã€�ã‚·ãƒ¼ãƒˆ B åˆ—ã‚’ openpyxl ã�§èª­ã‚�ã�ªã�„ã�Ÿã‚�ç„¡åŠ¹ï¼ˆå…¨ä»¶é…�å�°ï¼‰ã�§ã�™ã€‚",
    )


def _show_stage2_debug_dispatch_mode_dialog(task_ids_sorted: list[str]) -> None:
    """è¨­å®šã‚·ãƒ¼ãƒˆ B3ä»¥é™�ã�Œç©ºã�§ã�ªã�„ã�¨ã��ã� ã�‘å‘¼ã�¶ã€‚Windows ã�§ã�¯ MessageBoxã€�ã��ã‚Œä»¥å¤–ã�¯ WARNING ãƒ­ã‚°ã€‚"""
    if not task_ids_sorted:
        return
    preview_lines = task_ids_sorted[:30]
    preview = "\n".join(preview_lines)
    if len(task_ids_sorted) > 30:
        preview += "\nâ€¦"
    body = (
        "ãƒ‡ãƒ�ãƒƒã‚°ãƒ¢ãƒ¼ãƒ‰ã�§å®Ÿè¡Œã�—ã�¾ã�™ã€‚\n\n"
        "ã€Œè¨­å®šã€�ã‚·ãƒ¼ãƒˆ B3ä»¥é™�ã�«å…¥åŠ›ã�—ã�Ÿä¾�é ¼NOã�®ã�¿ã‚’é…�å�°å¯¾è±¡ã�¨ã�—ã�¾ã�™ã€‚\n\n"
        "å¯¾è±¡ä¾�é ¼NO:\n"
        + preview
    )
    title = "æ®µéšŽ2ï¼ˆé…�å�°ï¼‰â€” ãƒ‡ãƒ�ãƒƒã‚°ãƒ¢ãƒ¼ãƒ‰"
    if sys.platform != "win32":
        logging.warning("%s\n%s", title, body)
        return
    try:
        ctypes.windll.user32.MessageBoxW(0, body, title, 0x00000040)
    except Exception as ex:
        logging.warning(
            "ãƒ‡ãƒ�ãƒƒã‚°é…�å�°: ãƒ¡ãƒƒã‚»ãƒ¼ã‚¸ãƒœãƒƒã‚¯ã‚¹ã‚’è¡¨ç¤ºã�§ã��ã�¾ã�›ã‚“ (%s)ã€‚%s", ex, body
        )


def _extract_gemini_api_key_from_plain_dict(data: dict, json_path: str) -> str | None:
    key = data.get("gemini_api_key")
    if key is None or (isinstance(key, str) and not key.strip()):
        key = data.get("GEMINI_API_KEY")
    if key is None:
        logging.warning(
            "Gemini: èª�è¨¼ãƒ‡ãƒ¼ã‚¿ã�« gemini_api_keyï¼ˆã�¾ã�Ÿã�¯ GEMINI_API_KEYï¼‰ã�Œã�‚ã‚Šã�¾ã�›ã‚“ï¼ˆ%sï¼‰ã€‚",
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
            "Gemini: æš—å�·åŒ–èª�è¨¼ JSON ã�«ã�¯ cryptography ã�Œå¿…è¦�ã�§ã�™ï¼ˆpip install cryptographyï¼‰ã€‚"
        )
        return None
    token_s = (data.get("fernet_ciphertext") or "").strip()
    if not token_s:
        logging.warning(
            "Gemini: æš—å�·åŒ–èª�è¨¼ JSON ã�« fernet_ciphertext ã�Œã�‚ã‚Šã�¾ã�›ã‚“ï¼ˆ%sï¼‰ã€‚",
            json_path,
        )
        return None
    salt_b64 = (data.get("salt_b64") or "").strip()
    if not salt_b64:
        logging.warning(
            "Gemini: æš—å�·åŒ–èª�è¨¼ JSON ã�« salt_b64 ã�Œã�‚ã‚Šã�¾ã�›ã‚“ï¼ˆ%sï¼‰ã€‚",
            json_path,
        )
        return None
    try:
        salt = base64.standard_b64decode(salt_b64)
    except Exception as ex:
        logging.warning("Gemini: salt_b64 ã�®è§£é‡ˆã�«å¤±æ•—ã�—ã�¾ã�—ã�Ÿï¼ˆ%sï¼‰: %s", json_path, ex)
        return None
    iterations = int(data.get("iterations") or _GEMINI_CREDENTIALS_PBKDF2_ITERATIONS_DEFAULT)
    kdf_name = (data.get("kdf") or "pbkdf2_sha256").strip()
    if kdf_name != "pbkdf2_sha256":
        logging.warning("Gemini: æœªå¯¾å¿œã�® kdfï¼ˆ%sï¼‰: %s", kdf_name, json_path)
        return None
    try:
        fkey = _derive_fernet_key_from_passphrase(passphrase, salt, iterations)
        plain = Fernet(fkey).decrypt(token_s.encode("ascii"))
    except Exception:
        logging.debug("Gemini: æš—å�·åŒ–èª�è¨¼ã�®å¾©å�·å‡¦ç�†ã�«å¤±æ•—ã�—ã�¾ã�—ã�Ÿï¼ˆ%sï¼‰ã€‚", json_path)
        return None
    try:
        inner = json.loads(plain.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as ex:
        logging.warning(
            "Gemini: å¾©å�·å¾Œã�® JSON ã�Œä¸�æ­£ã�§ã�™ï¼ˆ%sï¼‰: %s",
            json_path,
            ex,
        )
        return None
    if not isinstance(inner, dict):
        logging.warning("Gemini: å¾©å�·å¾Œã�® JSON ã�¯ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�§ã�‚ã‚‹å¿…è¦�ã�Œã�‚ã‚Šã�¾ã�™ï¼ˆ%sï¼‰ã€‚", json_path)
        return None
    return inner


def _load_gemini_api_key_from_credentials_json(
    json_path: str, workbook_path: str | None = None
) -> tuple[str | None, bool]:
    """æˆ»ã‚Šå€¤: (api_key ã�¾ã�Ÿã�¯ None, æš—å�·åŒ–å½¢å¼�ã� ã�£ã�Ÿã�‹)ã€‚æš—å�·åŒ–æ™‚ã�¯ _GEMINI_CREDENTIALS_PASSPHRASE_FIXED ã�®ã�¿ã�§å¾©å�·ã€‚"""
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
    except OSError as ex:
        logging.warning("Gemini: èª�è¨¼ JSON ã‚’é–‹ã�‘ã�¾ã�›ã‚“: %s (%s)", json_path, ex)
        return None, False
    except json.JSONDecodeError as ex:
        logging.warning("Gemini: èª�è¨¼ JSON ã�®å½¢å¼�ã�Œä¸�æ­£ã�§ã�™: %s (%s)", json_path, ex)
        return None, False
    if not isinstance(data, dict):
        logging.warning("Gemini: èª�è¨¼ JSON ã�¯ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆå½¢å¼�ã�§ã�‚ã‚‹å¿…è¦�ã�Œã�‚ã‚Šã�¾ã�™: %s", json_path)
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
            logging.info("Gemini API ã‚­ãƒ¼: æš—å�·åŒ–èª�è¨¼ JSON ã�‹ã‚‰èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚")
        else:
            logging.info(
                "Gemini API ã‚­ãƒ¼: ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œ%sã€�B1 ã�®ãƒ‘ã‚¹ã�‹ã‚‰èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚",
                APP_CONFIG_SHEET_NAME,
            )
elif _cred_path:
    logging.warning(
        "Gemini: ã€Œ%sã€�B1 ã�§æŒ‡å®šã�•ã‚Œã�Ÿèª�è¨¼ JSON ã�Œè¦‹ã�¤ã�‹ã‚Šã�¾ã�›ã‚“ã€‚",
        APP_CONFIG_SHEET_NAME,
    )

# B1 ã�Œæš—å�·åŒ– JSON ã�ªã�®ã�«ã‚­ãƒ¼ã�Œå�–ã‚Œã�ªã�„ï¼ˆå¹³æ–‡ JSON ã�§ã‚­ãƒ¼æ¬ è�½ã�¨ã�®åŒºåˆ¥ï¼‰ã€‚åŽŸå› ã�®ç‰¹å®šã�¯ãƒ­ã‚°ã�«æ›¸ã�‹ã�šæ±Žç”¨ãƒ¡ãƒƒã‚»ãƒ¼ã‚¸ã�®ã�¿ã€‚
_encrypted_json_missing_key = (
    bool(_cred_path)
    and os.path.isfile(_cred_path)
    and _used_encrypted_credentials
    and not API_KEY
)
if _encrypted_json_missing_key:
    logging.error(
        "Gemini: ã€Œ%sã€�B1 ã�®èª�è¨¼ãƒ•ã‚¡ã‚¤ãƒ«ã�‹ã‚‰ API ã‚­ãƒ¼ã‚’åˆ©ç”¨ã�§ã��ã�¾ã�›ã‚“ã€‚"
        " ç¤¾å†…æ‰‹é †ã�«å¾“ã�„èª�è¨¼ã‚’å†�è¨­å®šã�™ã‚‹ã�‹ã€�ç®¡ç�†è€…ã�«å•�ã�„å�ˆã‚�ã�›ã�¦ã��ã� ã�•ã�„ã€‚",
        APP_CONFIG_SHEET_NAME,
    )

if not API_KEY:
    logging.warning(
        "Gemini API ã‚­ãƒ¼ã�Œæœªè¨­å®šã�§ã�™ã€‚ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã‚·ãƒ¼ãƒˆã‚’ç”¨æ„�ã�— B1 ã�«èª�è¨¼ JSON ã�®ãƒ•ãƒ«ãƒ‘ã‚¹ã‚’æ›¸ã�„ã�¦ã��ã� ã�•ã�„ã€‚"
        " å‚™è€ƒã�® AI è§£æž�ç­‰ã�¯ã‚¹ã‚­ãƒƒãƒ—ã�•ã‚Œã�¾ã�™ã€‚"
        " ã�²ã�ªåž‹: gemini_credentials.example.json / encrypt_gemini_credentials.pyï¼ˆæš—å�·åŒ–ï¼‰ã€‚",
        APP_CONFIG_SHEET_NAME,
    )

RESULT_SHEET_GANTT_NAME = "çµ�æžœ_è¨­å‚™ã‚¬ãƒ³ãƒˆ"

# ã‚¿ã‚¹ã‚¯åˆ—å��ï¼ˆãƒžã‚¯ãƒ­å®Ÿè¡Œãƒ–ãƒƒã‚¯ã€ŒåŠ å·¥è¨ˆç”»DATAã€�ï¼‰
TASK_COL_TASK_ID = "ä¾�é ¼NO"
TASK_COL_MACHINE = "å·¥ç¨‹å��"
TASK_COL_MACHINE_NAME = "æ©Ÿæ¢°å��"
TASK_COL_QTY = "æ�›ç®—æ•°é‡�"
TASK_COL_ORDER_QTY = "å�—æ³¨æ•°"
TASK_COL_SPEED = "åŠ å·¥é€Ÿåº¦"
TASK_COL_PRODUCT = "è£½å“�å��"
TASK_COL_ANSWER_DUE = "å›žç­”ç´�æœŸ"
TASK_COL_SPECIFIED_DUE = "æŒ‡å®šç´�æœŸ"
TASK_COL_RAW_INPUT_DATE = "åŽŸå��æŠ•å…¥æ—¥"
# åŠ å·¥è¨ˆç”»DATA ç”±æ�¥ã€‚é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã�§ã�¯åŽŸå��æŠ•å…¥æ—¥ã�®å�³éš£ï¼ˆSOURCE_BASE_COLUMNS é †ï¼‰ã€‚
TASK_COL_STOCK_LOCATION = "åœ¨åº«å ´æ‰€"
# å�Œä¸€ä¾�é ¼NOã�®å·¥ç¨‹é †ï¼ˆã‚«ãƒ³ãƒžåŒºåˆ‡ã‚Šã�®å·¥ç¨‹å��ï¼‰ã€‚åŠ å·¥è¨ˆç”»DATAï¼�é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€‚
TASK_COL_PROCESS_CONTENT = "åŠ å·¥å†…å®¹"
# æŠ•å…¥å�¯èƒ½æ—¥ã�®ç›®å®‰ã�¯ã€Œå›žç­”ç´�æœŸã€�ã€�æœªå…¥åŠ›æ™‚ã�¯ã€ŒæŒ‡å®šç´�æœŸã€�ï¼ˆå‰�æ—¥åŸºæº–ãƒ»å½“æ—¥/é�…ã‚Œã�¯æœ€å„ªå…ˆï¼‰ã€‚ã€ŒåŠ å·¥é–‹å§‹æ—¥ã€�åˆ—ã�¯å�‚ç…§ã�—ã�ªã�„ã€‚
# å®Œäº†åˆ¤å®šãƒ»é€²æ�—ï¼ˆåŠ å·¥è¨ˆç”»DATAï¼‰
TASK_COL_COMPLETION_FLAG = "åŠ å·¥å®Œäº†åŒºåˆ†"
TASK_COL_ACTUAL_DONE = "å®ŸåŠ å·¥æ•°"   # æ—§äº’æ�›ï¼ˆç›´æŽ¥ã�®åŠ å·¥æ¸ˆæ•°é‡�ï¼‰
TASK_COL_ACTUAL_OUTPUT = "å®Ÿå‡ºæ�¥é«˜"  # å®Œæˆ�å“�æ•°é‡�ï¼ˆæ�›ç®—ã�«ä½¿ã�†ï¼‰
TASK_COL_DATA_EXTRACTION_DT = "ãƒ‡ãƒ¼ã‚¿æŠ½å‡ºæ—¥"
AI_CACHE_TTL_SECONDS = 6 * 60 * 60  # 6æ™‚é–“
# json/ai_remarks_cache.json å†…ã�®ã‚­ãƒ¼æŽ¥é ­è¾žï¼ˆè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ãƒ»é…�å�°ä¸�èƒ½ãƒ­ã‚¸ãƒƒã‚¯ Dâ†’Eï¼‰
AI_CACHE_KEY_PREFIX_EXCLUDE_RULE_DE = "exclude_rule_de_v1"

# ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€ŒåŠ å·¥å®Ÿç¸¾DATAã€�ï¼ˆPower Query ç­‰ã�§å�–ã‚Šè¾¼ã�¿æƒ³å®šï¼‰
ACTUALS_SHEET_NAME = "åŠ å·¥å®Ÿç¸¾DATA"
ACT_COL_TASK_ID = "ä¾�é ¼NO"
ACT_COL_PROCESS = "å·¥ç¨‹å��"
ACT_COL_OPERATOR = "æ‹…å½“è€…"
ACT_COL_START_DT = "é–‹å§‹æ—¥æ™‚"
ACT_COL_END_DT = "çµ‚äº†æ—¥æ™‚"
ACT_COL_START_ALT = "å®Ÿç¸¾é–‹å§‹"
ACT_COL_END_ALT = "å®Ÿç¸¾çµ‚äº†"
ACT_COL_DAY = "æ—¥ä»˜"
ACT_COL_TIME_START = "é–‹å§‹æ™‚åˆ»"
ACT_COL_TIME_END = "çµ‚äº†æ™‚åˆ»"
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

# --- 2æ®µéšŽå‡¦ç�†: æ®µéšŽ1æŠ½å‡º â†’ ãƒ–ãƒƒã‚¯ã€Œé…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€�ç·¨é›† â†’ æ®µéšŽ2è¨ˆç”» ---
STAGE1_OUTPUT_FILENAME = "plan_input_tasks.xlsx"
PLAN_INPUT_SHEET_NAME = os.environ.get("TASK_PLAN_SHEET", "").strip() or "é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›"
PLAN_COL_SPEED_OVERRIDE = "åŠ å·¥é€Ÿåº¦_ä¸Šæ›¸ã��"
# ç©ºç™½ã�®ã�¨ã��ã�¯åˆ—ã€ŒåŽŸå��æŠ•å…¥æ—¥ã€�ï¼ˆåŠ å·¥è¨ˆç”»DATA ç”±æ�¥ï¼‰ã‚’ã��ã�®ã�¾ã�¾ä½¿ã�†ã€‚æ—¥ä»˜ã�‚ã‚Šã�®ã�¨ã��ã�¯é…�å�°ã�®åŽŸå��åˆ¶ç´„ãƒ»çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§è¡¨ç¤ºã�®ä¸¡æ–¹ã�§ã�“ã�®æ—¥ä»˜ã‚’æŽ¡ç”¨ã€‚
PLAN_COL_RAW_INPUT_DATE_OVERRIDE = "åŽŸå��æŠ•å…¥æ—¥_ä¸Šæ›¸ã��"
PLAN_COL_PREFERRED_OP = "æ‹…å½“OP_æŒ‡å®š"
PLAN_COL_SPECIAL_REMARK = "ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒ"
# å�‚ç…§åˆ—ã€Œï¼ˆå…ƒï¼‰é…�å�°ä¸�è¦�ã€�ã�¯ç½®ã�‹ã�ªã�„ï¼ˆå…ƒãƒ‡ãƒ¼ã‚¿ã�«ç›¸å½“ã�™ã‚‹ãƒžã‚¹ã‚¿åˆ—ã�Œç„¡ã�„ã�Ÿã‚�ï¼‰ã€‚
# ã‚»ãƒ«å€¤ã�®ä¾‹ï¼ˆé…�å�°ã�‹ã‚‰å¤–ã�™ï¼‰: Excel ã�® TRUE / æ•°å€¤ 1 / æ–‡å­—åˆ—ã€Œã�¯ã�„ã€�ã€Œyesã€�ã€Œtrueã€�ã€Œâ—‹ã€�ã€Œã€‡ã€�ã€Œâ—�ã€�ç­‰ã€‚
# ç©ºãƒ»FALSEãƒ»0ãƒ»ã€Œã�„ã�„ã�ˆã€�ç­‰ã�¯é…�å�°å¯¾è±¡ã€‚è©³ç´°ã�¯ _plan_row_exclude_from_assignmentã€‚
PLAN_COL_EXCLUDE_FROM_ASSIGNMENT = "é…�å�°ä¸�è¦�"
PLAN_COL_AI_PARSE = "AIç‰¹åˆ¥æŒ‡å®š_è§£æž�"
PLAN_COL_PROCESS_FACTOR = "åŠ å·¥å·¥ç¨‹ã�®æ±ºå®šãƒ—ãƒ­ã‚»ã‚¹ã�®å› å­�"
# 1ãƒ­ãƒ¼ãƒ«ã�‚ã�Ÿã‚Šã�®é•·ã�•ï¼ˆmï¼‰ã€‚é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã�«ã�®ã�¿å­˜åœ¨ï¼ˆåŠ å·¥è¨ˆç”»DATA ã�«ã�¯ç„¡ã�„ï¼‰ã€‚è£½å“�å��åˆ—ã�®å�³éš£ã�«é…�ç½®ã€‚
PLAN_COL_ROLL_UNIT_LENGTH = "ãƒ­ãƒ¼ãƒ«å�˜ä½�é•·ã�•"
DEBUG_TASK_ID = os.environ.get("DEBUG_TASK_ID", "Y3-26").strip()
# ä¾‹: set TRACE_TEAM_ASSIGN_TASK_ID=W3-14 â€¦ é…�å�°ãƒ«ãƒ¼ãƒ—ã�§ã€Œäººæ•°åˆ¥ã�®æœ€è‰¯å€™è£œã€�ã�¨æŽ¡ç”¨ç�†ç”±ã‚’ INFO ãƒ­ã‚°ã�«å‡ºã�™
TRACE_TEAM_ASSIGN_TASK_ID = os.environ.get("TRACE_TEAM_ASSIGN_TASK_ID", "").strip()
# é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹å¯¾è±¡ã�¯ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œè¨­å®šã€�ã‚·ãƒ¼ãƒˆ A åˆ— 3 è¡Œç›®ä»¥é™�ã�®ã�¿ï¼ˆgenerate_plan å†’é ­ã�§ç¢ºå®šï¼‰ã€‚ç’°å¢ƒå¤‰æ•°ã�¯ä½¿ã‚�ã�ªã�„ã€‚
TRACE_SCHEDULE_TASK_IDS: frozenset[str] = frozenset()
# æ®µéšŽ2ãƒ‡ãƒ�ãƒƒã‚°é…�å�°: ã€Œè¨­å®šã€�B åˆ— 3 è¡Œç›®ä»¥é™�ã�«ä¾�é ¼NOã�Œã�‚ã‚‹ã�¨ã��ã�®ã�¿ã€�ã��ã�®ä¾�é ¼ã�®è¡Œã� ã�‘é…�å�°ï¼ˆgenerate_plan å†’é ­ã�§ç¢ºå®šï¼‰ã€‚ç©ºã�ªã‚‰å…¨ä»¶ã€‚
DEBUG_DISPATCH_ONLY_TASK_IDS: frozenset[str] = frozenset()
# ç´�æœŸè¶…é�Žãƒªãƒˆãƒ©ã‚¤ã�®å¤–å�´ãƒ©ã‚¦ãƒ³ãƒ‰ï¼ˆ0=åˆ�å›žã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼é€šã�—ã€�ä»¥é™�ã�¯ while å…ˆé ­ã�§æ›´æ–°ï¼‰ã€‚é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹å‡ºåŠ›ã�®ãƒ•ã‚¡ã‚¤ãƒ«å��ãƒ»æŽ¥é ­è¾žã�«ä½¿ç”¨ã€‚
DISPATCH_TRACE_OUTER_ROUND: int = 0


def _trace_schedule_task_enabled(task_id) -> bool:
    if not TRACE_SCHEDULE_TASK_IDS:
        return False
    return str(task_id or "").strip() in TRACE_SCHEDULE_TASK_IDS


def _sanitize_dispatch_trace_filename_part(task_id: str) -> str:
    """ä¾�é ¼NOã‚’ log ãƒ•ã‚¡ã‚¤ãƒ«å��ã�«ä½¿ã�†ã�Ÿã‚�ã�®ç°¡æ˜“ã‚µãƒ‹ã‚¿ã‚¤ã‚ºï¼ˆWindows ç¦�æ­¢æ–‡å­—ã‚’é�¿ã�‘ã‚‹ï¼‰ã€‚"""
    s = "".join(
        c if (c.isalnum() or c in "-_.") else "_"
        for c in str(task_id or "").strip()
    )
    return s[:120] if s else "task"


def _reset_dispatch_trace_per_task_logfiles() -> None:
    """
    æ®µéšŽ2å®Ÿè¡Œã�®å†’é ­ã�§1å›žã€�log å†…ã�® dispatch_trace_*.txt ã‚’ã�™ã�¹ã�¦å‰Šé™¤ã�™ã‚‹ï¼ˆé�ŽåŽ»å®Ÿè¡Œã�®æ®‹éª¸ã‚’æ®‹ã�•ã�ªã�„ï¼‰ã€‚
    å�„å¤–å�´ãƒ©ã‚¦ãƒ³ãƒ‰ç”¨ãƒ•ã‚¡ã‚¤ãƒ«ã�¯ generate_plan ã�® while å…ˆé ­ã�§ _dispatch_trace_begin_outer_round ã�Œãƒ˜ãƒƒãƒ€ä»˜ã��æ–°è¦�ä½œæˆ�ã�™ã‚‹ã€‚
    execution_log.txt ã�¨ã�¯åˆ¥ãƒ•ã‚¡ã‚¤ãƒ«ã€‚å†…å®¹ã�¯ [é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=â€¦] è¡Œã‚’ _log_dispatch_trace_schedule ã�§è¿½è¨˜
    ï¼ˆæ—¥æ¬¡æ®‹ãƒ»ãƒ­ãƒ¼ãƒ«ç¢ºå®šã�®ä½™å‰°æœ‰ç„¡ãƒ»ä½™åŠ›è¿½è¨˜ãƒ»çµ‚äº†æ™‚ã‚µãƒžãƒªç­‰ï¼‰ã€‚
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
    """ç´�æœŸè¶…é�Žãƒªãƒˆãƒ©ã‚¤ã�®å¤–å�´ãƒ©ã‚¦ãƒ³ãƒ‰ç•ªå�·ã‚’ç¢ºå®šã�—ã€�å½“ãƒ©ã‚¦ãƒ³ãƒ‰ç”¨ dispatch_trace_*_rNN.txt ã�®ãƒ˜ãƒƒãƒ€ã‚’1å›žã� ã�‘æ›¸ã��ã€‚"""
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
                    "# é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ï¼ˆä¾�é ¼NOã�”ã�¨ãƒ»å¤–å�´ãƒ©ã‚¦ãƒ³ãƒ‰åˆ¥ï¼‰ã€‚å�Œä¸€è¡Œã�¯ log/execution_log.txt ã�«ã‚‚å‡ºåŠ›ã�•ã‚Œã�¾ã�™ã€‚\n"
                    f"# task_id={t}  outer_round={DISPATCH_TRACE_OUTER_ROUND}  "
                    "# ï¼ˆ0=åˆ�å›žã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼é€šã�—ã€�ä»¥é™�ã�¯ç´�æœŸè¶…é�Žãƒªãƒˆãƒ©ã‚¤ã�”ã�¨ã�« +1ï¼‰\n\n"
                )
        except OSError as ex:
            logging.warning("dispatch_trace ãƒ­ã‚°ã�®åˆ�æœŸåŒ–ã�«å¤±æ•—: %s (%s)", path, ex)


def _log_dispatch_trace_schedule(task_id, msg: str, *args) -> None:
    """[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=â€¦] ã‚’ execution_log ã�«å‡ºã�—ã�¤ã�¤ã€�å¯¾è±¡ä¾�é ¼NOå°‚ç”¨ãƒ•ã‚¡ã‚¤ãƒ«ã�«ã‚‚è¿½è¨˜ã�™ã‚‹ã€‚"""
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
            logging.warning("dispatch_trace å�´ãƒ•ã‚¡ã‚¤ãƒ«ã�¸ã�®è¿½è¨˜ã�«å¤±æ•—: %s (%s)", path, ex)
        except Exception:
            pass


# True: å¾“æ�¥ã�®ã€Œäººæ•°æœ€å„ªå…ˆã€�ã‚¿ãƒ—ãƒ« (-äººæ•°, é–‹å§‹, -å�˜ä½�æ•°, å„ªå…ˆåº¦å�ˆè¨ˆ)ã€‚False ã�®ã�¨ã��ä¸‹è¨˜ã‚¹ãƒ©ãƒƒã‚¯åˆ†ã�¨çµ„ã�¿å�ˆã‚�ã�›
TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF = os.environ.get(
    "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF", "0"
).strip().lower() not in ("0", "false", "no", "off", "ã�„ã�„ã�ˆ")


def _team_assign_start_slack_wait_minutes() -> int:
    """å…¨æ—¥å€™è£œã�®æœ€æ—©é–‹å§‹ã�‹ã‚‰ã�“ã�®åˆ†ä»¥å†…ã�®é�…ã‚Œã�ªã‚‰ã€�é–‹å§‹ã‚ˆã‚Šäººæ•°ã‚’å„ªå…ˆï¼ˆåˆ†ï¼‰ã€‚0 ã�§ç„¡åŠ¹ã€‚"""
    raw = os.environ.get("TEAM_ASSIGN_START_SLACK_WAIT_MINUTES", "60").strip()
    try:
        v = int(raw)
    except ValueError:
        v = 60
    return max(0, v)


TEAM_ASSIGN_START_SLACK_WAIT_MINUTES = _team_assign_start_slack_wait_minutes()

# True ã�®ã�¨ã�� need ã‚·ãƒ¼ãƒˆã€Œé…�å�°æ™‚è¿½åŠ äººæ•°ã€�è¡Œã‚’ç„¡è¦–ã�—ã€�ãƒ�ãƒ¼ãƒ äººæ•°ã�¯åŸºæœ¬å¿…è¦�äººæ•°ï¼ˆreq_numï¼‰ã�®ã�¿è©¦è¡Œã�—ã€�ãƒ¡ã‚¤ãƒ³å¾Œè¿½è¨˜ã‚‚ã�—ã�ªã�„ã€‚
TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW = (
    os.environ.get("TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW", "0")
    .strip()
    .lower()
    in ("1", "true", "yes", "on", "ã�¯ã�„")
)

# True: å¾“æ�¥ã�©ã�Šã‚Šãƒ¡ã‚¤ãƒ³å‰²ä»˜ã�®çµ„ã�¿å�ˆã‚�ã�›æŽ¢ç´¢ã�§ req_numã€œreq_num+è¿½åŠ äººæ•°ä¸Šé™�ã�¾ã�§è©¦ã�™ã€‚
# Falseï¼ˆæ—¢å®šï¼‰: ãƒ¡ã‚¤ãƒ³ã�¯ req_num ã�®ã�¿ã€‚è¿½åŠ äººæ•°ä¸Šé™�ã�¯å…¨ã‚·ãƒŸãƒ¥ãƒ¬ãƒ¼ã‚·ãƒ§ãƒ³å®Œäº†å¾Œã€�å½“è©²ãƒ–ãƒ­ãƒƒã‚¯æ™‚é–“ã�«
#     ä»–ã‚¿ã‚¹ã‚¯ã�¸æœªå‰²å½“ï¼ˆæ™‚é–“é‡�ã�ªã‚Šã�ªã�—ï¼‰ã�‹ã�¤ skills é�©å�ˆã�®è€…ã‚’ã‚µãƒ–ã�¨ã�—ã�¦è¿½è¨˜ï¼ˆappend_surplus_staff_after_main_dispatchï¼‰ã€‚
TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS = (
    os.environ.get("TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS", "")
    .strip()
    .lower()
    in ("1", "true", "yes", "on", "ã�¯ã�„")
)

# Trueï¼ˆæ—¢å®šï¼‰: ãƒ¡ã‚¤ãƒ³é…�å�°ã�®å¿…è¦�äººæ•°ã�¯ needï¼ˆåŸºæœ¬å¿…è¦�äººæ•°ï¼‹ç‰¹åˆ¥æŒ‡å®šï¼‰ã�®ã�¿ã€‚
# False ã�®ã�¨ã��ã�¯ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ AI ã�® required_op ã�®ã�¿è¨ˆç”»å�´ã�‹ã‚‰å�‚ç…§ã�—å¾—ã‚‹ï¼ˆã‚·ãƒ¼ãƒˆåˆ—ã€Œå¿…è¦�äººæ•°ã€�ã�¯å»ƒæ­¢æ¸ˆã�¿ï¼‰ã€‚
TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY = (
    os.environ.get("TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "ã�„ã�„ã�ˆ")
)
# Trueï¼ˆæ—¢å®šï¼‰: masterã€Œçµ„ã�¿å�ˆã‚�ã�›è¡¨ã€�ã�«è©²å½“è¡Œã�Œã�‚ã‚‹å·¥ç¨‹+æ©Ÿæ¢°ã�¯ã€�çµ„å�ˆã�›å„ªå…ˆåº¦ã�®æ˜‡é †ã�§
# æœ€åˆ�ã�«æˆ�ç«‹ã�—ã�Ÿãƒ¡ãƒ³ãƒ�ãƒ¼ç·¨æˆ�ã‚’æŽ¡ç”¨ã€‚ã�™ã�¹ã�¦ä¸�å�¯ã�ªã‚‰å¾“æ�¥ã�® itertools çµ„å�ˆã�›æŽ¢ç´¢ã€‚
TEAM_ASSIGN_USE_MASTER_COMBO_SHEET = (
    os.environ.get("TEAM_ASSIGN_USE_MASTER_COMBO_SHEET", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "ã�„ã�„ã�ˆ")
)

# Â§B-2 ç†±èž�ç�€æ¤œæŸ»ã‚’å�Œä¸€è¨­å‚™ï¼ˆå·¥ç¨‹åˆ—ã‚­ãƒ¼ï¼‰ã�§ã€Œé–‹å§‹æ¸ˆã�¿1ä»¶ã�«æ®‹ãƒ­ãƒ¼ãƒ«ã�Œã�‚ã‚‹é–“ã�¯ä»–ä¾�é ¼ã�®æ¤œæŸ»ã‚’è©¦ã�•ã�ªã�„ã€�ã�‹ã€‚
# 0 / false / no / off ã�§ç„¡åŠ¹ã�«ã�™ã‚‹ã�¨è¨­å‚™æ™‚é–“å‰²ä¸Šã�§ä¾�é ¼ã�Œæ··åœ¨ã�—å¾—ã‚‹ã�Œã€�å� æœ‰ã�«ã‚ˆã‚‹é•·æœŸãƒ–ãƒ­ãƒƒã‚¯ï¼ˆä¾‹: W3-14 åž‹ï¼‰ã‚’é�¿ã�‘ã‚‰ã‚Œã‚‹ã€‚
PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE = (
    os.environ.get("PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "ã�„ã�„ã�ˆ", "ç„¡åŠ¹")
)

# Â§B-2 / Â§B-3 å�Œä¸€ä¾�é ¼ã�§ EC ã�¨å¾Œç¶šï¼ˆæ¤œæŸ»ï¼�å·»è¿”ã�—ï¼‰ã�®æ‹…å½“è€…é›†å�ˆã‚’æŽ’ä»–ã�™ã‚‹ã�‹ã€‚
# 0 / false / no / off / ã�„ã�„ã�ˆ / ç„¡åŠ¹ ã�§ç„¡åŠ¹åŒ–ã�™ã‚‹ã�¨ã€�å±¥æ­´ãƒ™ãƒ¼ã‚¹ã�®ç›¸äº’é™¤å¤–ã‚’è¡Œã‚�ã�šå�Œä¸€äººç‰©ã�Œä¸¡å�´ã�®å€™è£œã�«æ®‹ã‚Šå¾—ã‚‹ã€‚
PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS = (
    os.environ.get("PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "ã�„ã�„ã�ˆ", "ç„¡åŠ¹")
)

# ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�: æ—¢å®šã�§ã�¯ openpyxl save ã‚’è©¦ã�•ã�š xlwings å�ŒæœŸâ†’Saveï¼ˆExcel å� æœ‰æ™‚ã�¯ openpyxl ã�Œå®Ÿè³ªå¤±æ•—ã�™ã‚‹ã�Ÿã‚�ï¼‰ã€‚å¤±æ•—æ™‚ã�¯ TSVâ†’VBA å��æ˜ ã€‚
# ã‚³ãƒžãƒ³ãƒ‰ç­‰ã�§ openpyxl ã‚’è©¦ã�™å ´å�ˆã�¯ EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1ã€‚
EXCLUDE_RULES_SHEET_NAME = "è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹"
EXCLUDE_RULES_SKIP_OPENPYXL_SAVE = os.environ.get(
    "EXCLUDE_RULES_TRY_OPENPYXL_SAVE", ""
).strip().lower() not in ("1", "true", "yes", "on")
EXCLUDE_RULE_COL_PROCESS = "å·¥ç¨‹å��"
EXCLUDE_RULE_COL_MACHINE = "æ©Ÿæ¢°å��"
EXCLUDE_RULE_COL_FLAG = "é…�å�°ä¸�è¦�"
EXCLUDE_RULE_COL_LOGIC_JA = "é…�å�°ä¸�èƒ½ãƒ­ã‚¸ãƒƒã‚¯"
EXCLUDE_RULE_COL_LOGIC_JSON = "ãƒ­ã‚¸ãƒƒã‚¯å¼�"
# å…ƒãƒ–ãƒƒã‚¯ã�Œãƒ­ãƒƒã‚¯ã�•ã‚Œåˆ¥å��ä¿�å­˜ã�—ã�Ÿå ´å�ˆã€�å�Œä¸€ãƒ—ãƒ­ã‚»ã‚¹å†…ã�®ãƒ«ãƒ¼ãƒ«èª­è¾¼ã�¯ã�“ã�®ãƒ‘ã‚¹ã‚’å„ªå…ˆ
_exclude_rules_effective_read_path: str | None = None
# ç›´å¾Œã�® apply_exclude_rulesï¼ˆå�Œä¸€ãƒ—ãƒ­ã‚»ã‚¹ï¼‰ç”¨: VBA å��æ˜ å‰�ã�§ã‚‚ E åˆ—ä»˜ã��ãƒ«ãƒ¼ãƒ«ã‚’ä½¿ã�†
_exclude_rules_rules_snapshot: list | None = None
_exclude_rules_snapshot_wb: str | None = None
# ãƒ«ãƒ¼ãƒ« JSON ã�® conditions ã�§å�‚ç…§å�¯èƒ½ã�ªåˆ—ï¼ˆAI ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆã�¨è©•ä¾¡å™¨ã‚’ä¸€è‡´ã�•ã�›ã‚‹ï¼‰
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

# è¨ˆç”»çµ�æžœãƒ–ãƒƒã‚¯ã€Œçµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã€�ã�®åˆ—é †ãƒ»è¡¨ç¤ºï¼ˆãƒžã‚¯ãƒ­å®Ÿè¡Œãƒ–ãƒƒã‚¯ã�®å�Œå��ã‚·ãƒ¼ãƒˆã�§ä¸Šæ›¸ã��å�¯ï¼‰
RESULT_TASK_SHEET_NAME = "çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§"
RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME = "çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²"
# ä½™åŠ›è¿½è¨˜å‰�ã�®ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ã‚’å�¯è¦–åŒ–ï¼ˆçµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ã�¨å�Œã�˜ 10 åˆ†æž ãƒ»åˆ—æ§‹é€ ï¼‰
TEMP_EQUIPMENT_SCHEDULE_SHEET_NAME = "TEMP_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²"
# è¨­å‚™ãƒ»äººã�®å� æœ‰ï¼ˆãƒ–ãƒ­ãƒƒã‚¯ï¼‰ã‚’ 10 åˆ†æž ã�§ä¸€è¦§ï¼ˆèª¿æŸ»ãƒ»æ¤œè¨¼ç”¨ï¼‰
BLOCK_TABLE_SHEET_NAME = "ãƒ–ãƒ­ãƒƒã‚¯ãƒ†ãƒ¼ãƒ–ãƒ«"
# å·¥ç¨‹å��+æ©Ÿæ¢°ã�®è¤‡å�ˆåˆ—ã�§ã�¯ã�ªã��ã€�æ©Ÿæ¢°å��å�˜ä½�ã�§å�„æž ã�®ä¾�é ¼NOã‚’æŠŠæ�¡ã�—ã‚„ã�™ã��ã�™ã‚‹
RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME = "çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²_æ©Ÿæ¢°å��æ¯Ž"
# master ãƒ¡ã‚¤ãƒ³ A15/B15 ã�®å®šå¸¸å¤–ã�®ã€Œæ—¥æ™‚å¸¯ã€�è¦‹å‡ºã�—ç�€è‰²ï¼ˆçµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ãƒ»çµ�æžœ_è¨­å‚™ã‚¬ãƒ³ãƒˆï¼‰
RESULT_OUTSIDE_REGULAR_TIME_FILL = "FCE4D6"
# çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²_æ©Ÿæ¢°å��æ¯Ž: é…�å�°æ¸ˆã�¿ä¾�é ¼NOã‚»ãƒ«ï¼ˆæ©Ÿæ¢°åˆ—ï¼‰ã�®è–„ã�„ã‚°ãƒªãƒ¼ãƒ³
# çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ï¼ˆã�Šã‚ˆã�³ TEMPï¼‰: åŠ å·¥å‰�æº–å‚™ãƒ»ä¾�é ¼åˆ‡æ›¿å¾Œå§‹æœ«ã�®è¨­å‚™ã‚»ãƒ«ã‚‚å�Œç³»è‰²
RESULT_DISPATCHED_REQUEST_FILL = "C6EFCE"
# çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²: masterã€Œæ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã€�å� æœ‰ã�¨é‡�ã�ªã‚‹è¨­å‚™ã‚»ãƒ«ï¼ˆ10åˆ†æž ï¼‰
RESULT_MACHINE_CALENDAR_BLOCK_FILL = "D4B3E8"
# çµ�æžœ_è¨­å‚™ã‚¬ãƒ³ãƒˆ: æ©Ÿæ¢°å��ã‚°ãƒ«ãƒ¼ãƒ—ï¼ˆæ©Ÿæ¢°å��åˆ—ã�®å�Œä¸€å��ç§°ï¼‰ã�”ã�¨ã�« Bã€œE åˆ—ã‚’åŒºåˆ¥ã�™ã‚‹æ·¡è‰²ï¼ˆé †ã�«å‰²å½“ãƒ»å¾ªç’°ï¼‰
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
# é…�å�°ã‚·ãƒŸãƒ¥ãƒ¬ãƒ¼ã‚·ãƒ§ãƒ³é–‹å§‹å‰�ï¼ˆåˆ�å›ž task_queue.sort å¾Œï¼‰ã�®ã‚­ãƒ¥ãƒ¼é †ã€‚1 å§‹ã�¾ã‚Šãƒ»å…¨æ—¥ç¨‹ã�§ä¸�å¤‰
RESULT_TASK_COL_DISPATCH_TRIAL_ORDER = "é…�å�°è©¦è¡Œé †ç•ª"
# é…�å®Œ_åŠ å·¥çµ‚äº†ã�Œã€Œå›žç­”ç´�æœŸ+16:00ã€�ã�¾ã�Ÿã�¯ã€ŒæŒ‡å®šç´�æœŸ+16:00ã€�ï¼ˆå›žç­”ã�Œç©ºã�®ã�¨ã��ï¼‰ä»¥å‰�ã�‹ã‚’è¡¨ç¤º
RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16 = "é…�å®Œ_å›žç­”æŒ‡å®š16æ™‚ã�¾ã�§"
# ãƒžã‚¹ã‚¿ skills ã�®å·¥ç¨‹+æ©Ÿæ¢°åˆ—ã�”ã�¨ã�® OP/AS å‰²å½“å�‚è€ƒé †ï¼ˆå„ªå…ˆåº¦å€¤ãƒ»æ°�å��é †ï¼‰ã�¨ãƒ�ãƒ¼ãƒ æŽ¡ç”¨ãƒ«ãƒ¼ãƒ«ã�®èª¬æ˜Ž
RESULT_MEMBER_PRIORITY_SHEET_NAME = "çµ�æžœ_äººå“¡é…�å�°å„ªå…ˆé †"
COLUMN_CONFIG_SHEET_NAME = "åˆ—è¨­å®š_çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§"
COLUMN_CONFIG_HEADER_COL = "åˆ—å��"
COLUMN_CONFIG_VISIBLE_COL = "è¡¨ç¤º"
# æ®µéšŽ2ã�®çµ�æžœ xlsx ç”Ÿæˆ�å¾Œã€�å…¥åŠ›ãƒ–ãƒƒã‚¯ã�®åˆ—è¨­å®šã‚·ãƒ¼ãƒˆä¸Šã�®å›³å½¢ï¼ˆãƒ•ã‚©ãƒ¼ãƒ ãƒœã‚¿ãƒ³ç­‰ï¼‰ã‚’ xlwings ã�§è¤‡è£½ã�™ã‚‹ï¼ˆæ—¢å®š ONã€‚ç„¡åŠ¹åŒ–ã�¯ STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT=0ï¼‰
STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT = os.environ.get(
    "STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT", "1"
).strip().lower() in ("1", "true", "yes", "on")
# çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®æ—¥ä»˜ç³»ï¼ˆyyyy/mm/dd æ–‡å­—åˆ—ï¼‰ã�«ä»˜ã�‘ã‚‹ãƒ•ã‚©ãƒ³ãƒˆè‰²ã€‚å±¥æ­´åˆ—ã�®ã€�æ—¥ä»˜ã€‘ã�¨æ�ƒã�ˆã‚‹
RESULT_TASK_DATE_STYLE_HEADERS = frozenset(
    {
        "å›žç­”ç´�æœŸ",
        "æŒ‡å®šç´�æœŸ",
        "è¨ˆç”»åŸºæº–ç´�æœŸ",
        TASK_COL_RAW_INPUT_DATE,
        "åŠ å·¥é–‹å§‹æ—¥",
        "é…�å®Œ_åŠ å·¥é–‹å§‹",
        "é…�å®Œ_åŠ å·¥çµ‚äº†",
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
# çŸ›ç›¾æ¤œå‡ºã�§ãƒªã‚»ãƒƒãƒˆå¯¾è±¡ã�«ã�™ã‚‹åˆ—ï¼ˆè¦‹å‡ºã�—è¡Œã�®æ–‡è¨€ã�¨ä¸€è‡´ã�™ã‚‹ã�“ã�¨ï¼‰
PLAN_CONFLICT_STYLABLE_COLS = tuple(PLAN_OVERRIDE_COLUMNS)
# æ®µéšŽ1å†�æŠ½å‡ºæ™‚ã€�æ—¢å­˜ã€Œé…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€�ã�‹ã‚‰ç¶™æ‰¿ã�™ã‚‹åˆ—ï¼ˆAIã�®è§£æž�çµ�æžœåˆ—ã�¯æ¯Žå›žç©ºã�«æˆ»ã�™ï¼‰
PLAN_STAGE1_MERGE_COLUMNS = tuple(c for c in PLAN_OVERRIDE_COLUMNS if c != PLAN_COL_AI_PARSE)
# ä¸Šæ›¸ã��ä»¥å¤–ã�§ã€�å†�æŠ½å‡ºæ™‚ã�«æ—§ã‚·ãƒ¼ãƒˆã�‹ã‚‰å¼•ã��ç¶™ã��åˆ—ï¼ˆã‚»ãƒ«ã�Œç©ºã�§ã�ªã�„ã�¨ã��ã�®ã�¿ï¼‰
# é…�å�°è©¦è¡Œé †ç•ªã�¯æ¯Žå›žç©ºã‚¯ãƒªã‚¢ã�®ã�†ã�ˆ fill_plan_dispatch_trial_order_column_stage1 ã�§ä»˜ã�‘ç›´ã�™ã�Ÿã‚�å¯¾è±¡å¤–ã€‚
PLAN_STAGE1_MERGE_EXTRA_COLUMNS = (PLAN_COL_ROLL_UNIT_LENGTH,)
# openpyxl ä¿�å­˜ã�Œãƒ–ãƒƒã‚¯ãƒ­ãƒƒã‚¯ã�§å¤±æ•—ã�—ã�Ÿã�¨ã��ã€�VBA ã�Œé–‹ã�„ã�¦ã�„ã‚‹ãƒ–ãƒƒã‚¯ã�¸æ›¸å¼�é�©ç”¨ã�™ã‚‹ã�Ÿã‚�ã�®æŒ‡ç¤ºãƒ•ã‚¡ã‚¤ãƒ«
PLANNING_CONFLICT_SIDECAR = "planning_conflict_highlight.tsv"
# é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã�¸ã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�ã€�ã‚’æ›¸ã��åˆ—ï¼ˆè¡¨ã�®å�³ç«¯ã‚ˆã‚Šå¤–å�´ã€‚1è¡Œç›®ã�‹ã‚‰ç¸¦ã�«ãƒ©ãƒ™ãƒ«ï¼�å€¤ï¼‰
# â˜… å�‚ç…§è¡¨ç¤ºã�®ã�¿: load_planning_tasks_df ç­‰ã�¯æœ¬åˆ—ã‚’ä¸€åˆ‡èª­ã�¾ã�ªã�„ã€‚é…�å�°é�©ç”¨ã�¯å¸¸ã�«ãƒ¡ã‚¤ãƒ³ã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�1çµŒè·¯ã�®ã�Ÿã‚�äºŒé‡�é�©ç”¨ã�«ã�ªã‚‰ã�ªã�„ã€‚
PLAN_SHEET_GLOBAL_PARSE_LABEL_COL = 50  # AX
PLAN_SHEET_GLOBAL_PARSE_VALUE_COL = 51  # AY
PLAN_SHEET_GLOBAL_PARSE_MAX_ROWS = 42


def plan_reference_column_name(override_col: str) -> str:
    """ä¸Šæ›¸ã��åˆ—ã�®å·¦éš£ã�«ç½®ã��å�‚ç…§åˆ—ã�®è¦‹å‡ºã�—ï¼ˆã‚»ãƒ«å€¤ã�¯æ‹¬å¼§ä»˜ã��ã�§å…ƒãƒ‡ãƒ¼ã‚¿ã‚’è¡¨ç¤ºï¼‰ã€‚"""
    return f"ï¼ˆå…ƒï¼‰{override_col}"


def plan_input_sheet_column_order():
    """
    é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã�®åˆ—é †ï¼ˆæ®µéšŽ1å‡ºåŠ›ãƒ»æ®µéšŽ2èª­è¾¼ã�§å…±é€šï¼‰ã€‚

    0. é…�å�°è©¦è¡Œé †ç•ªï¼ˆæ®µéšŽ1æŠ½å‡ºç›´å¾Œã�«ç©ºã‚¯ãƒªã‚¢â†’æ®µéšŽ2ã�¨å�Œè¶£æ—¨ã�«ä»˜ä¸Žã€‚æ®µéšŽ2ã�¯å…¨è¡Œã�«å€¤ã�Œã�‚ã‚‹ã�¨ã��ã�“ã�®é †ã‚’å„ªå…ˆï¼‰
    1. é…�å�°ä¸�è¦�ï¼ˆå�‚ç…§åˆ—ã�ªã�—ï¼‰
    2. åŠ å·¥è¨ˆç”»DATA ç”±æ�¥ï¼ˆSOURCE_BASE_COLUMNSï¼‰â€¦ ä¾�é ¼NOã€œå®Ÿå‡ºæ�¥é«˜ã�¾ã�§ï¼ˆè£½å“�å��ã�®ç›´å¾Œã�«ãƒ­ãƒ¼ãƒ«å�˜ä½�é•·ã�•ã€�åŽŸå��æŠ•å…¥æ—¥ã�®ç›´å¾Œã�«åœ¨åº«å ´æ‰€ï¼‰
    3. åŠ å·¥å·¥ç¨‹ã�®æ±ºå®šãƒ—ãƒ­ã‚»ã‚¹ã�®å› å­�
    4. ä¸Šæ›¸ã��åˆ—â€¦ å�„åˆ—ã�®ç›´å‰�ã�«ã€Œï¼ˆå…ƒï¼‰â€¦ã€�å�‚ç…§åˆ—ã€‚AIç‰¹åˆ¥æŒ‡å®š_è§£æž�ã�®ã�¿å�‚ç…§åˆ—ã�ªã�—ã€‚
       ï¼ˆæ—¥ä»˜ç³»ä¸Šæ›¸ã��ã�« åŽŸå��æŠ•å…¥æ—¥_ä¸Šæ›¸ã�� ã‚’å�«ã‚€ã€‚ç©ºç™½æ™‚ã�¯åˆ—ã€ŒåŽŸå��æŠ•å…¥æ—¥ã€�ã‚’é…�å�°ã�«ä½¿ç”¨ï¼‰

    global_speed_rules ç­‰ã�§å¤‰ã‚�ã‚‹å®ŸåŠ¹é€Ÿåº¦ã�¯ã‚·ãƒ¼ãƒˆåˆ—ã�§ã�¯æŒ�ã�Ÿã�šã€�é…�å�°å†…éƒ¨ã�®ã�¿ã�§å��æ˜ ã�™ã‚‹ã€‚
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
    """å�‚ç…§è¡¨ç¤ºç”¨: ç©ºã�¯ï¼ˆâ€•ï¼‰ã€�æ—¥ä»˜ãƒ»ã��ã�®ä»–ã�¯ï¼ˆå€¤ï¼‰ã€‚"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "ï¼ˆâ€•ï¼‰"
    if isinstance(val, datetime):
        d = val.date() if hasattr(val, "date") else val
        if isinstance(d, date):
            return f"ï¼ˆ{d.year}/{d.month}/{d.day}ï¼‰"
    if isinstance(val, date):
        return f"ï¼ˆ{val.year}/{val.month}/{val.day}ï¼‰"
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none"):
        return "ï¼ˆâ€•ï¼‰"
    return f"ï¼ˆ{s}ï¼‰"


def _reference_text_for_override_row(row, override_col: str, req_map: dict, need_rules: list) -> str:
    """1è¡Œåˆ†ã�®ä¸Šæ›¸ã��åˆ—ã�«å¯¾å¿œã�™ã‚‹å�‚ç…§æ–‡è¨€ï¼ˆæ‹¬å¼§ä»˜ã��ï¼‰ã€‚"""
    _ = (req_map, need_rules)  # æ—§ã€Œï¼ˆå…ƒï¼‰å¿…è¦�äººæ•°ã€�å�‚ç…§ã�§ä½¿ç”¨ã€‚åˆ—å»ƒæ­¢ã�«ã‚ˆã‚Šæœªä½¿ç”¨ã� ã�Œå‘¼ã�³å‡ºã�—äº’æ�›ã�®ã�Ÿã‚�æ®‹ã�™ã€‚
    if override_col == PLAN_COL_SPEED_OVERRIDE:
        v = row.get(TASK_COL_SPEED)
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "ï¼ˆâ€•ï¼‰"
        try:
            x = float(v)
            if abs(x - round(x)) < 1e-9:
                return f"ï¼ˆ{int(round(x))}ï¼‰"
            return f"ï¼ˆ{x}ï¼‰"
        except (TypeError, ValueError):
            return _format_paren_ref_scalar(v)
    if override_col in (PLAN_COL_PREFERRED_OP, PLAN_COL_SPECIAL_REMARK):
        return "ï¼ˆâ€•ï¼‰"
    if override_col == PLAN_COL_RAW_INPUT_DATE_OVERRIDE:
        return _format_paren_ref_scalar(
            parse_optional_date(_planning_df_cell_scalar(row, TASK_COL_RAW_INPUT_DATE))
        )
    return "ï¼ˆâ€•ï¼‰"


def _refresh_plan_reference_columns(df, req_map: dict, need_rules: list):
    """åŠ å·¥è¨ˆç”»DATAï¼�need ã�«åŸºã�¥ã��ã€Œï¼ˆå…ƒï¼‰â€¦ã€�åˆ—ã‚’å†�è¨ˆç®—ï¼ˆãƒžãƒ¼ã‚¸å¾Œã�«å¿…ã�šå‘¼ã�¶ï¼‰ã€‚"""
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


def _apply_plan_input_visual_format(path: str, sheet_name: str = "ã‚¿ã‚¹ã‚¯ä¸€è¦§"):
    """ä¸Šæ›¸ã��å…¥åŠ›åˆ—ã�«è–„ã�„é»„è‰²ã‚’ä»˜ä¸Žï¼ˆå�‚ç…§åˆ—ã�¯æœªç�€è‰²ã€‚AIè§£æž�åˆ—ã�¯é™¤å¤–ï¼‰ã€‚"""
    # è¦‹å‡ºã�—æ–‡å­—ã�®è¡¨è¨˜ã‚†ã‚Œã�§åˆ—å��æ¤œç´¢ã�«å¤±æ•—ã�—ã�Œã�¡ã�ªã�Ÿã‚�ã€�æ®µéšŽ1ã�®åˆ—é †ï¼ˆplan_input_sheet_column_orderï¼‰ã�®
    # 1-based åˆ—ç•ªå�·ã�§å¡—ã‚‹ï¼ˆto_excel ã�®åˆ—é †ã�¨ä¸€è‡´ã�•ã�›ã‚‹ï¼‰ã€‚
    fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    order = plan_input_sheet_column_order()
    col_1based = {name: i + 1 for i, name in enumerate(order)}
    if _workbook_should_skip_openpyxl_io(path):
        logging.info(
            "é…�å�°è¨ˆç”»ã�®è¦–è¦šæ•´å½¢: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� openpyxl ã�§ã�®ç�€è‰²ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚",
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
    Excel ã�Œãƒ–ãƒƒã‚¯ã‚’é–‹ã�„ã�Ÿã�¾ã�¾ã�®ã�¨ã��ä¿�å­˜ã�§ã��ã�ªã�„å ´å�ˆã�«ã€�VBA ç”¨ã�® TSV ã‚’ log ã�«æ›¸ã��ã€‚
    å½¢å¼�: V1 / ã‚·ãƒ¼ãƒˆå�� / ãƒ‡ãƒ¼ã‚¿è¡Œæ•° / ã‚¯ãƒªã‚¢åˆ—ã‚’ã‚¿ãƒ–çµ�å�ˆ / ä»¥é™� è¡Œç•ªå�·\\tåˆ—å��
    """
    path = _planning_conflict_sidecar_path()
    clear_cols = "\t".join(PLAN_CONFLICT_STYLABLE_COLS)
    lines = ["V1", sheet_name, str(int(num_data_rows)), clear_cols]
    for r in sorted(conflicts_by_row.keys()):
        for name in sorted(conflicts_by_row[r]):
            lines.append(f"{int(r)}\t{name}")
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")

# æ®µéšŽ1å‡ºåŠ›ãƒ»ãƒ–ãƒƒã‚¯å†…ã�®æ—¥ä»˜åˆ—ã‚’ Excel ä¸Šã€Œæ—¥ä»˜ã�®ã�¿ã€�(æ™‚åˆ»ã�ªã�—è¡¨ç¤º) ã�«æ•´ã�ˆã‚‹
STAGE1_SHEET_DATEONLY_HEADERS = frozenset(
    {
        TASK_COL_ANSWER_DUE,
        TASK_COL_SPECIFIED_DUE,
        TASK_COL_RAW_INPUT_DATE,
        PLAN_COL_RAW_INPUT_DATE_OVERRIDE,
    }
)


def _result_font(**kwargs):
    """çµ�æžœãƒ–ãƒƒã‚¯ç”¨ Fontï¼ˆå‘¼ã�³å‡ºã�—å�´ã�Œ name/size ç­‰ã‚’æŒ‡å®šï¼‰ã€‚"""
    return Font(**kwargs)


def _output_book_font(bold=False):
    return _result_font(bold=bold)


def _apply_output_font_to_result_sheet(ws):
    """çµ�æžœ_* ã�®ã�†ã�¡ã‚¬ãƒ³ãƒˆä»¥å¤–å�‘ã�‘: æ—¢å®šãƒ•ã‚©ãƒ³ãƒˆãƒ»1è¡Œç›®å¤ªå­—ã�®ã�¿ï¼ˆåˆ—å¹…ã�¯ VBA AutoFitï¼‰ã€‚"""
    base = _output_book_font(bold=False)
    hdr = _output_book_font(bold=True)
    mr, mc = ws.max_row or 1, ws.max_column or 1
    for row in ws.iter_rows(min_row=1, max_row=mr, min_col=1, max_col=mc):
        for cell in row:
            cell.font = base
    for cell in ws[1]:
        cell.font = hdr


def _apply_excel_date_columns_date_only_display(path, sheet_name, header_names=None):
    """openpyxl: æŒ‡å®šãƒ˜ãƒƒãƒ€ãƒ¼åˆ—ã‚’ yyyy/mm/dd ã�®æ—¥ä»˜è¡¨ç¤ºã�«ã�™ã‚‹ï¼ˆæ™‚åˆ»ã‚’è¡¨ç¤ºã�—ã�ªã�„ï¼‰ã€‚"""
    from openpyxl import load_workbook

    headers = header_names or STAGE1_SHEET_DATEONLY_HEADERS
    if _workbook_should_skip_openpyxl_io(path):
        logging.info(
            "æ—¥ä»˜åˆ—è¡¨ç¤ºæ•´å½¢: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� openpyxl ã�§ã�®å‡¦ç�†ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚",
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
    `åŠ å·¥è¨ˆç”»DATA` ã‚·ãƒ¼ãƒˆã�® `ãƒ‡ãƒ¼ã‚¿æŠ½å‡ºæ—¥` ã�‹ã‚‰ datetime ã‚’å�–å¾—ã�™ã‚‹ã€‚
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
    `åŠ å·¥è¨ˆç”»DATA` ã‚·ãƒ¼ãƒˆã�® `ãƒ‡ãƒ¼ã‚¿æŠ½å‡ºæ—¥` ã�‹ã‚‰ã€Œãƒ‡ãƒ¼ã‚¿å�¸å‡ºã�—æ—¥æ™‚ã€�ã‚’å�–å¾—ã�—ã�¦æ–‡å­—åˆ—åŒ–ã�™ã‚‹ã€‚
    """
    try:
        dt = _extract_data_extraction_datetime()
        if dt is None:
            return "â€”"
        return dt.strftime("%Y/%m/%d %H:%M:%S")
    except Exception:
        return "â€”"


def _weekday_jp(d):
    return "æœˆç�«æ°´æœ¨é‡‘åœŸæ—¥"[d.weekday()]


# ã‚¬ãƒ³ãƒˆã�®ä½œæ¥­ãƒ�ãƒ¼ï¼šã�„ã�šã‚Œã‚‚æ˜Žã‚‹ã�„åœ°è‰²ï¼‹é»’æ–‡å­—ã�Œèª­ã‚�ã‚‹ãƒˆãƒ¼ãƒ³ï¼ˆãƒ¢ãƒŽã‚¯ãƒ­å�°åˆ·ã�§ã‚‚æ¿ƒæ·¡ã�§è­˜åˆ¥ã�—ã‚„ã�™ã�„ï¼‰
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

# å®Ÿç¸¾ãƒ�ãƒ¼ç”¨ï¼ˆè¨ˆç”»ã�¨ä¸¦ã�¹ã�¦ã‚‚ãƒ¢ãƒŽã‚¯ãƒ­ã�§åŒºåˆ¥ã�—ã‚„ã�™ã�„ãƒˆãƒ¼ãƒ³ï¼‰
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

# è¨­å‚™ã‚¬ãƒ³ãƒˆ: æ—¥æ¬¡å§‹æ¥­æº–å‚™ï¼ˆmachine_daily_startupï¼‰ã�®å¸¯è‰²ï¼ˆé»„è‰²ç³»ï¼‰
_GANTT_DAILY_STARTUP_FILL = "FFEB9C"


def _gantt_bar_fill_for_task_id(task_id):
    """ä¾�é ¼NOã�”ã�¨ã�«ä¸Šè¨˜ãƒ‘ãƒ¬ãƒƒãƒˆã�‹ã‚‰1è‰²ï¼ˆRRGGBBï¼‰ã€‚æ¿ƒè‰²ï¼‹ç™½æ–‡å­—ã�®çµ„ã�¿å�ˆã‚�ã�›ã�¯ä½¿ã‚�ã�ªã�„ã€‚"""
    h = hashlib.md5(str(task_id).encode("utf-8")).hexdigest()
    i = int(h[:8], 16) % len(_GANTT_BAR_FILLS_PRINT_SAFE)
    return _GANTT_BAR_FILLS_PRINT_SAFE[i]


def _gantt_bar_fill_actual_for_task_id(task_id):
    h = hashlib.md5(str(task_id).encode("utf-8")).hexdigest()
    i = int(h[:8], 16) % len(_GANTT_BAR_FILLS_ACTUAL)
    return _GANTT_BAR_FILLS_ACTUAL[i]


# ã‚¬ãƒ³ãƒˆæ™‚åˆ»ã‚»ãƒ«ï¼ˆçµ�å�ˆå¸¯ã�®å…ˆé ­ã‚»ãƒ«ï¼‰: æ¯Žã‚»ã‚°ãƒ¡ãƒ³ãƒˆ new ã�—ã�ªã�„
_GANTT_TIMELINE_CELL_ALIGNMENT = Alignment(
    horizontal="left",
    vertical="center",
    wrap_text=False,
    shrink_to_fit=False,
    indent=1,
)
# ã‚¿ã‚¹ã‚¯å¸¯ã�®è‰²ã�¯ãƒ‘ãƒ¬ãƒƒãƒˆæœ‰é™�ã�ªã�®ã�§ PatternFill ã‚’ hex å�˜ä½�ã�§å…±æœ‰ï¼ˆopenpyxl ã�®ã‚¹ã‚¿ã‚¤ãƒ«å±•é–‹ã‚³ã‚¹ãƒˆå‰Šæ¸›ï¼‰
_GANTT_TASK_PATTERN_FILL_BY_HEX: dict[str, PatternFill] = {}


def _gantt_cached_pattern_fill(hex_rrggbb: str) -> PatternFill:
    fi = _GANTT_TASK_PATTERN_FILL_BY_HEX.get(hex_rrggbb)
    if fi is None:
        fi = PatternFill(fill_type="solid", start_color=hex_rrggbb, end_color=hex_rrggbb)
        _GANTT_TASK_PATTERN_FILL_BY_HEX[hex_rrggbb] = fi
    return fi


def _gantt_slot_state_tuple(evlist, slot_mid, task_fill_fn=None):
    """ã‚¹ãƒ­ãƒƒãƒˆä¸­å¤®æ™‚åˆ»ã�«ã�Šã�‘ã‚‹1ãƒžã‚¹åˆ†ã�®çŠ¶æ…‹ã€‚('idle',) | ('break',) | ('daily_startup', fill_hex) | ('task', tid, fill_hex, pct)"""
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
        # ã€Œãƒžã‚¯ãƒ­å®Ÿè¡Œæ™‚ç‚¹ã€�ã�®å®Œäº†çŽ‡ã‚’å„ªå…ˆï¼ˆpct_macro ã‚’ timeline_event ã�«æŒ�ã�Ÿã�›ã‚‹ï¼‰
        if active.get("pct_macro") is not None:
            pct = int(round(parse_float_safe(active.get("pct_macro"), 0.0)))
            pct = max(0, min(100, pct))
        else:
            # ãƒ•ã‚§ã‚¤ãƒ«ã‚»ãƒ¼ãƒ•ï¼ˆå¾“æ�¥ã�®æ“¬ä¼¼é€²æ�—è¨ˆç®—ï¼‰
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
    """çµ�å�ˆã‚»ã‚°ãƒ¡ãƒ³ãƒˆå¢ƒç•Œåˆ¤å®šï¼ˆæ¯Žã‚¹ãƒ­ãƒƒãƒˆ tuple ã‚’å‰²ã‚Šå½“ã�¦ã�ªã�„ï¼‰ã€‚"""
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
    æ™‚é–“è»¸ã‚’å¡—ã‚Šåˆ†ã�‘ã�Ÿã�†ã�ˆã�§ã€�å�Œä¸€çŠ¶æ…‹ã�Œé€£ç¶šã�™ã‚‹ã‚»ãƒ«ã‚’æ¨ªçµ�å�ˆã�—å¸¯çŠ¶ã�®ãƒ�ãƒ¼ã�«ã�™ã‚‹ã€‚
    ï¼ˆç´°ãƒžã‚¹å�˜ä½“ã�®å¡—ã‚Šã�§ã�¯ã�ªã��15åˆ†åˆ»ã�¿ï¼‹å�Œä¸€çŠ¶æ…‹ã�®ã‚»ãƒ«çµ�å�ˆã�§ã€�å¸¯çŠ¶ã�®ãƒ�ãƒ¼ã�¨ã�—ã�¦è¡¨ç�¾ã�™ã‚‹ï¼‰
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
                    c.value = "(æ—¥æ¬¡å§‹æ¥­æº–å‚™)"
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
    """å�Šé–‹åŒºé–“ [a_start, a_end) ã�¨ [b_start, b_end) ã�Œé‡�ã�ªã‚‹ã�‹ï¼ˆå�Œä¸€æ—¥å†…ï¼‰ã€‚"""

    def _sec(t: time) -> int:
        return t.hour * 3600 + t.minute * 60 + t.second

    return _sec(a_start) < _sec(b_end) and _sec(a_end) > _sec(b_start)


def _parse_equipment_schedule_time_band_cell(val) -> tuple[time | None, time | None]:
    """çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ã€Œæ—¥æ™‚å¸¯ã€�ã‚»ãƒ«ï¼ˆä¾‹ 08:45-09:00ï¼‰ã‚’è§£é‡ˆã€‚"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None, None
    s = str(val).strip()
    if not s or "â– " in s:
        return None, None
    for sep in ("-", "ï¼�", "~", "ã€œ"):
        if sep in s:
            left, right = s.split(sep, 1)
            left = left.strip().replace("ï¼š", ":")
            right = right.strip().replace("ï¼š", ":")
            t0 = parse_time_str(left, None)
            t1 = parse_time_str(right, None)
            if t0 is not None and t1 is not None and t0 < t1:
                return t0, t1
            return None, None
    return None, None


def _apply_equipment_schedule_outside_regular_fill(
    ws, reg_start: time, reg_end: time
) -> None:
    """ã€Œæ—¥æ™‚å¸¯ã€�åˆ—ã�§å®šå¸¸ [reg_start, reg_end) ã�¨é‡�ã�ªã‚‰ã�ªã�„è¡Œã�®ã‚»ãƒ«ã�«ç�€è‰²ã€‚"""
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_OUTSIDE_REGULAR_TIME_FILL,
        end_color=RESULT_OUTSIDE_REGULAR_TIME_FILL,
    )
    col_idx = None
    for i, c in enumerate(ws[1], start=1):
        if c.value is not None and str(c.value).strip() == "æ—¥æ™‚å¸¯":
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
    è¨­å‚™åˆ—ï¼ˆé€²åº¦åˆ—ã‚’é™¤ã��ï¼‰ã�§ã€�è¡¨ç¤ºã�«ã€Œæ—¥æ¬¡å§‹æ¥­æº–å‚™ã€�ã€ŒåŠ å·¥å‰�æº–å‚™ã€�ã€Œä¾�é ¼åˆ‡æ›¿å¾Œå§‹æœ«ã€�ã�Œå�«ã�¾ã‚Œã‚‹ã‚»ãƒ«ã‚’è–„ç·‘ã�«ã�™ã‚‹ã€‚
    çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰² / TEMP_è¨­å‚™æ¯Žã�®æ™‚é–“å‰² ã�® equip ã‚»ãƒ«ç”¨ï¼ˆæ—¥æ™‚å¸¯åˆ—ã�¯å¤‰æ›´ã�—ã�ªã�„ï¼‰ã€‚
    """
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_DISPATCHED_REQUEST_FILL,
        end_color=RESULT_DISPATCHED_REQUEST_FILL,
    )
    markers = ("(æ—¥æ¬¡å§‹æ¥­æº–å‚™)", "(åŠ å·¥å‰�æº–å‚™)", "(ä¾�é ¼åˆ‡æ›¿å¾Œå§‹æœ«)")
    col_tb = None
    equip_cols: list[int] = []
    for i, c in enumerate(ws[1], start=1):
        if c.value is None:
            continue
        h = str(c.value).strip()
        if h == "æ—¥æ™‚å¸¯":
            col_tb = i
            continue
        if h.endswith("é€²åº¦"):
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
    """æ—¥ä»˜è¦‹å‡ºã�—è¡Œã€Œâ–  YYYY/MM/DD â€¦ â– ã€�ã�‹ã‚‰æ—¥ä»˜ã‚’å�–ã‚‹ã€‚"""
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
    """å½“æ—¥ãƒ»å½“è©²è¨­å‚™åˆ—ã‚­ãƒ¼ã�«å¯¾å¿œã�™ã‚‹æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼å� æœ‰åŒºé–“ï¼ˆå·¥å ´ç¨¼åƒ�æž ã�§ã‚¯ãƒªãƒƒãƒ—æ¸ˆã�¿ï¼‰ã€‚"""
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
    çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²: æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼å� æœ‰ã�¨é‡�ã�ªã‚‹è¨­å‚™ã‚»ãƒ«ï¼ˆé€²åº¦åˆ—ä»¥å¤–ï¼‰ã‚’ç´«è‰²ã�§å¡—ã‚‹ã€‚
    10 åˆ†æž ã�®å�Šé–‹åŒºé–“ [slot_start, slot_end) ã�¨å� æœ‰ [bs, be) ã�Œé‡�ã�ªã‚Œã�°å¯¾è±¡ã€‚
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
        if c.value is not None and str(c.value).strip() == "æ—¥æ™‚å¸¯":
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
    çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²_æ©Ÿæ¢°å��æ¯Žã�®æ©Ÿæ¢°å��åˆ—ã�§ã€�ä¾�é ¼NOã�Œå…¥ã�£ã�¦ã�„ã‚‹ã‚»ãƒ«ã�«è–„ç·‘ã‚’ä»˜ä¸Žã�™ã‚‹ã€‚
    ã€Œï¼ˆä¼‘æ†©ï¼‰ã€�ã�®ã�¿ã�®ã‚»ãƒ«ã�¯å¯¾è±¡å¤–ã€‚è¦‹å‡ºã�—è¡Œãƒ»æ—¥æ™‚å¸¯åˆ—ã�¯å¤‰æ›´ã�—ã�ªã�„ã€‚
    """
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_DISPATCHED_REQUEST_FILL,
        end_color=RESULT_DISPATCHED_REQUEST_FILL,
    )
    col_tb = None
    for i, c in enumerate(ws[1], start=1):
        if c.value is not None and str(c.value).strip() == "æ—¥æ™‚å¸¯":
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
            if not s or s == "ï¼ˆä¼‘æ†©ï¼‰":
                continue
            cell.fill = fill


def _equipment_gantt_fills_by_machine_name(equipment_list) -> dict[str, PatternFill]:
    """
    çµ�æžœ_è¨­å‚™ã‚¬ãƒ³ãƒˆã�®å›ºå®šåˆ—ï¼ˆBã€œEã€�A ã�¯æ—¥ä»˜ç¸¦çµ�å�ˆï¼‰ç”¨ã€‚equipment_list å†…ã�®æ©Ÿæ¢°å��ï¼ˆ+ ç„¡ã�—æ™‚ã�¯è¡Œå…¨ä½“ã‚’æ©Ÿæ¢°å��ï¼‰ã�®å‡ºç�¾é †ã�§
    æ·¡è‰²ã‚’å‰²ã‚Šå½“ã�¦ã€�å�Œä¸€æ©Ÿæ¢°å��ã�¯å¸¸ã�«å�Œã�˜ PatternFill ã‚’å…±æœ‰ã�™ã‚‹ã€‚
    """
    order: list[str] = []
    seen: set[str] = set()
    for eq in equipment_list or []:
        _, mn = _split_equipment_line_process_machine(eq)
        key = (mn or "").strip() or "â€”"
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
    çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ã�¨å�Œä¸€ãƒ‡ãƒ¼ã‚¿æº�ï¼ˆtimeline_eventsï¼‰ã�«åŸºã�¥ã��ã€�
    è¨­å‚™Ã—æ¨ªè»¸æ™‚é–“ã�®ã‚¬ãƒ³ãƒ�ãƒ£ãƒ¼ãƒˆé¢¨ã‚·ãƒ¼ãƒˆã‚’è¿½åŠ ã�™ã‚‹ã€‚
    æ¨ªè»¸ã�¯15åˆ†åˆ»ã�¿ã€‚é€£ç¶šã�™ã‚‹å�Œä¸€ã‚¿ã‚¹ã‚¯ï¼�ä¼‘æ†©ï¼�ç©ºã��ã�¯ã‚»ãƒ«çµ�å�ˆã�—ã�¦å¸¯çŠ¶ã�«è¡¨ç¤ºã�™ã‚‹ã€‚
    actual_timeline_events ã�Œã�‚ã‚Œã�°è¨­å‚™ã�”ã�¨ã�«ã€Œå®Ÿç¸¾ã€�è¡Œã‚’è¨ˆç”»è¡Œã�®ä¸‹ã�¸è¿½åŠ ã�™ã‚‹ã€‚
    """
    wb = writer.book
    try:
        insert_at = wb.sheetnames.index("çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²") + 1
    except ValueError:
        insert_at = len(wb.sheetnames)
    ws = wb.create_sheet("çµ�æžœ_è¨­å‚™ã‚¬ãƒ³ãƒˆ", insert_at)
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

    # æ¨ªè»¸(10åˆ†åˆ»ã�¿)ã�¯æ—¥ä»˜ã�§å…±é€šã�®ã�Ÿã‚�ã€�slot_times ã‚’å…ˆã�«ç¢ºå®š
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
    n_fixed = 5  # A=æ—¥ä»˜ï¼ˆæ—¥ãƒ–ãƒ­ãƒƒã‚¯å†…ã�§ç¸¦çµ�å�ˆï¼‰/ Bã€œE=æ©Ÿæ¢°å��ãƒ»å·¥ç¨‹å��ãƒ»æ‹…å½“è€…ãƒ»ã‚¿ã‚¹ã‚¯æ¦‚è¦�
    last_col = n_fixed + n_slots
    fills_by_mach = _equipment_gantt_fills_by_machine_name(equipment_list)
    fb_gantt = "F5F5F5"
    fill_gantt_fallback = PatternFill(fill_type="solid", start_color=fb_gantt, end_color=fb_gantt)

    # ã‚¿ã‚¤ãƒˆãƒ«ï¼†æ—¥æ™‚ï¼ˆãƒšãƒ¼ã‚¸ä¸Šéƒ¨ï¼‰
    create_ts = base_dt.strftime("%Y/%m/%d %H:%M:%S")
    master_path = os.path.join(os.getcwd(), MASTER_FILE) if MASTER_FILE else ""

    def _fmt_mtime(p):
        try:
            if p and os.path.exists(p):
                return datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y/%m/%d %H:%M:%S")
        except Exception:
            pass
        return "â€”"

    master_mtime = _fmt_mtime(master_path)

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    tcell = ws.cell(row=row, column=1, value="æ¹–å�—å·¥å ´ åŠ å·¥è¨ˆç”»")
    tcell.font = title_font
    tcell.fill = title_fill
    # çµ�å�ˆã‚»ãƒ«ã�§ã‚‚å·¦ç«¯ã�‹ã‚‰è¡¨ç¤ºï¼ˆç¸®å°�ãƒ»æŠ˜ã‚Šè¿”ã�—ã�ªã�—ï¼‰
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
        f"ä½œæˆ�ã€€{create_ts}"
        f"ã€€ãƒ»ã€€ãƒ‡ãƒ¼ã‚¿å�¸å‡ºã�—ã€€{data_extract_dt_str or 'â€”'}"
        f"ã€€ãƒ»ã€€ãƒžã‚¹ã‚¿ï¼ˆmaster.xlsmï¼‰ã€€{master_mtime}"
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
    fixed_hdr = ["æ—¥ä»˜", "æ©Ÿæ¢°å��", "å·¥ç¨‹å��", "æ‹…å½“è€…", "ã‚¿ã‚¹ã‚¯æ¦‚è¦�"]
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
    # å…ˆé ­ãƒ‡ãƒ¼ã‚¿è¡Œã�®å·¦ä¸Šï¼�æ™‚åˆ»åˆ—å…ˆé ­ï¼ˆF4ï¼‰ã�§çª“æž å›ºå®šï¼ˆè¡Œ1ã€œ3ãƒ»åˆ—Aã€œEã‚’å›ºå®šï¼‰
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
            mk_key = (mach_nm or "").strip() or "â€”"
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
                task_sum = " ".join(tids) if tids else "â€”"
                member_disp = _gantt_row_member_names(evlist)
            else:
                task_sum = "â€”"
                member_disp = "â€”"

            c1 = ws.cell(row=row, column=2, value=mach_nm if mach_nm else "â€”")
            c2 = ws.cell(row=row, column=3, value=proc_nm if proc_nm else "â€”")
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
                    task_sum_a = " ".join(tids_a) if tids_a else "â€”"
                    member_disp_a = _gantt_row_member_names(evlist_a)
                else:
                    task_sum_a = "â€”"
                    member_disp_a = "â€”"

                lab_fill_a = fills_by_mach.get(mk_key) or fill_gantt_fallback

                if mach_nm:
                    act_mach = f"{mach_nm}ï¼ˆå®Ÿç¸¾ï¼‰"
                elif proc_nm:
                    act_mach = "ï¼ˆå®Ÿç¸¾ï¼‰"
                else:
                    act_mach = "â€”"
                ca1 = ws.cell(row=row, column=2, value=act_mach)
                ca2 = ws.cell(row=row, column=3, value=proc_nm if proc_nm else "â€”")
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
                value=f"ã€�{d.strftime('%Y/%m/%d')}ã€‘",
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

    # å‡¡ä¾‹ã�¯é«˜ã�•ç¢ºä¿�ã�®ã�Ÿã‚�çœ�ç•¥ï¼ˆãƒ¢ãƒŽã‚¯ãƒ­å�°åˆ·ã�¯è‰²ã�®æ¿ƒæ·¡/ã‚»ãƒ«ã�®æž ã�§è­˜åˆ¥ï¼‰
    # åˆ—å¹…ãƒ»æŠ˜ã‚Šè¿”ã�—ã�¯ VBA å�–ã‚Šè¾¼ã�¿æ™‚ï¼ˆçµ�æžœ_è¨­å‚™ã‚¬ãƒ³ãƒˆ_åˆ—å¹…ã‚’è¨­å®šï¼‰ã�§è¨­å®š

    try:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToHeight = False
        ws.page_setup.fitToWidth = 1
        # A3ï¼ˆopenpyxl ä¸Šã�§ paperSize=8 ã�Œ A3 ç›¸å½“ï¼‰
        ws.page_setup.paperSize = 8
        # ä½™ç™½ã‚’ç‹­ã‚�ã�¦æ¨ª1ãƒšãƒ¼ã‚¸ã�«å�Žã�¾ã‚Šã‚„ã�™ã��ã�™ã‚‹ï¼ˆå�˜ä½�: ã‚¤ãƒ³ãƒ�ï¼‰
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.2
        ws.page_margins.bottom = 0.2
        # ã‚¿ã‚¤ãƒˆãƒ«ãƒ»è¡¨ã‚’ãƒšãƒ¼ã‚¸å·¦åŸºæº–ã�«ï¼ˆãƒ¬ãƒ�ãƒ¼ãƒˆé¢¨ï¼‰
        ws.print_options.horizontalCentered = False
        ws.print_options.verticalCentered = False
        ws.print_options.gridLines = False
    except Exception:
        pass


def row_has_completion_keyword(row):
    """åŠ å·¥å®Œäº†åŒºåˆ†ã�«ã€Œå®Œäº†ã€�ã�®æ–‡å­—ã�Œå�«ã�¾ã‚Œã‚‹å ´å�ˆã�¯ã‚¿ã‚¹ã‚¯å®Œäº†ã�¨ã�¿ã�ªã�™ã€‚"""
    v = row.get(TASK_COL_COMPLETION_FLAG)
    if v is None or pd.isna(v):
        return False
    return "å®Œäº†" in str(v)


def _plan_row_exclude_from_assignment(row) -> bool:
    """
    ã€Œé…�å�°ä¸�è¦�ã€�åˆ—ã�Œã‚ªãƒ³ã�ªã‚‰ã€�ã��ã�®è¡Œã�¯é…�å�°ã‚­ãƒ¥ãƒ¼ã�¸å…¥ã‚Œã�šã€�ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã�® AI è§£æž�è¡Œã�‹ã‚‰ã‚‚é™¤ã��ã€‚

    é…�å�°ã�‹ã‚‰å¤–ã�™ï¼ˆçœŸï¼‰: è«–ç�†å€¤ Trueã€�æ•°å€¤ 1ã€�æ–‡å­—åˆ—ï¼ˆNFKC å¾Œãƒ»å°�æ–‡å­—ï¼‰
      true / 1 / yes / on / y / t / ã�¯ã�„ / â—‹ / ã€‡ / â—�
    é…�å�°å¯¾è±¡ï¼ˆå�½ï¼‰: ç©ºã€�Noneã€�Falseã€�0ã€�no / off / false / ã�„ã�„ã�ˆ / å�¦ ç­‰
    ä¸Šè¨˜ä»¥å¤–ã�®æ–‡å­—åˆ—ã�¯å�½ï¼ˆé…�å�°ã�™ã‚‹ï¼‰ã€‚ãƒ�ã‚§ãƒƒã‚¯ãƒœãƒƒã‚¯ã‚¹é€£å‹•ã‚»ãƒ«ã�¯é€šå¸¸ TRUE/FALSE ã�¾ã�Ÿã�¯ 1/0ã€‚
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
    if not s or s in ("nan", "none", "false", "0", "no", "off", "ã�„ã�„ã�ˆ", "å�¦"):
        return False
    if s in ("true", "1", "yes", "on", "ã�¯ã�„", "y", "t", "â—‹", "ã€‡", "â—�"):
        return True
    return False


def _coerce_plan_exclude_column_value_for_storage(v):
    """
    ã€Œé…�å�°ä¸�è¦�ã€�åˆ—ã�¸æ›¸ã��è¾¼ã‚€å€¤ã‚’ã€�StringDtype åˆ—ã�§ã‚‚ä»£å…¥ã‚¨ãƒ©ãƒ¼ã�«ã�ªã‚‰ã�ªã�„å½¢ã�«ã��ã‚�ã�ˆã‚‹ã€‚
    Excel å�–ã‚Šè¾¼ã�¿ã�® True / 1 / False / 0 ã�¨æ–‡å­—åˆ—ã‚’ä¿�æŒ�ã�—ã€�_plan_row_exclude_from_assignment ã�¨æ•´å�ˆã�™ã‚‹ã€‚
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
    åŠ å·¥æ¸ˆæ•°é‡�ï¼ˆå·¥ç¨‹æŠ•å…¥é‡�æ�›ç®—ï¼‰ã‚’è¿”ã�™ã€‚

    åŸºæœ¬å¼�:
      å®Ÿå‡ºæ�¥é«˜ Ã· (å�—æ³¨æ•° Ã· æ�›ç®—æ•°é‡�)
    = å®Ÿå‡ºæ�¥é«˜ * æ�›ç®—æ•°é‡� / å�—æ³¨æ•°

    å�—æ³¨æ•°ã�Œç„¡ã�„/ä¸�æ­£ã�ªå ´å�ˆã�¯ã€�æ—§åˆ—ã€Œå®ŸåŠ å·¥æ•°ã€�ã‚’äº’æ�›ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ã�¨ã�—ã�¦ä½¿ã�†ã€‚
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
    iterrows() 1è¡Œåˆ†ã�‹ã‚‰åˆ—å€¤ã‚’å�–ã‚‹ã€‚å�Œä¸€è¦‹å‡ºã�—ã�®é‡�è¤‡åˆ—ã�Œã�‚ã‚‹ã�¨ row.get ã�¯ Series ã�«ã�ªã‚Šã€�
    strâ†’to_datetime ã�§èª¤ã�£ã�Ÿæ—¥ä»˜ã�«ã�ªã‚‹ã�“ã�¨ã�Œã�‚ã‚‹ã�Ÿã‚�ã€�å…ˆé ­ã�®é�žæ¬ æ��ã‚¹ã‚«ãƒ©ãƒ¼ã‚’è¿”ã�™ã€‚
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
                    # æœŸé™�åˆ‡ã‚Œã‚¨ãƒ³ãƒˆãƒªã‚’é™¤åŽ»ï¼ˆ6æ™‚é–“ï¼‰
                    now_ts = time_module.time()
                    cleaned = {}
                    expired_count = 0
                    for k, v in data.items():
                        # æ–°å½¢å¼�: {"ts": epoch_seconds, "data": {...}}
                        if isinstance(v, dict) and "ts" in v and "data" in v:
                            ts = parse_float_safe(v.get("ts"), 0.0)
                            if ts > 0 and (now_ts - ts) <= AI_CACHE_TTL_SECONDS:
                                cleaned[k] = v
                            else:
                                expired_count += 1
                        # æ—§å½¢å¼�: å€¤ã�Œç›´æŽ¥AIçµ�æžœdictï¼ˆäº’æ�›ã�§èª­ã�¿å�–ã‚Šã€�å�³æ™‚ã�«æ–°å½¢å¼�ã�¸å†�ä¿�å­˜ã�•ã‚Œã‚‹ï¼‰
                        else:
                            cleaned[k] = {"ts": now_ts, "data": v}
                    if expired_count > 0:
                        logging.info(f"AIã‚­ãƒ£ãƒƒã‚·ãƒ¥æœŸé™�åˆ‡ã‚Œã‚’å‰Šé™¤: {expired_count}ä»¶")
                    return cleaned
    except Exception as e:
        logging.warning(f"AIã‚­ãƒ£ãƒƒã‚·ãƒ¥èª­ã�¿è¾¼ã�¿å¤±æ•—: {e}")
    return {}

def save_ai_cache(cache_obj):
    try:
        with open(ai_cache_path, "w", encoding="utf-8") as f:
            json.dump(cache_obj, f, ensure_ascii=False)
    except Exception as e:
        logging.warning(f"AIã‚­ãƒ£ãƒƒã‚·ãƒ¥ä¿�å­˜å¤±æ•—: {e}")

def get_cached_ai_result(cache_obj, cache_key, content_key=None):
    """
    content_key: ã‚ªãƒ—ã‚·ãƒ§ãƒ³ã€‚ä¿�å­˜æ™‚ã�¨å�Œä¸€ã�®æ–‡å­—åˆ—ã�§ã�ªã�„ãƒ’ãƒƒãƒˆã�¯ç„¡åŠ¹åŒ–ã�™ã‚‹ï¼ˆç‰¹åˆ¥æŒ‡å®šãƒ»ç…§å�ˆç”¨ã�®äºŒæ¬¡ãƒ�ã‚§ãƒƒã‚¯ï¼‰ã€‚
    æ—§ã‚¨ãƒ³ãƒˆãƒªã�« content_key ã�Œç„¡ã�„å ´å�ˆã�¯ SHA256 ã‚­ãƒ¼ä¸€è‡´ã�®ã�¿ã�§å¾“æ�¥ã�©ã�Šã‚Šãƒ’ãƒƒãƒˆã�¨ã�¿ã�ªã�™ã€‚
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
                "AIã‚­ãƒ£ãƒƒã‚·ãƒ¥: ã‚­ãƒ¼ã�¯ä¸€è‡´ã�—ã�¾ã�™ã�Œ content_key ã�Œç�¾è¡Œå…¥åŠ›ã�¨ç•°ã�ªã‚‹ã�Ÿã‚�ç„¡åŠ¹åŒ–ã�—ã�¾ã�™ã€‚"
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
    # ä¾‹: "Please retry in 57.089735313s."
    m = re.search(r"retry in ([0-9]+(?:\.[0-9]+)?)s", err_text, re.IGNORECASE)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    # ä¾‹: "'retryDelay': '57s'"
    m = re.search(r"retryDelay'\s*:\s*'([0-9]+)s'", err_text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return None


def infer_unit_m_from_product_name(product_name, fallback_unit):
    """
    è£½å“�å��æ–‡å­—åˆ—ã�‹ã‚‰åŠ å·¥å�˜ä½�(m)ã‚’æŽ¨å®šã�™ã‚‹æš«å®šãƒ«ãƒ¼ãƒ«ã€‚
    ä¾‹: 15020-JX5R- 770X300F-A   R -> 300
    â€» ãƒ�ãƒªã‚¨ãƒ¼ã‚·ãƒ§ãƒ³ã�Œå¤šã�„å‰�æ��ã�®ã�Ÿã‚�ã€�ã�“ã�“ã‚’éƒ½åº¦èª¿æ•´ã�§ã��ã‚‹ã‚ˆã�†é–¢æ•°åŒ–ã�—ã�¦ã�„ã‚‹ã€‚
    """
    if product_name is None or pd.isna(product_name):
        return fallback_unit
    s = str(product_name)
    # "770X300..." ã�®ã‚ˆã�†ã�ªãƒ‘ã‚¿ãƒ¼ãƒ³ã�‹ã‚‰ X ã�®å¾Œã�®æ•°å€¤ã‚’æ‹¾ã�†ï¼ˆæœ€å¾Œã�«è¦‹ã�¤ã�‹ã�£ã�ŸXã‚’å„ªå…ˆï¼‰
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
    ã‚¿ã‚¹ã‚¯å…¥åŠ›ã‚’å�–å¾—ã�™ã‚‹ï¼ˆtasks.xlsx ã�¯ä½¿ç”¨ã�—ã�ªã�„ï¼‰ã€‚
    å¿…é ˆ: ç’°å¢ƒå¤‰æ•° TASK_INPUT_WORKBOOK ã�«ãƒžã‚¯ãƒ­å®Ÿè¡Œãƒ–ãƒƒã‚¯ã�®ãƒ•ãƒ«ãƒ‘ã‚¹ï¼ˆVBA ã�Œè¨­å®šï¼‰
         ã‚·ãƒ¼ãƒˆã€ŒåŠ å·¥è¨ˆç”»DATAã€�ã‚’èª­ã�¿è¾¼ã‚€ï¼ˆæŠ•å…¥ç›®å®‰ã�¯ã€Œå›žç­”ç´�æœŸã€�ã€�æœªå…¥åŠ›æ™‚ã�¯ã€ŒæŒ‡å®šç´�æœŸã€�ï¼‰ã€‚
    """
    if not TASKS_INPUT_WORKBOOK:
        raise FileNotFoundError(
            "TASK_INPUT_WORKBOOK ã�Œæœªè¨­å®šã�§ã�™ã€‚VBA ã�® RunPython ã�§ãƒžã‚¯ãƒ­å®Ÿè¡Œãƒ–ãƒƒã‚¯ã�®ãƒ‘ã‚¹ã‚’æ¸¡ã�—ã�¦ã��ã� ã�•ã�„ã€‚"
        )
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        raise FileNotFoundError(f"TASK_INPUT_WORKBOOK ã�Œå­˜åœ¨ã�—ã�¾ã�›ã‚“: {TASKS_INPUT_WORKBOOK}")
    df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=TASKS_SHEET_NAME)
    df.columns = df.columns.str.strip()
    logging.info(f"ã‚¿ã‚¹ã‚¯å…¥åŠ›: '{TASKS_INPUT_WORKBOOK}' ã�® '{TASKS_SHEET_NAME}' ã‚’èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚")
    return df


def _nfkc_column_aliases(canonical_name):
    """è¦‹å‡ºã�—ã�®è¡¨è¨˜ã‚†ã‚Œï¼ˆå…¨è§’è¨˜å�·ãƒ»äº’æ�›æ–‡å­—ï¼‰ã‚’å�¸å�Žã�™ã‚‹ã�Ÿã‚�ã�®æ¯”è¼ƒã‚­ãƒ¼ã€‚"""
    return unicodedata.normalize("NFKC", str(canonical_name).strip())


def _align_dataframe_headers_to_canonical(df, canonical_names):
    """åˆ—å��ã‚’ NFKC ä¸€è‡´ã�§ canonical ã�«å¯„ã�›ã‚‹ï¼ˆExcel å�´ã�Œå…¨è§’ '_' ç­‰ã�§ã‚‚èª­ã‚�ã‚‹ã‚ˆã�†ã�«ï¼‰ã€‚"""
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
    å·¥ç¨‹å��ï¼ˆè¨­å‚™å��ï¼‰ã�®ç…§å�ˆç”¨ã‚­ãƒ¼ã€‚
    NFKCãƒ»å‰�å¾Œç©ºç™½ãƒ»é€£ç¶šç©ºç™½ãƒ»NBSP/å…¨è§’ã‚¹ãƒšãƒ¼ã‚¹ãƒ»ã‚¼ãƒ­å¹…æ–‡å­—ã‚’æ­£è¦�åŒ–ã�™ã‚‹ã€‚
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    t = unicodedata.normalize("NFKC", str(val))
    t = t.replace("\u00a0", " ").replace("\u3000", " ")
    t = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _equipment_line_key_to_physical_occupancy_key(eq_line: str) -> str:
    """è¨­å‚™åˆ—ã‚­ãƒ¼ï¼ˆå·¥ç¨‹+æ©Ÿæ¢° ç­‰ï¼‰ã�‹ã‚‰ã€�ç‰©ç�†æ©Ÿæ¢°ã�®å� æœ‰ã�«ç”¨ã�„ã‚‹ã‚­ãƒ¼ï¼ˆæ©Ÿæ¢°å��å�´ãƒ»æ­£è¦�åŒ–ï¼‰ã‚’å¾—ã‚‹ã€‚"""
    s = str(eq_line or "").strip()
    if not s:
        return ""
    nk = _normalize_equipment_match_key(s)
    if "+" in nk:
        return _normalize_equipment_match_key(nk.split("+", 1)[1])
    return nk


def _physical_machine_occupancy_key_for_task(task: dict) -> str:
    """
    è¨­å‚™ã�®å£�æ™‚è¨ˆå� æœ‰ï¼ˆmachine_avail_dtãƒ»é–“éš”ãƒŸãƒ©ãƒ¼ï¼‰ã�«ç”¨ã�„ã‚‹ã‚­ãƒ¼ã€‚
    æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼åˆ—ã�¯ equipment_line_key ã�®ã€Œå·¥ç¨‹+æ©Ÿæ¢°ã€�ã�¨ä¸€è‡´ã�™ã‚‹ã�Ÿã‚�ã€�
    æ­£è¦�åŒ–å¾Œã�«ã€Œ+ã€�ã‚’å�«ã‚€ã�¨ã��ã�¯ **machine_name ã‚ˆã‚Šå…ˆã�«** ã��ã�“ã�‹ã‚‰ç‰©ç�†æ©Ÿæ¢°å��ã‚’æŽ¡ç”¨ã�™ã‚‹ã€‚
    ï¼ˆmachine_name ã�«å·¥ç¨‹å��ã�®ã�¿ã�ªã�©ã�Œå…¥ã‚Šã€�åºŠã‚­ãƒ¼ã€Œç†±èž�ç�€æ©Ÿ æ¹–å�—ã€�ã�¨ã�šã‚Œã�¦å€™è£œå¤–ã�—æ¼�ã‚Œã�™ã‚‹ã�®ã‚’é˜²ã��ï¼‰
    å�˜ä¸€å��ã�®ã�¨ã��ã�¯å¾“æ�¥ã�©ã�Šã‚Š machine_name ã‚’å„ªå…ˆã�—ã€�ç„¡ã�‘ã‚Œã�° equipment_line_key / machine ã�‹ã‚‰æŽ¨å®šã�™ã‚‹ã€‚
    machine_name ã�«ã€Œå·¥ç¨‹+æ©Ÿæ¢°ã€�ã�¨å…¥ã�£ã�¦ã�„ã‚‹å ´å�ˆã�§ã‚‚ã€�å� æœ‰ã�¯ç‰©ç�†æ©Ÿæ¢°å��ï¼ˆ+ ã�®å�³å�´ï¼‰ã�«å¯„ã�›ã‚‹ã€‚
    å…¨è§’ã€Œï¼‹ã€�ã�®ã�¿ã�®åˆ—ã�¯ NFKC å¾Œã�«å�Šè§’ã€Œ+ã€�ã�«ã�ªã‚‹ã�Ÿã‚�ã€�åˆ†å‰²åˆ¤å®šã�¯æ­£è¦�åŒ–å¾Œã�«è¡Œã�†ã€‚
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
    machine_avail_dtãƒ»æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼åºŠã�¨æ•´å�ˆã�™ã‚‹å� æœ‰ã‚­ãƒ¼ï¼ˆåŽŸå‰‡: ç‰©ç�†æ©Ÿæ¢°å��ï¼‰ã€‚
    task ã�‹ã‚‰å�–ã‚Œã�ªã�„ã�¨ã��ã�¯ eq_lineï¼ˆå·¥ç¨‹+æ©Ÿæ¢°ï¼‰ã�‹ã‚‰æ©Ÿæ¢°å��å�´ã‚’æŽ¨å®šã�—ã€�æœ€å¾Œã�®æ‰‹æ®µã�§ eq_lineã€‚
    ã€Œâ€¦ or eq_lineã€�ã�«ã‚ˆã‚‹å·¥ç¨‹+æ©Ÿæ¢°ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ã�¯æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ç‰©ç�†ã‚­ãƒ¼ã�¨ä¸�ä¸€è‡´ã�«ã�ªã‚Šå¾—ã‚‹ã�Ÿã‚�ç¦�æ­¢ã€‚
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
    """æ­£è¦�åŒ–ã‚­ãƒ¼ â†’ master ã‚¹ã‚­ãƒ«ã‚·ãƒ¼ãƒˆä¸Šã�®åˆ—å��ï¼ˆcanonical è¡¨è¨˜ï¼‰ã€‚"""
    lookup = {}
    for eq in equipment_list:
        k = _normalize_equipment_match_key(eq)
        if k and k not in lookup:
            lookup[k] = eq
    # å·¥ç¨‹å��ã�®ã�¿ã�®ç…§å�ˆï¼ˆåŠ å·¥å®Ÿç¸¾DATAç­‰ï¼‰: å�Œä¸€å·¥ç¨‹ã�®å…ˆé ­åˆ—ï¼ˆå·¥ç¨‹+æ©Ÿæ¢°ï¼‰ã�¸å¯„ã�›ã‚‹
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
    çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ãƒ»çµ�æžœ_è¨­å‚™ã‚¬ãƒ³ãƒˆã�®è¡Œï¼�åˆ—è¦‹å‡ºã�—ç”¨ã€‚
    å†…éƒ¨ã‚­ãƒ¼ã�Œã€Œå·¥ç¨‹+æ©Ÿæ¢°ã€�ã�®ã�¨ã��ã�¯æ©Ÿæ¢°å��ã‚’è¡¨ç¤ºã�—ã€�æ©Ÿæ¢°å��ã�®é‡�è¤‡æ™‚ã�®ã�¿å·¥ç¨‹ã‚’æ‹¬å¼§ã�§è£œã�†ã€‚
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
                out.append(f"{r}ï¼ˆ{p}ï¼‰" if p else r)
            else:
                out.append(r)
        else:
            out.append(r)
    return out


def _split_equipment_line_process_machine(eq_line: str) -> tuple[str, str]:
    """
    è¨­å‚™ãƒžã‚¹ã‚¿ã�®åˆ—ã‚­ãƒ¼ã€Œå·¥ç¨‹+æ©Ÿæ¢°ã€�ã‚’ (å·¥ç¨‹å��, æ©Ÿæ¢°å��) ã�«åˆ†å‰²ã�™ã‚‹ã€‚
    '+' ã�Œç„¡ã�„ã�¨ã��ã�¯æ©Ÿæ¢°å��ã�®ã�¿ã�¨ã�¿ã�ªã�—ã€�å·¥ç¨‹å��ã�¯ç©ºæ–‡å­—ã€‚
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
    è¨­å‚™ã‚¬ãƒ³ãƒˆã�®æ‹…å½“è€…ã‚»ãƒ«ç”¨ã€‚å�Šè§’ï¼�å…¨è§’ç©ºç™½ã�Œã�‚ã‚Œã�°æ‰‹å‰�ã‚’å§“ã�¨ã�¿ã�ªã�—ã€�ç„¡ã�„ã�¨ã��ã�¯å…¨ä½“ã‚’è¡¨ç¤º
    ï¼ˆæ°�å��ã�Œ1ãƒˆãƒ¼ã‚¯ãƒ³ã�®ã�¿ã�®ã�¨ã��ã�¯å§“ã�®åˆ‡ã‚Šå‡ºã�—ä¸�å�¯ã�®ã�Ÿã‚�ã��ã�®ã�¾ã�¾ï¼‰ã€‚NFKCãƒ»å¯Œç”°/å†¨ç”°å¯„ã�›ã�¯å§“ç”¨ã�¨å�Œã�˜ã€‚
    """
    sei, mei = _split_person_sei_mei(raw)
    if not sei:
        return ""
    n = _normalize_sei_for_match(sei)
    return n if n else sei


def _gantt_row_member_names(evlist) -> str:
    """è¨­å‚™ã‚¬ãƒ³ãƒˆè¡Œç”¨: ä¸»æ‹…å½“(op)ã�¨ã‚µãƒ–(sub)ã‚’å‡ºç�¾é †ã�§é‡�è¤‡é™¤åŽ»ã�—ã€�å§“ã�®ã�¿ã‚’ã‚«ãƒ³ãƒž+å�Šè§’ã‚¹ãƒšãƒ¼ã‚¹åŒºåˆ‡ã‚Šã€‚"""
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
        for seg in re.split(r"[,ã€�]", sub_raw):
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
    return ", ".join(labels) if labels else "â€”"


def _resolve_equipment_line_key_for_task(task: dict, equipment_list: list | None) -> str:
    """
    è¨­å‚™æ™‚é–“å‰²ãƒ»è¨­å‚™å°‚æœ‰ç©ºã��ã�®åˆ—ã‚­ãƒ¼ï¼ˆskills / need ã�¨å�Œã�˜ã€Œå·¥ç¨‹+æ©Ÿæ¢°ã€�ã‚’åŸºæœ¬ã�¨ã�™ã‚‹ï¼‰ã€‚
    æ©Ÿæ¢°å��ã�Œç©ºã�§ãƒžã‚¹ã‚¿ã�«å½“è©²å·¥ç¨‹ã�®åˆ—ã�Œ1ã�¤ã� ã�‘ã�ªã‚‰ã��ã�®è¤‡å�ˆã‚­ãƒ¼ã�¸å¯„ã�›ã‚‹ã€‚
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
    é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã‚’ DataFrame åŒ–ã�—ã�Ÿç›´å¾Œã�®å…±é€šå‡¦ç�†ã€‚
    æ®µéšŽ2ã�® ``load_planning_tasks_df`` ã�¨å�Œã�˜ï¼ˆè¨­å®šã‚·ãƒ¼ãƒˆãƒ»åˆ†å‰²è¡Œãƒ»é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«ï¼‰ã€‚
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
        logging.exception("%s: è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã�®ä¿�å®ˆã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰", log_prefix)
    try:
        _apply_auto_exclude_bunkatsu_duplicate_machine(df, log_prefix=log_prefix)
    except Exception as ex:
        logging.warning(
            "%s: åˆ†å‰²è¡Œã�®é…�å�°ä¸�è¦�è‡ªå‹•è¨­å®šã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰: %s",
            log_prefix,
            ex,
        )
    try:
        apply_exclude_rules_config_to_plan_df(df, wb_path, log_prefix)
    except Exception as ex:
        logging.warning(
            "%s: è¨­å®šã‚·ãƒ¼ãƒˆã�«ã‚ˆã‚‹é…�å�°ä¸�è¦�é�©ç”¨ã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰: %s",
            log_prefix,
            ex,
        )

def load_planning_tasks_df():
    """
    2æ®µéšŽç›®ç”¨: ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ä¸Šã�®ã€Œé…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€�ã‚·ãƒ¼ãƒˆã‚’èª­ã�¿è¾¼ã‚€ã€‚

    ã€Œæ‹…å½“OP_æŒ‡å®šã€�åˆ—ã�¾ã�Ÿã�¯ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒã�® AI å‡ºåŠ› preferred_operator ã�§ä¸»æ‹…å½“ OP ã‚’æŒ‡å��ã�§ã��ã‚‹ï¼ˆskills ã�®ãƒ¡ãƒ³ãƒ�ãƒ¼å��ã�¨ã�‚ã�„ã�¾ã�„ä¸€è‡´ï¼‰ã€‚
    ãƒ¡ã‚¤ãƒ³ã€Œå†�å„ªå…ˆç‰¹åˆ¥è¨˜è¼‰ã€�ã�® task_preferred_operators ã�¯ generate_plan å�´ã�§æœ€å„ªå…ˆãƒžãƒ¼ã‚¸ã�•ã‚Œã‚‹ã€‚
    ã€Œé…�å�°ä¸�è¦�ã€�ã�Œã‚ªãƒ³ï¼ˆTRUE/1/ã�¯ã�„ ç­‰ï¼‰ã�®è¡Œã�¯é…�å�°å¯¾è±¡å¤–ã€‚
    èª­ã�¿è¾¼ã�¿å¾Œã€�å�Œä¸€ä¾�é ¼NOãƒ»é‡�è¤‡æ©Ÿæ¢°å��ã�Œã�‚ã‚‹ã‚°ãƒ«ãƒ¼ãƒ—ã�®å·¥ç¨‹ã€Œåˆ†å‰²ã€�è¡Œã�¸ç©ºã�ªã‚‰ã€Œé…�å�°ä¸�è¦�ã€�=yesï¼ˆæ®µéšŽ1ã�¨å�Œã�˜ï¼‰ã€‚
    ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�ã�§å·¥ç¨‹+æ©Ÿæ¢°ã�®çµ„ã‚’å�ŒæœŸã�—ã€�C/D/E ã�«åŸºã�¥ã��é…�å�°ä¸�è¦�ã‚’å��æ˜ ã�™ã‚‹ï¼ˆã‚·ãƒ¼ãƒˆä½œæˆ�ã�¯ VBAï¼‰ã€‚
    """
    if not TASKS_INPUT_WORKBOOK:
        raise FileNotFoundError(
            "TASK_INPUT_WORKBOOK ã�Œæœªè¨­å®šã�§ã�™ã€‚VBA ã�® RunPython ã�§ãƒžã‚¯ãƒ­å®Ÿè¡Œãƒ–ãƒƒã‚¯ã�®ãƒ‘ã‚¹ã‚’æ¸¡ã�—ã�¦ã��ã� ã�•ã�„ã€‚"
        )
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        raise FileNotFoundError(f"TASK_INPUT_WORKBOOK ã�Œå­˜åœ¨ã�—ã�¾ã�›ã‚“: {TASKS_INPUT_WORKBOOK}")
    df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=PLAN_INPUT_SHEET_NAME)
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(
        df, plan_input_sheet_column_order()
    )
    for c in plan_input_sheet_column_order():
        if c not in df.columns:
            df[c] = ""
    _apply_planning_sheet_post_load_mutations(df, TASKS_INPUT_WORKBOOK, "é…�å�°ã‚·ãƒ¼ãƒˆèª­è¾¼")
    logging.info(
        f"è¨ˆç”»ã‚¿ã‚¹ã‚¯å…¥åŠ›: '{TASKS_INPUT_WORKBOOK}' ã�® '{PLAN_INPUT_SHEET_NAME}' ã‚’èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚"
    )
    return df


def _main_sheet_cell_is_global_comment_label(val) -> bool:
    """ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆä¸Šã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�è¦‹å‡ºã�—ã‚»ãƒ«ã�‹ï¼ˆè¡¨è¨˜ã‚†ã‚Œè¨±å®¹ï¼‰ã€‚"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    s = unicodedata.normalize("NFKC", str(val).strip())
    if not s:
        return False
    if _nfkc_column_aliases(s) == _nfkc_column_aliases("ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆ"):
        return True
    if "ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«" in s and "ã‚³ãƒ¡ãƒ³ãƒˆ" in s:
        return True
    return False


def load_main_sheet_global_priority_override_text() -> str:
    """
    TASK_INPUT_WORKBOOK ã�®ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆã�§ã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�ã�¨æ›¸ã�‹ã‚Œã�Ÿã‚»ãƒ«ã�® **ç›´ä¸‹** ã‚’èª­ã‚€ã€‚
    ã‚·ãƒ¼ãƒˆå��: ã€Œãƒ¡ã‚¤ãƒ³ã€�ã€Œãƒ¡ã‚¤ãƒ³_ã€�ã€ŒMainã€�ã�®ã�„ã�šã‚Œã�‹ã€�ã�¾ã�Ÿã�¯å��å‰�ã�«ã€Œãƒ¡ã‚¤ãƒ³ã€�ã‚’å�«ã‚€ï¼ˆVBA GetMainWorksheet ã�¨å�Œè¶£æ—¨ï¼‰ã€‚

    å†…å®¹ã�¯ **Gemini ã�§ä¸€æ‹¬è§£é‡ˆ**ï¼ˆ`analyze_global_priority_override_comment`ï¼‰ã€‚å·¥å ´ä¼‘æ¥­æ—¥ãƒ»å†�å„ªå…ˆãƒ•ãƒ©ã‚°ãƒ»æœªå®Ÿè£…æŒ‡ç¤ºã�®ãƒ¡ãƒ¢ã‚’ JSON åŒ–ã�™ã‚‹ã€‚
    API ã‚­ãƒ¼ã�Œç„¡ã�„å ´å�ˆã�®ã�¿ã€�å·¥å ´ä¼‘æ¥­æ—¥ã�¯ãƒ«ãƒ¼ãƒ«ãƒ™ãƒ¼ã‚¹ã�® `parse_factory_closure_dates_from_global_comment` ã�§è£œå®Œã�™ã‚‹ã€‚
    """
    wb_path = TASKS_INPUT_WORKBOOK.strip() if TASKS_INPUT_WORKBOOK else ""
    if not wb_path or not os.path.exists(wb_path):
        return ""
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� openpyxl ã�§ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã‚’èª­ã�¿ã�¾ã�›ã‚“ã€‚",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return ""
    try:
        wb = load_workbook(wb_path, data_only=True, read_only=False)
    except Exception as e:
        logging.warning("ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: ãƒ–ãƒƒã‚¯ã‚’é–‹ã�‘ã�¾ã�›ã‚“ã�§ã�—ã�Ÿ: %s", e)
        return ""
    try:
        ws = None
        for name in ("ãƒ¡ã‚¤ãƒ³", "ãƒ¡ã‚¤ãƒ³_", "Main"):
            if name in wb.sheetnames:
                ws = wb[name]
                break
        if ws is None:
            for sn in wb.sheetnames:
                if "ãƒ¡ã‚¤ãƒ³" in sn:
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
    ãƒ¡ã‚¤ãƒ³ã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�ã�®æ–­ç‰‡ã�Œã€�å·¥å ´å�˜ä½�ã�®ä¼‘æ¥­ãƒ»é�žç¨¼åƒ�ã‚’æ„�å‘³ã�™ã‚‹ã�‹ï¼ˆå€‹äººä¼‘ã�¿ã� ã�‘ã‚’èª¤æ¤œå‡ºã�—ã�ªã�„ï¼‰ã€‚
    """
    c = unicodedata.normalize("NFKC", str(chunk or ""))
    if not c.strip():
        return False
    if re.search(r"è‡¨æ™‚\s*ä¼‘æ¥­", c):
        return True
    if "ä¼‘å ´" in c:
        return True
    if re.search(r"å·¥å ´", c) and re.search(r"ä¼‘|ä¼‘æ¥­|ä¼‘ã�¿|å�œæ­¢|ã�Šä¼‘ã�¿", c):
        return True
    if re.search(r"(?:å…¨ç¤¾|å…¨é¤¨|å…¨å·¥å ´).{0,15}(?:ä¼‘|ä¼‘æ¥­|å�œæ­¢)", c):
        return True
    if re.search(r"(?:ç¨¼åƒ�|ç”Ÿç”£|ãƒ©ã‚¤ãƒ³).{0,12}(?:å�œæ­¢|ã�ªã�—|ç„¡ã�—)", c):
        return True
    if re.search(r"åŠ å·¥.{0,15}(?:ã�—ã�ªã�„|ç„¡ã�—|ã�ªã�—|ã�Šä¼‘ã�¿)", c):
        return True
    if "ä¼‘æ¥­" in c and re.search(
        r"(?:å·¥å ´|å…¨ç¤¾|æœ¬ç¤¾|å½“æ—¥|å¼Šç¤¾|å½“ç¤¾|å…¨å“¡|ç¤¾å…¨ä½“)", c
    ):
        return True
    return False


def _md_slash_is_likely_fraction_not_date(t: str, start: int, end: int, mo: int, day: int) -> bool:
    """
    ã€ŒåŠ å·¥é€Ÿåº¦ã�¯1/3ã�¨ã�—ã�¾ã�™ã€�ã�® 1/3 ã‚’ 1æœˆ3æ—¥ ã�¨èª¤èª�ã�—ã�ªã�„ã€‚
    ã€Œ4/1ã�¯å·¥å ´ã‚’ä¼‘ã�¿ã€�ã�® 4/1 ã�¯æ—¥ä»˜ã�®ã�¾ã�¾ï¼ˆç›´å¾Œã�Œã€Œã�¯ã€�ã�ªã‚‰åˆ†æ•°æ‰±ã�„ã�«ã�—ã�ªã�„ï¼‰ã€‚
    """
    if mo <= 0 or day <= 0:
        return True
    before = t[max(0, start - 32) : start]
    after = t[end : min(len(t), end + 14)]
    after_st = after.lstrip()
    if after_st.startswith("ã�¯"):
        return False
    if re.search(
        r"(?:åŠ å·¥é€Ÿåº¦|åŠ å·¥\s*ã‚¹ãƒ”ãƒ¼ãƒ‰|é€Ÿåº¦|å€�çŽ‡|ã‚¹ãƒ”ãƒ¼ãƒ‰|åŠ¹çŽ‡|å‰²å�ˆ)(?:\s*ã�¯)?\s*$",
        before,
    ):
        return True
    # 1/2ãƒ»1/3ãƒ»2/3 ç­‰ + ã€Œã�¨ã�—ã�¾ã�™ã€�ã€Œå€�ã€�â€¦ ã�¯åˆ†æ•°ãƒ»æ¯”çŽ‡å¯„ã‚Šï¼ˆã€Œ3/1ã�§ã�™ã€�ç­‰ã�®æ—¥ä»˜ã‚’èª¤ã‚¹ã‚­ãƒƒãƒ—ã�—ã�ªã�„ã‚ˆã�† ã�§ã�™/ã�§ã�‚ã‚‹ ã�¯å�«ã‚�ã�ªã�„ï¼‰
    frac_pat = re.compile(
        r"^(?:ã�¨ã�—ã�¾ã�™?|ã�¨ã�™ã‚‹|å€�|å‰²å�ˆ|ã�«ã�™ã‚‹|ã�«è¨­å®š|ã��ã‚‰ã�„|ç¨‹åº¦|ã�«å›ºå®š|ã�«å¤‰æ›´)"
    )
    if mo <= 12 and day <= 12 and frac_pat.match(after_st):
        if mo <= 2 or (mo == 3 and day <= 3):
            return True
    # ã€Œ1/2ã�§ã�™ã€�ã€Œ1/10ã�§ã�™ã€�ã�®ã‚ˆã�†ã�ªåˆ†æ¯�è¡¨ç�¾ï¼ˆå…ˆé ­ã�Œ 1/ ã�®ã�¿ï¼‰
    if (
        mo == 1
        and 2 <= day <= 12
        and re.match(r"^ã�§ã�™|ã�§ã�‚ã‚‹\b", after_st)
    ):
        return True
    return False


def _extract_calendar_dates_from_text(s: str, default_year: int) -> list[date]:
    """ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆå†…ã�®æ—¥ä»˜è¡¨è¨˜ã‚’ date ã�«å¤‰æ�›ï¼ˆåŸºæº–å¹´ã�¯è¨ˆç”»ã�®åŸºæº–å¹´ï¼‰ã€‚"""
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
        r"(\d{4})\s*å¹´\s*(\d{1,2})\s*æœˆ\s*(\d{1,2})\s*æ—¥?",
        t,
    ):
        add(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    for m in re.finditer(
        r"(\d{4})\s*[/\-\.ï¼�]\s*(\d{1,2})\s*[/\-\.ï¼�]\s*(\d{1,2})",
        t,
    ):
        add(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    for m in re.finditer(r"(\d{1,2})\s*æœˆ\s*(\d{1,2})\s*æ—¥", t):
        add(int(default_year), int(m.group(1)), int(m.group(2)))
    for m in re.finditer(
        r"(?<!\d)(\d{1,2})\s*[/ï¼�]\s*(\d{1,2})(?!\d)",
        t,
    ):
        mo_i, d_i = int(m.group(1)), int(m.group(2))
        if _md_slash_is_likely_fraction_not_date(t, m.start(), m.end(), mo_i, d_i):
            continue
        add(int(default_year), mo_i, d_i)
    return found


def _split_global_comment_into_chunks(blob: str) -> list[str]:
    """
    ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã‚’ã€Œç‹¬ç«‹ã�—ã�ŸæŒ‡ç¤ºã€�ã�®å¡Šã�«åˆ†ã�‘ã‚‹ã€‚
    æ”¹è¡Œï¼ˆExcel ã�® Alt+Enterãƒ»Unicode æ”¹è¡Œå�«ã‚€ï¼‰ã�§å¿…ã�šåˆ†å‰²ã�—ã€�å�Œä¸€è¡Œå†…ã�¯ ã€‚;ï¼› ã�§ç¶šã�‘ã�¦åˆ†å‰²ã€‚
    """
    t = unicodedata.normalize("NFKC", str(blob or "").strip())
    if not t:
        return []
    lines = [ln.strip() for ln in re.split(r"[\n\r\v\f\u2028\u2029]+", t) if ln.strip()]
    if not lines:
        return []
    chunks: list[str] = []
    for line in lines:
        subs = [c.strip() for c in re.split(r"[ã€‚;ï¼›]+", line) if c.strip()]
        if subs:
            chunks.extend(subs)
        else:
            chunks.append(line)
    return chunks


def parse_factory_closure_dates_from_global_comment(
    text: str, default_year: int
) -> set[date]:
    """
    ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�ã�«ã€�å·¥å ´è‡¨æ™‚ä¼‘æ¥­ã�ªã�©ã�¨æ—¥ä»˜ã�Œæ›¸ã�‹ã‚Œã�¦ã�„ã‚‹å ´å�ˆã�«
    ã��ã�®æ—¥ã‚’å·¥å ´ä¼‘ã�¿ï¼ˆå…¨å“¡é�žç¨¼åƒ�ãƒ»é…�å�°ã�§åŠ å·¥ã�—ã�ªã�„ï¼‰ã�¨ã�—ã�¦æ‰±ã�†æ—¥ä»˜é›†å�ˆã‚’è¿”ã�™ã€‚
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
    """å·¥å ´ä¼‘æ¥­æ—¥: å‹¤æ€ ä¸Šã�¯å…¨å“¡ is_working=False ã�¨ã�—ã€�ã��ã�®æ—¥ã�¯è¨­å‚™å‰²ä»˜ã‚’è¡Œã‚�ã�ªã�„ã€‚"""
    if not closure_dates or not attendance_data:
        return
    tag = "å·¥å ´ä¼‘æ¥­ï¼ˆãƒ¡ã‚¤ãƒ³ãƒ»ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆï¼‰"
    for d in sorted(closure_dates):
        if d not in attendance_data:
            logging.warning(
                "ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã�®å·¥å ´ä¼‘æ¥­æ—¥ %s ã�¯ãƒžã‚¹ã‚¿å‹¤æ€ ã�«è¡Œã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚"
                " ã��ã�®æ—¥ã�¯è¨ˆç”»ãƒ«ãƒ¼ãƒ—ã�«å�«ã�¾ã‚Œã�ªã�„å ´å�ˆã€�é…�å�°ä¸Šã�®åŠ¹æžœã�Œé™�å®šçš„ã�§ã�™ã€‚",
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
    ã€Œåˆ¶é™�æ’¤å»ƒã€�ã€Œã�‚ã‚‰ã‚†ã‚‹æ�¡ä»¶ã€�ç­‰: è¨­å‚™å°‚æœ‰ãƒ»æ™‚åˆ»ã‚¬ãƒ¼ãƒ‰ã�¾ã�§å�«ã‚�é…�å�°åˆ¶ç´„ã‚’ç·©ã‚�ã‚‹ï¼ˆabolish_all_scheduling_limitsï¼‰ã€‚
    """
    b = unicodedata.normalize("NFKC", str(blob or ""))
    strong = (
        "åˆ¶é™�æ’¤å»ƒ",
        "åˆ¶é™�ã‚’æ’¤å»ƒ",
        "ã�™ã�¹ã�¦ã�®åˆ¶é™�",
        "å…¨ã�¦ã�®åˆ¶é™�",
        "ã�‚ã‚‰ã‚†ã‚‹åˆ¶é™�",
        "ã�‚ã‚‰ã‚†ã‚‹æ�¡ä»¶",
        "ã�™ã�¹ã�¦ã�®æ�¡ä»¶",
        "å…¨ã�¦ã�®æ�¡ä»¶",
        "æ’¤å»ƒã�—ã�¦",
        "æ’¤å»ƒã�—",
    )
    if any(k in b for k in strong):
        out = dict(coerced)
        out["abolish_all_scheduling_limits"] = True
        out["ignore_skill_requirements"] = True
        out["ignore_need_minimum"] = True
        logging.warning(
            "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: åˆ¶é™�æ’¤å»ƒã‚­ãƒ¼ãƒ¯ãƒ¼ãƒ‰ã‚’æ¤œå‡ºã€‚è¨­å‚™å°‚æœ‰ãƒ»æ™‚åˆ»ã‚¬ãƒ¼ãƒ‰ã‚’å�«ã‚�é…�å�°ä¸Šã�®åˆ¶ç´„ã‚’ç·©ã‚�ã�¾ã�™ã€‚"
        )
        return out
    return coerced


def _maybe_fill_global_speed_rules_from_scheduler_notes(coerced: dict) -> dict:
    """
    AI ã�Œ global_speed_rules ã‚’ç©ºã�«ã�—ã�Ÿã�Œ scheduler_notes ã�«å…·ä½“ãƒ‘ã‚¿ãƒ¼ãƒ³ã�Œã�‚ã‚‹å ´å�ˆã�®è£œå®Œã€‚
    åºƒã��æŽ¨æ¸¬ã�—ã�ªã�„ï¼ˆç†±èž�ç�€ï¼‹æ¤œæŸ»ï¼‹1/3 ç³»ã�®ã�¿ï¼‰ã€‚
    """
    if not isinstance(coerced, dict):
        return coerced
    if coerced.get("global_speed_rules"):
        return coerced
    sn = str(coerced.get("scheduler_notes_ja") or "")
    t = unicodedata.normalize("NFKC", sn)
    if "ç†±èž�ç�€" not in t or "æ¤œæŸ»" not in t:
        return coerced
    if not re.search(r"(?:1\s*/\s*3|ï¼‘\s*/\s*3|ä¸‰åˆ†ã�®ä¸€|3\s*åˆ†ã�®\s*1)", t):
        return coerced
    out = dict(coerced)
    out["global_speed_rules"] = [
        {
            "process_contains": "ç†±èž�ç�€",
            "machine_contains": "æ¤œæŸ»",
            "speed_multiplier": 1.0 / 3.0,
        }
    ]
    logging.info(
        "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: scheduler_notes_ja ã�‹ã‚‰ global_speed_rules ã‚’è£œå®Œï¼ˆç†±èž�ç�€ãƒ»æ¤œæŸ»ãƒ»1/3ï¼‰"
    )
    return out


def _finalize_global_priority_override(blob: str, coerced: dict) -> dict:
    """ã‚½ãƒ­è£œæ­£ã�®å¾Œã€�abolish ã�Œ true ã�ªã‚‰ã‚¹ã‚­ãƒ«ãƒ»äººæ•°ã‚‚å¼·åˆ¶ã‚ªãƒ³ã€‚"""
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
    ã€Œä¸€äººã�§æ‹…å½“ã€�ã€Œå�˜ç‹¬ã€�ç­‰ã�§äººæ•°ã� ã�‘ç·©ã‚�ã�¦ã‚‚ã€�æŒ‡å��ãƒ¡ãƒ³ãƒ�ãƒ¼ã�Œã‚¹ã‚­ãƒ«é�žè©²å½“ã� ã�¨é…�å�°ã�•ã‚Œã�ªã�„ã€‚
    ã��ã�®å ´å�ˆã�¯ã‚¹ã‚­ãƒ«ç„¡è¦–ã‚’å�Œæ™‚ã�«ç«‹ã�¦ã‚‹ã€‚
    """
    if not coerced.get("ignore_need_minimum") or coerced.get("ignore_skill_requirements"):
        return coerced
    b = unicodedata.normalize("NFKC", str(blob or ""))
    solo_kw = ("ä¸€äºº", "ã�²ã�¨ã‚Š", "å�˜ç‹¬", "ï¼‘äºº", "1äºº", "ç‹¬è‡ª", "å�˜èº«")
    if any(k in b for k in solo_kw):
        out = dict(coerced)
        out["ignore_skill_requirements"] = True
        logging.info(
            "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: å�˜ç‹¬ç³»ã‚­ãƒ¼ãƒ¯ãƒ¼ãƒ‰ã�®ã�Ÿã‚� ignore_skill_requirements ã‚’è£œåŠ©çš„ã�« true ã�«ã�—ã�¾ã�—ã�Ÿã€‚"
        )
        return out
    return coerced


def _coerce_task_preferred_operators_dict(raw_val) -> dict:
    """AI ã�® task_preferred_operators ã‚’ {ä¾�é ¼NO: æ°�å��} ã�«æ­£è¦�åŒ–ã€‚"""
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
    AI ã�¾ã�Ÿã�¯ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ã�®æ—¥ä»˜ãƒªã‚¹ãƒˆã‚’ YYYY-MM-DD æ–‡å­—åˆ—ã�®æ˜‡é †ãƒ¦ãƒ‹ãƒ¼ã‚¯ã�«æ­£è¦�åŒ–ã€‚
    è¦�ç´ ã�¯ ISO æ–‡å­—åˆ—ãƒ»Excel æ—¥ä»˜ãƒ»ã€Œ4/1ã€�ç¨‹åº¦ã�®çŸ­æ–‡ã�§ã‚‚å�¯ã€‚
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
    Gemini ã�® global_speed_rules ã‚’æ­£è¦�åŒ–ã€‚
    å�„è¦�ç´ : process_contains / machine_containsï¼ˆã�„ã�šã‚Œã�‹å¿…é ˆãƒ»éƒ¨åˆ†ä¸€è‡´ç”¨ï¼‰, speed_multiplierï¼ˆæ—¢å­˜é€Ÿåº¦ã�«ä¹—ç®—ã€�0è¶…ã€œ10ä»¥ä¸‹ï¼‰ã€‚
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
    """sub ã�Œç©ºã�§ã�ªã�‘ã‚Œã�°ã€�å·¥ç¨‹å��ã�¾ã�Ÿã�¯æ©Ÿæ¢°å��ã�®ã�„ã�šã‚Œã�‹ã�«éƒ¨åˆ†ä¸€è‡´ã�™ã‚Œã�° Trueã€‚"""
    if not sub_nfkc:
        return True
    return sub_nfkc in pnorm or sub_nfkc in mnorm


def _global_speed_multiplier_for_row(process_name: str, machine_name: str, rules: list) -> float:
    """
    å·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ã�«ä¸€è‡´ã�™ã‚‹ãƒ«ãƒ¼ãƒ«ã�® speed_multiplier ã‚’æŽ›ã�‘å�ˆã‚�ã�›ã‚‹ï¼ˆä¸€è‡´ã�ªã�—ã�¯ 1.0ï¼‰ã€‚

    process_contains / machine_contains ã�¯ã��ã‚Œã�žã‚Œ **å·¥ç¨‹å��ã�¾ã�Ÿã�¯æ©Ÿæ¢°å��ã�®ã�©ã�¡ã‚‰ã�‹** ã�«å�«ã�¾ã‚Œã‚Œã�°ã‚ˆã�„ã€‚
    ä¸¡æ–¹æŒ‡å®šæ™‚ã�¯ ANDï¼ˆä¾‹: ã€Œç†±èž�ç�€ã€�ã�¨ã€Œæ¤œæŸ»ã€�ã�Œã€�åˆ—ã�®çµ„ã�¿å�ˆã‚�ã�›ã�§ä¸¡æ–¹ç�¾ã‚Œã‚‹è¡Œã�«ãƒžãƒƒãƒ�ã€‚
    ãƒžã‚¹ã‚¿ä¸Šã�§å·¥ç¨‹=æ¤œæŸ»ãƒ»æ©Ÿæ¢°=ç†±èž�ç�€æ©Ÿ ã�®ã‚ˆã�†ã�«ã‚­ãƒ¼ãƒ¯ãƒ¼ãƒ‰ã�Œé€†å�´ã�®åˆ—ã�«ã�‚ã�£ã�¦ã‚‚å�Œã�˜ãƒ«ãƒ¼ãƒ«ã�§åŠ¹ã��ã€‚
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
    Gemini ã�Œ task_preferred_operators ã�«èª¤ã�£ã�¦é•·æ–‡ã‚’å…¥ã‚Œã�Ÿå ´å�ˆã�ªã�©ã€�
    è‡ªç„¶è¨€èªžæ–­ç‰‡ã�‹ã‚‰ global_day_process_operator_rules ç›¸å½“ã‚’æŽ¨å®šã�™ã‚‹ï¼ˆä¿�å®ˆçš„ï¼‰ã€‚
    ä¾‹: ã€Œ2026/4/4 å·¥ç¨‹å��:EC æ£®ä¸‹ã�¨å®®å³¶ã‚’é…�å�°ã€�
    """
    t = unicodedata.normalize("NFKC", str(text or "")).strip()
    if len(t) < 6:
        return []
    dates = _extract_calendar_dates_from_text(t, int(ref_y))
    if not dates:
        return []
    d0 = dates[0]
    proc_m = re.search(
        r"å·¥ç¨‹å��?\s*[:ï¼š]?\s*([A-Za-z0-9ä¸€-é¾¯ãƒ¼ãƒ»ã€†ã€…]+)",
        t,
    )
    pc = proc_m.group(1).strip() if proc_m else ""
    if not pc:
        m2 = re.search(r"([\dA-Za-zä¸€-é¾¯ãƒ¼ãƒ»ã€†ã€…]{1,12})\s*å·¥ç¨‹", t)
        pc = m2.group(1).strip() if m2 else ""
    if not pc:
        return []
    names: list[str] = []
    for m in re.finditer(
        r"([\u3040-\u9FFFã€…ãƒ¼ãƒ»A-Za-zãƒ»ã€†ã€…]{1,16}?)\s*ã�¨\s*([\u3040-\u9FFFã€…ãƒ¼ãƒ»A-Za-zãƒ»ã€†ã€…]{1,16}?)\s*ã‚’?\s*(?:é…�å�°|é…�å±ž|çµ„ã�¾ã�›|å�Œä¸€ãƒ�ãƒ¼ãƒ )",
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
    Gemini ã�Œ task_preferred_operators ã�« **é…�åˆ—**ã‚„èª¤ã‚¹ã‚­ãƒ¼ãƒžï¼ˆworkstation_id ç­‰ï¼‰ã‚’è¿”ã�—ã�Ÿã�¨ã��ã€�
    æ�¨ã�¦ã�šã�« global_day_process_operator_rules / scheduler_notes_ja ã�¸æ•‘æ¸ˆã�™ã‚‹ã€‚
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
            if isinstance(v, str) and len(v) > 35 and ("é…�" in v or "å·¥ç¨‹" in v):
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
    """Gemini æˆ»ã‚Šã‚’é…�å�°ç”¨ãƒ•ãƒ©ã‚°ãƒ»å·¥å ´ä¼‘æ¥­æ—¥ãƒªã‚¹ãƒˆã�«æ­£è¦�åŒ–ã€‚"""
    y0 = int(reference_year) if reference_year is not None else date.today().year

    def as_bool(v):
        if v is True:
            return True
        if v is False:
            return False
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        s = unicodedata.normalize("NFKC", str(v).strip()).lower()
        return s in ("true", "1", "yes", "ã�¯ã�„", "on")

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
    """Gemini å¿œç­”ã�‹ã‚‰ JSON ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆ1ã�¤ã‚’å�–ã‚Šå‡ºã�™ï¼ˆ```json ãƒ•ã‚§ãƒ³ã‚¹ä»˜ã��ã�§ã‚‚å�¯ï¼‰ã€‚"""
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
    """Gemini æœªä½¿ç”¨ãƒ»å¿œç­”è§£é‡ˆå¤±æ•—æ™‚: ãƒ«ãƒ¼ãƒ«ãƒ™ãƒ¼ã‚¹ã�§å·¥å ´ä¼‘æ¥­æ—¥ã� ã�‘è£œå®Œï¼ˆå¾“æ�¥äº’æ�›ï¼‰ã€‚"""
    out = dict(coerced)
    rx = parse_factory_closure_dates_from_global_comment(blob, ref_y)
    out["factory_closure_dates"] = sorted({d.isoformat() for d in rx})
    return out


def analyze_global_priority_override_comment(
    text: str, members: list, reference_year: int, ai_sheet_sink: dict | None = None
) -> dict:
    """
    ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�ï¼ˆUI ä¸Šã�®è‡ªç”±è¨˜è¿°ï¼‰ã‚’ **Gemini ã�§ä¸€æ‹¬è§£é‡ˆ**ã�—ã€�é…�å�°ã�«åŠ¹ã�� JSON ã�«è�½ã�¨ã�™ã€‚
    è‡ªç„¶è¨€èªžã�®æ–‡è„ˆåˆ‡ã‚Šåˆ†ã�‘ãƒ»æ”¹è¡Œã�®åˆ¥æŒ‡ç¤ºè§£é‡ˆã�¯ AI ã�«ä»»ã�›ã€�æˆ»ã‚Šå€¤ã�®ã‚­ãƒ¼ã� ã�‘ã‚·ã‚¹ãƒ†ãƒ ã�Œæ©Ÿæ¢°é�©ç”¨ã�™ã‚‹ã€‚

    - factory_closure_dates: **å·¥å ´å…¨ä½“**ã�§ç¨¼åƒ�ã�—ã�ªã�„æ—¥ï¼ˆå…¨å“¡é�žç¨¼åƒ�æ‰±ã�„ï¼‰ã�® YYYY-MM-DD æ–‡å­—åˆ—ã�®é…�åˆ—ã€‚è©²å½“ã�ªã�—ã�¯ []ã€‚
    - ignore_skill_requirements / ignore_need_minimum / abolish_all_scheduling_limits / task_preferred_operators: å¾“æ�¥ã�©ã�Šã‚Šã€‚
    - global_speed_rules: **å·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��**ã�¸ã�®éƒ¨åˆ†ä¸€è‡´ï¼ˆå�„ã‚­ãƒ¼ãƒ¯ãƒ¼ãƒ‰ã�¯ **ã�©ã�¡ã‚‰ã�®åˆ—ã�«ã�‚ã�£ã�¦ã‚‚å�¯**ï¼‰ã�§ã€�æ—¢å­˜ã�®åŠ å·¥é€Ÿåº¦ï¼ˆã‚·ãƒ¼ãƒˆï¼�ä¸Šæ›¸ã��å¾Œï¼‰ã�« **ä¹—ç®—**ã�™ã‚‹ãƒ«ãƒ¼ãƒ«ã�®é…�åˆ—ã€‚è©²å½“ã�ªã�—ã�¯ []ã€‚
    - global_day_process_operator_rules: **æ—¥ä»˜ï¼‹å·¥ç¨‹å��ã�®éƒ¨åˆ†ä¸€è‡´ï¼‹è¤‡æ•°ãƒ¡ãƒ³ãƒ�ãƒ¼**ã‚’ã€�å½“æ—¥ã��ã�®å·¥ç¨‹ã�®ã‚¿ã‚¹ã‚¯ã�®**ãƒ�ãƒ¼ãƒ å…¨å“¡ã�«å¿…ã�šå�«ã‚�ã‚‹**ãƒ«ãƒ¼ãƒ«ã�®é…�åˆ—ã€‚è©²å½“ã�ªã�—ã�¯ []ã€‚
    - scheduler_notes_ja: ä¸Šè¨˜ã�«è�½ã�¨ã�—ã��ã‚Œã�ªã�„è£œè¶³ã‚„é�‹ç”¨ãƒ¡ãƒ¢ï¼ˆé€Ÿåº¦ã�¯å�¯èƒ½ã�ªã‚‰ global_speed_rules ã‚‚ä½µè¨˜ï¼‰ã€‚

    API ã‚­ãƒ¼ç„¡ã�—ãƒ»JSON è§£é‡ˆå¤±æ•—æ™‚: ä¸Šè¨˜ãƒ–ãƒ¼ãƒ«ãƒ»æŒ‡å��ã�¯æ—¢å®šå€¤ã€�å·¥å ´ä¼‘æ¥­æ—¥ã�®ã�¿å¾“æ�¥ã�®ãƒ«ãƒ¼ãƒ«ãƒ™ãƒ¼ã‚¹è§£æž�ã�§è£œå®Œã€‚
    """
    ref_y = int(reference_year) if reference_year is not None else date.today().year
    empty = _coerce_global_priority_override_dict({}, ref_y)
    if not text or not str(text).strip():
        if ai_sheet_sink is not None:
            ai_sheet_sink["ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜_AI_API"] = "ã‚¹ã‚­ãƒƒãƒ—ï¼ˆãƒ¡ã‚¤ãƒ³åŽŸæ–‡ã�ªã�—ï¼‰"
        return empty
    blob = str(text).strip()
    mem_sig = ",".join(sorted(str(m).strip() for m in (members or []) if m))
    cache_fingerprint = f"{GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX}{ref_y}\n{blob}\n{mem_sig}"
    cache_key = hashlib.sha256(cache_fingerprint.encode("utf-8")).hexdigest()
    ai_cache = load_ai_cache()
    cached = get_cached_ai_result(ai_cache, cache_key, content_key=cache_fingerprint)
    if cached is not None:
        logging.info("ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: ã‚­ãƒ£ãƒƒã‚·ãƒ¥ãƒ’ãƒƒãƒˆï¼ˆGemini ã�¯å‘¼ã�³ã�¾ã�›ã‚“ï¼‰ã€‚")
        if ai_sheet_sink is not None:
            ai_sheet_sink["ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜_AI_API"] = "ã�ªã�—ï¼ˆã‚­ãƒ£ãƒƒã‚·ãƒ¥ä½¿ç”¨ï¼‰"
        return _finalize_global_priority_override(
            blob, _coerce_global_priority_override_dict(cached, ref_y)
        )

    if not API_KEY:
        logging.info("GEMINI_API_KEY æœªè¨­å®šã�®ã�Ÿã‚�ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜ã�® AI è§£æž�ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚")
        if ai_sheet_sink is not None:
            ai_sheet_sink["ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜_AI_API"] = "ã�ªã�—ï¼ˆAPIã‚­ãƒ¼æœªè¨­å®šãƒ»å·¥å ´ä¼‘æ¥­ã�®ã�¿ãƒ«ãƒ¼ãƒ«è£œå®Œï¼‰"
        coerced = _apply_regex_factory_closure_fallback(
            _coerce_global_priority_override_dict({}, ref_y), blob, ref_y
        )
        return _finalize_global_priority_override(blob, coerced)

    member_sample = ", ".join(str(m) for m in (members or [])[:80])
    if len(members or []) > 80:
        member_sample += " â€¦"

    prompt = f"""ã�‚ã�ªã�Ÿã�¯å·¥å ´ã�®é…�å�°è¨ˆç”»ã‚·ã‚¹ãƒ†ãƒ ç”¨ã‚¢ã‚·ã‚¹ã‚¿ãƒ³ãƒˆã�§ã�™ã€‚
Excel ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆã�® **ã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�**ï¼ˆè‡ªç”±è¨˜è¿°ãƒ»è‡ªç„¶è¨€èªžï¼‰ã�® **å…¨æ–‡** ã‚’èª­ã�¿ã€�æ¬¡ã�®ã‚­ãƒ¼ã� ã�‘ã‚’æŒ�ã�¤ JSON ã‚’1ã�¤è¿”ã�—ã�¦ã��ã� ã�•ã�„ã€‚

ã€�å½¹å‰²ã€‘
ãƒ¦ãƒ¼ã‚¶ãƒ¼ã�¯æ”¹è¡Œã‚„å�¥ç‚¹ã�§è¤‡æ•°ã�®æŒ‡ç¤ºã‚’æ›¸ã��ã�“ã�¨ã�Œã�‚ã‚Šã�¾ã�™ã€‚**æ–‡è„ˆã‚’èª­ã�¿åˆ†ã�‘**ã€�é…�å�°ã‚·ã‚¹ãƒ†ãƒ ã�Œ **æ©Ÿæ¢°çš„ã�«é�©ç”¨ã�§ã��ã‚‹å€¤** ã�«è�½ã�¨ã�—è¾¼ã‚“ã�§ã��ã� ã�•ã�„ã€‚
æŽ¨æ¸¬ã�§ãƒ–ãƒ¼ãƒ«ã‚’ true ã�«ã�—ã�ªã�„ã�“ã�¨ã€‚æ ¹æ‹ ã�Œæ˜Žç¢ºã�ªã�¨ã��ã� ã�‘ trueã€‚

ã€�æœ€å„ªå…ˆã€‘
ã�“ã�®æ¬„ã�®å†…å®¹ã�¯ãƒžã‚¹ã‚¿ãƒ»ã‚¹ã‚­ãƒ«ãƒ»needãƒ»ã‚¿ã‚¹ã‚¯è¡Œãƒ»ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã�® AI æŒ‡å��ã‚ˆã‚Šå„ªå…ˆã�•ã‚Œã‚‹ä¾‹å¤–æŒ‡ç¤ºã�¨ã�—ã�¦æ‰±ã‚�ã‚Œã�¾ã�™ã€‚

ã€�æ”¹è¡Œãƒ»è¤‡æ•°è¡Œã€‘
å�„è¡Œãƒ»å�„æ–‡ã�¯ **åŽŸå‰‡ã�¨ã�—ã�¦ç‹¬ç«‹ã�—ã�ŸæŒ‡ç¤º** ã�§ã�™ã€‚è¡Œã‚’ã�¾ã�Ÿã�„ã�§1ã�¤ã�«ã�¾ã�¨ã‚�ã�Ÿã‚Šã€�**å‰²å�ˆè¡¨ç�¾ï¼ˆä¾‹ 1/3ï¼‰ã‚’æ—¥ä»˜ã�¨çµ�ã�³ä»˜ã�‘ã�Ÿã‚Šã�—ã�ªã�„**ã�“ã�¨ã€‚

ã€�ã‚­ãƒ¼åˆ¥ãƒ«ãƒ¼ãƒ«ã€‘

A) **factory_closure_dates** ï¼ˆé…�åˆ—ãƒ»å¿…é ˆï¼‰
   - **å·¥å ´å…¨ä½“**ã�Œç¨¼åƒ�ã�—ã�ªã�„æ—¥ï¼ˆè‡¨æ™‚ä¼‘æ¥­ãƒ»å…¨å·¥å ´ä¼‘ã�¿ãƒ»ã��ã�®æ—¥ã�¯åŠ å·¥ã�—ã�ªã�„ç­‰ï¼‰ã�®æ—¥ä»˜ã‚’ **YYYY-MM-DD** ã�®æ–‡å­—åˆ—ã�§åˆ—æŒ™ã€‚
   - **å€‹äººã�®ä¼‘ã�¿ãƒ»ç‰¹å®šãƒ©ã‚¤ãƒ³ã� ã�‘**ã�®å�œæ­¢ã�¯ã�“ã�“ã�« **å�«ã‚�ã�ªã�„**ï¼ˆ[]ï¼‰ã€‚
   - è©²å½“ã�Œã�ªã�‘ã‚Œã�° **ç©ºé…�åˆ— []**ï¼ˆã‚­ãƒ¼çœ�ç•¥ä¸�å�¯ï¼‰ã€‚
   - å¹´ã�Œçœ�ç•¥ã�•ã‚Œã�¦ã�„ã‚Œã�°è¥¿æš¦ {ref_y} å¹´ã�¨ã�—ã�¦è§£é‡ˆã€‚

B) **ignore_skill_requirements** / **ignore_need_minimum** / **abolish_all_scheduling_limits** / **task_preferred_operators**
   - å¾“æ�¥ã�©ã�Šã‚Šï¼ˆé…�å�°ã�®ã‚¹ã‚­ãƒ«ç„¡è¦–ãƒ»äººæ•°1å›ºå®šãƒ»åˆ¶é™�æ’¤å»ƒãƒ»ä¾�é ¼NOâ†’ä¸»æ‹…å½“OPæŒ‡å��ï¼‰ã€‚è©²å½“ã�ªã�‘ã‚Œã�° false ã�¾ã�Ÿã�¯ {{}}ã€‚

C) **global_speed_rules** ï¼ˆé…�åˆ—ãƒ»å¿…é ˆï¼‰
   - ç‰¹å®šã�® **å·¥ç¨‹å��**ï¼ˆExcelã€Œå·¥ç¨‹å��ã€�åˆ—ï¼‰ã‚„ **æ©Ÿæ¢°å��**ï¼ˆã€Œæ©Ÿæ¢°å��ã€�åˆ—ï¼‰ã�«å¯¾ã�—ã€�**æ—¢å­˜ã�®åŠ å·¥é€Ÿåº¦ã�«æŽ›ã�‘ã‚‹å€�çŽ‡** ã‚’æŒ‡å®šã�™ã‚‹ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�®ãƒªã‚¹ãƒˆã€‚
   - å�„ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�®ã‚­ãƒ¼:
     - "process_contains": æ–‡å­—åˆ—ï¼ˆçœ�ç•¥å�¯ï¼‰ã€‚**å·¥ç¨‹å��ã�¾ã�Ÿã�¯æ©Ÿæ¢°å��ã�®ã�„ã�šã‚Œã�‹**ã�« **éƒ¨åˆ†ä¸€è‡´**ï¼ˆNFKC æƒ³å®šï¼‰ã€‚
     - "machine_contains": æ–‡å­—åˆ—ï¼ˆçœ�ç•¥å�¯ï¼‰ã€‚**å·¥ç¨‹å��ã�¾ã�Ÿã�¯æ©Ÿæ¢°å��ã�®ã�„ã�šã‚Œã�‹**ã�« **éƒ¨åˆ†ä¸€è‡´**ã€‚
     - "speed_multiplier": æ­£ã�®æ•°ã€‚**1/3 ã�®é€Ÿåº¦**ã�ªã‚‰ç´„ **0.333333**ï¼ˆæ—¢å­˜é€Ÿåº¦ Ã— ã�“ã�®å€¤ï¼‰ã€‚**2å€�é€Ÿ**ã�ªã‚‰ 2.0ã€‚
   - **ä¸¡æ–¹æŒ‡å®šæ™‚ã�¯ AND**ï¼ˆ2ã�¤ã�®ã‚­ãƒ¼ãƒ¯ãƒ¼ãƒ‰ã�Œã€�**ä¸¡æ–¹ã�¨ã‚‚**ã€Œå·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ã�®ã�©ã�¡ã‚‰ã�‹ã€�ã�«ç�¾ã‚Œã‚‹è¡Œï¼‰ã€‚ä¾‹: å·¥ç¨‹=æ¤œæŸ»ãƒ»æ©Ÿæ¢°=ç†±èž�ç�€æ©Ÿ ã�§ã‚‚ã€�å·¥ç¨‹=ç†±èž�ç�€ãƒ»æ©Ÿæ¢°=æ¤œæŸ»ç”¨è¨­å‚™ ã�§ã‚‚ãƒžãƒƒãƒ�ã�—ã�†ã‚‹ã€‚
   - ã�©ã�¡ã‚‰ã�‹ä¸€æ–¹ã� ã�‘æŒ‡å®šã�™ã‚Œã�°ã€�ã��ã�®ã‚­ãƒ¼ãƒ¯ãƒ¼ãƒ‰ã�Œå·¥ç¨‹å��ã�¾ã�Ÿã�¯æ©Ÿæ¢°å��ã�®ã�©ã�¡ã‚‰ã�‹ã�«ã�‚ã‚Œã�°ãƒžãƒƒãƒ�ã€‚
   - è©²å½“æŒ‡ç¤ºã�Œã�ªã�‘ã‚Œã�° **ç©ºé…�åˆ— []**ã€‚
   - ä¾‹: ã€Œç†±èž�ç�€ã‚’ä½¿ã�†æ¤œæŸ»ã�®åŠ å·¥é€Ÿåº¦ã�¯1/3ã€�â†’
     [{{"process_contains":"ç†±èž�ç�€","machine_contains":"æ¤œæŸ»","speed_multiplier":0.333333}}]
     ï¼ˆã€Œç†±èž�ç�€ã€�ã�¨ã€Œæ¤œæŸ»ã€�ã�Œå·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ã�®çµ„ã�¿å�ˆã‚�ã�›ã�§æ�ƒã�†ã‚¿ã‚¹ã‚¯ã�®é€Ÿåº¦ã�Œç´„1/3ã�«ã�ªã‚‹ï¼‰

D) **scheduler_notes_ja** ï¼ˆæ–‡å­—åˆ—ãƒ»å¿…é ˆï¼‰
   - ä¸Šè¨˜ã‚­ãƒ¼ã�«è�½ã�¨ã�—ã��ã‚Œã�ªã�„è£œè¶³ã€‚é€Ÿåº¦ã�¯ **global_speed_rules ã�§æ§‹é€ åŒ–ã�§ã��ã‚‹ã�¨ã��ã�¯å¿…ã�šã��ã�¡ã‚‰ã�«ã‚‚å‡ºã�™**ï¼ˆã�“ã�“ã�¯äººé–“å�‘ã�‘è¦�ç´„ã�§ã‚‚ã‚ˆã�„ï¼‰ã€‚ç„¡ã�‘ã‚Œã�° ""ã€‚

E) **interpretation_ja** ï¼ˆæ–‡å­—åˆ—ãƒ»å¿…é ˆï¼‰
   - åŽŸæ–‡ã�®è¦�ç´„ã‚’1æ–‡ï¼ˆ200æ–‡å­—ä»¥å†…ï¼‰ã€‚

F) **global_day_process_operator_rules** ï¼ˆé…�åˆ—ãƒ»å¿…é ˆï¼‰
   - **ç‰¹å®šã�®ç¨¼åƒ�æ—¥**ã�‹ã�¤ **å·¥ç¨‹å��ï¼ˆExcelã€Œå·¥ç¨‹å��ã€�åˆ—ï¼‰ã�®éƒ¨åˆ†ä¸€è‡´** ã�«å½“ã�¦ã�¯ã�¾ã‚‹ã‚¿ã‚¹ã‚¯ã�«ã�¤ã�„ã�¦ã€�
     åˆ—æŒ™ã�—ã�Ÿ **å…¨ãƒ¡ãƒ³ãƒ�ãƒ¼ã‚’å�Œä¸€ãƒ�ãƒ¼ãƒ ã�«å¿…ã�šå�«ã‚�ã‚‹** ãƒ«ãƒ¼ãƒ«ï¼ˆ**OP/AS ã�©ã�¡ã‚‰ã�®ã‚¹ã‚­ãƒ«ã�§ã‚‚å�¯**ã€‚æ°�å��è§£æ±ºã�¯ **æ‹…å½“OPæŒ‡å��ã�¨å�Œã�˜**ï¼‰ã€‚
   - **ä¾�é ¼NOã�Œåˆ†ã�‹ã‚‹ä¸»æ‹…å½“ã�®1å��æŒ‡å��**ã�¯ **task_preferred_operators** ã‚’ä½¿ã�†ã�“ã�¨ã€‚åŽŸæ–‡ã�Œ **ã€Œâ—¯æœˆâ—¯æ—¥ã�®â–³å·¥ç¨‹ã�«ï¼¡ã�¨ï¼¢ã‚’é…�å�°ã€�** ã�®ã‚ˆã�†ã�« **æ—¥ä»˜ãƒ»å·¥ç¨‹ãƒ»è¤‡æ•°å��**ã�®ã�¨ã��ã�¯ **æœ¬é…�åˆ—**ã�¸è�½ã�¨ã�™ã€‚
   - å�„ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�®ã‚­ãƒ¼:
     - "date": **YYYY-MM-DD**ï¼ˆã��ã�®æ—¥ã�«å‰²ã‚Šå½“ã�¦ã‚‹ãƒ­ãƒ¼ãƒ«ã�«é�©ç”¨ï¼‰
     - "process_contains": å·¥ç¨‹å��ã�« **éƒ¨åˆ†ä¸€è‡´**ï¼ˆNFKC æƒ³å®šï¼‰ã€‚ä¾‹: "EC"
     - "operator_names": æ°�å��ã�®é…�åˆ—ï¼ˆä¾‹: ["æ£®ä¸‹", "å®®å³¶ã€€èŠ±å­�"]ï¼‰
   - è©²å½“æŒ‡ç¤ºã�Œã�ªã�‘ã‚Œã�° **ç©ºé…�åˆ— []**ã€‚

ã€�è¿”ç­”å½¢å¼�ã€‘
å…ˆé ­ã�Œ {{ ã�§çµ‚ã‚�ã‚Šã�Œ }} ã�® **JSON ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆ1ã�¤ã�®ã�¿**ï¼ˆèª¬æ˜Žæ–‡ãƒ»ãƒžãƒ¼ã‚¯ãƒ€ã‚¦ãƒ³ç¦�æ­¢ï¼‰ã€‚

å¿…é ˆã‚­ãƒ¼ä¸€è¦§:
- "factory_closure_dates": string ã�®é…�åˆ—ï¼ˆYYYY-MM-DDï¼‰
- "ignore_skill_requirements": true ã�¾ã�Ÿã�¯ false
- "ignore_need_minimum": true ã�¾ã�Ÿã�¯ false
- "abolish_all_scheduling_limits": true ã�¾ã�Ÿã�¯ false
- "task_preferred_operators": **JSON ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�®ã�¿**ï¼ˆã‚­ãƒ¼=ä¾�é ¼NOãƒ»å€¤=ä¸»æ‹…å½“æ°�å��ï¼‰ã€‚**é…�åˆ—ã�«ã�—ã�¦ã�¯ã�ªã‚‰ã�ªã�„**ã€‚è©²å½“ã�ªã�—ã�¯ {{}}
- "global_speed_rules": ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�®é…�åˆ—ï¼ˆè©²å½“ã�ªã�—ã�¯ []ï¼‰
- "global_day_process_operator_rules": ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�®é…�åˆ—ï¼ˆè©²å½“ã�ªã�—ã�¯ []ï¼‰
- "scheduler_notes_ja": æ–‡å­—åˆ—
- "interpretation_ja": æ–‡å­—åˆ—

ã€�åŸºæº–å¹´ã€‘ æ—¥ä»˜è¨€å�Šã�Œã�‚ã‚Œã�°è¥¿æš¦ {ref_y} å¹´ã�¨ã�—ã�¦è§£é‡ˆã�—ã�¦ã‚ˆã�„ã€‚

ã€�ç™»éŒ²ãƒ¡ãƒ³ãƒ�ãƒ¼å��ã�®å�‚è€ƒã€‘ï¼ˆç…§å�ˆç”¨ã€‚JSON ã‚­ãƒ¼ã�«ã�¯å�«ã‚�ã�ªã�„ï¼‰
{member_sample}

ã€�ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆãƒ»åŽŸæ–‡ã€‘
{blob}
"""
    try:
        ppath = os.path.join(log_dir, "ai_global_priority_override_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆå…¨æ–‡ â†’ %s", ppath)
    except OSError as ex:
        logging.warning("ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆä¿�å­˜å¤±æ•—: %s", ex)

    client = genai.Client(api_key=API_KEY)
    try:
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        parsed = _parse_global_priority_override_gemini_response(res)
        if parsed is None:
            logging.warning(
                "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: AI å¿œç­”ã�‹ã‚‰ JSON ã‚’è§£é‡ˆã�§ã��ã�¾ã�›ã‚“ã�§ã�—ã�Ÿã€‚ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã�›ã�šã€�æ¬¡å›žå†�è©¦è¡Œã�•ã‚Œã�¾ã�™ã€‚"
            )
            try:
                rpath = os.path.join(log_dir, "ai_global_priority_override_last_response.txt")
                with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                    rf.write(_gemini_result_text(res) or "")
            except OSError:
                pass
            if ai_sheet_sink is not None:
                ai_sheet_sink["ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜_AI_API"] = "ã�‚ã‚Šï¼ˆJSONè§£é‡ˆå¤±æ•—ãƒ»å·¥å ´ä¼‘æ¥­ã�¯ãƒ«ãƒ¼ãƒ«è£œå®Œï¼‰"
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
            "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: AI è§£é‡ˆ factoryä¼‘æ¥­=%sæ—¥ é€Ÿåº¦ãƒ«ãƒ¼ãƒ«=%sä»¶ æ—¥Ã—å·¥ç¨‹ãƒ�ãƒ¼ãƒ =%sä»¶ skill=%s need1=%s abolish=%s task_pref=%sä»¶ â€” %s",
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
            ai_sheet_sink["ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜_AI_API"] = "ã�‚ã‚Š"
        return coerced
    except Exception as e:
        logging.warning("ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: Gemini å‘¼ã�³å‡ºã�—å¤±æ•—: %s", e)
        if ai_sheet_sink is not None:
            ai_sheet_sink["ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜_AI_API"] = f"å¤±æ•—: {e}"[:500]
        coerced = _apply_regex_factory_closure_fallback(
            _coerce_global_priority_override_dict({}, ref_y), blob, ref_y
        )
        return _finalize_global_priority_override(blob, coerced)


def default_result_task_sheet_column_order(max_history_len: int) -> list:
    """çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®æ—¢å®šåˆ—é †ï¼ˆå±¥æ­´åˆ—æ•°ã�¯å®Ÿè¡Œæ™‚ã�«æ±ºã�¾ã‚‹ï¼‰ã€‚"""
    hist = [f"å±¥æ­´{i+1}" for i in range(max_history_len)]
    return [
        "ã‚¹ãƒ†ãƒ¼ã‚¿ã‚¹",
        "ã‚¿ã‚¹ã‚¯ID",
        "å·¥ç¨‹å��",
        "æ©Ÿæ¢°å��",
        "å„ªå…ˆåº¦",
        RESULT_TASK_COL_DISPATCH_TRIAL_ORDER,
        *hist,
        "å¿…è¦�OP(ä¸Šæ›¸)",
        "ã‚¿ã‚¹ã‚¯åŠ¹çŽ‡",
        "åŠ å·¥é€”ä¸­",
        "ç‰¹åˆ¥æŒ‡å®šã�‚ã‚Š",
        "æ‹…å½“OPæŒ‡å��",
        "å›žç­”ç´�æœŸ",
        "æŒ‡å®šç´�æœŸ",
        "è¨ˆç”»åŸºæº–ç´�æœŸ",
        TASK_COL_RAW_INPUT_DATE,
        "ç´�æœŸç·Šæ€¥",
        "åŠ å·¥é–‹å§‹æ—¥",
        "é…�å®Œ_åŠ å·¥é–‹å§‹",
        "é…�å®Œ_åŠ å·¥çµ‚äº†",
        RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16,
        "ç·�åŠ å·¥é‡�",
        "æ®‹åŠ å·¥é‡�",
        "å®Œäº†çŽ‡(å®Ÿè¡Œæ™‚ç‚¹)",
        "ç‰¹åˆ¥æŒ‡å®š_AI",
    ]


def _task_date_key_for_result_sheet_sort(val):
    """çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®ä¸¦ã�¹æ›¿ã�ˆç”¨ã€‚æ¬ æ��ãƒ»è§£é‡ˆä¸�èƒ½ã�¯æœ€å¾Œï¼ˆdate.maxï¼‰ã€‚"""
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
    """å›žç­”ç´�æœŸãƒ»æŒ‡å®šç´�æœŸã�ªã�©ã‚’ date ã�«æ­£è¦�åŒ–ï¼ˆæ¬ æ��ã�¯ Noneï¼‰ã€‚"""
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
    çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ç”¨: ã€Œé…�å®Œ_åŠ å·¥çµ‚äº†ã€�ç›¸å½“ã�®æœ€çµ‚çµ‚äº†ã�Œã€�
    å›žç­”ç´�æœŸã�®æ—¥ä»˜ + PLAN_DUE_DAY_COMPLETION_TIMEï¼ˆæ—¢å®š 16:00ï¼‰ä»¥ä¸‹ã�‹ã‚’åˆ¤å®šã€‚
    å›žç­”ç´�æœŸã�Œç„¡ã�„è¡Œã�¯æŒ‡å®šç´�æœŸã�®æ—¥ä»˜ + 16:00 ã�§åˆ¤å®šã€‚
    ä¸¡æ–¹ç„¡ã�„å ´å�ˆã�¯ã€Œç´�æœŸã�ªã�—ã€�ã€‚
    """
    if not plan_window or len(plan_window) < 2:
        return "æœªå‰²å½“"
    _pe = plan_window[1]
    if _pe is None:
        return "æœªå‰²å½“"
    dd = _coerce_planning_date_for_deadline(answer_due)
    if dd is None:
        dd = _coerce_planning_date_for_deadline(specified_due)
    if dd is None:
        return "ç´�æœŸã�ªã�—"
    try:
        deadline_dt = datetime.combine(dd, PLAN_DUE_DAY_COMPLETION_TIME)
        if _pe <= deadline_dt:
            return "ã�¯ã�„"
        return "ã�„ã�„ã�ˆ"
    except Exception:
        return "åˆ¤å®šä¸�èƒ½"


def _result_task_sheet_sort_key(t: dict):
    """
    çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®è¡¨ç¤ºé †ã€‚â‘ é…�å�°è©¦è¡Œé †ç•ªï¼ˆgenerate_plan å†’é ­ã�§ã‚­ãƒ¥ãƒ¼é †ã�«ä»˜ä¸Žã�—ã�Ÿ 1..nï¼‰æ˜‡é †ã€‚
    æ¬ æ��ãƒ»é�žæ•°ã�¯æœ€å¾Œã€‚å�Œä¸€è©¦è¡Œé †å†…ã�¯ä¾�é ¼NOãƒ»æ©Ÿæ¢°å��ã€�ç¶šã�‘ã�¦åŠ å·¥é–‹å§‹æ—¥ãƒ»ç´�æœŸã�§å®‰å®šåŒ–ã€‚
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
    """åˆ—è¨­å®šã‚·ãƒ¼ãƒˆã�§ã€Œå±¥æ­´ã€�1è¡Œã‚’ç½®ã��ã�¨å±¥æ­´1ï½žn ã‚’ã��ã�®ä½�ç½®ã�«å±•é–‹ã�™ã‚‹ã€‚"""
    if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
        return False
    s = unicodedata.normalize("NFKC", str(cell_val).strip())
    return s in ("å±¥æ­´", "å±¥æ­´*")


def _result_task_column_alias_map(df_columns) -> dict:
    """è¦‹å‡ºã�—ã�® NFKC æ­£è¦�åŒ–ã‚­ãƒ¼ â†’ DataFrame ä¸Šã�®å®Ÿåˆ—å��ã€‚"""
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
    # æ—§åˆ—å��ï¼ˆè¨ˆç”»åŸºæº–ç´�æœŸãƒ™ãƒ¼ã‚¹ï¼‰â†’ é…�å®Œ_å›žç­”æŒ‡å®š16æ™‚ã�¾ã�§
    if nk == _nfkc_column_aliases("é…�å®Œ_åŸºæº–16æ™‚ã�¾ã�§"):
        return col_by_norm.get(
            _nfkc_column_aliases(RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16)
        )
    return None


def _parse_column_visible_cell(val) -> bool:
    """è¡¨ç¤ºåˆ—: ç©ºãƒ»æœªè¨˜å…¥ã�¯ Trueï¼ˆè¡¨ç¤ºï¼‰ã€‚FALSE/0/ã�„ã�„ã�ˆ ç­‰ã�§é�žè¡¨ç¤ºã€‚"""
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
    if s in ("", "true", "1", "ã�¯ã�„", "yes", "on", "è¡¨ç¤º", "â—‹"):
        return True
    if s in ("false", "flase", "0", "ã�„ã�„ã�ˆ", "no", "off", "é�žè¡¨ç¤º", "éš ã�™", "Ã—"):
        return False
    return True


def parse_result_task_column_config_dataframe(
    df_cfg: pd.DataFrame | None, max_history_len: int
) -> list | None:
    """
    ã€Œåˆ—è¨­å®š_çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã€�ç›¸å½“ã�® DataFrame ã�‹ã‚‰ (åˆ—ãƒ©ãƒ™ãƒ«, è¡¨ç¤º) ã‚’ä¸Šã�‹ã‚‰èª­ã‚€ã€‚
    è¦‹å‡ºã�—ã€Œåˆ—å��ã€�ã�¨ã€Œè¡¨ç¤ºã€�ï¼ˆç„¡ã�„å ´å�ˆã�¯è¡¨ç¤ºã�¯ã�™ã�¹ã�¦ Trueï¼‰ã€‚
    ã€Œå±¥æ­´ã€�ã€Œå±¥æ­´*ã€�ã�®1è¡Œã�¯å±¥æ­´1ï½žå±¥æ­´n ã�«å±•é–‹ã�—ã€�å�Œä¸€è¡Œã�®è¡¨ç¤ºãƒ•ãƒ©ã‚°ã‚’å…±æœ‰ã�™ã‚‹ã€‚
    å�Œä¸€åˆ—å��ï¼ˆNFKCãƒ»åˆ¥å��æ­£è¦�åŒ–å¾Œï¼‰ã�Œè¤‡æ•°è¡Œã�‚ã‚‹å ´å�ˆã�¯å…ˆé ­è¡Œã�®ã�¿æŽ¡ç”¨ã�—ã€�ä»¥é™�ã�¯ãƒ­ã‚°ã�«å‡ºã�—ã�¦æ�¨ã�¦ã‚‹ã€‚
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
                "åˆ—è¨­å®šã€Œ%sã€�: é‡�è¤‡åˆ—å��ã€Œ%sã€�ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿï¼ˆä¸Šã�®è¡Œã‚’å„ªå…ˆï¼‰ã€‚",
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
                _try_add(f"å±¥æ­´{j+1}", vis)
            continue
        if raw is None or (isinstance(raw, float) and pd.isna(raw)):
            continue
        s = unicodedata.normalize("NFKC", str(raw).strip())
        if not s or s.lower() in ("nan", "none"):
            continue
        _try_add(s, vis)
    return out or None


def _xlwings_write_column_config_sheet_ab(xw_sheet, rows: list[tuple[str, bool]]) -> None:
    """åˆ—è¨­å®šã‚·ãƒ¼ãƒˆã�® A:B ã‚’ åˆ—å��ãƒ»è¡¨ç¤º ã�®ã�¿ã�§ä¸Šæ›¸ã��ï¼ˆ1è¡Œç›®è¦‹å‡ºã�—ï¼‹ãƒ‡ãƒ¼ã‚¿ï¼‰ã€‚"""
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
    TASK_INPUT_WORKBOOK ã�®ã€Œåˆ—è¨­å®š_çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã€�ã‚·ãƒ¼ãƒˆã�‹ã‚‰ (åˆ—ãƒ©ãƒ™ãƒ«, è¡¨ç¤º) ã‚’ä¸Šã�‹ã‚‰èª­ã‚€ã€‚
    """
    wb = TASKS_INPUT_WORKBOOK
    if not wb or not os.path.exists(wb):
        return None
    if _workbook_should_skip_openpyxl_io(wb):
        logging.info(
            "åˆ—è¨­å®š: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� pandas(openpyxl) ã�§ã�®ã€Œ%sã€�èª­è¾¼ã‚’ã‚¹ã‚­ãƒƒãƒ—ï¼ˆæ—¢å®šåˆ—é †ã‚’ä½¿ã�„ã�¾ã�™ï¼‰ã€‚",
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
            "ã‚·ãƒ¼ãƒˆã€Œ%sã€�: èª­ã�¿è¾¼ã�¿ã�«å¤±æ•—ã�—ã�Ÿã�Ÿã‚�æ—¢å®šã�®åˆ—é †ã‚’ä½¿ã�„ã�¾ã�™ (%s)",
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
    åˆ—è¨­å®šã‚·ãƒ¼ãƒˆã�Œã�‚ã‚Œã�°ã��ã�®é †ãƒ»è¡¨ç¤ºã‚’å„ªå…ˆã�—ã€�ç„¡ã�„åˆ—ã�¯æ—¢å®šé †ã�§å¾Œã‚�ã�«è¿½è¨˜ï¼ˆè¡¨ç¤ºã�¯ Trueï¼‰ã€‚
    config_dataframe ã‚’æ¸¡ã�—ã�Ÿå ´å�ˆã�¯ãƒ•ã‚¡ã‚¤ãƒ«ã‚’èª­ã�¾ã�šã��ã�®å†…å®¹ã‚’åˆ—è¨­å®šã�¨ã�¿ã�ªã�™ï¼ˆxlwings å®Ÿè¡Œæ™‚ç”¨ï¼‰ã€‚
    æˆ»ã‚Šå€¤: (ä¸¦ã�¹æ›¿ã�ˆå¾Œ DataFrame, å®Ÿéš›ã�®åˆ—å��ãƒªã‚¹ãƒˆ, è¨­å®šã‚½ãƒ¼ã‚¹èª¬æ˜Žæ–‡å­—åˆ—, åˆ—å��â†’è¡¨ç¤ºbool)
    """
    default_order = default_result_task_sheet_column_order(max_history_len)
    if config_dataframe is not None:
        user_rows = parse_result_task_column_config_dataframe(config_dataframe, max_history_len)
    else:
        user_rows = load_result_task_column_rows_from_input_workbook(max_history_len)
    if user_rows:
        primary = user_rows
        source = (
            f"ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œ{COLUMN_CONFIG_SHEET_NAME}ã€�"
            if config_dataframe is None
            else f"ã‚·ãƒ¼ãƒˆã€Œ{COLUMN_CONFIG_SHEET_NAME}ã€�ï¼ˆå®Ÿè¡Œä¸­ãƒ–ãƒƒã‚¯ï¼‰"
        )
    else:
        primary = [(n, True) for n in default_order]
        source = "æ—¢å®š"

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
            "åˆ—è¨­å®š: çµ�æžœã�«ç„¡ã�„åˆ—å��ã‚’ç„¡è¦–ã�—ã�¾ã�—ã�Ÿï¼ˆæœ€å¤§20ä»¶ï¼‰: %s",
            ", ".join(unknown[:20]) + (" â€¦" if len(unknown) > 20 else ""),
        )
    logging.info("çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®åˆ—é †ã‚½ãƒ¼ã‚¹: %sï¼ˆ%s åˆ—ï¼‰", source, len(ordered))
    if not user_rows and config_dataframe is None:
        logging.info(
            "åˆ—é †ãƒ»è¡¨ç¤ºã�®ã‚«ã‚¹ã‚¿ãƒžã‚¤ã‚º: ãƒžã‚¯ãƒ­å®Ÿè¡Œãƒ–ãƒƒã‚¯ã�«ã‚·ãƒ¼ãƒˆã€Œ%sã€�ã‚’è¿½åŠ ã€‚"
            " è¦‹å‡ºã�—ã€Œ%sã€�ã€Œ%sã€�â€¦ è¡¨ç¤ºã�Œ FALSE ã�®åˆ—ã�¯çµ�æžœã‚·ãƒ¼ãƒˆã�§é�žè¡¨ç¤ºã€‚"
            " 1è¡Œã€Œå±¥æ­´ã€�ã�§å±¥æ­´1ï½žn ã‚’æŒ¿å…¥ã€‚VBA ã�®ã€Œåˆ—è¨­å®š_çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§_ãƒ�ã‚§ãƒƒã‚¯ãƒœãƒƒã‚¯ã‚¹ã‚’é…�ç½®ã€�ã�§ãƒ�ã‚§ãƒƒã‚¯ãƒœãƒƒã‚¯ã‚¹ã‚’è¡¨ç¤ºåˆ—ã�«é€£å‹•å�¯èƒ½ã€‚",
            COLUMN_CONFIG_SHEET_NAME,
            COLUMN_CONFIG_HEADER_COL,
            COLUMN_CONFIG_VISIBLE_COL,
        )
    return df[ordered], ordered, source, vis_map


def _xlwings_sheet_to_matrix(sheet) -> list:
    """xlwings Sheet ã�® UsedRange ã‚’çŸ©å½¢ã�® list[list] ã�«ã�™ã‚‹ï¼ˆ1è¡Œã�®ã�¿ã�§ã‚‚2æ¬¡å…ƒï¼‰ã€‚"""
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
    """1è¡Œç›®ã‚’åˆ—å��ã�¨ã�¿ã�ªã�— DataFrame ã‚’è¿”ã�™ã€‚ç©ºã�ªã‚‰ Noneã€‚"""
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
    """çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®ã€Œå±¥æ­´nã€�åˆ—ã�‹ã‚‰ n ã�®æœ€å¤§ã‚’è¿”ã�™ï¼ˆç„¡ã�‘ã‚Œã�° 1ï¼‰ã€‚"""
    imax = 0
    for c in columns:
        m = re.match(r"^å±¥æ­´(\d+)$", str(c).strip())
        if m:
            imax = max(imax, int(m.group(1)))
    return max(imax, 1)


def apply_result_task_column_layout_via_xlwings(workbook_path: str | None = None) -> bool:
    """
    Excel ã�§é–‹ã�„ã�¦ã�„ã‚‹ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã�«ã�¤ã�„ã�¦ã€�
    ã€Œåˆ—è¨­å®š_çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã€�ã�®å†…å®¹ã�«å�ˆã‚�ã�›ã�¦ã€Œçµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã€�ã�®åˆ—é †ã�¨åˆ—é�žè¡¨ç¤ºã‚’æ›´æ–°ã�™ã‚‹ã€‚
    ãƒ–ãƒƒã‚¯ã�¯äº‹å‰�ã�«ä¿�å­˜ã�—ã€�æœ¬å‡¦ç�†ä¸­ã‚‚ Excel ä¸Šã�§é–‹ã�„ã�Ÿã�¾ã�¾ã�«ã�™ã‚‹ã�“ã�¨ï¼ˆxlwings ã�ŒæŽ¥ç¶šã�™ã‚‹ï¼‰ã€‚
    """
    path = (workbook_path or "").strip() or TASKS_INPUT_WORKBOOK.strip()
    if not path:
        logging.error("çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ åˆ—é�©ç”¨: ãƒ–ãƒƒã‚¯ãƒ‘ã‚¹ã�Œç©ºã�§ã�™ï¼ˆTASK_INPUT_WORKBOOK ã‚’è¨­å®šã�—ã�¦ã��ã� ã�•ã�„ï¼‰ã€‚")
        return False
    try:
        import xlwings as xw
    except ImportError:
        logging.error("çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ åˆ—é�©ç”¨: xlwings ã�Œ import ã�§ã��ã�¾ã�›ã‚“ã€‚pip install xlwings ã‚’ç¢ºèª�ã�—ã�¦ã��ã� ã�•ã�„ã€‚")
        return False

    try:
        wb = xw.Book(path)
    except Exception as e:
        logging.error("çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ åˆ—é�©ç”¨: ãƒ–ãƒƒã‚¯ã�«æŽ¥ç¶šã�§ã��ã�¾ã�›ã‚“: %s", e)
        return False

    try:
        ws_res = wb.sheets[RESULT_TASK_SHEET_NAME]
        ws_cfg = wb.sheets[COLUMN_CONFIG_SHEET_NAME]
    except Exception as e:
        logging.error("çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ åˆ—é�©ç”¨: å¿…è¦�ã‚·ãƒ¼ãƒˆã�Œè¦‹ã�¤ã�‹ã‚Šã�¾ã�›ã‚“: %s", e)
        return False

    mat_res = _xlwings_sheet_to_matrix(ws_res)
    mat_cfg = _xlwings_sheet_to_matrix(ws_cfg)
    df_res = _matrix_to_dataframe_header_first(mat_res)
    df_cfg = _matrix_to_dataframe_header_first(mat_cfg)
    if df_res is None or df_res.empty:
        logging.error("çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ åˆ—é�©ç”¨: ã€Œ%sã€�ã�«ãƒ‡ãƒ¼ã‚¿ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚", RESULT_TASK_SHEET_NAME)
        return False
    if df_cfg is None:
        logging.error("çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ åˆ—é�©ç”¨: ã€Œ%sã€�ã�®è¦‹å‡ºã�—ã‚’èª­ã‚�ã�¾ã�›ã‚“ã€‚", COLUMN_CONFIG_SHEET_NAME)
        return False

    max_h = _max_history_len_from_result_task_df_columns(df_res.columns)
    rows_cfg = parse_result_task_column_config_dataframe(df_cfg, max_h)
    if not rows_cfg:
        logging.error(
            "çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ åˆ—é�©ç”¨: ã€Œ%sã€�ã�«æœ‰åŠ¹ã�ªåˆ—å��è¡Œã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚",
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
                logging.warning("åˆ—é�žè¡¨ç¤ºã�«å¤±æ•—ï¼ˆåˆ—%s %sï¼‰: %s", ci, col_name, e)

    try:
        wb.save()
    except Exception as e:
        logging.warning("çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ åˆ—é�©ç”¨: ä¿�å­˜ã�§è­¦å‘Šï¼ˆãƒ‡ãƒ¼ã‚¿ã�¯ã‚·ãƒ¼ãƒˆä¸Šã�¯æ›´æ–°æ¸ˆã�¿ã�®å�¯èƒ½æ€§ï¼‰: %s", e)

    logging.info(
        "çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ åˆ—é�©ç”¨å®Œäº†: %sï¼ˆ%s åˆ—ã€�é�žè¡¨ç¤º=%sï¼‰",
        source,
        len(ordered),
        sum(1 for c in ordered if not vis_map.get(c, True)),
    )
    return True


def apply_result_task_column_layout_only() -> bool:
    """ç’°å¢ƒå¤‰æ•° TASK_INPUT_WORKBOOK ã�®ãƒ–ãƒƒã‚¯ã�«å¯¾ã�—åˆ—è¨­å®šã‚’é�©ç”¨ã�™ã‚‹ï¼ˆVBA ãƒœã‚¿ãƒ³ç”¨ï¼‰ã€‚"""
    p = os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK
    return apply_result_task_column_layout_via_xlwings(p)

_PLAN_INPUT_XLWINGS_ORIG_ROW = "__orig_sheet_row__"


def _plan_input_dispatch_trial_order_local_only_from_env() -> bool:
    """ç’°å¢ƒå¤‰æ•° PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY ã�ŒçœŸã�ªã‚‰ post_load ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�™ã‚‹ã€‚"""
    v = (os.environ.get("PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY") or "").strip().lower()
    return v in ("1", "true", "yes", "on", "y")


def refresh_plan_input_dispatch_trial_order_via_xlwings(
    workbook_path: str | None = None,
    *,
    apply_post_load_mutations: bool = True,
) -> bool:
    """
    Excel ã�§é–‹ã�„ã�Ÿãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯å†…ã�®ã€Œé…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€�ã�«ã�¤ã�„ã�¦ã€�
    æ®µéšŽ2 ã�¨å�Œã�˜ ``fill_plan_dispatch_trial_order_column_stage1`` ã�§ã€Œé…�å�°è©¦è¡Œé †ç•ªã€�ã‚’
    å†�ä»˜ä¸Žã�—ã€�æ®µéšŽ1 å‡ºåŠ›ç›´å‰�ã�¨å�Œã�˜æ‰‹é †ã�§è¡Œã‚’ä¸¦ã�¹æ›¿ã�ˆã‚‹ã€‚
    ï¼ˆæœªä¿�å­˜ã�®ç·¨é›†åˆ†ã‚‚ xlwings ã�§å��æ˜ ã�•ã�›ã‚‹ã�Ÿã‚� read_excel ã�¯ä½¿ã‚�ã�ªã�„ï¼‰

    apply_post_load_mutations=False ã�®ã�¨ã��ã�¯ ``_apply_planning_sheet_post_load_mutations`` ã‚’å‘¼ã�°ã�ªã�„ã€‚
    åŠ å·¥è¨ˆç”»DATA ã�‹ã‚‰ã�®å†�å�–ã‚Šè¾¼ã�¿ã�¯è¡Œã‚�ã�ªã�„ã�Œã€�True ã�®ã�¨ã��ã�¯æ®µéšŽ2 èª­è¾¼ã�¨å�Œæ§˜ã�«
    è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã�®ä¿�å®ˆãƒ»åˆ†å‰²è¡Œã�®è‡ªå‹•é…�å�°ä¸�è¦�ãƒ»è¨­å®šãƒ«ãƒ¼ãƒ«ã�«ã‚ˆã‚‹é…�å�°ä¸�è¦�ã�®å†�é�©ç”¨ã�Œèµ°ã‚‹ã€‚
    æ‰‹å‹•ã�§ã€Œé…�å�°ä¸�è¦�ã€�ã‚’å¤–ã�—ã�Ÿã‚Šã€ŒåŽŸå��æŠ•å…¥æ—¥_ä¸Šæ›¸ã��ã€�ç­‰ã�®ã�¿å¤‰æ›´ã�—ã�Ÿå†…å®¹ã‚’ã��ã�®ã�¾ã�¾è©¦è¡Œé †ã�«å��æ˜ ã�—ã�Ÿã�„å ´å�ˆã�¯ Falseã€‚
    """
    path = (workbook_path or "").strip() or os.environ.get(
        "TASK_INPUT_WORKBOOK", ""
    ).strip() or TASKS_INPUT_WORKBOOK.strip()
    if not path:
        logging.error("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: ãƒ–ãƒƒã‚¯ãƒ‘ã‚¹ã�Œç©ºã�§ã�™ã€‚")
        return False
    try:
        import xlwings as xw
    except ImportError:
        logging.error("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: xlwings ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚")
        return False
    try:
        wb = xw.Book(path)
        ws = wb.sheets[PLAN_INPUT_SHEET_NAME]
    except Exception as e:
        logging.error("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: ã‚·ãƒ¼ãƒˆæŽ¥ç¶šã�«å¤±æ•—: %s", e)
        return False

    mat = _xlwings_sheet_to_matrix(ws)
    df = _matrix_to_dataframe_header_first(mat)
    if df is None or df.empty:
        logging.warning("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: ãƒ‡ãƒ¼ã‚¿è¡Œã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚")
        return False

    df = df.copy()
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(df, plan_input_sheet_column_order())
    for c in plan_input_sheet_column_order():
        if c not in df.columns:
            df[c] = ""

    df.insert(0, _PLAN_INPUT_XLWINGS_ORIG_ROW, range(len(df)))

    if apply_post_load_mutations:
        _apply_planning_sheet_post_load_mutations(df, path, "é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°")
    else:
        logging.info(
            "é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: ã‚·ãƒ¼ãƒˆå†…å®¹ã�®ã�¿ãƒ¢ãƒ¼ãƒ‰ï¼ˆè¨­å®šã‚·ãƒ¼ãƒˆå†�é�©ç”¨ãƒ»åˆ†å‰²è¡Œã�®è‡ªå‹•é…�å�°ä¸�è¦�ã‚’ã‚¹ã‚­ãƒƒãƒ—ï¼‰"
        )

    dto_col = RESULT_TASK_COL_DISPATCH_TRIAL_ORDER
    if dto_col not in df.columns:
        logging.error("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: åˆ—ã€Œ%sã€�ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚", dto_col)
        return False

    _dto_loc = df.columns.get_loc(dto_col)
    if isinstance(_dto_loc, slice):
        logging.error("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: åˆ—ã€Œ%sã€�ã�Œè¤‡æ•°ã�‚ã‚Šã�¾ã�™ã€‚", dto_col)
        return False
    # Excel ç”±æ�¥ã�§åˆ—ã�Œ float64 ã�®ã�¨ã�� "" ã‚’å…¥ã‚Œã‚‹ã�¨ pandas ã�Œæ‹’å�¦ã�™ã‚‹ã�Ÿã‚�ã€�ã‚¯ãƒªã‚¢ã�¯ NaN
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
        logging.exception("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: master èª­è¾¼ã�«å¤±æ•—: %s", e)
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
        logging.exception("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: è©¦è¡Œé †è¨ˆç®—ã�«å¤±æ•—: %s", e)
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
        logging.exception("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: ã‚·ãƒ¼ãƒˆæ›¸è¾¼ã�«å¤±æ•—: %s", e)
        return False

    try:
        wb.save()
    except Exception as e:
        logging.warning("é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: Save è­¦å‘Š: %s", e)

    logging.info(
        "é…�å�°è©¦è¡Œé †ç•ªæ›´æ–°: ã€Œ%sã€�ã‚’ %s è¡Œã�§æ›´æ–°ã�—ã�¾ã�—ã�Ÿã€‚",
        PLAN_INPUT_SHEET_NAME,
        len(df_sorted),
    )
    return True


def refresh_plan_input_dispatch_trial_order_only() -> bool:
    """TASK_INPUT_WORKBOOK ã�«å¯¾ã�™ã‚‹é…�å�°è©¦è¡Œé †ç•ªå†�è¨ˆç®—ï¼ˆVBA / cmd çµŒç”±ã�®ã‚¨ãƒ³ãƒˆãƒªï¼‰ã€‚
    ç’°å¢ƒå¤‰æ•° PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY=1 ç­‰ã�§ã‚·ãƒ¼ãƒˆä¸Šã�®ã‚»ãƒ«å€¤ã�®ã�¿ã‚’å…¥åŠ›ã�¨ã�™ã‚‹ã€‚
    """
    p = os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK
    local = _plan_input_dispatch_trial_order_local_only_from_env()
    return refresh_plan_input_dispatch_trial_order_via_xlwings(
        p, apply_post_load_mutations=not local
    )


def apply_plan_input_column_layout_only() -> bool:
    """
    é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã�®åˆ—é †ãƒ»è¡¨ç¤ºã�®ã�¿ã‚’é�©ç”¨ã�™ã‚‹äºˆå®šï¼ˆVBA ç”¨ï¼‰ã€‚
    æœªå®Ÿè£…ã€‚åˆ—ã�®ä¸¦ã�³ã�¯æ®µéšŽ1å‡ºåŠ›ã�¾ã�Ÿã�¯æ‰‹å‹•æ•´ç�†ã‚’ä½¿ç”¨ã�—ã�¦ã��ã� ã�•ã�„ã€‚
    """
    logging.warning("apply_plan_input_column_layout_only: not implemented")
    return False



def dedupe_result_task_column_config_sheet_via_xlwings(workbook_path: str | None = None) -> bool:
    """
    ã€Œåˆ—è¨­å®š_çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã€�ã�® A:B ã� ã�‘ã‚’ã€�é‡�è¤‡åˆ—å��ã‚’é™¤ã�„ã�Ÿä¸€è¦§ã�§æ›¸ã��ç›´ã�™ï¼ˆå…ˆã�®è¡Œã‚’å„ªå…ˆï¼‰ã€‚
    ã€Œçµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã€�ã�Œã�‚ã‚Œã�°å±¥æ­´åˆ—æ•°ã�®è§£é‡ˆã�«ä½¿ã�†ã€‚çµ�æžœã‚·ãƒ¼ãƒˆã�¯å¤‰æ›´ã�—ã�ªã�„ã€‚
    """
    path = (workbook_path or "").strip() or TASKS_INPUT_WORKBOOK.strip()
    if not path:
        logging.error("åˆ—è¨­å®š é‡�è¤‡æ•´ç�†: ãƒ–ãƒƒã‚¯ãƒ‘ã‚¹ã�Œç©ºã�§ã�™ã€‚")
        return False
    try:
        import xlwings as xw
    except ImportError:
        logging.error("åˆ—è¨­å®š é‡�è¤‡æ•´ç�†: xlwings ã�Œ import ã�§ã��ã�¾ã�›ã‚“ã€‚")
        return False
    try:
        wb = xw.Book(path)
        ws_cfg = wb.sheets[COLUMN_CONFIG_SHEET_NAME]
    except Exception as e:
        logging.error("åˆ—è¨­å®š é‡�è¤‡æ•´ç�†: æŽ¥ç¶šã�¾ã�Ÿã�¯ã‚·ãƒ¼ãƒˆå�–å¾—ã�«å¤±æ•—: %s", e)
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
        logging.error("åˆ—è¨­å®š é‡�è¤‡æ•´ç�†: ã€Œ%sã€�ã�®è¦‹å‡ºã�—ã‚’èª­ã‚�ã�¾ã�›ã‚“ã€‚", COLUMN_CONFIG_SHEET_NAME)
        return False
    rows = parse_result_task_column_config_dataframe(df_cfg, max_h)
    if not rows:
        logging.warning("åˆ—è¨­å®š é‡�è¤‡æ•´ç�†: æœ‰åŠ¹ã�ªãƒ‡ãƒ¼ã‚¿è¡Œã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚")
        return False
    _xlwings_write_column_config_sheet_ab(ws_cfg, rows)
    try:
        wb.save()
    except Exception as e:
        logging.warning("åˆ—è¨­å®š é‡�è¤‡æ•´ç�†: ä¿�å­˜è­¦å‘Š: %s", e)
    logging.info(
        "åˆ—è¨­å®šã€Œ%sã€�ã‚’é‡�è¤‡é™¤åŽ»æ¸ˆã�¿ã�§ %s è¡Œã�«æ•´ç�†ã�—ã�¾ã�—ã�Ÿï¼ˆå±¥æ­´å±•é–‹å¾Œã�®è¡Œæ•°ï¼‰ã€‚",
        COLUMN_CONFIG_SHEET_NAME,
        len(rows),
    )
    return True


def dedupe_result_task_column_config_sheet_only() -> bool:
    """ç’°å¢ƒå¤‰æ•° TASK_INPUT_WORKBOOK ã�®ãƒ–ãƒƒã‚¯ã�®åˆ—è¨­å®šã‚·ãƒ¼ãƒˆã� ã�‘é‡�è¤‡æ•´ç�†ï¼ˆVBA ç”¨ï¼‰ã€‚"""
    p = os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK
    return dedupe_result_task_column_config_sheet_via_xlwings(p)


def _apply_result_task_sheet_column_visibility(worksheet, column_names: list, vis_map: dict):
    """çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�§ã€�vis_map ã�Œ False ã�®åˆ—ã‚’é�žè¡¨ç¤ºã�«ã�™ã‚‹ã€‚"""
    for idx, col_name in enumerate(column_names, 1):
        if not vis_map.get(col_name, True):
            worksheet.column_dimensions[get_column_letter(idx)].hidden = True


def _norm_history_member_label(name: str) -> str:
    """å±¥æ­´ã�®æ‹…å½“å��æ¯”è¼ƒç”¨ï¼ˆå…¨è§’ç©ºç™½ã‚’å�Šè§’1å€‹åŒ–ãƒ»å‰�å¾Œtrimãƒ»é€£ç¶šç©ºç™½ã�®åœ§ç¸®ï¼‰ã€‚"""
    t = str(name or "").replace("\u3000", " ").strip()
    return " ".join(t.split())


def _history_team_text_main_assignment_only(h: dict) -> str:
    """
    çµ�æžœã‚·ãƒ¼ãƒˆã€Œæ‹…å½“ã€�æ¬„ç”¨: ãƒ¡ã‚¤ãƒ³å‰²ä»˜ç¢ºå®šæ™‚ç‚¹ã�®å��å‰�ï¼ˆä½™åŠ›è¿½è¨˜ã‚µãƒ–ã�¯å�«ã‚�ã�ªã�„ï¼‰ã€‚
    append_surplus å¾Œã�® h['team'] ã�‹ã‚‰ post_dispatch_surplus_names ã‚’é™¤å¤–ã�™ã‚‹ã€‚
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
    """çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®å±¥æ­´ã‚»ãƒ«æ–‡å­—åˆ—ï¼ˆçµ„å�ˆã�›è¡¨ã�®æŽ¡ç”¨è¡ŒIDãƒ»ãƒ¡ã‚¤ãƒ³è¿½åŠ äººæ•°ãƒ»ä½™åŠ›è¿½è¨˜ã�®æ˜Žç¤ºã‚’å�«ã‚€ï¼‰ã€‚"""
    um = task.get("unit_m") or 0
    try:
        done_r = int(h["done_m"] / um) if um else 0
    except (TypeError, ValueError, ZeroDivisionError):
        done_r = 0
    dm = h.get("done_m", 0)
    parts_out: list[str] = [f"ãƒ»ã€�{h.get('date', '')}ã€‘ï¼š{done_r}R ({dm}m)"]
    cid = h.get("combo_sheet_row_id")
    if cid is not None:
        try:
            parts_out.append(f"çµ„å�ˆã�›è¡¨#{int(cid)}")
        except (TypeError, ValueError):
            parts_out.append(f"çµ„å�ˆã�›è¡¨#{cid}")
    parts_out.append(f"æ‹…å½“[{_history_team_text_main_assignment_only(h)}]")
    sm = h.get("surplus_member_names") or []
    if sm:
        parts_out.append(f"è¿½åŠ [{','.join(str(x) for x in sm)}]")
    ps = h.get("post_dispatch_surplus_names") or []
    if ps:
        parts_out.append(f"ä½™åŠ›è¿½è¨˜[{','.join(str(x) for x in ps)}]")
    return " ".join(parts_out)


_RESULT_TASK_HISTORY_RICH_HEAD_RE = re.compile(r"^ãƒ»(ã€�[^ã€‘]*ã€‘)(.*)$", re.DOTALL)


def _apply_result_task_history_rich_text(worksheet, column_names: list):
    """
    å±¥æ­´åˆ—: ã€Œãƒ»ã€�æ—¥ä»˜ã€‘ï¼šâ€¦ã€�ã�®æ—¥ä»˜æ‹¬å¼§éƒ¨åˆ†ã‚’é�’è‰²ãƒªãƒƒãƒ�ãƒ†ã‚­ã‚¹ãƒˆã�«ã�™ã‚‹ã€‚
    openpyxl 3.1 æœªæº€ã�§ã�¯ã‚¹ã‚­ãƒƒãƒ—ï¼ˆæ–‡å­—åˆ—ã�®ã€�ã€‘ã�®ã�¿ï¼‰ã€‚
    """
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        from openpyxl.styles.colors import Color
    except ImportError:
        return

    hist_cols = [
        i + 1 for i, c in enumerate(column_names) if str(c).startswith("å±¥æ­´")
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
            if not isinstance(v, str) or not v.startswith("ãƒ»ã€�"):
                continue
            m = _RESULT_TASK_HISTORY_RICH_HEAD_RE.match(v)
            if not m:
                continue
            bracketed, rest = m.group(1), m.group(2)
            cell.value = CellRichText(
                TextBlock(plain_if, "ãƒ»"),
                TextBlock(blue_if, bracketed),
                TextBlock(plain_if, rest),
            )
            cell.alignment = top


def _apply_result_task_date_columns_blue_font(worksheet, column_names: list):
    """
    çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§: å›žç­”ç´�æœŸãƒ»æŒ‡å®šç´�æœŸãƒ»è¨ˆç”»åŸºæº–ç´�æœŸãƒ»åŽŸå��æŠ•å…¥æ—¥ãƒ»åŠ å·¥é–‹å§‹æ—¥ã�®ã‚»ãƒ«ã‚’é�’è‰²ã�«ã�™ã‚‹ã€‚
    ï¼ˆå±¥æ­´åˆ—ã�®ã€�æ—¥ä»˜ã€‘ã�¯ _apply_result_task_history_rich_text å�´ã€‚è‰²ã�¯ 0070C0 ã�§çµ±ä¸€ï¼‰
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
    needã€Œé…�å�°æ™‚è¿½åŠ äººæ•°ã€�ç›¸å½“ã�§åŸºæœ¬å¿…è¦�äººæ•°ã‚’è¶…ã�ˆã�¦æŽ¡ç”¨ã�—ã�Ÿãƒ–ãƒ­ãƒƒã‚¯ã€�ã�¾ã�Ÿã�¯
    ãƒ¡ã‚¤ãƒ³å®Œäº†å¾Œã�®ä½™åŠ›è¿½è¨˜ã�§ã‚µãƒ–ã�Œå¢—ã�ˆã�Ÿãƒ–ãƒ­ãƒƒã‚¯ã�«å¯¾å¿œã�™ã‚‹ã€Œå±¥æ­´nã€�ã‚»ãƒ«ã‚’è–„é»„ã�«å¡—ã‚‹ã€‚
    """
    hist_cols: list[tuple[int, int]] = []
    for col_idx, col_name in enumerate(column_names, 1):
        m = re.match(r"^å±¥æ­´(\d+)$", str(col_name).strip())
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
    åŠ å·¥å†…å®¹ã�«å·¥ç¨‹å��ã�Œå�«ã�¾ã‚Œã�ªã�„è¡Œã�®ã€Œã‚¿ã‚¹ã‚¯IDã€�ã‚»ãƒ«ã‚’èµ¤èƒŒæ™¯ãƒ»ç™½æ–‡å­—ã�«ã�™ã‚‹ï¼ˆå…ƒãƒ‡ãƒ¼ã‚¿ä¸�æ•´å�ˆã�®è¦–èª�ç”¨ï¼‰ã€‚
    """
    task_id_col_idx = None
    for col_idx, col_name in enumerate(column_names, 1):
        if str(col_name) == "ã‚¿ã‚¹ã‚¯ID":
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
    åˆ—ã€Œé…�å®Œ_å›žç­”æŒ‡å®š16æ™‚ã�¾ã�§ã€�ã�Œã€Œã�„ã�„ã�ˆã€�ã�®ã‚»ãƒ«ã‚’èµ¤èƒŒæ™¯ãƒ»ç™½æ–‡å­—ãƒ»å¤ªå­—ã�«ã�™ã‚‹ã€‚
    åˆ—è¨­å®šã�§æ—§å��ã€Œé…�å®Œ_åŸºæº–16æ™‚ã�¾ã�§ã€�ã�®ã�¾ã�¾ã�®è¦‹å‡ºã�—ã�«ã‚‚å¯¾å¿œã€‚
    """
    target_names = frozenset(
        {
            RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16,
            "é…�å®Œ_åŸºæº–16æ™‚ã�¾ã�§",
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
        if s != "ã�„ã�„ã�ˆ":
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
    çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®ã€Œã‚¿ã‚¹ã‚¯IDã€�ã‚»ãƒ«ã�«ã€�çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ã�§å½“è©²ã‚¿ã‚¹ã‚¯ã�Œæœ€åˆ�ã�«ç�¾ã‚Œã‚‹ã‚»ãƒ«ã�¸ã�®å†…éƒ¨ãƒ�ã‚¤ãƒ‘ãƒ¼ãƒªãƒ³ã‚¯ã‚’ä»˜ä¸Žã�™ã‚‹ã€‚
    æ™‚é–“å‰²ã�«ç�¾ã‚Œã�ªã�„ã‚¿ã‚¹ã‚¯ï¼ˆæœªå‰²å½“ã�®ã�¿ç­‰ï¼‰ã�¯ãƒªãƒ³ã‚¯ã�ªã�—ã€‚
    """
    if not task_id_to_schedule_cell or worksheet_tasks.max_row < 2:
        return
    task_id_col_idx = None
    for col_idx, col_name in enumerate(column_names, 1):
        if str(col_name) == "ã‚¿ã‚¹ã‚¯ID":
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
    """è¡¨ç¤ºåˆ—ã�« TRUE/FALSE ãƒªã‚¹ãƒˆï¼ˆãƒ�ã‚§ãƒƒã‚¯ã�®ä»£ã‚�ã‚Šã�«ãƒ—ãƒ«ãƒ€ã‚¦ãƒ³ï¼‰ã‚’ä»˜ä¸Žã€‚"""
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
    pandas/openpyxl ã�§æ–°è¦�ä½œæˆ�ã�—ã�Ÿçµ�æžœãƒ–ãƒƒã‚¯ã�«ã�¯å›³å½¢ã�Œå�«ã�¾ã‚Œã�ªã�„ã€‚
    æ—¢å®šã�§æœ‰åŠ¹ï¼ˆç’°å¢ƒå¤‰æ•°ã�§ 0/false/no/off ã�®ã�¨ã��ç„¡åŠ¹ï¼‰ã€‚å…¥åŠ›ãƒ–ãƒƒã‚¯ã�®
    ã€Œåˆ—è¨­å®š_çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã€�ä¸Šã�® Shapes ã‚’çµ�æžœãƒ–ãƒƒã‚¯ã�®å�Œå��ã‚·ãƒ¼ãƒˆã�¸ã‚³ãƒ”ãƒ¼ã�—ã€�
    å�„å›³å½¢ã�® Left/Top/Width/Heightï¼ˆã�Šã‚ˆã�³å�–ã‚Œã‚‹ã�¨ã�� Placementï¼‰ã‚’å…¥åŠ›å�´ã�¨å�Œã�˜ã�«æˆ»ã�™ã€‚
    openpyxl ã�«ã‚ˆã‚‹å½“è©²ãƒ–ãƒƒã‚¯ã�¸ã�®ä¿�å­˜ã�Œã�™ã�¹ã�¦çµ‚ã‚�ã�£ã�Ÿå¾Œã�«å‘¼ã�¶ã�“ã�¨ã€‚
    """
    if not STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT:
        return
    rp = (result_path or "").strip()
    ip = (input_path or "").strip()
    if not rp or not os.path.isfile(rp):
        logging.warning(
            "åˆ—è¨­å®šã‚·ãƒ¼ãƒˆå›³å½¢ã‚³ãƒ”ãƒ¼: çµ�æžœãƒ‘ã‚¹ã�Œç„¡åŠ¹ã�®ã�Ÿã‚�ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚"
        )
        return
    if not ip or not os.path.isfile(ip):
        logging.warning(
            "åˆ—è¨­å®šã‚·ãƒ¼ãƒˆå›³å½¢ã‚³ãƒ”ãƒ¼: TASK_INPUT_WORKBOOK ã�Œç„¡åŠ¹ã�®ã�Ÿã‚�ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚"
        )
        return
    try:
        import xlwings as xw
    except ImportError:
        logging.warning(
            "åˆ—è¨­å®šã‚·ãƒ¼ãƒˆå›³å½¢ã‚³ãƒ”ãƒ¼: xlwings ã�Œ import ã�§ã��ã�¾ã�›ã‚“ã€‚"
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
                "åˆ—è¨­å®šã‚·ãƒ¼ãƒˆå›³å½¢ã‚³ãƒ”ãƒ¼: çµ�æžœãƒ–ãƒƒã‚¯ã�«ã‚·ãƒ¼ãƒˆã€Œ%sã€�ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚",
                COLUMN_CONFIG_SHEET_NAME,
            )
            return
        try:
            ws_in = wb_in.sheets[COLUMN_CONFIG_SHEET_NAME]
        except Exception:
            logging.warning(
                "åˆ—è¨­å®šã‚·ãƒ¼ãƒˆå›³å½¢ã‚³ãƒ”ãƒ¼: å…¥åŠ›ãƒ–ãƒƒã‚¯ã�«ã‚·ãƒ¼ãƒˆã€Œ%sã€�ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚",
                COLUMN_CONFIG_SHEET_NAME,
            )
            return
        n_shapes = int(ws_in.api.Shapes.Count)
        if n_shapes <= 0:
            logging.info(
                "åˆ—è¨­å®šã‚·ãƒ¼ãƒˆå›³å½¢ã‚³ãƒ”ãƒ¼: å…¥åŠ›å�´ã�«å›³å½¢ã�Œã�‚ã‚Šã�¾ã�›ã‚“ï¼ˆã‚¹ã‚­ãƒƒãƒ—ï¼‰ã€‚"
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
            "åˆ—è¨­å®šã‚·ãƒ¼ãƒˆå›³å½¢ã‚³ãƒ”ãƒ¼: å…¥åŠ›ã�‹ã‚‰ %s å€‹ã�®å›³å½¢ã‚’çµ�æžœãƒ–ãƒƒã‚¯ã�¸è¤‡è£½ã�—ã�¾ã�—ã�Ÿã€‚",
            n_shapes,
        )
    except Exception as e:
        logging.warning(
            "åˆ—è¨­å®šã‚·ãƒ¼ãƒˆå›³å½¢ã‚³ãƒ”ãƒ¼: å¤±æ•—ã�—ã�¾ã�—ã�Ÿï¼ˆ%sï¼‰ã€‚Excel å� æœ‰ãƒ»COM ã‚¨ãƒ©ãƒ¼ç­‰ã�®å�¯èƒ½æ€§ã�Œã�‚ã‚Šã�¾ã�™ã€‚",
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
    """åŠ å·¥å®Ÿç¸¾DATA ã�®1è¡Œã�‹ã‚‰ (é–‹å§‹, çµ‚äº†) ã‚’å¾—ã‚‹ã€‚è§£ã�‘ã�ªã�‘ã‚Œã�° (None, None)ã€‚"""
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
    ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã�®ã€ŒåŠ å·¥å®Ÿç¸¾DATAã€�ã‚’èª­ã‚€ï¼ˆç„¡ã�‘ã‚Œã�°ç©º DataFrameï¼‰ã€‚
    Power Query ç­‰ã�§ç”¨æ„�ã�—ã�Ÿã‚·ãƒ¼ãƒˆã‚’æƒ³å®šã€‚
    """
    if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
        return pd.DataFrame()
    try:
        df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=ACTUALS_SHEET_NAME)
    except ValueError:
        logging.info(
            f"ã‚·ãƒ¼ãƒˆã€Œ{ACTUALS_SHEET_NAME}ã€�ã�Œç„¡ã�„ã�Ÿã‚�ã€�ã‚¬ãƒ³ãƒˆã�®å®Ÿç¸¾è¡Œã�¯å‡ºåŠ›ã�—ã�¾ã�›ã‚“ã€‚"
        )
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(df, ACTUAL_HEADER_CANONICAL)
    logging.info(
        f"åŠ å·¥å®Ÿç¸¾: '{TASKS_INPUT_WORKBOOK}' ã�® '{ACTUALS_SHEET_NAME}' ã‚’ {len(df)} è¡Œèª­ã�¿è¾¼ã�¿ã€‚"
    )
    return df


def build_actual_timeline_events(df, equipment_list, sorted_dates):
    """
    å®Ÿç¸¾ã‚·ãƒ¼ãƒˆã�®å�„è¡Œã‚’ã‚¬ãƒ³ãƒˆç”¨ã‚¤ãƒ™ãƒ³ãƒˆã�¸å¤‰æ�›ã€‚
    è¨ˆç”»è¡¨ç¤ºæ—¥ï¼ˆsorted_datesï¼‰ã�‹ã�¤è¨­å‚™ãƒžã‚¹ã‚¿ã�«ä¸€è‡´ã�™ã‚‹ã€Œå·¥ç¨‹å��ã€�ã� ã�‘å¯¾è±¡ã€‚
    å·¥ç¨‹å��ã�¯ NFKCãƒ»ç©ºç™½æ­£è¦�åŒ–å¾Œã�«ãƒžã‚¹ã‚¿åˆ—å��ã�¸ãƒžãƒƒãƒ”ãƒ³ã‚°ã�™ã‚‹ã€‚
    æ™‚åˆ»ã�¯ DEFAULT_START_TIME / DEFAULT_END_TIME ã�®æž å†…ã�«ã‚¯ãƒªãƒƒãƒ—ã€‚
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
            f"åŠ å·¥å®Ÿç¸¾DATA: å·¥ç¨‹å��ã�Œãƒžã‚¹ã‚¿è¨­å‚™ã�¨ä¸€è‡´ã�—ã�ªã�„è¡Œã‚’ {bad_eq} ä»¶ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿï¼ˆç©ºç™½ç­‰ã�¯æ­£è¦�åŒ–æ¸ˆã�¿ï¼‰ã€‚"
        )
        if mismatch_norm_samples:
            logging.info(
                "  ä¸�ä¸€è‡´ã�¨ã�ªã�£ã�Ÿå·¥ç¨‹å��ã�®æ­£è¦�åŒ–å¾Œã‚µãƒ³ãƒ—ãƒ«: "
                + " | ".join(mismatch_norm_samples[:12])
            )
    if bad_time:
        logging.info(
            f"åŠ å·¥å®Ÿç¸¾DATA: é–‹å§‹/çµ‚äº†æ—¥æ™‚ã�Œè§£é‡ˆã�§ã��ã�ªã�„è¡Œã‚’ {bad_time} ä»¶ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚"
        )
    if no_plan_overlap and sorted_dates:
        logging.info(
            f"åŠ å·¥å®Ÿç¸¾DATA: è¨­å‚™ãƒ»æ—¥æ™‚ã�¯æœ‰åŠ¹ã� ã�Œã€�è¨ˆç”»å¯¾è±¡æ—¥ï¼ˆå½“æ—¥ä»¥é™�ã�®å‹¤æ€ æ—¥Ã—{DEFAULT_START_TIME}ï½ž{DEFAULT_END_TIME}ï¼‰ã�¨é‡�ã�ªã‚‰ã�ªã�„è¡Œã�Œ {no_plan_overlap} ä»¶ã�‚ã‚Šã�¾ã�—ã�Ÿã€‚"
        )
    if not events and len(df) > 0:
        logging.info(
            "åŠ å·¥å®Ÿç¸¾DATA: ã‚¬ãƒ³ãƒˆç”¨ã‚»ã‚°ãƒ¡ãƒ³ãƒˆã�Œ0ä»¶ã�§ã�™ã€‚é�ŽåŽ»æ—¥ã�®å®Ÿç¸¾ã�®ã�¿ã�®å ´å�ˆã€�è¨ˆç”»ã�®è¡¨ç¤ºæ—¥ï¼ˆsorted_datesï¼‰ã�«å�«ã�¾ã‚Œã�ªã�„ã�Ÿã‚�æ��ç”»ã�•ã‚Œã�¾ã�›ã‚“ã€‚"
        )
    logging.info(f"åŠ å·¥å®Ÿç¸¾DATA ã�‹ã‚‰ã‚¬ãƒ³ãƒˆç”¨ã‚»ã‚°ãƒ¡ãƒ³ãƒˆ {len(events)} ä»¶ã‚’ç”Ÿæˆ�ã�—ã�¾ã�—ã�Ÿã€‚")
    return events


TASK_SPECIAL_AI_LAST_RESPONSE_FILE = "ai_task_special_remark_last.txt"
# å‹¤æ€ å‚™è€ƒã‚­ãƒ£ãƒƒã‚·ãƒ¥ã�¨ã‚­ãƒ¼ç©ºé–“ã‚’åˆ†é›¢ï¼ˆå�Œä¸€SHAè¡�çª�ã‚’é�¿ã�‘ã‚‹ï¼‰ã€‚æŒ‡ç´‹ã�«åŸºæº–å¹´ã‚’å�«ã‚�æ—¥ä»˜è§£é‡ˆã�®ã‚ºãƒ¬ã‚’é˜²ã��ã€‚
TASK_SPECIAL_CACHE_KEY_PREFIX = "TASK_SPECIAL_v3|"
# ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�ä¸‹ã�®è‡ªç”±è¨˜è¿° â†’ Gemini è§£é‡ˆï¼ˆé…�å�°ã�®æœ€å„ªå…ˆã‚ªãƒ¼ãƒ�ãƒ¼ãƒ©ã‚¤ãƒ‰ï¼‰
GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX = "GLOBAL_PRIO_v8|"


def _normalize_special_task_id_for_ai(val):
    """
    ä¾�é ¼NOã‚’ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã‚­ãƒ¼ãƒ»ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆè¡Œã�§ä¸€è²«ã�•ã�›ã‚‹ã€‚
    Excel ã�®æ•°å€¤ã‚»ãƒ«ã�¯ float ã�«ã�ªã‚Šã�Œã�¡ã�ªã�®ã�§ 12345.0 â†’ \"12345\" ã�«æ�ƒã�ˆã‚‹ã€‚
    æ–‡å­—åˆ—ã�¯ NFKCï¼ˆå…¨è§’è‹±æ•°å­—ã�ªã�©ï¼‰ã�§è¡¨è¨˜ã‚†ã‚Œã‚’å�¸å�Žï¼ˆå�Œä¸€å®Ÿä½“ã�®å†�APIå‘¼ã�³å‡ºã�—ã‚’æ¸›ã‚‰ã�™ï¼‰ã€‚
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
    # æ–‡å­—åˆ—ã�¨ã�—ã�¦ã�® "20010.0" ç­‰ï¼ˆExcelãƒ»CSVï¼‰ã‚’æ•´æ•°è¡¨è¨˜ã�®ä¾�é ¼NOã�«å¯„ã�›ã‚‹
    if re.fullmatch(r"-?\d+\.0+", s):
        try:
            return str(int(float(s)))
        except ValueError:
            pass
    return s or None


def planning_task_id_str_from_scalar(val) -> str:
    """é…�å�°ãƒ»æ®µéšŽ1ãƒžãƒ¼ã‚¸ãƒ»ã‚­ãƒ¥ãƒ¼æ§‹ç¯‰ã�§ç”¨ã�„ã‚‹ä¾�é ¼NOã�®å®‰å®šæ–‡å­—åˆ—ï¼ˆç©ºã�ªã‚‰ \"\"ï¼‰ã€‚"""
    return _normalize_special_task_id_for_ai(val) or ""


def planning_task_id_str_from_plan_row(row) -> str:
    """é‡�è¤‡è¦‹å‡ºã�—åˆ—ã�§ã‚‚å…ˆé ­ã‚¹ã‚«ãƒ©ãƒ¼ã‚’æ‹¾ã�„ã€�ä¾�é ¼NOã‚’ planning_task_id_str_from_scalar ã�«æ¸¡ã�™ã€‚"""
    return planning_task_id_str_from_scalar(_planning_df_cell_scalar(row, TASK_COL_TASK_ID))


def _cell_text_task_special_remark(val):
    """
    ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã‚’ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆç”¨ã�«å�–ã‚Šå‡ºã�™ã€‚ä»•æ§˜ã�©ã�Šã‚Š **strip ã�®ã�¿**
    ï¼ˆå…ˆé ­æœ«å°¾ã�®ç©ºç™½ãƒ»Excel ã�®å�½ç©ºç™½ã‚’é™¤ã��ã€�æ–‡ä¸­ã�®æ”¹è¡Œãƒ»ã‚¹ãƒšãƒ¼ã‚¹ã�¯ä¿�æŒ�ã€‚æ•°å€¤ã‚»ãƒ«ã�¯è¡¨è¨˜ã‚’å›ºå®šï¼‰ã€‚
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
        # å‚™è€ƒåˆ—ã�«æ•°å€¤ã� ã�‘å…¥ã�£ã�¦ã�„ã‚‹å ´å�ˆã�®è¡¨è¨˜ã‚†ã‚Œã‚’æ¸›ã‚‰ã�™
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
    """ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆã�«è¼‰ã�›ã‚‹è¡Œãƒªã‚¹ãƒˆï¼ˆã‚½ãƒ¼ãƒˆå‰�ï¼‰ã€‚æ­£è¦�åŒ–ã�¯ä¸Šè¨˜ãƒ˜ãƒ«ãƒ‘ãƒ¼ã�«çµ±ä¸€ã€‚"""
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
        proc_disp = proc if proc else "ï¼ˆç©ºï¼‰"
        macn_disp = macn if macn else "ï¼ˆç©ºï¼‰"
        lines.append(
            f"- ä¾�é ¼NOã€�{tid}ã€‘| å·¥ç¨‹å��ã€Œ{proc_disp}ã€� | æ©Ÿæ¢°å��ã€Œ{macn_disp}ã€� | å‚™è€ƒæœ¬æ–‡: {rem}"
        )
    return lines


def _repair_task_special_ai_wrong_top_level_keys(parsed: dict, tasks_df) -> dict:
    """
    å‚™è€ƒã�Œå“�ç•ªãƒ»åŽŸå��ã‚³ãƒ¼ãƒ‰ï¼ˆä¾‹: 20010 ã�§å§‹ã�¾ã‚‹æ•°å­—åˆ—ï¼‰ã�§å§‹ã�¾ã‚‹ã�¨ã€�ãƒ¢ãƒ‡ãƒ«ã�Œã��ã�®åˆ—ã‚’ JSON ãƒˆãƒƒãƒ—ã‚­ãƒ¼ã�«
    èª¤ç”¨ã�™ã‚‹ã�“ã�¨ã�Œã�‚ã‚‹ã€‚ä¾�é ¼NOã€�â€¦ã€‘ã�¨ä¸€è‡´ã�—ã�ªã�„æ•°å­—ã�®ã�¿ã�®ã‚­ãƒ¼ã‚’ã€�å½“è©²å‚™è€ƒã‚’æŒ�ã�¤è¡Œã�®ä¾�é ¼NOã�¸ä»˜ã�‘æ›¿ã�ˆã‚‹ã€‚
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
                or r.startswith(sk + "ãƒ¼")
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
                "ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: JSON ãƒˆãƒƒãƒ—ã‚­ãƒ¼èª¤ã‚Šã‚’ä¿®å¾©ï¼ˆ%r ã�¯ä¾�é ¼NOã�§ã�¯ã�ªã�„ â†’ %rï¼‰",
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
    restrict ã�Œç„¡ã�„ãƒ»ç©ºã�ªã‚‰åˆ¶é™�ã�ªã�—ï¼ˆTrueï¼‰ã€‚
    é�žç©ºã�ªã‚‰ Excel å�´ã�®å€¤ã�¨ã�‚ã�„ã�¾ã�„ä¸€è‡´ï¼ˆéƒ¨åˆ†ä¸€è‡´å�¯ï¼‰ã€‚
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
    """restrict_to_* ã�Œç„¡ã�„ã�¨ã��ã�¯å�Œä¸€ä¾�é ¼NOã�®å…¨è¡Œã�«é�©ç”¨ã€‚"""
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
    JSON ã�® process_name / machine_name ã�Œã€�å½“è©² Excel è¡Œã�®å·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ã�¨ä¸€è‡´ã�™ã‚‹ã�‹ã€‚
    ï¼ˆãƒ—ãƒ­ãƒ³ãƒ—ãƒˆã�§æ¸¡ã�—ã�Ÿã€Œå‚™è€ƒã�Œã�‚ã�£ã�Ÿè¡Œã€�ã�¨å¯¾å¿œã�¥ã�‘ã‚‹ã€‚ç‰‡æ–¹ã� ã�‘ä¸€è‡´ã�§ã‚‚å�¯ï¼‰
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
    """restrict_to_* ã�Œç„¡ã�„ãƒ»ç©ºï¼�å�Œä¸€ä¾�é ¼NOã�®å…¨å·¥ç¨‹è¡Œã�«åŠ¹ã�‹ã�›ã‚‹æŒ‡å®šã€‚"""
    if not isinstance(entry, dict):
        return False
    a = _normalize_task_special_scope_str(entry.get("restrict_to_process_name"))
    b = _normalize_task_special_scope_str(entry.get("restrict_to_machine_name"))
    return not a and not b


def _select_ai_task_special_entry_for_tid_value(val, row):
    """1ä¾�é ¼NOã�«å¯¾ã�™ã‚‹å€¤ã�Œ dict ã�¾ã�Ÿã�¯ dict ã�®é…�åˆ—ã�®ã�©ã�¡ã‚‰ã�§ã‚‚è¡Œã�«å�ˆã�†è¦�ç´ ã‚’è¿”ã�™ã€‚"""
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
    analyze_task_special_remarks ã�®æˆ»ã‚Šã�‹ã‚‰å½“è©²è¡Œã�®ã‚¨ãƒ³ãƒˆãƒªã‚’å�–ã‚‹ã€‚
    ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆã‚­ãƒ¼ã�¯æ­£è¦�åŒ–æ¸ˆã�¿ä¾�é ¼NOã�ªã�®ã�§ã€�Excel ã�Œ 12345.0 ã�§ã‚‚ãƒ’ãƒƒãƒˆã�™ã‚‹ã€‚
    restrict_to_process_name / restrict_to_machine_name ã�Œç„¡ã�„ãƒ»ç©ºã�®ã�¨ã��ã�¯
    å�Œä¸€ä¾�é ¼NOã�®å·¥ç¨‹ãƒ»æ©Ÿæ¢°ã�Œç•°ã�ªã‚‹å…¨è¡Œã�«å�Œã�˜æŒ‡ç¤ºã‚’é�©ç”¨ã�™ã‚‹ã€‚
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


# 1 å›žã�® Python å®Ÿè¡Œï¼ˆæ®µéšŽ1 ã�¾ã�Ÿã�¯ æ®µéšŽ2ï¼‰å�˜ä½�ã�§ãƒªã‚»ãƒƒãƒˆã�™ã‚‹
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
    """API_Payment å†…ã�®ç´¯è¨ˆ JSON ã‚’èª­ã‚€ã€‚ç„¡ã�„ãƒ»å£Šã‚Œã�¦ã�„ã‚Œã�°åˆ�æœŸå½¢ã‚’è¿”ã�™ã€‚"""
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
        logging.debug("Gemini ç´¯è¨ˆ JSON ã�®ä¿�å­˜ã�«å¤±æ•—: %s", ex)


def _gemini_buckets_ensure_structure(data: dict) -> None:
    """ç´¯è¨ˆ JSON ã�«æœŸé–“åˆ¥ãƒ�ã‚±ãƒƒãƒˆç”¨ã�®è¾žæ›¸ã‚’ç”¨æ„�ã�™ã‚‹ï¼ˆæ—¢å­˜ v1 ãƒ•ã‚¡ã‚¤ãƒ«ã‚‚ãƒžãƒ¼ã‚¸ï¼‰ã€‚"""
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
        "period_key ã�¯ PC ãƒ­ãƒ¼ã‚«ãƒ«æ™‚åˆ»ï¼ˆdatetime.nowï¼‰ã�§ä»˜ä¸Žã€‚ä»– PC ã�¨ã�®é›†è¨ˆã�¯æ··ã�œã�ªã�„ã�§ã��ã� ã�•ã�„ã€‚",
    )


def _gemini_time_bucket_keys(dt: datetime) -> tuple[str, str, str, str, str]:
    """å¹´ãƒ»æœˆãƒ»ISOé€±ãƒ»æ—¥ãƒ»æ™‚ ã�®ã‚­ãƒ¼ï¼ˆæ–‡å­—åˆ—ã‚½ãƒ¼ãƒˆã�§æ™‚ç³»åˆ—æ¯”è¼ƒã�—ã‚„ã�™ã�„å½¢ï¼‰ã€‚"""
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
    """1 å›žã�® API å‘¼å‡ºã�—ã‚’å¹´ãƒ»æœˆãƒ»é€±ãƒ»æ—¥ãƒ»æ™‚ã�®å�„ãƒ�ã‚±ãƒƒãƒˆã�«åŠ ç®—ã�™ã‚‹ã€‚"""
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
    """1 å›žã�® API å¿œç­”åˆ†ã‚’ç´¯è¨ˆ JSON ã�«åŠ ç®—ã�™ã‚‹ï¼ˆãƒ­ã‚°ã�«å�˜ç™ºæ–™é‡‘ã�¯å‡ºã�•ã�ªã�„ï¼‰ã€‚"""
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
    """generate_content ã�®å¿œç­”ã�‹ã‚‰ usage_metadata ã‚’é›†è¨ˆã�™ã‚‹ï¼ˆã‚»ãƒƒã‚·ãƒ§ãƒ³ï¼‹ç´¯è¨ˆ JSONï¼‰ã€‚"""
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
        logging.debug("Gemini ç´¯è¨ˆã�®æ›´æ–°ã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰: %s", ex)


def _gemini_estimate_cost_usd(
    model_id: str, prompt_tok: int, cand_tok: int, thoughts_tok: int
) -> float | None:
    m = str(model_id).strip().lower()
    rin, rout = None, None
    if "flash" in m:
        rin, rout = _GEMINI_FLASH_IN_PER_M, _GEMINI_FLASH_OUT_PER_M
    elif "pro" in m:
        # ç›®å®‰ï¼ˆæœªä½¿ç”¨ãƒ¢ãƒ‡ãƒ«å�‘ã�‘ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ï¼‰
        rin, rout = 1.25, 5.0
    if rin is None:
        return None
    out_equiv = cand_tok + thoughts_tok
    return (prompt_tok / 1_000_000.0) * rin + (out_equiv / 1_000_000.0) * rout


def _gemini_daily_trend_series(
    cum: dict, *, max_days: int | None = None
) -> tuple[list[str], list[float], str] | None:
    """ç´¯è¨ˆ JSON ã�® by_day ã�‹ã‚‰ã€�æ—¥ä»˜ã‚­ãƒ¼ï¼ˆå�¤â†’æ–°ï¼‰ãƒ»å€¤ãƒ»ç³»åˆ—å��ã€‚ç„¡ã�‘ã‚Œã�° Noneã€‚"""
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
    label = "å‘¼å‡ºã�—å›žæ•°" if use_calls else "æŽ¨å®šUSD"
    return (keys, series, label)


def _gemini_daily_total_tokens_for_days(cum: dict, day_keys: list[str]) -> list[int]:
    """by_day ã�®å�„ã‚­ãƒ¼ã�«ã�¤ã�„ã�¦ã€�total_tokensï¼ˆç„¡ã�‘ã‚Œã�° prompt+candidates+thoughtsï¼‰ã‚’è¿”ã�™ã€‚"""
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
    """ãƒ†ã‚­ã‚¹ãƒˆå�´ã�¯ã‚°ãƒ©ãƒ•å�‚ç…§ã�¨ CSV æ¡ˆå†…ã�®ã�¿ï¼ˆASCII ã‚¹ãƒ‘ãƒ¼ã‚¯ãƒ©ã‚¤ãƒ³ã�¯å‡ºã�•ã�ªã�„ï¼‰ã€‚"""
    ser = _gemini_daily_trend_series(cum)
    if ser is None:
        return []
    keys, _, label = ser
    b = cum.get("buckets")
    lines = [
        "ã€�æŽ¨ç§»ã‚°ãƒ©ãƒ•ã€‘æ–™é‡‘ãƒ»å‘¼å‡ºã�—: Qã€œR åˆ—ï¼�ãƒˆãƒ¼ã‚¯ãƒ³é‡�: Sã€œT åˆ—ï¼ˆå�„ã‚°ãƒ©ãƒ•ãƒ»è‡ªå‹•æ›´æ–°ï¼‰ã‚’å�‚ç…§",
        f"  ç³»åˆ—1: æ—¥æ¬¡ {label}ï¼ˆ{keys[0]} ï½ž {keys[-1]}ï¼‰",
        "  ç³»åˆ—2: æ—¥æ¬¡ å�ˆè¨ˆãƒˆãƒ¼ã‚¯ãƒ³ï¼ˆAPI å ±å‘Š total ã�¾ã�Ÿã�¯å†…è¨³å�ˆè¨ˆï¼‰",
        f"  å¹´ãƒ»æœˆãƒ»é€±ãƒ»æ™‚ã�ªã�©ã�®å†…è¨³: log\\{GEMINI_USAGE_BUCKETS_CSV_FILE}ï¼ˆExcel ã�§ã‚°ãƒ©ãƒ•å�¯ï¼‰",
    ]
    if isinstance(b, dict):
        note = b.get("timezone_note")
        if note:
            lines.append(f"  ï¼ˆ{note}ï¼‰")
    return lines


def _gemini_resolve_main_sheet_xlwings(book) -> object | None:
    """xlwings Book ã�‹ã‚‰ãƒ¡ã‚¤ãƒ³ç›¸å½“ã‚·ãƒ¼ãƒˆã‚’è¿”ã�™ã€‚ç„¡ã�‘ã‚Œã�° Noneã€‚"""
    for name in ("ãƒ¡ã‚¤ãƒ³", "ãƒ¡ã‚¤ãƒ³_", "Main"):
        try:
            return book.sheets[name]
        except Exception:
            continue
    try:
        for sht in book.sheets:
            try:
                if "ãƒ¡ã‚¤ãƒ³" in str(sht.name):
                    return sht
            except Exception:
                continue
    except Exception:
        pass
    return None


def _strip_gemini_usage_charts_xlwings(ws) -> None:
    """å½“æ©Ÿèƒ½ã�Œç®¡ç�†ã�™ã‚‹æŠ˜ã‚Œç·šï¼ˆå��å‰�ã�¾ã�Ÿã�¯ã‚°ãƒ©ãƒ•ã‚¿ã‚¤ãƒˆãƒ«ï¼‰ã‚’å‰Šé™¤ã�™ã‚‹ã€‚"""
    managed_names = (
        GEMINI_USAGE_XLW_CHART_NAME,
        GEMINI_USAGE_XLW_CHART_TOKENS_NAME,
    )
    title_markers = (
        "Gemini API æ—¥æ¬¡æŽ¨ç§»",
        "Gemini API æ—¥æ¬¡ãƒˆãƒ¼ã‚¯ãƒ³",
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
    """é–‹ã�„ã�Ÿãƒ–ãƒƒã‚¯ä¸Šã�§ Qã€œRãƒ»Sã€œT ã‚’åŸ‹ã‚�ã€�æŠ˜ã‚Œç·šã‚°ãƒ©ãƒ•ã‚’ 2 æœ¬ã�¾ã�§ç½®ã��ï¼ˆxlwingsï¼‰ã€‚"""
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

    ws.range((hr, cdt)).value = "æ—¥ä»˜"
    ws.range((hr, cvl)).value = val_label
    for i, (dk, val) in enumerate(zip(day_keys, values)):
        r = hr + 1 + i
        ws.range((r, cdt)).value = dk
        ws.range((r, cvl)).value = val
    try:
        vrng = ws.range((hr + 1, cvl), (hr + n, cvl))
        vrng.number_format = "0.000000" if val_label == "æŽ¨å®šUSD" else "0"
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
        ca.ChartTitle.Text = "Gemini API æ—¥æ¬¡æŽ¨ç§»"
        ca.HasLegend = False
    except Exception:
        pass

    tok_vals = _gemini_daily_total_tokens_for_days(cum, day_keys)
    if not tok_vals or max(tok_vals) <= 0:
        return

    tok_label = "å�ˆè¨ˆãƒˆãƒ¼ã‚¯ãƒ³"
    ws.range((hr, cts)).value = "æ—¥ä»˜"
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
        ca2.ChartTitle.Text = "Gemini API æ—¥æ¬¡ãƒˆãƒ¼ã‚¯ãƒ³"
        ca2.HasLegend = False
    except Exception:
        pass


def _write_main_sheet_gemini_usage_via_xlwings(
    macro_wb_path: str, text: str, log_prefix: str
) -> bool:
    """Excel ã�§ãƒ–ãƒƒã‚¯ã�Œé–‹ã�„ã�¦ã�„ã‚‹ã�¨ã��ã€�ãƒ¡ã‚¤ãƒ³ P åˆ—ãƒ»Qã€œTãƒ»æŽ¨ç§»ã‚°ãƒ©ãƒ•ï¼ˆæœ€å¤§2æœ¬ï¼‰ã‚’ xlwings ã�§æ›´æ–°ã�—ã�¦ Saveã€‚"""
    attached = _xlwings_attach_open_macro_workbook(macro_wb_path, log_prefix)
    if attached is None:
        logging.info(
            "%s: xlwings ã�§ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã�«æŽ¥ç¶šã�§ã��ã�šã€�ãƒ¡ã‚¤ãƒ³ AI ã‚µãƒžãƒªã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚",
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
                "%s: ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆã�Œç„¡ã�„ã�Ÿã‚� xlwings ã�§ã�® AI ã‚µãƒžãƒªã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚",
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
                "%s: ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆ P%d ä»¥é™�ãƒ»Gemini æŽ¨ç§»ã‚°ãƒ©ãƒ•ï¼ˆæ–™é‡‘/å‘¼å‡ºã�—ãƒ»ãƒˆãƒ¼ã‚¯ãƒ³ï¼‰ã‚’ xlwings ã�§ä¿�å­˜ã�—ã�¾ã�—ã�Ÿã€‚",
                log_prefix,
                start_r,
            )
        finally:
            _xlwings_app_save_perf_state_pop(xw_book.app, _perf_snap)
    except Exception as ex:
        logging.warning(
            "%s: ãƒ¡ã‚¤ãƒ³ AI ã‚µãƒžãƒªã�® xlwings ä¿�å­˜ã�«å¤±æ•—: %s", log_prefix, ex
        )
        ok = False
    finally:
        _xlwings_release_book_after_mutation(xw_book, info, ok)
    return ok


def _gemini_kv_table_lines(title: str, rows: list[tuple[str, str]]) -> list[str]:
    """ç´¯è¨ˆãƒ»å½“å®Ÿè¡Œå�‘ã�‘ã�® 2 åˆ—ãƒ†ã‚­ã‚¹ãƒˆè¡¨ï¼ˆå±¥æ­´è¡Œã�¯å�«ã‚�ã�ªã�„ï¼‰ã€‚"""
    out = [title]
    if not rows:
        return out
    lw = min(22, max(len(a) for a, _ in rows))
    sep = "  " + ("â”€" * (lw + 2 + 28))
    out.append(sep)
    for a, b in rows:
        out.append(f"  {a:<{lw}}  {b}")
    return out


def _export_gemini_buckets_csv_for_charts(cum: dict) -> None:
    """Excel æŠ˜ã‚Œç·šãƒ»æ£’ã‚°ãƒ©ãƒ•å�‘ã�‘ã�«é•·å½¢å¼� CSV ã‚’ log ã�«æ›¸ã��å‡ºã�™ã€‚"""
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
        logging.debug("Gemini ãƒ�ã‚±ãƒƒãƒˆ CSV ã�®ä¿�å­˜ã�«å¤±æ•—: %s", ex)


def build_gemini_usage_summary_text() -> str:
    """ãƒ¡ã‚¤ãƒ³è¡¨ç¤ºãƒ»çµ�æžœãƒ­ã‚°ç”¨ã�®è¤‡æ•°è¡Œãƒ†ã‚­ã‚¹ãƒˆï¼ˆã�“ã�®å®Ÿè¡Œåˆ†ï¼‹ç´¯è¨ˆ JSONï¼‰ã€‚"""
    cum = _load_gemini_cumulative_payload()
    ct_tot = int(cum.get("calls_total") or 0)
    if not _gemini_usage_session and ct_tot <= 0:
        return ""

    lines: list[str] = []
    ts = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    if _gemini_usage_session:
        lines.append(f"é›†è¨ˆæ™‚åˆ»: {ts}ï¼ˆã�“ã�®å®Ÿè¡Œã�§ã�® Gemini APIï¼‰")
        tot_calls = sum(b["calls"] for b in _gemini_usage_session.values())
        tot_p = sum(b["prompt"] for b in _gemini_usage_session.values())
        tot_c = sum(b["candidates"] for b in _gemini_usage_session.values())
        tot_th = sum(b["thoughts"] for b in _gemini_usage_session.values())
        tot_t = sum(b["total"] for b in _gemini_usage_session.values())
        sess_rows: list[tuple[str, str]] = [
            ("å‘¼å‡ºã�—", f"{tot_calls:,} å›ž"),
            ("å…¥åŠ›ãƒˆãƒ¼ã‚¯ãƒ³", f"{tot_p:,}"),
            ("å‡ºåŠ›ãƒˆãƒ¼ã‚¯ãƒ³", f"{tot_c:,}"),
        ]
        if tot_th:
            sess_rows.append(("æ€�è€ƒãƒˆãƒ¼ã‚¯ãƒ³", f"{tot_th:,}"))
        sess_rows.append(("total å ±å‘Š", f"{tot_t:,}"))
        lines.extend(_gemini_kv_table_lines("ã€�ã�“ã�®å®Ÿè¡Œã€‘", sess_rows))
        grand_usd = 0.0
        any_price = False
        for mid in sorted(_gemini_usage_session.keys()):
            b = _gemini_usage_session[mid]
            mrows: list[tuple[str, str]] = [
                ("ãƒ¢ãƒ‡ãƒ«", mid),
                ("å‘¼å‡ºã�—", f"{b['calls']:,} å›ž"),
                ("å…¥åŠ›ãƒˆãƒ¼ã‚¯ãƒ³", f"{b['prompt']:,}"),
                ("å‡ºåŠ›ãƒˆãƒ¼ã‚¯ãƒ³", f"{b['candidates']:,}"),
            ]
            if b.get("thoughts", 0):
                mrows.append(("æ€�è€ƒãƒˆãƒ¼ã‚¯ãƒ³", f"{b['thoughts']:,}"))
            mrows.append(("total_token_count", f"{b['total']:,}"))
            est = _gemini_estimate_cost_usd(
                mid, b["prompt"], b["candidates"], b.get("thoughts", 0)
            )
            if est is not None:
                any_price = True
                grand_usd += est
                mrows.append(("æŽ¨å®šUSD", f"${est:.6f}"))
                mrows.append(
                    (
                        "æŽ¨å®šJPY",
                        f"Â¥{est * GEMINI_JPY_PER_USD:.2f}ï¼ˆ{GEMINI_JPY_PER_USD:.0f}å††/USDï¼‰",
                    )
                )
            else:
                mrows.append(("æŽ¨å®šæ–™é‡‘", "ï¼ˆå�˜ä¾¡æœªç™»éŒ²ãƒ¢ãƒ‡ãƒ«ï¼‰"))
            lines.append("")
            lines.extend(_gemini_kv_table_lines(f"ã€�ã�“ã�®å®Ÿè¡Œãƒ»ãƒ¢ãƒ‡ãƒ«åˆ¥ã€‘", mrows))
        if any_price:
            lines.append("")
            lines.extend(
                _gemini_kv_table_lines(
                    "ã€�ã�“ã�®å®Ÿè¡Œãƒ»æŽ¨å®šæ–™é‡‘å�ˆè¨ˆã€‘",
                    [
                        ("USD", f"${grand_usd:.6f}"),
                        (
                            "JPY",
                            f"Â¥{grand_usd * GEMINI_JPY_PER_USD:.2f}ï¼ˆ{GEMINI_JPY_PER_USD:.0f}å††/USDï¼‰",
                        ),
                    ],
                )
            )
    else:
        lines.append(f"é›†è¨ˆæ™‚åˆ»: {ts}")
        lines.append("ï¼ˆã�“ã�®å®Ÿè¡Œã�§ã�® Gemini API å‘¼å‡ºã�—ã�¯ã�‚ã‚Šã�¾ã�›ã‚“ï¼‰")
    lines.append("â€» ãƒˆãƒ¼ã‚¯ãƒ³ã�¯ API ã�® usage_metadata ã�«åŸºã�¥ã��ã�¾ã�™ã€‚")
    lines.append(
        "â€» USD å�˜ä¾¡ã�¯ã‚³ãƒ¼ãƒ‰ï¼�ç’°å¢ƒå¤‰æ•°ã�®ç›®å®‰ã�§ã�™ã€‚å®Ÿèª²é‡‘ã�¯ Google ã�®è«‹æ±‚ã‚’å�‚ç…§ã�—ã�¦ã��ã� ã�•ã�„ã€‚"
    )
    lines.append(
        "â€» å�„ API å‘¼å‡ºã�—ã�”ã�¨ã�®æŽ¨å®šæ–™é‡‘ã�¯ã‚³ãƒ³ã‚½ãƒ¼ãƒ«ã�«å‡ºã�•ã�šã€�ä¸‹è¨˜ç´¯è¨ˆ JSON ã�«ã�®ã�¿ç©�ã�¿ä¸Šã�’ã�¾ã�™ã€‚"
    )

    if ct_tot > 0:
        lines.append("")
        cum_hdr = (
            f"ã€�ç´¯è¨ˆã€‘{GEMINI_USAGE_CUMULATIVE_JSON_FILE} "
            "ï¼ˆAPI_Payment ãƒ•ã‚©ãƒ«ãƒ€ãƒ»å…¨å®Ÿè¡Œã�®æŽ¨å®šå€¤ï¼‰"
        )
        pt0 = int(cum.get("prompt_total") or 0)
        cc0 = int(cum.get("candidates_total") or 0)
        th0 = int(cum.get("thoughts_total") or 0)
        tt0 = int(cum.get("total_tokens_reported") or 0)
        cum_rows: list[tuple[str, str]] = [
            ("æœ€çµ‚æ›´æ–°", str(cum.get("updated_at") or "â€”")),
            ("å‘¼å‡ºã�—", f"{ct_tot:,} å›ž"),
            ("å…¥åŠ›ãƒˆãƒ¼ã‚¯ãƒ³", f"{pt0:,}"),
            ("å‡ºåŠ›ãƒˆãƒ¼ã‚¯ãƒ³", f"{cc0:,}"),
        ]
        if th0:
            cum_rows.append(("æ€�è€ƒãƒˆãƒ¼ã‚¯ãƒ³", f"{th0:,}"))
        cum_rows.append(("total å ±å‘Š", f"{tt0:,}"))
        usd_all = float(cum.get("estimated_cost_usd_total") or 0.0)
        if usd_all > 0:
            cum_rows.append(("æŽ¨å®šUSD ç´¯è¨ˆ", f"${usd_all:.6f}"))
            cum_rows.append(
                (
                    "æŽ¨å®šJPY ç´¯è¨ˆ",
                    f"Â¥{usd_all * GEMINI_JPY_PER_USD:.2f}ï¼ˆ{GEMINI_JPY_PER_USD:.0f}å††/USDï¼‰",
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
                    ("ãƒ¢ãƒ‡ãƒ«", mid),
                    ("å‘¼å‡ºã�—", f"{int(m.get('calls') or 0):,} å›ž"),
                    (
                        "å…¥åŠ› / å‡ºåŠ›",
                        f"{int(m.get('prompt') or 0):,} / {int(m.get('candidates') or 0):,}",
                    ),
                ]
                if int(m.get("thoughts") or 0):
                    mrows2.append(("æ€�è€ƒãƒˆãƒ¼ã‚¯ãƒ³", f"{int(m.get('thoughts') or 0):,}"))
                mud = float(m.get("estimated_cost_usd") or 0.0)
                if mud > 0:
                    mrows2.append(("æŽ¨å®šUSD ç´¯è¨ˆ", f"${mud:.6f}"))
                    mrows2.append(
                        ("æŽ¨å®šJPY ç´¯è¨ˆ", f"Â¥{mud * GEMINI_JPY_PER_USD:.2f}")
                    )
                lines.append("")
                lines.extend(_gemini_kv_table_lines("ã€�ç´¯è¨ˆãƒ»ãƒ¢ãƒ‡ãƒ«åˆ¥ã€‘", mrows2))
        trend = _gemini_usage_trend_caption_lines(cum)
        if trend:
            lines.append("")
            lines.extend(trend)
    return "\n".join(lines)


def write_main_sheet_gemini_usage_summary(wb_path: str, log_prefix: str) -> None:
    """Gemini åˆ©ç”¨ã‚µãƒžãƒªã‚’ log ã�«æ›¸ã��ã€�xlwings ã�§ãƒ¡ã‚¤ãƒ³ P åˆ—ãƒ»æŽ¨ç§»ã‚°ãƒ©ãƒ•ã�¸ä¿�å­˜ï¼ˆé–‹ã�„ã�¦ã�„ã‚‹ãƒ–ãƒƒã‚¯å�‘ã�‘ï¼‰ã€‚"""
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
                "%s: AI ã‚µãƒžãƒªã�® xlwings æ›¸ã��è¾¼ã�¿ã�§ä¾‹å¤–: %s", log_prefix, ex
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
        logging.debug("Gemini ãƒ�ã‚±ãƒƒãƒˆ CSV å‡ºåŠ›ã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰: %s", ex)
    if xw_ok:
        return
    if text.strip():
        logging.info(
            "%s: ãƒ¡ã‚¤ãƒ³ P åˆ—ãƒ»ã‚°ãƒ©ãƒ•ã‚’ xlwings ã�§ä¿�å­˜ã�§ã��ã�¾ã�›ã‚“ã�§ã�—ã�Ÿã€‚"
            " %s ã�«å‡ºåŠ›æ¸ˆã�¿ â†’ ãƒžã‚¯ãƒ­ã€Œãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆ_Geminiåˆ©ç”¨ã‚µãƒžãƒªã‚’Påˆ—ã�«å��æ˜ ã€�ã�§ P åˆ—ã�®ã�¿å��æ˜ ã�§ã��ã�¾ã�™ã€‚",
            log_prefix,
            path,
        )
    else:
        logging.info(
            "%s: Gemini æœªä½¿ç”¨: ã‚µãƒžãƒªã‚’ç©ºã�§ %s ã�«å‡ºåŠ›ã€‚",
            log_prefix,
            path,
        )


def _try_write_main_sheet_gemini_usage_summary(phase: str) -> None:
    try:
        write_main_sheet_gemini_usage_summary(TASKS_INPUT_WORKBOOK, phase)
    except Exception as ex:
        logging.warning(
            "%s: ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆã�¸ã�® AI åˆ©ç”¨ã‚µãƒžãƒªæ›¸ã��è¾¼ã�¿ã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰: %s", phase, ex
        )


def _plan_sheet_write_global_parse_block_to_ws(
    ws,
    global_priority_override: dict,
    when_str: str,
) -> None:
    """æ—¢ã�«é–‹ã�„ã�¦ã�„ã‚‹ã€Œé…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€�ç›¸å½“ã‚·ãƒ¼ãƒˆã�¸ AX:AY ã�®ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è§£æž�ãƒ–ãƒ­ãƒƒã‚¯ã‚’æ›¸ã��ã€‚"""
    gpo = global_priority_override or {}
    lc = PLAN_SHEET_GLOBAL_PARSE_LABEL_COL
    vc = PLAN_SHEET_GLOBAL_PARSE_VALUE_COL
    max_r = PLAN_SHEET_GLOBAL_PARSE_MAX_ROWS
    for i in range(max_r):
        ws.cell(row=1 + i, column=lc, value=None)
        ws.cell(row=1 + i, column=vc, value=None)
    align_top = Alignment(wrap_text=True, vertical="top")
    pairs: list[tuple[str, str]] = [
        ("ã€�ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�ã€‘", "å�‚ç…§ç”¨ãƒ»æ®µéšŽ2ã�§è‡ªå‹•è¨˜éŒ²"),
        (
            "â€»äºŒé‡�é�©ç”¨ã�«ã�¤ã�„ã�¦",
            "é…�å�°ã�¸ã�®å��æ˜ ã�¯ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�ã�‹ã‚‰ã�®ã�¿è¡Œã‚�ã‚Œã�¾ã�™ã€‚"
            "ã�“ã�®AXã€œAYåˆ—ã�¯èª­ã�¿å�–ã‚‰ã‚Œã�¾ã�›ã‚“ã€‚ç·¨é›†ã�—ã�¦ã‚‚æ¬¡å›žå®Ÿè¡Œã�¾ã�§é…�å�°ã�«åŠ¹ã��ã�¾ã�›ã‚“ã€‚"
            "åŽŸæ–‡ã�¯ãƒ¡ã‚¤ãƒ³æ¬„ã‚’å�‚ç…§ã�—ã�¦ã��ã� ã�•ã�„ã€‚",
        ),
        ("è¨ˆç”»åŸºæº–æ—¥æ™‚", (when_str or "").strip() or "â€•"),
        (
            "å·¥å ´ä¼‘æ¥­æ—¥",
            ", ".join(str(x) for x in (gpo.get("factory_closure_dates") or []))
            if gpo.get("factory_closure_dates")
            else "ï¼ˆã�ªã�—ï¼‰",
        ),
        (
            "ã‚¹ã‚­ãƒ«è¦�ä»¶ã‚’ç„¡è¦–",
            "ã�¯ã�„" if gpo.get("ignore_skill_requirements") else "ã�„ã�„ã�ˆ",
        ),
        (
            "needäººæ•°1å›ºå®š",
            "ã�¯ã�„" if gpo.get("ignore_need_minimum") else "ã�„ã�„ã�ˆ",
        ),
        (
            "é…�å�°åˆ¶é™�ã�®æ’¤å»ƒ",
            "ã�¯ã�„" if gpo.get("abolish_all_scheduling_limits") else "ã�„ã�„ã�ˆ",
        ),
        (
            "ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«OPæŒ‡å��",
            json.dumps(gpo.get("task_preferred_operators") or {}, ensure_ascii=False)
            if gpo.get("task_preferred_operators")
            else "ï¼ˆã�ªã�—ï¼‰",
        ),
        (
            "æ—¥ä»˜Ã—å·¥ç¨‹ãƒ�ãƒ¼ãƒ æŒ‡å��",
            json.dumps(
                gpo.get("global_day_process_operator_rules") or [],
                ensure_ascii=False,
            )
            if gpo.get("global_day_process_operator_rules")
            else "ï¼ˆã�ªã�—ï¼‰",
        ),
        (
            "ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«é€Ÿåº¦ãƒ«ãƒ¼ãƒ«",
            json.dumps(gpo.get("global_speed_rules") or [], ensure_ascii=False)
            if gpo.get("global_speed_rules")
            else "ï¼ˆã�ªã�—ï¼‰",
        ),
        (
            "æœªé�©ç”¨ãƒ¡ãƒ¢(AI)",
            str(gpo.get("scheduler_notes_ja") or "").strip() or "ï¼ˆã�ªã�—ï¼‰",
        ),
        (
            "AIè¦�ç´„",
            str(gpo.get("interpretation_ja") or "").strip() or "ï¼ˆã�ªã�—ï¼‰",
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
    log_prefix: str = "æ®µéšŽ2",
) -> bool:
    """
    ã€Œé…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€�ã‚·ãƒ¼ãƒˆã�®å�³ç«¯ä»˜è¿‘ï¼ˆAX:AYï¼‰ã�«ã€�ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã�®è§£æž�çµ�æžœã‚’æ›¸ã��è¾¼ã‚€ã€‚
    ãƒ¡ã‚¤ãƒ³åŽŸæ–‡ã�¯ã�“ã�“ã�«è»¢è¨˜ã�—ã�ªã�„ï¼ˆãƒ¡ã‚¤ãƒ³æ¬„ã�¨ã�®é‡�è¤‡ãƒ»èª¤è§£ã‚’é�¿ã�‘ã‚‹ï¼‰ã€‚æœ¬åˆ—ã�¯å†�èª­è¾¼ã�•ã‚Œã�šå�‚ç…§å°‚ç”¨ã€‚
    Excel ã�§ãƒ–ãƒƒã‚¯ã‚’é–‹ã�„ã�Ÿã�¾ã�¾ã� ã�¨ä¿�å­˜ã�«å¤±æ•—ã�™ã‚‹ã�“ã�¨ã�Œã�‚ã‚‹ï¼ˆä»–ã�® openpyxl æ›¸è¾¼ã�¨å�Œæ§˜ï¼‰ã€‚
    """
    if not wb_path or not os.path.isfile(wb_path):
        return False
    gpo = global_priority_override or {}
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = None
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "%s: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� openpyxl ã�§ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�ã‚’é…�å�°ã‚·ãƒ¼ãƒˆã�¸æ›¸ã��è¾¼ã�¿ã�¾ã�›ã‚“ã€‚",
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
            "%s: ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�ã�®é…�å�°ã‚·ãƒ¼ãƒˆæ›¸è¾¼ã�®ã�Ÿã‚�ãƒ–ãƒƒã‚¯ã‚’é–‹ã�‘ã�¾ã�›ã‚“: %s",
            log_prefix,
            ex,
        )
        return False
    try:
        if sheet_name not in wb.sheetnames:
            logging.info(
                "%s: ã‚·ãƒ¼ãƒˆ '%s' ã�Œç„¡ã�„ã�Ÿã‚�ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�ã�®å��æ˜ ã‚’ã‚¹ã‚­ãƒƒãƒ—ã€‚",
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
            "%s: ã€Œ%sã€�%s:%s åˆ—ã�«ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�ã‚’ä¿�å­˜ã�—ã�¾ã�—ã�Ÿã€‚",
            log_prefix,
            sheet_name,
            get_column_letter(lc),
            get_column_letter(vc),
        )
        return True
    except OSError as ex:
        logging.warning(
            "%s: ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�ã‚’é…�å�°ã‚·ãƒ¼ãƒˆã�¸ä¿�å­˜ã�§ã��ã�¾ã�›ã‚“ã�§ã�—ã�Ÿï¼ˆExcel ã�§é–‹ã�„ã�Ÿã�¾ã�¾ç­‰ï¼‰: %s",
            log_prefix,
            ex,
        )
        return False
    except Exception as ex:
        logging.warning(
            "%s: ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�ã�®é…�å�°ã‚·ãƒ¼ãƒˆæ›¸è¾¼ã�§ä¾‹å¤–: %s", log_prefix, ex
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
            log_prefix="æ®µéšŽ2",
        )
    except Exception as ex:
        logging.warning(
            "æ®µéšŽ2: é…�å�°ã‚·ãƒ¼ãƒˆã�¸ã�®ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�æ›¸ã��è¾¼ã�¿ã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰: %s",
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
            log_prefix="æ®µéšŽ2",
        )
    except Exception as ex:
        logging.warning(
            "æ®µéšŽ2: é…�å�°ã‚·ãƒ¼ãƒˆã�¸ã�®ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è§£æž�ï¼‹çŸ›ç›¾ç�€è‰²ï¼ˆ1å›žä¿�å­˜ï¼‰ã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰: %s",
            ex,
        )


def _log_task_special_ai_response(raw_text, parsed, extracted_json_str, prompt_text=None):
    """ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒå�‘ã�‘ Gemini ã�®ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆãƒ»ç”Ÿãƒ†ã‚­ã‚¹ãƒˆãƒ»æŠ½å‡ºJSONãƒ»ãƒ‘ãƒ¼ã‚¹çµ�æžœã‚’1ãƒ•ã‚¡ã‚¤ãƒ«ã�«æ®‹ã�™ã€‚"""
    path = os.path.join(log_dir, TASK_SPECIAL_AI_LAST_RESPONSE_FILE)
    try:
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            if prompt_text is not None and str(prompt_text).strip():
                f.write("=== Gemini ã�¸é€�ä¿¡ã�—ã�Ÿãƒ—ãƒ­ãƒ³ãƒ—ãƒˆï¼ˆå…¨æ–‡ï¼‰ ===\n")
                f.write(str(prompt_text).strip())
                f.write("\n\n")
            f.write("=== Gemini è¿”å�´ãƒ†ã‚­ã‚¹ãƒˆï¼ˆãƒ¢ãƒ‡ãƒ«å‡ºåŠ›ã��ã�®ã�¾ã�¾ï¼‰ ===\n")
            f.write(raw_text or "")
            f.write(
                "\n\n=== AI ã�Œè¿”ã�—ã�Ÿãƒ†ã‚­ã‚¹ãƒˆã�‹ã‚‰ã‚¯ãƒ©ã‚¤ã‚¢ãƒ³ãƒˆã�Œåˆ‡ã‚Šå‡ºã�—ã�Ÿ JSON æ–‡å­—åˆ— ===\n"
                "ï¼ˆâ€»ãƒ¦ãƒ¼ã‚¶ãƒ¼ç‰¹åˆ¥æŒ‡å®šã�®è§£æž�ã�«æ­£è¦�è¡¨ç�¾ã�¯ä½¿ã�£ã�¦ã�„ã�¾ã�›ã‚“ã€‚ãƒ¢ãƒ‡ãƒ«å¿œç­”ã�®ãƒ‘ãƒ¼ã‚¹ç”¨ã�§ã�™ï¼‰\n"
            )
            f.write(extracted_json_str if extracted_json_str else "(æŠ½å‡ºã�ªã�—)")
            f.write("\n\n=== json.loads å¾Œï¼ˆä¾�é ¼NOã‚­ãƒ¼ï¼‰ ===\n")
            if isinstance(parsed, dict):
                f.write(json.dumps(parsed, ensure_ascii=False, indent=2))
            else:
                f.write("(ãƒ‘ãƒ¼ã‚¹ã�§ã��ã�š)")
        logging.info(
            "ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆï¼‹AIå¿œç­”ã�®è©³ç´° â†’ %s",
            path,
        )
    except OSError as ex:
        logging.warning("ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: AIå¿œç­”ãƒ•ã‚¡ã‚¤ãƒ«ä¿�å­˜ã�«å¤±æ•—: %s", ex)
    if isinstance(parsed, dict) and parsed:
        logging.info(
            "ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: è§£æž�ã�•ã‚Œã�Ÿä¾�é ¼NO: %s",
            ", ".join(sorted(parsed.keys(), key=lambda x: str(x))),
        )
        for tid_k in sorted(parsed.keys(), key=lambda x: str(x)):
            logging.info(
                "  ä¾�é ¼NO [%s] AIè§£æž�ãƒ•ã‚£ãƒ¼ãƒ«ãƒ‰: %s",
                tid_k,
                json.dumps(parsed[tid_k], ensure_ascii=False),
            )


def _parse_and_log_task_special_gemini_response(res, prompt_text=None):
    """
    API ãƒ¬ã‚¹ãƒ�ãƒ³ã‚¹ã‚’ JSON åŒ–ã�—ãƒ­ã‚°ï¼�ãƒ•ã‚¡ã‚¤ãƒ«ã�¸è¨˜éŒ²ã€‚å¤±æ•—æ™‚ã�¯ Noneã€‚
    ãƒ¦ãƒ¼ã‚¶ãƒ¼ã�®ç‰¹åˆ¥æŒ‡å®šæ–‡è¨€ã�«ã�¯è§¦ã‚Œã�šã€�ãƒ¢ãƒ‡ãƒ«å‡ºåŠ›ã�‹ã‚‰ JSON ãƒ–ãƒ­ãƒƒã‚¯ã‚’å�–ã‚Šå‡ºã�™å‡¦ç�†ã�®ã�¿ã€‚
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
            "ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: AIå¿œç­”ã�‹ã‚‰ JSON ã‚’æŠ½å‡ºã�§ã��ã�¾ã�›ã‚“ã�§ã�—ã�Ÿã€‚ç”Ÿãƒ†ã‚­ã‚¹ãƒˆå…ˆé ­ 3000 æ–‡å­—:\n%s",
            (raw[:3000] if raw else "(ç©º)"),
        )
        return None
    extracted = match.group(0)
    try:
        parsed = json.loads(extracted)
    except json.JSONDecodeError as je:
        _log_task_special_ai_response(raw, None, extracted, prompt_text)
        logging.warning("ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: JSON ãƒ‘ãƒ¼ã‚¹å¤±æ•—: %s", je)
        return None
    if not isinstance(parsed, dict):
        _log_task_special_ai_response(raw, None, extracted, prompt_text)
        logging.warning("ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: ãƒˆãƒƒãƒ—ãƒ¬ãƒ™ãƒ«ã�Œ JSON ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�§ã�¯ã�‚ã‚Šã�¾ã�›ã‚“ã€‚")
        return None
    _log_task_special_ai_response(raw, parsed, extracted, prompt_text)
    return parsed


def analyze_task_special_remarks(tasks_df, reference_year=None, ai_sheet_sink: dict | None = None):
    """
    ã€Œé…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€�ã�®ã€Œç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã€�ã‚’ AI ã�§æ§‹é€ åŒ–ï¼ˆã‚»ãƒ«ã�«å€¤ã�Œã�‚ã‚‹é …ç›®ã�¯å¾Œæ®µã�§ã‚»ãƒ«ã‚’å„ªå…ˆï¼‰ã€‚
    ã€Œé…�å�°ä¸�è¦�ã€�ã�Œã‚ªãƒ³ã�ªè¡Œã�¯ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆã�«è¼‰ã�›ã�ªã�„ï¼ˆAPI ç¯€ç´„ãƒ»å½“è©²è¡Œã�¯é…�å�°ã�—ã�ªã�„ã�Ÿã‚�ï¼‰ã€‚
    æ‹…å½“OPæŒ‡å��ã�¯ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆã�®è¿”å�´å¥‘ç´„ã�§ãƒ¢ãƒ‡ãƒ«ã�« preferred_operator ã‚’å‡ºåŠ›ã�•ã�›ã‚‹ï¼ˆå‚™è€ƒã‚’æ­£è¦�è¡¨ç�¾ã�§åˆ‡ã‚Šå‡ºã�™å‡¦ç�†ã�¯è¡Œã‚�ã�ªã�„ï¼‰ã€‚
    json/ai_remarks_cache.json ã�« TTL AI_CACHE_TTL_SECONDS ã�§ã‚­ãƒ£ãƒƒã‚·ãƒ¥ï¼ˆå�Œä¸€å…¥åŠ›ãƒ»å�Œä¸€åŸºæº–å¹´ã�ªã‚‰ API ã‚’å‘¼ã�°ã�ªã�„ï¼‰ã€‚
    ä¾�é ¼NOã�¯æ•°å€¤è¡¨è¨˜ãƒ»å…¨è§’ã�ªã�©ã‚’æ­£è¦�åŒ–ã�—ã�¦ã‚­ãƒ¼ã‚’å®‰å®šåŒ–ã�—ã€�åŸºæº–å¹´ã�¯æŒ‡ç´‹ã�«å�«ã‚�ã�¦æ—¥ä»˜è§£é‡ˆã�®å¤‰åŒ–ã�¨ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã�®é£Ÿã�„é�•ã�„ã‚’é˜²ã��ã€‚

    æˆ»ã‚Šå€¤ã�®ä¾‹: ä¾�é ¼NO -> ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã€�ã�¾ã�Ÿã�¯å�Œä¸€ä¾�é ¼NOã�«å‚™è€ƒè¡Œã�Œè¤‡æ•°ã�‚ã‚‹å ´å�ˆã�¯ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�®é…�åˆ—ã€‚
      process_name, machine_name â€¦ å½“è©²å‚™è€ƒã‚»ãƒ«ã�Œã�‚ã‚‹è¡Œã�®å·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ï¼ˆãƒ—ãƒ­ãƒ³ãƒ—ãƒˆã�®è¡Œã�¨ä¸€è‡´ï¼‰
      restrict_to_process_name, restrict_to_machine_name â€¦ çœ�ç•¥ã�¾ã�Ÿã�¯ç©ºã�ªã‚‰å�Œä¸€ä¾�é ¼NOã�®å…¨å·¥ç¨‹ãƒ»å…¨æ©Ÿæ¢°è¡Œã�«é�©ç”¨ã€‚
      ã��ã�®ä»– required_op, speed_override, task_efficiency, priority, start_date, start_time,
      target_completion_date, ship_by_date, preferred_operator ã�ªã�©ã€‚
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
            "ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: AI è§£æž�å¯¾è±¡ã�Œã�‚ã‚Šã�¾ã�›ã‚“ï¼ˆã€Œ%sã€�åˆ—ã�¯%sï¼‰ã€‚"
            "æ€»è¡Œæ•°=%sã€�ä¾�é ¼NOã�®ã�‚ã‚‹è¡Œ=%sã€�å‚™è€ƒã�Œå…¥ã�£ã�¦ã�„ã‚‹è¡Œ=%sã€‚"
            "æ®µéšŽ2å®Ÿè¡Œå‰�ã�«ãƒ–ãƒƒã‚¯ã‚’ä¿�å­˜ã�—ã€�æœ¬å½“ã�«ã€Œ%sã€�åˆ—ã�¸å…¥åŠ›ã�—ã�¦ã�„ã‚‹ã�‹ç¢ºèª�ã�—ã�¦ã��ã� ã�•ã�„ã€‚",
            PLAN_COL_SPECIAL_REMARK,
            "è¦‹ã�¤ã�‹ã‚Šã�¾ã�›ã‚“" if miss_col else "ç©ºã�®å�¯èƒ½æ€§ã�Œã�‚ã‚Šã�¾ã�™",
            n_rows,
            n_tid_raw,
            n_rem_only,
            PLAN_COL_SPECIAL_REMARK,
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ_AI_API"] = "ã‚¹ã‚­ãƒƒãƒ—ï¼ˆå¯¾è±¡è¡Œã�ªã�—ï¼‰"
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
            "ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: ã‚­ãƒ£ãƒƒã‚·ãƒ¥ãƒ’ãƒƒãƒˆï¼ˆ%s ä»¶ãƒ»åŸºæº–å¹´=%sï¼‰ã€‚Gemini ã�¯å‘¼ã�³ã�¾ã�›ã‚“ã€‚",
            len(lines),
            ref_y,
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ_AI_API"] = "ã�ªã�—ï¼ˆã‚­ãƒ£ãƒƒã‚·ãƒ¥ä½¿ç”¨ï¼‰"
        out = copy.deepcopy(cached_parsed)
        if isinstance(out, dict):
            _repair_task_special_ai_wrong_top_level_keys(out, tasks_df)
        return out

    logging.info(
        "ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã�ªã�—ã€‚Gemini ã�§ %s ä»¶ã�®å‚™è€ƒã‚’è§£æž�ã�—ã�¾ã�™ï¼ˆåŸºæº–å¹´=%sï¼‰ã€‚",
        len(lines),
        ref_y,
    )

    if not API_KEY:
        logging.info("GEMINI_API_KEY æœªè¨­å®šã�®ã�Ÿã‚�ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®šã�®AIè§£æž�ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚")
        if ai_sheet_sink is not None:
            ai_sheet_sink["ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ_AI_API"] = "ã�ªã�—ï¼ˆAPIã‚­ãƒ¼æœªè¨­å®šï¼‰"
        return {}

    prompt = f"""
ã�‚ã�ªã�Ÿã�¯å·¥å ´ã�®é…�å�°è¨ˆç”»å�‘ã�‘ã�«ã€�Excelã€Œç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã€�æ¬„ã�¸ã�®è‡ªç”±è¨˜è¿°ã‚’èª­ã�¿ã€�é…�å�°ãƒ­ã‚¸ãƒƒã‚¯ã�Œä½¿ã�ˆã‚‹ãƒ•ã‚£ãƒ¼ãƒ«ãƒ‰ã� ã�‘ã�«è�½ã�¨ã�—è¾¼ã‚€ã‚¢ã‚·ã‚¹ã‚¿ãƒ³ãƒˆã�§ã�™ã€‚

ã€�æœ€é‡�è¦�ã€‘
1) ã€�ç‰¹åˆ¥æŒ‡å®šåŽŸæ–‡ã€‘ã�®å�„è¡Œã�¯ã€�ãƒ¦ãƒ¼ã‚¶ãƒ¼ã�Œã‚»ãƒ«ã�«å…¥åŠ›ã�—ã�Ÿæ–‡å­—åˆ—ã‚’ **æ”¹å¤‰ãƒ»è¦�ç´„ãƒ»æ–­ã�¡åˆ‡ã‚Šã�¯ã�—ã�¦ã�Šã‚‰ã�š**ï¼ˆå…ˆé ­æœ«å°¾ã�®ç©ºç™½ã�®ã�¿é™¤åŽ»ï¼‰ã€�ã��ã�®ã�¾ã�¾æ¸¡ã�—ã�¦ã�„ã�¾ã�™ã€‚**åŽŸæ–‡ã�®äº‹å®Ÿã‚„æ„�å›³ã‚’åˆ¥ã�®æ–‡è¨€ã�«ç½®ã��æ�›ã�ˆã�ªã�„ã�§ã��ã� ã�•ã�„ã€‚**
2) ã�‚ã�ªã�Ÿã�®å¿œç­”ã�¯ **1å€‹ã�® JSON ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�®ã�¿**ï¼ˆå…ˆé ­ã�Œ {{ ã€�æœ«å°¾ã�Œ }} ï¼‰ã€‚èª¬æ˜Žæ–‡ãƒ»ãƒžãƒ¼ã‚¯ãƒ€ã‚¦ãƒ³ãƒ»ã‚³ãƒ¼ãƒ‰ãƒ•ã‚§ãƒ³ã‚¹ã�¯ç¦�æ­¢ã€‚
3) JSON ã�®ãƒˆãƒƒãƒ—ãƒ¬ãƒ™ãƒ«ã‚­ãƒ¼ã�¯ã€�å�„è¡Œã�® **ä¾�é ¼NOã€�ã�¨ã€‘ã�®é–“ã�®æ–‡å­—åˆ—ã�®ã�¿** ã�¨ **å®Œå…¨ä¸€è‡´** ã�•ã�›ã‚‹ã�“ã�¨ã€‚**å‚™è€ƒæœ¬æ–‡**ã�«æ›¸ã�‹ã‚Œã�Ÿå“�ç•ªãƒ»åŽŸå��å��ãƒ»è£½å“�ã‚³ãƒ¼ãƒ‰ï¼ˆä¾‹: 20010 ã�§å§‹ã�¾ã‚‹ç•ªå�·åˆ—ï¼‰ã‚’ã‚­ãƒ¼ã�«ã�—ã�¦ã�¯ã�ªã‚‰ã�ªã�„ã€‚å‚™è€ƒã�Œã��ã�®ã‚ˆã�†ã�ªç•ªå�·ã�§å§‹ã�¾ã�£ã�¦ã�„ã�¦ã‚‚ã€�ã‚­ãƒ¼ã�¯å¿…ã�šã€�ã€‘å†…ã�®ä¾�é ¼NOã� ã�‘ã�¨ã�™ã‚‹ã€‚

ã€�è¿”å�´JSONã�®å¥‘ç´„ï¼ˆã�“ã�®ç¯€ã�©ã�Šã‚Šã�«å‡ºåŠ›ã�™ã‚‹ã�“ã�¨ï¼‰ã€‘
â–  ãƒˆãƒƒãƒ—ãƒ¬ãƒ™ãƒ«
- ã‚­ãƒ¼: ä¸Šè¨˜ã€�ç‰¹åˆ¥æŒ‡å®šåŽŸæ–‡ã€‘ã�® **ä¾�é ¼NOã€�â€¦ã€‘ã�®æ‹¬å¼§å†…** ã�®æ–‡å­—åˆ—ã�¨ **å®Œå…¨ä¸€è‡´**ï¼ˆè¡¨è¨˜ãƒ»ãƒ�ã‚¤ãƒ•ãƒ³ãƒ»è‹±å¤§æ–‡å­—å°�æ–‡å­—ã‚’åŽŸæ–‡ã�©ã�Šã‚Šï¼‰ã€‚å‚™è€ƒæœ¬æ–‡ä¸­ã�®æ•°å­—åˆ—ã‚’ã‚­ãƒ¼ã�«ã�—ã�ªã�„ã€‚
- å€¤: æ¬¡ã�®ã�„ã�šã‚Œã�‹ã€‚
  (A) **JSONã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆ1ã�¤** â€¦ å½“è©²ä¾�é ¼NOã�®å‚™è€ƒã�Œãƒ—ãƒ­ãƒ³ãƒ—ãƒˆä¸Š **1è¡Œã� ã�‘** ã�®ã�¨ã��ã€‚
  (B) **JSONé…�åˆ—**ï¼ˆè¦�ç´ ã�¯ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆï¼‰â€¦ å�Œä¸€ä¾�é ¼NOã�§å·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ã�Œç•°ã�ªã‚‹å‚™è€ƒè¡Œã�Œ **è¤‡æ•°** ã�‚ã‚‹ã�¨ã��ã€‚è¦�ç´ ã�®é †ã�¯ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆã�®è¡Œé †ã�¨å¯¾å¿œã�•ã�›ã‚‹ã€‚

â–  process_nameï¼ˆæ–‡å­—åˆ—ï¼‰ãƒ»machine_nameï¼ˆæ–‡å­—åˆ—ï¼‰â€” **å¿…é ˆ**
- å½“è©²å‚™è€ƒã�«å¯¾å¿œã�™ã‚‹ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆè¡Œã�® **å·¥ç¨‹å��ã€Œâ€¦ã€�**ãƒ»**æ©Ÿæ¢°å��ã€Œâ€¦ã€�** ã�®å€¤ã�¨ **ä¸€è‡´** ã�•ã�›ã‚‹ï¼ˆã€Œï¼ˆç©ºï¼‰ã€�ã�®ã�¨ã��ã�¯ç©ºæ–‡å­—åˆ— ""ï¼‰ã€‚
- ãƒ­ã‚°ãƒ»ãƒˆãƒ¬ãƒ¼ã‚¹ç”¨ã€‚çœ�ç•¥ä¸�å�¯ã€‚

â–  restrict_to_process_nameï¼ˆæ–‡å­—åˆ—ï¼‰ãƒ»restrict_to_machine_nameï¼ˆæ–‡å­—åˆ—ï¼‰â€” **ä»»æ„�**
- **åŽŸæ–‡ã�Œã€Œç‰¹å®šã�®å·¥ç¨‹ã� ã�‘ã€�ã€Œã�“ã�®æ©Ÿæ¢°ã� ã�‘ã€�ã�ªã�©ã€�é�©ç”¨ç¯„å›²ã‚’çµžã�£ã�¦ã�„ã‚‹ã�¨ã��ã� ã�‘** å‡ºåŠ›ã�™ã‚‹ã€‚
- **åŽŸæ–‡ã�«å·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ã�®é™�å®šã�Œç„¡ã�„**ï¼ˆä¾�é ¼å…¨ä½“ãƒ»å…¨è¡Œç¨‹ã�¸ã�®æŒ‡ç¤ºï¼‰ã�¨ã��ã�¯ **ä¸¡æ–¹ã�¨ã‚‚çœ�ç•¥** ã�™ã‚‹ã�‹ **ç©ºæ–‡å­—åˆ— ""** ã�¨ã�™ã‚‹ã€‚
- ã��ã�®å ´å�ˆã€�é…�å�°ãƒ­ã‚¸ãƒƒã‚¯ã�¯ **å�Œä¸€ä¾�é ¼NOã�®åˆ¥è¡Œï¼ˆä¾‹: ã‚¨ãƒ³ãƒœã‚¹è¡Œã�¨åˆ†å‰²è¡Œï¼‰ã�«ã‚‚å�Œã�˜æŒ‡ç¤ºã‚’é�©ç”¨** ã�™ã‚‹ã€‚
- çµžã‚‹å ´å�ˆã�¯ã€�åŽŸæ–‡ã�§ç¤ºã�•ã‚Œã�Ÿè­˜åˆ¥å��ã‚’å…¥ã‚Œã‚‹ï¼ˆExcel ã�®å·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ã�¨ç…§å�ˆã�—ã‚„ã�™ã�„è¡¨è¨˜ï¼‰ã€‚

â–  preferred_operatorï¼ˆæ–‡å­—åˆ—ï¼‰â€” æ�¡ä»¶ä»˜ã��**å¿…é ˆ**
- **å¿…è¦�æ�¡ä»¶**: å½“è©²ä¾�é ¼ã�®åŽŸæ–‡ã‚’èª­ã�¿ã€�ã€Œ**èª°ã�Œã�“ã�®åŠ å·¥ãƒ»ä½œæ¥­ã�®ä¸»æ‹…å½“ï¼ˆOPï¼‰ã�¨ã�—ã�¦å‰²ã‚Šå½“ã�¦ã�Ÿã�„ã�‹**ã€�ã�Œ **æ„�å‘³ã�¨ã�—ã�¦** èª­ã�¿å�–ã‚Œã‚‹ã�¨ã��ã€‚
  ä¾‹: ç‰¹å®šã�®äººã�«ã‚„ã�£ã�¦ã‚‚ã‚‰ã�†ï¼�ã��ã�®äººã�«ä»»ã�›ã‚‹ï¼�æ‹…å½“ã�¯ã�‚ã�®äººï¼�OPã�¯ã€œï¼�ã€œã�•ã‚“ï¼ˆæ°�å��ï¼‰ã�«ä¾�é ¼ã€�ã�ªã�©ã€‚**è¡¨ç�¾ã�®åž‹ã�«ä¾�å­˜ã�›ã�š**ã€�æ–‡ã�®æ„�å‘³ã�§åˆ¤æ–­ã�™ã‚‹ã€‚
- **æº€ã�Ÿã�—ã�Ÿã�¨ã��ã�®å‡ºåŠ›ç¾©å‹™**: ä¸Šè¨˜ã�®æ„�å‘³ã�Œæˆ�ç«‹ã�™ã‚‹ã�¨åˆ¤æ–­ã�—ã�Ÿã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�§ã�¯ã€�**å¿…ã�š** ã‚­ãƒ¼ `preferred_operator` ã‚’å�«ã‚�ã€�å€¤ã�¯ **ç©ºã�§ã�ªã�„æ–‡å­—åˆ—** ã�¨ã�™ã‚‹ã€‚ä½µã�›ã�¦ **process_name / machine_name ã�¯å¿…é ˆ**ï¼ˆä¾‹: `{{"process_name":"â€¦","machine_name":"â€¦","preferred_operator":"â€¦"}}`ï¼‰ã€‚
- **å€¤ã�®å½¢å¼�**: åŽŸæ–‡ã�§ç¤ºã�•ã‚Œã�Ÿ **æ‹…å½“è€…ã�®è­˜åˆ¥å��ã‚’1å��åˆ†**ï¼ˆå§“ãƒ»å��ãƒ»ãƒ‹ãƒƒã‚¯ãƒ�ãƒ¼ãƒ ç­‰ã€�åŽŸæ–‡ã�«ç�¾ã‚Œã�Ÿè¡¨è¨˜ã‚’ç¶­æŒ�ï¼‰ã€‚æœ«å°¾ã�®æ•¬ç§°ï¼ˆã�•ã‚“ãƒ»å�›ãƒ»æ°�ï¼‰ã�®ã�¿é™¤åŽ»ã€‚ä¾‹:ã€Œæ£®å²¡ã�•ã‚“ã�«ã‚„ã�£ã�¦ã‚‚ã‚‰ã�„ã�¾ã�™ã€�â†’ `"æ£®å²¡"`ã€‚
- **å‡ºåŠ›ã�—ã�¦ã�¯ã�„ã�‘ã�ªã�„ã�¨ã��**: åŽŸæ–‡ã�«æ‹…å½“è€…ã�®æŒ‡æ„�ã�Œ **ä¸€åˆ‡ã�ªã�„** ã�¨åˆ¤æ–­ã�—ã�Ÿä¾�é ¼NOã�§ã�¯ `preferred_operator` ã‚­ãƒ¼è‡ªä½“ã‚’ **çœ�ç•¥** ã�™ã‚‹ï¼ˆç©ºæ–‡å­—åˆ—ã‚‚ä»˜ã�‘ã�ªã�„ï¼‰ã€‚

â–  ã��ã�®ä»–ãƒ•ã‚£ãƒ¼ãƒ«ãƒ‰ï¼ˆrequired_op, speed_override, task_efficiency, priority, start_date, start_time, target_completion_date, ship_by_dateï¼‰
- åŽŸæ–‡ã�‹ã‚‰ **æ˜Žç¢ºã�«** èª­ã�¿å�–ã‚Œã‚‹å ´å�ˆã�®ã�¿å‡ºåŠ›ã€‚èª­ã�¿å�–ã‚Œã�ªã�„æ•°å€¤ãƒ»æ—¥ä»˜ã�¯ **çœ�ç•¥**ï¼ˆæŽ¨æ¸¬ã�§åŸ‹ã‚�ã�ªã�„ï¼‰ã€‚

ã€�å�Œä¸€ä¾�é ¼NOãƒ»è¤‡æ•°å·¥ç¨‹ã�®ä¾‹ã€‘
ä¾�é ¼NO Y4-2 ã�«ã€Œã‚¨ãƒ³ãƒœã‚¹ã€�ã�¨ã€Œåˆ†å‰²ã€�ã�®è¡Œã�Œã�‚ã‚Šã€�å‚™è€ƒã�Œã€Œ4/5ã�¾ã�§ã�«çµ‚ã‚�ã‚‰ã�›ã‚‹ã€�ã�®ã�¿ã�§å·¥ç¨‹ã�®é™�å®šã�Œç„¡ã�„å ´å�ˆ:
- process_name / machine_name ã�¯ **å‚™è€ƒã�Œæ›¸ã�‹ã‚Œã�Ÿè¡Œ** ã�®å€¤ã‚’å…¥ã‚Œã‚‹ã€‚
- restrict_to_* ã�¯ **å‡ºã�•ã�ªã�„ã�‹ç©º** ã�«ã�—ã€�**ã‚¨ãƒ³ãƒœã‚¹è¡Œãƒ»åˆ†å‰²è¡Œã�®ä¸¡æ–¹** ã�«å�Œã�˜å„ªå…ˆåº¦ãƒ»æ—¥ä»˜ç­‰ã�ŒåŠ¹ã��ã‚ˆã�†ã�«ã�™ã‚‹ã€‚

ã€�åŸºæº–å¹´ï¼ˆå¹´ã�ªã�—æ—¥ä»˜ç”¨ï¼‰ã€‘
ã€Œ4/5ã€�ã€Œ4/5ã�«å‡ºè�·ã€�ã�®ã‚ˆã�†ã�« **å¹´ã�Œç„¡ã�„** æ—¥ä»˜ã�¯åŽŸå‰‡ **è¥¿æš¦ {ref_y} å¹´** ã�¨ã�—ã€�YYYY-MM-DD ã�§å‡ºåŠ›ã€‚

ã€�ãƒ•ã‚£ãƒ¼ãƒ«ãƒ‰ä¸€è¦§ï¼ˆåž‹ã�®å�‚è€ƒï¼‰ã€‘
- process_name, machine_name: æ–‡å­—åˆ—ï¼ˆå¿…é ˆã€‚ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆè¡Œã�¨ä¸€è‡´ï¼‰
- restrict_to_process_name, restrict_to_machine_name: æ–‡å­—åˆ—ï¼ˆä»»æ„�ã€‚é™�å®šã�ªã‚‰ï¼‰
- preferred_operator: æ–‡å­—åˆ—ï¼ˆä¸Šè¨˜å¥‘ç´„ã�«å¾“ã�†ï¼‰
- required_op: æ­£ã�®æ•´æ•°
- speed_override: æ­£ã�®æ•°ï¼ˆm/åˆ†ï¼‰
- task_efficiency: 0ã€œ1
- priority: æ•´æ•°ï¼ˆå°�ã�•ã�„ã�»ã�©å…ˆã�«å‰²ä»˜ï¼‰
- start_date: YYYY-MM-DD / start_time: HH:MM
- target_completion_date, ship_by_date: YYYY-MM-DD

ã€�è§£é‡ˆã�®æŒ‡é‡�ã€‘
- ã€Œé–“ã�«å�ˆã�†ã‚ˆã�†ã�«ã€�ã€Œç¹°ã‚Šä¸Šã�’ã‚‹ã€�â†’ priority ã‚’ä¸Šã�’ã‚‹ï¼ˆæ•°å€¤ã‚’ä¸‹ã�’ã‚‹ï¼‰ã€‚æ—¥ä»˜ã�Œæ–‡ä¸­ã�«ã�‚ã‚Œã�° target_completion_date ã�¾ã�Ÿã�¯ ship_by_date ã�«å…¥ã‚Œã‚‹ã€‚
- æ‹…å½“è€…æŒ‡å��ã�¯ **æ„�å‘³ç�†è§£** ã�§ preferred_operator ã‚’æ±ºã‚�ã‚‹ï¼ˆç‰¹å®šã�®ã‚­ãƒ¼ãƒ¯ãƒ¼ãƒ‰åˆ—æŒ™ã�«é ¼ã‚‰ã�ªã�„ï¼‰ã€‚
- æ•°å€¤ãƒ»æ—¥ä»˜ã�¯æŽ¨æ¸¬ã�§è£œã‚�ã�ªã�„ã€‚
- **å‚™è€ƒã�Œç‰¹å®šã�®å·¥ç¨‹ãƒ»æ©Ÿæ¢°ã�«ã� ã�‘è¨€å�Šã�—ã�¦ã�„ã�ªã�„é™�ã‚Š**ã€�restrict_to_* ã�¯ç©ºã�«ã�—ã€�å�Œä¸€ä¾�é ¼NOã�®ä»–è¡Œã�«ã‚‚é�©ç”¨ã�•ã‚Œã‚‹å½¢ã�«ã�™ã‚‹ã€‚

ã€�å‡ºåŠ›ç›´å‰�ã�®è‡ªå·±æ¤œè¨¼ï¼ˆå¿…ã�šå®Ÿè¡Œã�—ã�¦ã�‹ã‚‰ JSON ã‚’é–‰ã�˜ã‚‹ï¼‰ã€‘
- ã€�ç‰¹åˆ¥æŒ‡å®šåŽŸæ–‡ã€‘ã�® **å�„è¡Œ** ã�«ã�¤ã�„ã�¦ã€�å¯¾å¿œã�™ã‚‹ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã�« **process_name** ã�¨ **machine_name** ã�Œã�‚ã‚‹ã�‹ã€‚
- å�Œä¸€ä¾�é ¼NOã�Œè¤‡æ•°è¡Œã�‚ã‚‹ã�¨ã��ã�¯ **é…�åˆ—** ã�§å�„è¡Œã�«1ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã€�ã�¾ã�Ÿã�¯é�©åˆ‡ã�«ãƒžãƒ¼ã‚¸ã�—ã�Ÿå�˜ä¸€ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆï¼‹restrict ã�®é�‹ç”¨ã‚’ä¸€è²«ã�•ã�›ã‚‹ã€‚
- ã€Œä¸»æ‹…å½“OPã�®æŒ‡æ„�ã€�ã�Œã�‚ã‚‹è¡Œã�§ã�¯ **é�žç©ºã�® preferred_operator** ã‚’ä»˜ã�‘ã‚‹ã€‚

ã€�å‡ºåŠ›å½¢å¼�ã�®ä¾‹ã€‘ï¼ˆä¾�é ¼NOãƒ»å€¤ã�¯å®Ÿãƒ‡ãƒ¼ã‚¿ã�«å�ˆã‚�ã�›æ›¿ã�ˆã‚‹ã�“ã�¨ï¼‰
{{
  "W3-14": {{
    "process_name": "æ¤œæŸ»",
    "machine_name": "ãƒ©ã‚¤ãƒ³A",
    "preferred_operator": "æ£®å²¡"
  }},
  "Y3-26": {{
    "process_name": "ã‚³ãƒ¼ãƒ†ã‚£ãƒ³ã‚°",
    "machine_name": "",
    "priority": 1,
    "ship_by_date": "{ref_y}-04-05",
    "target_completion_date": "{ref_y}-04-05"
  }},
  "Y4-2": {{
    "process_name": "ã‚¨ãƒ³ãƒœã‚¹",
    "machine_name": "E1",
    "priority": 2,
    "restrict_to_process_name": "",
    "restrict_to_machine_name": ""
  }}
}}

ã€�ç‰¹åˆ¥æŒ‡å®šåŽŸæ–‡ã€‘ï¼ˆExcel ã�‹ã‚‰ã��ã�®ã�¾ã�¾ã€‚1è¡Œï¼�ä¾�é ¼NOã�¨å‚™è€ƒã�®ãƒšã‚¢ï¼‰
{blob}
"""
    try:
        ppath = os.path.join(log_dir, "ai_task_special_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: ä»Šå›ž Gemini ã�«æ¸¡ã�—ã�Ÿãƒ—ãƒ­ãƒ³ãƒ—ãƒˆå…¨æ–‡ â†’ %s", ppath)
    except OSError as ex:
        logging.warning("ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆä¿�å­˜å¤±æ•—: %s", ex)

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
            logging.info("ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: AIè§£æž�ã�Œå®Œäº†ã�—ã�¾ã�—ã�Ÿã€‚")
            if ai_sheet_sink is not None:
                ai_sheet_sink["ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ_AI_API"] = "ã�‚ã‚Š"
            return parsed
        if ai_sheet_sink is not None:
            ai_sheet_sink["ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ_AI_API"] = "ã�‚ã‚Šï¼ˆJSONè§£é‡ˆå¤±æ•—ï¼‰"
        return {}
    except Exception as e:
        err_text = str(e)
        is_quota = ("429" in err_text) or ("RESOURCE_EXHAUSTED" in err_text)
        is_unavailable = ("503" in err_text) or ("UNAVAILABLE" in err_text)
        retry_sec = extract_retry_seconds(err_text) if is_quota else None
        if is_quota and retry_sec is not None:
            wait_sec = min(max(retry_sec, 1.0), 90.0)
            logging.warning(f"ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š AI 429ã€‚{wait_sec:.1f}ç§’å¾…æ©Ÿã�—ã�¦å†�è©¦è¡Œã�—ã�¾ã�™ã€‚")
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
                        ai_sheet_sink["ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ_AI_API"] = "ã�‚ã‚Šï¼ˆ429å†�è©¦è¡Œå¾Œï¼‰"
                    return parsed
            except Exception as e2:
                logging.warning(f"ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š AI å†�è©¦è¡Œå¤±æ•—: {e2}")
        elif is_unavailable:
            wait_sec = 8.0
            logging.warning(
                f"ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š AI 503/UNAVAILABLEã€‚{wait_sec:.1f}ç§’å¾…æ©Ÿã�—ã�¦å†�è©¦è¡Œã�—ã�¾ã�™ã€‚"
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
                    logging.info("ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: AIå†�è©¦è¡Œã�§è§£æž�ã�Œå®Œäº†ã�—ã�¾ã�—ã�Ÿã€‚")
                    if ai_sheet_sink is not None:
                        ai_sheet_sink["ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ_AI_API"] = "ã�‚ã‚Šï¼ˆ503å†�è©¦è¡Œå¾Œï¼‰"
                    return parsed
                logging.warning("ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š AI 503å†�è©¦è¡Œ: JSON æŠ½å‡ºã�«å¤±æ•—ã�—ã�¾ã�—ã�Ÿã€‚")
            except Exception as e2:
                logging.warning(f"ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š AI 503å†�è©¦è¡Œå¤±æ•—: {e2}")
        else:
            logging.warning(f"ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š AI ã‚¨ãƒ©ãƒ¼: {e}")
        logging.warning(
            "ã‚¿ã‚¹ã‚¯ç‰¹åˆ¥æŒ‡å®š: AIè§£æž�çµ�æžœã‚’å�–å¾—ã�§ã��ã�ªã�‹ã�£ã�Ÿã�Ÿã‚�ã€�ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã�®é–‹å§‹æ—¥/å„ªå…ˆæŒ‡ç¤ºã�¯å��æ˜ ã�•ã‚Œã�¾ã�›ã‚“ã€‚"
            "ï¼ˆåˆ—ã€ŒåŠ å·¥é–‹å§‹æ—¥_æŒ‡å®šã€�ã€ŒæŒ‡å®šç´�æœŸ_ä¸Šæ›¸ã��ã€�ã�¯å»ƒæ­¢æ¸ˆã�¿ã€‚å‚™è€ƒã�®å†�è¨˜è¼‰ã�¾ã�Ÿã�¯å¾Œã�‹ã‚‰ AI å†�å®Ÿè¡Œã‚’æ¤œè¨Žã�—ã�¦ã��ã� ã�•ã�„ã€‚ï¼‰"
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ_AI_API"] = f"å¤±æ•—: {e}"[:500]
        return {}


def _merge_preferred_operator_cell_and_ai(row, ai_for_tid):
    """Excelã€Œæ‹…å½“OP_æŒ‡å®šã€�ã‚’å„ªå…ˆã�—ã€�ç©ºã�ªã‚‰ AI ã�® preferred_operatorã€‚"""
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
    ãƒ¡ã‚¤ãƒ³ã€Œå†�å„ªå…ˆç‰¹åˆ¥è¨˜è¼‰ã€�ã�® task_preferred_operatorsã€‚
    ã‚­ãƒ¼ã�¯ä¾�é ¼NOï¼ˆå¤§æ–‡å­—ãƒ»å°�æ–‡å­—ã�®å·®ã�¯ç„¡è¦–ï¼‰ã€‚
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
    ä¸Šæ›¸ã��åˆ—ã�¯åŠ å·¥é€Ÿåº¦_ä¸Šæ›¸ã��ãƒ»åŽŸå��æŠ•å…¥æ—¥_ä¸Šæ›¸ã��ç­‰ã�®ã�¿ï¼ˆè¨ˆç”»ã‚·ãƒ¼ãƒˆï¼‰ã€‚ã��ã�®ä»–ã�¯ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒ AI ã�‹ã‚‰ã€‚
    allow_ai_dispatch_priority_from_remark ã�Œ False ã�®ã�¨ã��ã€�AI ã�® required_op / task_efficiency / priority /
    start_date / start_time ã�¯æŽ¡ç”¨ã�—ã�ªã�„ï¼ˆå‚™è€ƒã�«ç´�æœŸç³»æ–‡è¨€ã�Œç„¡ã�„è¡Œå�‘ã�‘ï¼‰ã€‚
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
    ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã�«ä¾�ã‚‹ AI è§£æž�çµ�æžœã�¨ã€�æ˜Žç¤ºã‚»ãƒ«ã�®ä¸¡æ–¹ã�«å€¤ã�Œã�‚ã‚Šé£Ÿã�„é�•ã�†åˆ—ã‚’è¿”ã�™ã€‚
    å‚™è€ƒãƒ»AIã�„ã�šã‚Œã�‹æ¬ ã�‘ã‚‹å ´å�ˆã�¯ç©ºé›†å�ˆã€‚
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
    """Excel è¡Œç•ªå�·(1å§‹ã�¾ã‚Šãƒ»ãƒ˜ãƒƒãƒ€ãƒ¼=1è¡Œç›®) -> çŸ›ç›¾ã�Œã�‚ã�£ã�Ÿåˆ—å��ã�®é›†å�ˆ"""
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
    """æ—¢ã�«é–‹ã�„ã�¦ã�„ã‚‹é…�å�°è¨ˆç”»ã‚·ãƒ¼ãƒˆã�¸ã€�çŸ›ç›¾åˆ—ã�®ç�€è‰²ï¼ˆè–„é»„ãƒªã‚»ãƒƒãƒˆâ†’èµ¤ï¼‰ã‚’é�©ç”¨ã�™ã‚‹ã€‚ä¿�å­˜ã�¯å‘¼ã�³å‡ºã�—å�´ã€‚"""
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
            # ãƒ•ã‚©ãƒ³ãƒˆã�¯ä¸Šæ›¸ã��ã�—ã�ªã�„ï¼ˆãƒ–ãƒƒã‚¯æ—¢å®šãƒ»ãƒ¦ãƒ¼ã‚¶ãƒ¼è¨­å®šã‚’ç¶­æŒ�ï¼‰

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
    log_prefix: str = "æ®µéšŽ2",
) -> bool:
    """
    æ®µéšŽ2å�‘ã�‘: ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆè§£æž�ãƒ–ãƒ­ãƒƒã‚¯ï¼ˆAX:AYï¼‰ã�¨çŸ›ç›¾ãƒ�ã‚¤ãƒ©ã‚¤ãƒˆã‚’ **1å›žã�® load/save** ã�§å��æ˜ ã�™ã‚‹ã€‚
    å¾“æ�¥ã�¯åˆ¥é–¢æ•°ã�§ãƒ–ãƒƒã‚¯ã‚’2å›žé–‹ã�„ã�¦ã�„ã�Ÿã�Ÿã‚�ã€�.xlsm ã�Œå¤§ã��ã�„ç’°å¢ƒã�§å��æ•°ç§’å�˜ä½�ã�®çŸ­ç¸®ã�«ã�ªã‚‹ã€‚
    """
    if not wb_path or not os.path.isfile(wb_path):
        return False
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "%s: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� openpyxl ã�§ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è§£æž�ãƒ»çŸ›ç›¾ç�€è‰²ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚",
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
            "%s: é…�å�°ã‚·ãƒ¼ãƒˆä¸€æ‹¬æ›¸è¾¼ã�®ã�Ÿã‚�ãƒ–ãƒƒã‚¯ã‚’é–‹ã�‘ã�¾ã�›ã‚“: %s",
            log_prefix,
            ex,
        )
        return False
    try:
        if sheet_name not in wb.sheetnames:
            logging.info(
                "%s: ã‚·ãƒ¼ãƒˆ '%s' ã�Œç„¡ã�„ã�Ÿã‚�ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è§£æž�ãƒ»çŸ›ç›¾ç�€è‰²ã‚’ã‚¹ã‚­ãƒƒãƒ—ã€‚",
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
                "%s: é…�å�°ã‚·ãƒ¼ãƒˆã�¸ã�®ä¸€æ‹¬ä¿�å­˜ã�«å¤±æ•—ï¼ˆExcel ã�§é–‹ã�„ã�Ÿã�¾ã�¾ç­‰ï¼‰ã€‚"
                " çŸ›ç›¾ãƒ�ã‚¤ãƒ©ã‚¤ãƒˆã�¯ '%s' ã�«æ›¸ã��å‡ºã�—ã�¾ã�—ã�Ÿã€‚ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è§£æž�ã�¯æœªä¿�å­˜ã�®å�¯èƒ½æ€§ã�Œã�‚ã‚Šã�¾ã�™ã€‚ (%s)",
                log_prefix,
                _planning_conflict_sidecar_path(),
                e,
            )
            return False
        _remove_planning_conflict_sidecar_safe()
        _n_conf = len(conflicts_by_row) if conflicts_by_row else 0
        if _n_conf:
            logging.info(
                "%s: ã€Œ%sã€�%s:%s åˆ—ã�«ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è§£æž�ã‚’ä¿�å­˜ã�—ã€�"
                "ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã�¨åˆ—ã�®çŸ›ç›¾ %s è¡Œã‚’å�Œã�˜ä¿�å­˜ã�§ãƒ�ã‚¤ãƒ©ã‚¤ãƒˆã�—ã�¾ã�—ã�Ÿã€‚",
                log_prefix,
                sheet_name,
                get_column_letter(lc),
                get_column_letter(vc),
                _n_conf,
            )
        else:
            logging.info(
                "%s: ã€Œ%sã€�%s:%s åˆ—ã�«ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è§£æž�ã‚’ä¿�å­˜ã�—ã�¾ã�—ã�Ÿï¼ˆçŸ›ç›¾è¡Œã�ªã�—ï¼‰ã€‚",
                log_prefix,
                sheet_name,
                get_column_letter(lc),
                get_column_letter(vc),
            )
        return True
    except OSError as ex:
        logging.warning(
            "%s: é…�å�°ã‚·ãƒ¼ãƒˆä¸€æ‹¬ä¿�å­˜ã�§ OSError: %s",
            log_prefix,
            ex,
        )
        return False
    except Exception as ex:
        logging.warning(
            "%s: é…�å�°ã‚·ãƒ¼ãƒˆã�¸ã�®ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è§£æž�ï¼‹çŸ›ç›¾ç�€è‰²ï¼ˆä¸€æ‹¬ï¼‰ã�§ä¾‹å¤–: %s",
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
    é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã‚·ãƒ¼ãƒˆã�®ãƒ‡ãƒ¼ã‚¿è¡Œã‚’ã€�çŸ›ç›¾åˆ—ã�®ã�¿èµ¤åœ°ãƒ»ç™½å¤ªå­—ã�«ã�™ã‚‹ã€‚
    äº‹å‰�ãƒ‘ã‚¹ã�§ã�¯ä¸Šæ›¸ã��å…¥åŠ›åˆ—ã‚’æ®µéšŽ1ã�¨å�Œã�˜è–„é»„è‰²ã�«æˆ»ã�—ã€�ãƒ•ã‚©ãƒ³ãƒˆã�¯å¤‰æ›´ã�—ã�ªã�„ï¼ˆä½“è£�ç¶­æŒ�ï¼‰ã€‚
    AIè§£æž�åˆ—ã�¯ç�€è‰²ã�—ã�ªã�„ï¼ˆæ®µéšŽ1ã�®ä»•æ§˜ã�«å�ˆã‚�ã�›ã‚‹ï¼‰ã€‚
    .xlsm ã�¯ keep_vba=True ã�§ä¿�å­˜ã�™ã‚‹ã€‚
    """
    if not wb_path or not os.path.exists(wb_path):
        return
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "çŸ›ç›¾æ›¸å¼�: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� openpyxl ã�§ã�®ãƒ�ã‚¤ãƒ©ã‚¤ãƒˆã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = load_workbook(wb_path, keep_vba=keep_vba)
    try:
        if sheet_name not in wb.sheetnames:
            logging.warning(f"çŸ›ç›¾æ›¸å¼�: ã‚·ãƒ¼ãƒˆ '{sheet_name}' ã�Œè¦‹ã�¤ã�‹ã‚Šã�¾ã�›ã‚“ã€‚")
            return
        ws = wb[sheet_name]
        _plan_sheet_apply_conflict_styles_to_ws(ws, num_data_rows, conflicts_by_row)

        try:
            wb.save(wb_path)
        except OSError as e:
            write_planning_conflict_highlight_sidecar(sheet_name, num_data_rows, conflicts_by_row)
            logging.warning(
                "é…�å�°ã‚·ãƒ¼ãƒˆã�¸ã�®çŸ›ç›¾ãƒ�ã‚¤ãƒ©ã‚¤ãƒˆã‚’ãƒ•ã‚¡ã‚¤ãƒ«ä¿�å­˜ã�§ã��ã�¾ã�›ã‚“ã�§ã�—ã�Ÿï¼ˆExcel ã�§ãƒ–ãƒƒã‚¯ã‚’é–‹ã�„ã�Ÿã�¾ã�¾ç­‰ï¼‰ã€‚"
                " '%s' ã�«æŒ‡ç¤ºã‚’æ›¸ã��å‡ºã�—ã�¾ã�—ã�Ÿã€‚ãƒžã‚¯ãƒ­ã�Œã‚·ãƒ¼ãƒˆä¸Šã�«ç›´æŽ¥é�©ç”¨ã�—ã�¾ã�™ã€‚ (%s)",
                _planning_conflict_sidecar_path(),
                e,
            )
        else:
            _remove_planning_conflict_sidecar_safe()
            if conflicts_by_row:
                logging.info(
                    f"ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã�¨åˆ—ã�®çŸ›ç›¾: {len(conflicts_by_row)} è¡Œã‚’ '{sheet_name}' ã�§ãƒ�ã‚¤ãƒ©ã‚¤ãƒˆã�—ã�¾ã�—ã�Ÿã€‚"
                )
    finally:
        wb.close()


def _ai_planning_target_due_date(ai_dict):
    """AI JSON ã�®å®Œäº†ãƒ»å‡ºè�·ç›®æ¨™æ—¥ã�‹ã‚‰ã€�é…�å�°ã�®ç›®æ¨™æ—¥1ã�¤ã‚’æ±ºã‚�ã‚‹ï¼ˆè¤‡æ•°ã�‚ã‚Œã�°æœ€ã‚‚æ—©ã�„æ—¥ï¼�åŽ³ã�—ã�„æ–¹ï¼‰ã€‚"""
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
    ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã�«ã€�ç´�æœŸãƒ»æœŸé™�ãƒ»æœ€å„ªå…ˆã�ªã�©ã€Œé…�å�°è©¦è¡Œã‚’å‰�ã�«å‡ºã�™ã€�æ„�å›³ã�®æ–‡è¨€ã�Œã�‚ã‚‹ã�¨ã�� Trueã€‚
    å‚™è€ƒã�Œè¨˜å…¥ã�•ã‚Œã�¦ã�„ã‚‹ã� ã�‘ã�§ã�¯ True ã�«ã�—ã�ªã�„ï¼ˆAI ç”±æ�¥ã�®ç›®æ¨™æ—¥ãƒ»é–‹å§‹æ—¥ãƒ»å„ªå…ˆåº¦ã�¯ä½¿ã‚�ã�ªã�„ï¼‰ã€‚
    """
    if not remark_raw:
        return False
    s = str(remark_raw).strip()
    if not s or s.lower() in ("nan", "none"):
        return False
    n = unicodedata.normalize("NFKC", s)
    n_lower = n.casefold()
    needles = (
        "ç´�æœŸ",
        "æŒ‡å®šç´�æœŸ",
        "å›žç­”ç´�æœŸ",
        "è¨ˆç”»åŸºæº–",
        "æœŸæ—¥",
        "ç· åˆ‡",
        "ç· ã‚�åˆ‡ã‚Š",
        "æœŸé™�",
        "æœ€å„ªå…ˆ",
        "è‡³æ€¥",
        "æ€¥ã�Ž",
        "ç›´ã�¡ã�«",
        "æ—©æ€¥",
        "å‡ºè�·",
        "ç´�å…¥",
        "å¿…ç�€",
        "deadline",
        "ãƒ‡ãƒƒãƒ‰ãƒ©ã‚¤ãƒ³",
        "å‰�å€’ã�—",
        "æ—©ã‚�ã�«",
        "åŽ³å®ˆ",
        "ã�¾ã�§ã�«",
        "é–“ã�«å�ˆã‚�",
        "é–“ã�«å�ˆã�„",
        "é�…ã‚Œã�ªã�„",
        "é�…å»¶ä¸�å�¯",
        "å„ªå…ˆé…�å�°",
        "å…ˆã�«é…�å�°",
        "å®Œäº†äºˆå®š",
        "æœ¬ç´�æœŸ",
        "å›žç­”æœŸé™�",
    )
    return any(w.casefold() in n_lower for w in needles)


def _task_id_same_machine_due_tiebreak_key(task_id) -> tuple:
    """
    ç´�æœŸåŸºæº–ï¼ˆå›žç­”â†’æŒ‡å®šï¼‰ãƒ»æ©Ÿæ¢°å��ã�Œå�Œã�˜å¸¯ã�§ã�®è©¦è¡Œé †ã€‚
    Y3-24 ã�¯æœ«å°¾ã�®æ•°å€¤ã€‚Y4-1-1 ã�®ã‚ˆã�†ã�«ãƒ�ã‚¤ãƒ•ãƒ³ã�Œ2ã�¤ä»¥ä¸Šã�‚ã‚‹ã�¨ã��ã�¯ã€Œæœ€åˆ�ã�® - ã�®ç›´å¾Œã€�ã�®æ•°å€¤éƒ¨ã‚’æŽ¡ç”¨ã€‚
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
# é…�å�°ç”¨ã‚¿ã‚¹ã‚¯ã‚­ãƒ¥ãƒ¼
#   é…�å�°è¨ˆç”» DataFrame 1è¡Œ â†’ å‰²ä»˜ã‚¢ãƒ«ã‚´ãƒªã‚ºãƒ ç”¨ dict ã�¸ã�®å¤‰æ�›ï¼ˆå„ªå…ˆåº¦ãƒ»ç´�æœŸãƒ»AI ä¸Šæ›¸ã��ã‚’é›†ç´„ï¼‰
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
    ``generate_plan`` å†…ã�§å‘¼ã�°ã‚Œã‚‹ã€‚å®Œäº†æ¸ˆã�¿ãƒ»é…�å�°ä¸�è¦�è¡Œã‚’é™¤ã��ã€�æ®‹ã‚Šã‚’ task_queue ã�«ç©�ã‚€ã€‚
    ai_by_tid ã�Œ None ã�®ã�¨ã��ã� ã�‘å†…éƒ¨ã�§ analyze_task_special_remarks ã‚’å®Ÿè¡Œã�™ã‚‹ã€‚
    """
    if ai_by_tid is None:
        ai_by_tid = analyze_task_special_remarks(tasks_df, reference_year=run_date.year)
    task_queue = []
    n_exclude_plan = 0
    seq_by_tid = _collect_process_content_order_by_task_id(tasks_df)
    same_tid_line_seq = defaultdict(int)
    # ä¾�é ¼NOç›´åˆ—é…�å�°ã�®é †åº�ç”¨: iterrows ã�®èª­ã�¿è¾¼ã�¿é †ï¼ˆ0 å§‹ã�¾ã‚Šï¼‰ã€‚task_queue.sort å¾Œã‚‚ä¸�å¤‰ã€‚
    planning_sheet_row_seq = 0

    for planning_df_iloc, (_, row) in enumerate(tasks_df.iterrows()):
        if row_has_completion_keyword(row):
            continue
        if _plan_row_exclude_from_assignment(row):
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
        # ç´�æœŸåŸºæº–: â‘ å›žç­”ç´�æœŸï¼ˆç©ºã�§ã�ªã�‘ã‚Œã�°ï¼‰â‘¡åˆ—ã€ŒæŒ‡å®šç´�æœŸã€�ï¼ˆåˆ—ã€ŒæŒ‡å®šç´�æœŸ_ä¸Šæ›¸ã��ã€�ã�¯å»ƒæ­¢æ¸ˆã�¿ï¼‰
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
                "åŽŸå��æŠ•å…¥æ—¥_ä¸Šæ›¸ã��ã‚’æŽ¡ç”¨: ä¾�é ¼NO=%s ã‚·ãƒ¼ãƒˆåŽŸå��æŠ•å…¥æ—¥=%s ä¸Šæ›¸ã��=%s",
                task_id,
                raw_input_sheet,
                raw_input_date_ov,
            )

        qty = max(0.0, qty_total - done_qty)
        speed = parse_float_safe(speed_raw, 1.0)
        if speed <= 0:
            speed = 1.0

        if qty <= 0 or not machine or not task_id:
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
                "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: ä¾�é ¼NO=%s ã�®æ‹…å½“OPã‚’ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«æŒ‡å��ã�§ä¸Šæ›¸ã�� %rï¼ˆã‚»ãƒ«ãƒ»ç‰¹åˆ¥æŒ‡å®šå‚™è€ƒAIã‚ˆã‚Šå„ªå…ˆï¼‰",
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
                "ãƒ¡ã‚¤ãƒ³ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«: ä¾�é ¼NO=%s å·¥ç¨‹=%r æ©Ÿæ¢°å��=%r ã�« speed_multiplier ç´¯ç©�=%s ã‚’é�©ç”¨ï¼ˆé€Ÿåº¦ %s â†’ %sï¼‰",
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

        # ç´�æœŸã�¯å„ªå…ˆé †ä½�ãƒ»ç·Šæ€¥åº¦ã�«ã�¯ä½¿ã�†ã�Œã€�é–‹å§‹æ—¥ã�®ä¸‹é™�ã�«ã�¯ä½¿ã‚�ã�ªã�„ï¼ˆä½™åŠ›ã�Œã�‚ã‚Œã�°å‰�å€’ã�—é–‹å§‹ã�™ã‚‹ã�Ÿã‚�ï¼‰ã€‚
        if due_basis is None:
            due_urgent = False
        else:
            due_urgent = due_basis <= run_date

        # é–‹å§‹æ—¥ãƒ«ãƒ¼ãƒ«:
        # 1) åŽŸå��æŠ•å…¥æ—¥ã�Œã�‚ã‚‹ã�¨ã��ã�¯ã€ŒåŽŸå��æŠ•å…¥æ—¥ 13:00 ä»¥é™�ã€�ã‚’é–‹å§‹å�¯èƒ½æ—¥æ™‚ã�®ä¸‹é™�ã�¨ã�™ã‚‹ã€‚
        #    ï¼ˆæ—¥ä»˜ä¸‹é™�: max(run_date, raw_input_date)ã€�å�Œæ—¥æ™‚é–“ä¸‹é™�: 13:00ï¼‰
        # 2) ç‰¹åˆ¥æŒ‡å®šï¼ˆã‚»ãƒ«/AIï¼‰ã�®é–‹å§‹æ—¥ã�Œã�‚ã‚‹å ´å�ˆã‚‚ã€�åŽŸå��æŠ•å…¥æ—¥ã‚ˆã‚Šå‰�å€’ã�—ã�«ã�¯ã�—ã�ªã�„ï¼ˆdate ä¸‹é™�ã‚’ç¶­æŒ�ï¼‰
        # 3) åŽŸå��ã�Œç„¡ã�„ã�¨ã��ã�¯ run_date
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
                    "é–‹å§‹æ—¥ä¸Šæ›¸ã��ã�¯åŽŸå��æŠ•å…¥æ—¥ã‚ˆã‚Šå‰�å€’ã�—ä¸�å�¯: ä¾�é ¼NO=%s æŒ‡å®šé–‹å§‹æ—¥=%s åŽŸå��æŠ•å…¥æ—¥=%s æŽ¡ç”¨é–‹å§‹æ—¥=%s",
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
                # ç´�æœŸå¾Œã‚�å€’ã�—å†�è©¦è¡Œã�§ due_basis_date ã‚’å†…éƒ¨ +1 ã�—ã�¦ã‚‚ã€�çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®å½“åˆ—ï¼ˆåˆ—å��ã�¯äº’æ�›ã�§ã€Œè¨ˆç”»åŸºæº–ç´�æœŸã€�ï¼‰ã�¯ã�“ã�®å€¤ã�®ã�¾ã�¾
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
        "task_queue æ§‹ç¯‰å®Œäº†: total=%sï¼ˆé…�å�°ä¸�è¦�ã�«ã‚ˆã‚Šã‚¹ã‚­ãƒƒãƒ— %s è¡Œï¼‰",
        len(task_queue),
        n_exclude_plan,
    )
    return task_queue


def _task_id_priority_key(task_id):
    """
    ä¾�é ¼NOã�®å�Œæ�¡ä»¶ã‚¿ã‚¤ãƒ–ãƒ¬ãƒ¼ã‚¯ç”¨ã‚­ãƒ¼ã€‚
    ä¾‹: Y3-24, Y3-34 ã�®ã‚ˆã�†ã�ªå ´å�ˆã�¯ãƒ�ã‚¤ãƒ•ãƒ³å¾Œå�Šã�®æ•°å€¤ã�Œå°�ã�•ã�„æ–¹ã‚’å„ªå…ˆã€‚
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
    ä¾�é ¼NOç›´åˆ—é…�å�°ã�®å‡¦ç�†é †ã€‚å�„ä¾�é ¼NOã�«ã�¤ã�„ã�¦ **é…�å�°è©¦è¡Œé †ç•ªã�®æœ€å°�å€¤** ã�Œå°�ã�•ã�„ä¾�é ¼ã‚’å…ˆã�«å®Œèµ°ã�•ã�›ã‚‹
    ï¼ˆå�Œä¸€ä¾�é ¼å†…ã�®è¤‡æ•°è¡Œã�¯æœ€å°�å¹…ã�®è©¦è¡Œé †ã�§ä»£è¡¨ï¼‰ã€‚ã‚¿ã‚¤ãƒ–ãƒ¬ãƒ¼ã‚¯ã�¯è¨ˆç”»ã‚·ãƒ¼ãƒˆä¸Šã�®å…ˆè¡Œè¡Œ
    ï¼ˆplanning_sheet_row_seqï¼‰ã�¨ä¾�é ¼NOã‚­ãƒ¼ã€‚
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
    æ—¢å­˜ã‚·ãƒ¼ãƒˆï¼ˆread_excelï¼‰ç”±æ�¥ã�®ã‚¹ã‚«ãƒ©ãƒ¼ã‚’ã€�é…�å�°è¨ˆç”» DataFrame ã�®æ–‡å­—åˆ—åˆ—ï¼ˆStringDtypeï¼‰ã�¸
    ä»£å…¥ã�§ã��ã‚‹ str ã�«æ­£è¦�åŒ–ã�™ã‚‹ã€‚Excel ã�Œæ•°å€¤ã�¨ã�—ã�¦ä¿�æŒ�ã�—ã�Ÿå„ªå…ˆåº¦ 1 â†’ \"1\" ã�ªã�©ã€‚
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
    ãƒ–ãƒƒã‚¯å†…ã�®ã€Œé…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€�ã�«ãƒ¦ãƒ¼ã‚¶ãƒ¼ã�Œå…¥åŠ›ã�—ã�Ÿä¸Šæ›¸ã��åˆ—ã‚’ã€�
    æ®µéšŽ1ã�®æŠ½å‡ºçµ�æžœã�¸ (ä¾�é ¼NO, å·¥ç¨‹å��) å�˜ä½�ã�§å¼•ã��ç¶™ã��ã€‚
    ç©ºã�®ã‚»ãƒ«ã�¯ãƒžãƒ¼ã‚¸ã�—ã�ªã�„ï¼ˆæ–°è¦�æŠ½å‡ºå�´ã�®ç©ºã�®ã�¾ã�¾ï¼‰ã€‚
    """
    if out_df is None or out_df.empty:
        return out_df
    if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
        return out_df
    try:
        df_old = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=PLAN_INPUT_SHEET_NAME)
    except Exception as e:
        logging.info("æ®µéšŽ1: æ—¢å­˜ã�®é…�å�°ã‚·ãƒ¼ãƒˆã‚’èª­ã‚�ã�ªã�„ã�Ÿã‚�ä¸Šæ›¸ã��ç¶™æ‰¿ã‚’ã‚¹ã‚­ãƒƒãƒ— (%s)", e)
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
            "æ®µéšŽ1: æ—¢å­˜ã‚·ãƒ¼ãƒˆã�‹ã‚‰ä¸Šæ›¸ã��åˆ—ã‚’ %s è¡Œã�¸å¼•ã��ç¶™ã�Žã�¾ã�—ã�Ÿï¼ˆã‚­ãƒ¼: ä¾�é ¼NO+å·¥ç¨‹å��ï¼‰ã€‚",
            merged_rows,
        )
    return out_df


# ---------------------------------------------------------------------------
# é…�å�°ä¸�è¦�ï¼ˆ2ç³»çµ±ï¼‰
#   (A) DataFrame ä¸Šã�®ãƒ«ãƒ¼ãƒ« â€¦ å�Œä¸€ä¾�é ¼NOÃ—å�Œä¸€æ©Ÿæ¢°ã�§ã€Œåˆ†å‰²ã€�è¡Œã�« yesï¼ˆæ‰‹å…¥åŠ›ã�¯ä¸Šæ›¸ã��ã�—ã�ªã�„ï¼‰
#   (B) ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�â€¦ å·¥ç¨‹+æ©Ÿæ¢°ã�”ã�¨ã�® C/D/E åˆ—ã€�Gemini ã�§ Dâ†’Eã€�
#       ä¿�å­˜ãƒ­ãƒƒã‚¯æ™‚ã�¯ xlwings ã�§ A:E å�ŒæœŸâ†’Save ã�®ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ã�‚ã‚Š
#   ã�„ã�šã‚Œã‚‚ apply_exclude_rules_config_to_plan_df ã�§è¨ˆç”» DataFrame ã�«å��æ˜ ã�•ã‚Œã‚‹ã€‚
# ---------------------------------------------------------------------------

def _auto_exclude_cell_empty_for_autofill(v) -> bool:
    """é…�å�°ä¸�è¦�ã‚»ãƒ«ã�Œæœªå…¥åŠ›ã�®ã�¨ã��ã� ã�‘è‡ªå‹•ã�§ yes ã‚’æ›¸ã��è¾¼ã‚€ã€‚"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return True
    if isinstance(v, str):
        s = str(v).strip()
        return not s or s.lower() in ("nan", "none")
    return False


def _normalize_task_id_for_dup_grouping(raw) -> str:
    """å�Œä¸€ä¾�é ¼NOã�®ã‚°ãƒ«ãƒ¼ãƒ”ãƒ³ã‚°ç”¨ï¼ˆè¡¨è¨˜ã‚†ã‚Œãƒ»è‹±å­—ã�®å¤§å°�ã‚’å¯„ã�›ã‚‹ï¼‰ã€‚"""
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
    """å·¥ç¨‹å��ã�Œã€Œåˆ†å‰²ã€�ï¼ˆç©ºç™½é™¤åŽ»ãƒ»NFKC å¾Œï¼‰ã€‚"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\sã€€]+", "", t)
    return t == "åˆ†å‰²"


def _apply_auto_exclude_bunkatsu_duplicate_machine(
    df: pd.DataFrame, log_prefix: str = "æ®µéšŽ1"
) -> pd.DataFrame:
    """
    å�Œä¸€ä¾�é ¼NOã�Œ2è¡Œä»¥ä¸Šã�‚ã‚Šã€�ã�‹ã�¤ç©ºã�§ã�ªã�„å�Œä¸€æ©Ÿæ¢°å��ã�Œ2è¡Œä»¥ä¸Šã�‚ã‚‹ã‚°ãƒ«ãƒ¼ãƒ—ã�§ã�¯ã€�
    å·¥ç¨‹å��ã�Œã€Œåˆ†å‰²ã€�ã�®è¡Œã�®ã€Œé…�å�°ä¸�è¦�ã€�ã�« yes ã‚’å…¥ã‚Œã‚‹ï¼ˆã‚»ãƒ«ã�Œç©ºã�®ã�¨ã��ã�®ã�¿ï¼‰ã€‚
    æ©Ÿæ¢°å��ã�¯ _normalize_equipment_match_key ã�§é‡�è¤‡åˆ¤å®šã€‚
    """
    if df is None or df.empty:
        return df
    need_cols = (TASK_COL_TASK_ID, TASK_COL_MACHINE, TASK_COL_MACHINE_NAME)
    for c in need_cols:
        if c not in df.columns:
            return df
    if PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns:
        df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = ""
    # read_excel ç­‰ã�§ StringDtype ã�«ã�ªã‚‹ã�¨æ•°å€¤ãƒ»çœŸå�½ã�® .at ä»£å…¥ã�§ TypeError ã�«ã�ªã‚‹ã�Ÿã‚� object ã�«å¯„ã�›ã‚‹
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
        for i in idx_list:
            if not _process_name_is_bunkatsu_for_auto_exclude(df.at[i, TASK_COL_MACHINE]):
                continue
            if not _auto_exclude_cell_empty_for_autofill(
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT]
            ):
                continue
            # åˆ—ã�Œ StringDtype ã�®ã�¨ã�� int ä»£å…¥ã�§ TypeError ã�«ã�ªã‚‹ã�Ÿã‚�æ–‡å­—åˆ—ã�«ã�™ã‚‹ï¼ˆ_plan_row_exclude_from_assignment ã�¯ yes ã‚’çœŸã�¨ã�¿ã�ªã�™ï¼‰
            df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
            n_set += 1

    if n_set:
        logging.info(
            "%s: å�Œä¸€ä¾�é ¼NOã�‹ã�¤å�Œä¸€æ©Ÿæ¢°å��ã�Œè¤‡æ•°è¡Œã�‚ã‚‹ã‚°ãƒ«ãƒ¼ãƒ—ã�§ã€�å·¥ç¨‹å��ã€Œåˆ†å‰²ã€�ã�®è¡Œ %s ä»¶ã�«ã€Œé…�å�°ä¸�è¦�ã€�=yes ã‚’è‡ªå‹•è¨­å®šã�—ã�¾ã�—ã�Ÿã€‚",
            log_prefix,
            n_set,
        )
    return df


def _normalize_process_name_for_rule_match(raw) -> str:
    """å·¥ç¨‹å��ã�®ãƒ«ãƒ¼ãƒ«ç…§å�ˆï¼ˆNFKCãƒ»ç©ºç™½é™¤åŽ»ï¼‰ã€‚"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\sã€€]+", "", t)
    return t


def _exclude_rules_sheet_header_map(ws) -> dict:
    """1è¡Œç›®è¦‹å‡ºã�— â†’ åˆ—ç•ªå�·(1å§‹ã�¾ã‚Š)ã€‚
    openpyxl ã�¯æ–°è¦�ã‚·ãƒ¼ãƒˆç›´å¾Œã�« max_column ã�Œ 0 ã�®ã�¾ã�¾ã�®ã�“ã�¨ã�Œã�‚ã‚Šã€�è¦‹å‡ºã�—ã�Œèª­ã‚�ã�šä¿�å­˜å‰�ã�« return ã�—ã�¦ã�—ã�¾ã�†ã€‚
    ã��ã�®ã�Ÿã‚�æœ€ä½Ž Aï½žE åˆ—ã�¯å¿…ã�šèµ°æŸ»ã�™ã‚‹ã€‚
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
    1è¡Œç›®ã�«æ¨™æº–è¦‹å‡ºã�—ï¼ˆå·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ãƒ»é…�å�°ä¸�è¦�ãƒ»é…�å�°ä¸�èƒ½ãƒ­ã‚¸ãƒƒã‚¯ãƒ»ãƒ­ã‚¸ãƒƒã‚¯å¼�ï¼‰ã�Œã�‚ã‚‹ã�“ã�¨ã‚’ä¿�è¨¼ã�™ã‚‹ã€‚
    æ‰‹å‹•ã�§ç©ºã‚·ãƒ¼ãƒˆã� ã�‘è¿½åŠ ã�—ã�Ÿå ´å�ˆã�¯ A1:E1 ã�Œç©ºã�®ã�Ÿã‚�ã€�ã�“ã�“ã�§æ›¸ã��è¾¼ã‚“ã�§åˆ—ç•ªå�·ã‚’è¿”ã�™ã€‚
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
        "%s: ã€Œ%sã€�ã�®è¦‹å‡ºã�—ã�Œç„¡ã�„ï¼�åˆ—å��ã�Œä¸€è‡´ã�—ã�ªã�„ã�Ÿã‚�ã€�æ¨™æº–ã�®1è¡Œç›®ï¼ˆA1:E1ï¼‰ã‚’è¨­å®šã�—ã�¾ã�—ã�Ÿã€‚",
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
    2 è¡Œç›®ä»¥é™�ã�‹ã‚‰ã€Œç©ºè¡Œã€�ã‚’é™¤ã�„ã�¦ä¸Šã�«è©°ã‚�ã‚‹ï¼ˆå…ƒã�®ä¸¦ã�³ã�¯ç¶­æŒ�ã€�ã‚½ãƒ¼ãƒˆã�—ã�ªã�„ï¼‰ã€‚
    ç©ºè¡Œ: å·¥ç¨‹å��ã�Œç©ºã€�ã�¾ã�Ÿã�¯ Aï½žE ç›¸å½“ã�®5ã‚»ãƒ«ã�Œã�™ã�¹ã�¦ç©ºç™½ç›¸å½“ã€‚
    Returns (æ®‹ã�—ã�Ÿãƒ‡ãƒ¼ã‚¿è¡Œæ•°, å‰Šé™¤ã�—ã�Ÿè¡Œæ•°).
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
                "%s: ã€Œ%sã€�ã�¯æœ‰åŠ¹ã�ªãƒ‡ãƒ¼ã‚¿è¡Œã�Œç„¡ã�‹ã�£ã�Ÿã�Ÿã‚�ã€�ãƒ‡ãƒ¼ã‚¿è¡Œ %s è¡Œã‚’å‰Šé™¤ã�—ã�¾ã�—ã�Ÿã€‚",
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
            "%s: ã€Œ%sã€�ã�‹ã‚‰ç©ºè¡Œã‚’ %s ä»¶å‰Šé™¤ã�—ã€�%s è¡Œã�«è©°ã‚�ã�¾ã�—ã�Ÿï¼ˆä¸¦ã�³é †ã�¯ç¶­æŒ�ï¼‰ã€‚",
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
    """Cåˆ—ã€Œé…�å�°ä¸�è¦�ã€�ã�Œã‚ªãƒ³ï¼ˆã�“ã�®å·¥ç¨‹+æ©Ÿæ¢°ãƒ‘ã‚¿ãƒ¼ãƒ³ã�¯å¸¸ã�«é…�å�°ä¸�è¦�ï¼‰ã€‚"""
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
    return s in ("yes", "true", "1", "y", "ã�¯ã�„", "â—‹", "ã€‡", "â—�")


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
    """åŠ å·¥è¨ˆç”»DATA ã�‹ã‚‰ã€�æ®µéšŽ1ã�¨å�Œã�˜æŠ½å‡ºæ�¡ä»¶ã�§ (å·¥ç¨‹å��, æ©Ÿæ¢°å��) ã�®ä¸€è¦§ï¼ˆé‡�è¤‡é™¤ã��ãƒ»é †åº�ç¶­æŒ�ï¼‰ã€‚"""
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
    """Geminiï¼�Eåˆ—ã�‹ã‚‰å¾—ã�Ÿ dict ã�Œé…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«ã�¨ã�—ã�¦æœ‰åŠ¹ã�‹ã€‚"""
    if not isinstance(o, dict):
        return None
    if int(o.get("version") or 0) != 1:
        return None
    mode = str(o.get("mode") or "").strip().lower()
    if mode not in ("always_exclude", "conditions"):
        return None
    return o


def _exclude_rule_de_cache_key(stripped_blob: str) -> str:
    """ã€Œé…�å�°ä¸�èƒ½ãƒ­ã‚¸ãƒƒã‚¯ã€�æ–‡è¨€ï¼ˆæ­£è¦�åŒ–æ¸ˆã�¿ï¼‰ã�«å¯¾ã�™ã‚‹ ai_remarks_cache ç”¨ã‚­ãƒ¼ã€‚"""
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
        "ã€�ã‚¹ã‚­ãƒ¼ãƒž version ã�¯å¿…ã�š 1ã€‘\n"
        "1) å¸¸ã�«é…�å�°ä¸�è¦�ï¼ˆèª¬æ˜Žã�Œæ�¡ä»¶ã�ªã�—ã�§å¤–ã�™æ„�å‘³ï¼‰ã�®ã�¨ã��:\n"
        '{"version":1,"mode":"always_exclude"}\n\n'
        "2) åˆ—ã�®æ�¡ä»¶ã�§é…�å�°ä¸�è¦�ã�¨ã�™ã‚‹ã�¨ã��:\n"
        '{"version":1,"mode":"conditions","require_all": true ã�¾ã�Ÿã�¯ false,"conditions":[ ... ]}\n\n'
        "conditions ã�®å�„è¦�ç´ :\n"
        "- {\"column\":\"åˆ—å��\",\"op\":\"empty\"} â€¦ ã‚»ãƒ«ã�Œç©º\n"
        "- {\"column\":\"åˆ—å��\",\"op\":\"not_empty\"}\n"
        "- {\"column\":\"åˆ—å��\",\"op\":\"eq\",\"value\":\"æ–‡å­—åˆ—\"} / ne / contains / not_contains / regexï¼ˆæ­£è¦�è¡¨ç�¾ï¼‰\n"
        "- {\"column\":\"åˆ—å��\",\"op\":\"gt\"|\"gte\"|\"lt\"|\"lte\",\"value\":æ•°å€¤} â€¦ æ•°å€¤æ¯”è¼ƒï¼ˆåˆ—ã�¯æ•°ã�¨ã�—ã�¦è§£é‡ˆï¼‰\n\n"
        f"ã€�ä½¿ç”¨å�¯èƒ½ã�ªåˆ—å��ã�®ã�¿ã€‘ï¼ˆã�“ã‚Œä»¥å¤–ã�¯ä½¿ã‚�ã�ªã�„ï¼‰:\n{allowed}\n"
    )


def _parse_exclude_rule_json_array_response(text: str) -> list | None:
    """ãƒ¢ãƒ‡ãƒ«å¿œç­”ã�‹ã‚‰ JSON é…�åˆ—ã‚’å�–ã‚Šå‡ºã�™ï¼ˆ```json ãƒ•ã‚§ãƒ³ã‚¹ä»˜ã��å�¯ï¼‰ã€‚"""
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
        logging.warning("é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«: æœªå¯¾å¿œã�®åˆ—å��ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿ: %s", col)
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
    Eåˆ—ã�® JSONï¼ˆversion=1ï¼‰ã‚’è©•ä¾¡ã�—ã€�å½“è©²ã‚¿ã‚¹ã‚¯è¡Œã‚’é…�å�°ä¸�è¦�ã�¨ã�™ã�¹ã��ã�ªã‚‰ Trueã€‚
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
    Dåˆ—ã�®è‡ªç„¶è¨€èªžã‚’ Gemini ã�§ JSON ãƒ«ãƒ¼ãƒ«ã�«å¤‰æ�›ã€‚å¤±æ•—æ™‚ Noneã€‚
    json/ai_remarks_cache.json ã�« TTL ä»˜ã��ã�§ã‚­ãƒ£ãƒƒã‚·ãƒ¥ï¼ˆå�Œä¸€æ–‡è¨€ã�ªã‚‰ API ã‚’å‘¼ã�°ã�ªã�„ï¼‰ã€‚
    """
    blob = str(natural_language or "").strip()
    if not blob:
        return None
    ai_cache = load_ai_cache()
    hit = _cache_get_exclude_rule_de_parsed(ai_cache, blob)
    if hit is not None:
        logging.info("é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«: AIã‚­ãƒ£ãƒƒã‚·ãƒ¥ãƒ’ãƒƒãƒˆï¼ˆé…�å�°ä¸�èƒ½ãƒ­ã‚¸ãƒƒã‚¯â†’JSONï¼‰")
        return hit
    if not API_KEY:
        return None
    schema = _exclude_rule_logic_gemini_schema_instructions()
    prompt = (
        "ã�‚ã�ªã�Ÿã�¯å·¥å ´ã�®é…�å�°ã‚·ã‚¹ãƒ†ãƒ ç”¨ã�§ã�™ã€‚æ¬¡ã�®ã€Œé…�å�°ä¸�èƒ½ã�®èª¬æ˜Žã€�ã‚’ã€�ã‚¿ã‚¹ã‚¯1è¡Œã‚’åˆ¤å®šã�™ã‚‹æ©Ÿæ¢°å�¯èª­ãƒ«ãƒ¼ãƒ«ã�«å¤‰æ�›ã�—ã�¦ã��ã� ã�•ã�„ã€‚\n\n"
        "ã€�å‡ºåŠ›ã€‘å…ˆé ­ã�Œ { ã�§çµ‚ã‚�ã‚Šã�Œ } ã�® JSON ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆ1ã�¤ã�®ã�¿ï¼ˆèª¬æ˜Žãƒ»ãƒžãƒ¼ã‚¯ãƒ€ã‚¦ãƒ³ç¦�æ­¢ï¼‰ã€‚\n\n"
        f"{schema}\n"
        f"ã€�èª¬æ˜Žæ–‡ã€‘\n{blob}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«: ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆ â†’ %s", ppath)
    except OSError as ex:
        logging.warning("é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«: ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆä¿�å­˜å¤±æ•—: %s", ex)
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
        logging.warning("é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«: Gemini å¤‰æ�›å¤±æ•—: %s", e)
        return None


def _ai_compile_exclude_rule_logics_batch(blobs: list[str]) -> list[dict | None]:
    """
    è¤‡æ•°ã�® D åˆ—æ–‡è¨€ã‚’ 1 å›žã�® Gemini å‘¼ã�³å‡ºã�—ã�§ JSON åŒ–ã€‚å¤±æ•—ãƒ»è¦�ç´ æ•°ä¸�ä¸€è‡´æ™‚ã�¯ 1 ä»¶ã�šã�¤ã�«ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ã€‚
    json/ai_remarks_cache.json ã�«ãƒ’ãƒƒãƒˆã�—ã�Ÿæ–‡è¨€ã�¯ API ã‚’å‘¼ã�°ã�ªã�„ã€‚
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
            "é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«: AIã‚­ãƒ£ãƒƒã‚·ãƒ¥ã�®ã�¿ã�§ Dâ†’E ãƒ�ãƒƒãƒ� %s ä»¶ã‚’å®Œçµ�ï¼ˆAPI å‘¼ã�³å‡ºã�—ã�ªã�—ï¼‰ã€‚",
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
        "ã�‚ã�ªã�Ÿã�¯å·¥å ´ã�®é…�å�°ã‚·ã‚¹ãƒ†ãƒ ç”¨ã�§ã�™ã€‚ä»¥ä¸‹ã�® N å€‹ã�®ã€Œé…�å�°ä¸�èƒ½ã�®èª¬æ˜Žã€�ã‚’ã€�ä¸Žã�ˆã�Ÿé †åº�ã�§ã��ã‚Œã�žã‚Œ JSON ãƒ«ãƒ¼ãƒ«ã�«å¤‰æ�›ã�—ã�¦ã��ã� ã�•ã�„ã€‚\n\n"
        f"ã€�å‡ºåŠ›ã€‘JSON é…�åˆ—ã�®ã�¿ã€‚å…ˆé ­ã�Œ [ ã�§çµ‚ã‚�ã‚Šã�Œ ] ã€‚è¦�ç´ æ•°ã�¯å¿…ã�š {m}ï¼ˆMarkdownãƒ»èª¬æ˜Žç¦�æ­¢ï¼‰ã€‚\n"
        f"é…�åˆ—ã�®å…ˆé ­è¦�ç´ ã�Œ [1]ã€�2 ç•ªç›®ã�Œ [2] â€¦ ã�«å¯¾å¿œã�—ã�¾ã�™ã€‚\n\n"
        f"{schema}\n"
        f"ã€�èª¬æ˜Žæ–‡ã€‘\n{numbered}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_batch_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«(ãƒ�ãƒƒãƒ�): ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆ â†’ %s", ppath)
    except OSError as ex:
        logging.warning("é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«(ãƒ�ãƒƒãƒ�): ãƒ—ãƒ­ãƒ³ãƒ—ãƒˆä¿�å­˜å¤±æ•—: %s", ex)
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
                "é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«: ãƒ�ãƒƒãƒ�å¿œç­”ã�Œä¸�æ­£ï¼ˆè¦�ç´ æ•° %sã€�æœŸå¾… %sï¼‰ã€‚1 ä»¶ã�šã�¤å†�è©¦è¡Œã�—ã�¾ã�™ã€‚",
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
        logging.warning("é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«: ãƒ�ãƒƒãƒ� Gemini å¤±æ•—ã€�å�˜ç™ºã�«ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯: %s", e)
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
    ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�ã�®ä¿�å®ˆå‡¦ç�†ã�®ã‚¤ãƒ™ãƒ³ãƒˆãƒ­ã‚°ã€‚

    è¨­å®šã‚·ãƒ¼ãƒˆå‡¦ç�†ã�®æˆ�å�¦ã‚’ log/exclude_rules_sheet_debug.txt ã�«è¿½è¨˜ã�—ã€�execution_log ã�«ã‚‚ã‚¿ã‚°ä»˜ã��ã�§å‡ºåŠ›ã�™ã‚‹ã€‚
    event ä¾‹: START, OPEN_OK, OPEN_RETRY, OPEN_FAIL, HEADER_FIX, SYNC_ROWS, OPENPYXL_SAVE_OK, OPENPYXL_SAVE_FAIL,
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
        logging.warning("exclude_rules_sheet_debug.txt ã�¸æ›¸ã�‘ã�¾ã�›ã‚“: %s", wex)

    tag = "[è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹]"
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
    """ãƒ‡ã‚£ã‚¹ã‚¯ãƒ‘ã‚¹ã�¨ xlwings Book.full_name ã�Œå�Œä¸€ãƒ•ã‚¡ã‚¤ãƒ«ã‚’æŒ‡ã�™ã�‹ï¼ˆè¡¨è¨˜ã‚†ã‚Œã‚’å¤šå°‘å�¸å�Žï¼‰ã€‚"""
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
    """èµ·å‹•ä¸­ã�® Excel ã�‹ã‚‰ãƒ‘ã‚¹ä¸€è‡´ã�™ã‚‹ xlwings Book ã‚’è¿”ã�™ã€‚ç„¡ã�‘ã‚Œã�° Noneã€‚"""
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
    """æ—¢å­˜ã�® Excel.App ã�§ Workbooks.Open ã‚’è©¦ã�™ã€‚æˆ�åŠŸæ™‚ Bookã€�å¤±æ•—æ™‚ Noneã€‚"""
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
    """å°‚ç”¨èµ·å‹•ã�—ã�Ÿ Excel ã�¯çµ‚äº†ã�™ã‚‹ã€‚å®Ÿè¡Œä¸­ Excel ã�§ã� ã�‘ Open ã�—ã�Ÿãƒ–ãƒƒã‚¯ã�¯å¤±æ•—æ™‚ã�®ã�¿é–‰ã�˜ã‚‹ã€‚"""
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
    ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã‚’ xlwings ã�§å�–å¾—ã�™ã‚‹ï¼ˆæœ¬ç•ªãƒ»ãƒ†ã‚¹ãƒˆå…±é€šï¼‰ã€‚
    æˆ»ã‚Šå€¤: (Book, release_info) / å¤±æ•—æ™‚ Noneã€‚
    release_info: mode ã�Œ keep ã�¾ã�Ÿã�¯ quit_excelã€�opened_wb_here ã�Œ boolã€‚
    """
    try:
        import xlwings as xw  # noqa: F401
    except ImportError:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_UNAVAILABLE",
            log_prefix,
            "xlwings ã�Œ import ã�§ã��ã�¾ã�›ã‚“ï¼ˆpip install xlwings ã‚’ç¢ºèª�ï¼‰ã€‚",
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
            "xlwings ã�§ãƒ–ãƒƒã‚¯ã‚’é–‹ã�‘ã�¾ã�›ã‚“ã�§ã�—ã�Ÿã€‚",
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
    æ¤œè¨¼ã‚¹ã‚¯ãƒªãƒ—ãƒˆç”¨: èµ·å‹•ä¸­ãƒ–ãƒƒã‚¯ã‚’å„ªå…ˆã�—ã€�å¿…è¦�ã�ªã‚‰è¡¨ç¤ºä»˜ã�� Excel ã�§é–‹ã��ã€‚
    æˆ»ã‚Šå€¤: (Book, info, èª¬æ˜Žæ–‡å­—åˆ—) ã�¾ã�Ÿã�¯ Noneã€‚
    """
    abs_path = os.path.abspath(book_path)
    book = _xlwings_find_book_on_running_instances(abs_path)
    if book is not None:
        return book, {"mode": "keep", "opened_wb_here": False}, f"{label}:æ—¢å­˜ã‚¤ãƒ³ã‚¹ã‚¿ãƒ³ã‚¹"
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
    """VBA å�´ã�®ã‚¹ãƒ—ãƒ©ãƒƒã‚·ãƒ¥ãƒ�ãƒ¼ãƒªãƒ³ã‚°ã�¨ç«¶å�ˆã�—ã�«ã��ã��ã�™ã‚‹ã�Ÿã‚�ã€�å�ŒæœŸãƒ»ä¿�å­˜ã�®çŸ­æ™‚é–“ã� ã�‘ Excel ã‚’é�™ã�‹ã�«ã�™ã‚‹ã€‚"""
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
    openpyxl ã�§ä¿�å­˜ã�§ã��ã�ªã�„ã�¨ã��ã€�xlwings ã�§ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�A:E ã‚’ãƒ¡ãƒ¢ãƒªä¸Šã�®å€¤ã�§ä¸Šæ›¸ã��ã�— Saveã€‚

    è¡¨ç¤ºä¸­ã‚·ãƒ¼ãƒˆã�«å¯¾ã�™ã‚‹ä¸€æ‹¬ .value ã� ã�‘ã� ã�¨ã€�ã‚¹ãƒ—ãƒ©ãƒƒã‚·ãƒ¥ï¼‹ãƒ�ãƒ¼ãƒªãƒ³ã‚°ï¼ˆD3=trueï¼‰ä¸‹ã�§
    Range ä»£å…¥ã�Œæ•°åˆ†ã�‹ã�‹ã‚‹è¨ˆæ¸¬ã�Œã�‚ã‚Šå¾—ã‚‹ã€‚å�ŒæœŸä¸­ã�®ã�¿ã‚·ãƒ¼ãƒˆã‚’ä¸€æ™‚é�žè¡¨ç¤ºã�«ã�— api.Value2 ã�§æ›¸ã��ã€‚
    """
    global _exclude_rules_effective_read_path

    attached = _xlwings_attach_open_macro_workbook(wb_path, log_prefix)
    if attached is None:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_SKIP",
            log_prefix,
            "xlwings ã�§ãƒ–ãƒƒã‚¯ã�«æŽ¥ç¶šã�§ã��ã�š A:E å�ŒæœŸã‚’ã‚¹ã‚­ãƒƒãƒ—ã€‚",
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
        # å…¨ã‚·ãƒ¼ãƒˆå��ã‚’åˆ—æŒ™ã�™ã‚‹ã�¨ã‚·ãƒ¼ãƒˆæ•°åˆ†ã�® COM å¾€å¾©ã�«ã�ªã‚Šã€�D3=true æ™‚ã�¯ VBA ãƒ�ãƒ¼ãƒªãƒ³ã‚°ã�¨ç«¶å�ˆã�—ã�¦
        # 1 ã‚·ãƒ¼ãƒˆæ•°ç§’ã€œå��æ•°ç§’ã�‹ã�‹ã‚‹ã�“ã�¨ã�Œã�‚ã‚‹ï¼ˆè¨ˆæ¸¬ã�§ 40 ã‚·ãƒ¼ãƒˆâ‰ˆ213sï¼‰ã€‚å��å‰�ã�§ç›´æŽ¥è§£æ±ºã�™ã‚‹ã€‚
        try:
            sht = xw_book.sheets[EXCLUDE_RULES_SHEET_NAME]
        except Exception:
            _log_exclude_rules_sheet_debug(
                "XLWINGS_SYNC_SKIP",
                log_prefix,
                f"xlwings å�´ã�«ã‚·ãƒ¼ãƒˆã€Œ{EXCLUDE_RULES_SHEET_NAME}ã€�ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚",
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
                    sht.api.Visible = 0  # xlSheetHiddenï¼ˆå�ŒæœŸä¸­ã� ã�‘ã€‚å†�æ��ç”»ãƒ»ã‚¦ã‚£ãƒ³ãƒ‰ã‚¦æ›´æ–°è² è�·ã‚’æŠ‘ã�ˆã‚‹ï¼‰
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
            "xlwings çµŒç”±ã�§è¨­å®šã‚·ãƒ¼ãƒˆ Aã€œE ã‚’å�ŒæœŸã�—ãƒ–ãƒƒã‚¯ã‚’ä¿�å­˜ã�—ã�¾ã�—ã�Ÿã€‚",
            details=f"path={wb_path} rows={max_r}",
        )
        logging.info(
            "%s: è¨­å®šã‚·ãƒ¼ãƒˆã‚’ xlwings ã�§ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã�«ä¿�å­˜ã�—ã�¾ã�—ã�Ÿï¼ˆAã€œEï¼‰ã€‚",
            log_prefix,
        )
        return True
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_FAIL",
            log_prefix,
            "xlwings ã�§ã�® A:E å�ŒæœŸã�¾ã�Ÿã�¯ Save ã�«å¤±æ•—ã�—ã�¾ã�—ã�Ÿã€‚",
            details=f"path={wb_path}",
            exc=ex,
        )
        return False
    finally:
        _xlwings_release_book_after_mutation(xw_book, info, ok)


# è¨­å®šã‚·ãƒ¼ãƒˆã�®åˆ—ç¯„å›²ï¼ˆAã€œEï¼‰ã€‚xlwings å�ŒæœŸãƒ»VBA è¡Œåˆ— TSV å‡ºåŠ›ã�§ã‚‚ä½¿ç”¨ã€‚
EXCLUDE_RULES_SHEET_COM_SYNC_MAX_COL = 5
EXCLUDE_RULES_MATRIX_CLIP_MAX_COL = 5


def _persist_exclude_rules_workbook(_wb, wb_path: str, ws, log_prefix: str) -> bool:
    """
    è¨­å®šã‚·ãƒ¼ãƒˆã�®ãƒ‡ã‚£ã‚¹ã‚¯å��æ˜ ã€‚æ—¢å®šã�¯ xlwings ã�§ A:E å�ŒæœŸâ†’Saveï¼ˆEXCLUDE_RULES_TRY_OPENPYXL_SAVE=1 ã�®ã�¨ã��ã�®ã�¿ openpyxl save ã‚’è©¦è¡Œï¼‰ã€‚
    ä¿�å­˜ã�§ã��ã�ªã�„ã�¨ã��ã�¯ log ã�«è¡Œåˆ— TSV ã‚’å‡ºã�—ã€�VBAã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹_Aã�‹ã‚‰E_TSVã�‹ã‚‰å��æ˜ ã€�ã�§å��æ˜ ã�™ã‚‹ã€‚

    _wb â€¦ ç·¨é›†æ¸ˆã�¿ openpyxl ãƒ–ãƒƒã‚¯ï¼ˆopenpyxl çµŒè·¯æ™‚ã�®ã�¿ save ã�«ä½¿ç”¨ï¼‰ã€‚
    """
    global _exclude_rules_effective_read_path

    def _openpyxl_persist_ok(which: str) -> bool:
        try:
            _wb.save(wb_path)
        except Exception as ex:
            _log_exclude_rules_sheet_debug(
                "OPENPYXL_SAVE_FAIL",
                log_prefix,
                f"openpyxl ã�§ã�® .xlsm ä¿�å­˜ã�«å¤±æ•—ã�—ã�¾ã�—ã�Ÿ {which}ï¼ˆExcel ã�§é–‹ã��ã�£ã�±ã�ªã�—ãƒ»ãƒ­ãƒƒã‚¯ã�®å�¯èƒ½æ€§ï¼‰ã€‚",
                details=f"path={wb_path}",
                exc=ex,
            )
            return False
        _exclude_rules_effective_read_path = wb_path
        _clear_exclude_rules_e_apply_files()
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_OK",
            log_prefix,
            "openpyxl ã�§è¨­å®šã‚·ãƒ¼ãƒˆã‚’å�«ã‚€ãƒ–ãƒƒã‚¯ã‚’ä¿�å­˜ã�—ã�¾ã�—ã�Ÿï¼ˆAã€œEï¼‰ã€‚",
            details=f"path={wb_path} {which}",
        )
        logging.info(
            "%s: è¨­å®šã‚·ãƒ¼ãƒˆã‚’ openpyxl ã�§ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã�«ä¿�å­˜ã�—ã�¾ã�—ã�Ÿã€‚%s",
            log_prefix,
            which,
        )
        return True

    saved_openpyxl = False
    if EXCLUDE_RULES_SKIP_OPENPYXL_SAVE:
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_SKIPPED_EXCLUDE_RULES_POLICY",
            log_prefix,
            "è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã�®ä¿�å­˜ã�§ã�¯ openpyxl save ã‚’è©¦è¡Œã�—ã�¾ã�›ã‚“ï¼ˆxlwings å�ŒæœŸã‚’å…ˆè¡Œã€‚å†�è©¦è¡Œã�™ã‚‹å ´å�ˆã�¯ EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1ï¼‰ã€‚",
            details=f"path={wb_path}",
        )
        logging.info(
            "%s: è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã�¯ openpyxl ã‚’è©¦ã�•ã�š xlwings å�ŒæœŸâ†’Save ã‚’è©¦ã�¿ã�¾ã�™ï¼ˆä¸�å�¯ã�ªã‚‰ VBA ç”¨è¡Œåˆ— TSVï¼‰ã€‚",
            log_prefix,
        )
    elif not _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "%s: è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã�¯ openpyxl ã�§ä¿�å­˜ã�—ã�¾ã�™ï¼ˆä¸�å�¯ã�®ã�¨ã��ã�¯ xlwings å�ŒæœŸâ†’Saveã€�ã��ã‚Œã‚‚ä¸�å�¯ã�ªã‚‰ VBA ç”¨è¡Œåˆ— TSVï¼‰ã€‚",
            log_prefix,
        )
        labels = ("(1/4)", "(2/4)", "(3/4)", "(4/4)")
        for i, label in enumerate(labels):
            if i:
                _log_exclude_rules_sheet_debug(
                    "OPENPYXL_RETRY_WAIT",
                    log_prefix,
                    f"openpyxl å†�ä¿�å­˜ã�¾ã�§ 2 ç§’å¾…ã�¡ã�¾ã�™ {label}ã€‚",
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
            f"ãƒ–ãƒƒã‚¯ã�«ã€Œ{OPENPYXL_INCOMPATIBLE_SHEET_MARKER}ã€�ã�Œã�‚ã‚‹ã�Ÿã‚� openpyxl ã�§ã�®ä¿�å­˜ã‚’è©¦ã�¿ã�¾ã�›ã‚“ã€‚",
            details=f"path={wb_path}",
        )
        logging.info(
            "%s: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� openpyxl save ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã€�xlwings ã�¾ã�Ÿã�¯è¡Œåˆ— TSV ã�«åˆ‡ã‚Šæ›¿ã�ˆã�¾ã�™ã€‚",
            log_prefix,
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )

    if saved_openpyxl:
        return True

    if _xlwings_sync_exclude_rules_sheet_from_openpyxl(wb_path, ws, log_prefix):
        return True

    if _write_exclude_rules_matrix_vba_tsv(wb_path, ws, log_prefix):
        logging.warning(
            "%s: è¨­å®šã‚·ãƒ¼ãƒˆã‚’ log\\%s ã�«å‡ºåŠ›ã�—ã�¾ã�—ã�Ÿã€‚"
            " Excel ã�§ãƒžã‚¯ãƒ­ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹_Aã�‹ã‚‰E_TSVã�‹ã‚‰å��æ˜ ã€�ã‚’å®Ÿè¡Œã�—ã�¦ã��ã� ã�•ã�„ã€‚",
            log_prefix,
            EXCLUDE_RULES_MATRIX_VBA_FILENAME,
        )

    _log_exclude_rules_sheet_debug(
        "OPENPYXL_VBA_FALLBACK",
        log_prefix,
        "openpyxl ä¿�å­˜ã�«å¤±æ•—ã�—ã�Ÿã�Ÿã‚� VBA ç”¨è¡Œåˆ— TSV ã‚’å‡ºåŠ›ã�—ã�¾ã�—ã�Ÿï¼ˆãƒ–ãƒƒã‚¯ã�¯ Excel ä¸Šã�§æ‰‹å‹•å��æ˜ ã�Œå¿…è¦�ã�ªå ´å�ˆã�Œã�‚ã‚Šã�¾ã�™ï¼‰ã€‚",
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
    """VBA ç”¨: è¨­å®šã‚·ãƒ¼ãƒˆ 1 è¡Œç›®ã€œ max_row ã�® Aã€œE ã‚’ Base64(UTF-8) ä»˜ã�� TSV ã�§å‡ºåŠ›ã�™ã‚‹ã€‚"""
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
            "è¨­å®šã‚·ãƒ¼ãƒˆ Aã€œE ã‚’ VBA å��æ˜ ç”¨ TSV ã�«æ›¸ã��å‡ºã�—ã�¾ã�—ã�Ÿï¼ˆopenpyxl ä¿�å­˜ä¸�å�¯æ™‚ï¼‰ã€‚",
            details=f"path={path} rows={max_r}",
        )
        return True
    except OSError as ex:
        logging.warning("%s: è¡Œåˆ— VBA ç”¨ TSV ã‚’æ›¸ã�‘ã�¾ã�›ã‚“: %s", log_prefix, ex)
        return False


def _build_exclude_rules_list_from_openpyxl_ws(
    ws, c_proc: int, c_mach: int, c_flag: int, c_e: int
) -> list[dict]:
    """openpyxl ä¸Šã�®è¨­å®šã‚·ãƒ¼ãƒˆã�‹ã‚‰ _load_exclude_rules_from_workbook ã�¨å�Œå½¢ã�®ãƒªã‚¹ãƒˆã‚’æ§‹ç¯‰ã€‚"""
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
    """VBA ç”¨: è¡Œç•ªå�·ã�¨ Base64(UTF-8) ã‚»ãƒ«æ–‡å­—åˆ—ã�® TSVã€‚"""
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
            "E åˆ—ã‚’ VBA å��æ˜ ç”¨ TSV ã�«æ›¸ã��å‡ºã�—ã�¾ã�—ã�Ÿï¼ˆä¿�å­˜å¤±æ•—æ™‚ã�®ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ç”¨ï¼‰ã€‚",
            details=f"path={path_tsv} cells={len(cells)}",
        )
    except OSError as ex:
        logging.warning("%s: E åˆ— VBA ç”¨ TSV ã‚’æ›¸ã�‘ã�¾ã�›ã‚“: %s", log_prefix, ex)


def _write_exclude_rules_e_apply_artifacts(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> None:
    """
    E åˆ—ï¼ˆé�žç©ºï¼‰ã‚’ JSON ã‚µã‚¤ãƒ‰ã‚«ãƒ¼ãƒ‰ã�¨ VBA ç”¨ TSV ã�«æ›¸ã��ã€‚ç©ºã�ªã‚‰ä¸¡ãƒ•ã‚¡ã‚¤ãƒ«ã‚’å‰Šé™¤ã€‚
    Python æ¬¡å›žèµ·å‹•æ™‚ã�® E å¾©å…ƒç”¨ JSON ã�¨ã€�ãƒžã‚¯ãƒ­ã�‹ã‚‰ã�® E æ›¸è¾¼ã�¿ç”¨ TSVã€‚
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
        logging.warning("%s: E åˆ— JSON ã‚’æ›¸ã�‘ã�¾ã�›ã‚“: %s", log_prefix, ex)
    _write_exclude_rules_e_vba_tsv_from_cells(wb_path, c_e, cells, log_prefix)
    _log_exclude_rules_sheet_debug(
        "E_APPLY_FILES_WRITTEN",
        log_prefix,
        "E åˆ—ã‚’ JSON ã�¨ VBA ç”¨ TSV ã�«æ›¸ã��å‡ºã�—ã�¾ã�—ã�Ÿï¼ˆãƒžã‚¯ãƒ­ã�§ E åˆ—ã‚’å��æ˜ å¾Œã€�ãƒ•ã‚¡ã‚¤ãƒ«å‰Šé™¤ï¼‰ã€‚",
        details=f"cells={len(cells)}",
    )


def _try_apply_pending_exclude_rules_e_column(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> int:
    """
    å‰�å›žä¿�å­˜ã�«å¤±æ•—ã�—ã�Ÿã�¨ã��æ›¸ã��å‡ºã�—ã�Ÿ JSON ã�‹ã‚‰ E åˆ—ã‚’å¾©å…ƒã�™ã‚‹ã€‚
    ãƒ–ãƒƒã‚¯ãƒ‘ã‚¹ã�Œä¸€è‡´ã�—ã�ªã�‘ã‚Œã�°ä½•ã‚‚ã�—ã�ªã�„ã€‚é�©ç”¨å¾Œã�¯ã‚µã‚¤ãƒ‰ã‚«ãƒ¼ãƒ‰ã‚’å‰Šé™¤ã�™ã‚‹ã€‚
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
            f"æœªä¿�å­˜ã� ã�£ã�Ÿ E åˆ—ã‚’ã‚µã‚¤ãƒ‰ã‚«ãƒ¼ãƒ‰ã�‹ã‚‰ {n} ã‚»ãƒ«å¾©å…ƒã�—ã�¾ã�—ã�Ÿã€‚",
            details=path_sc,
        )
        logging.info(
            "%s: %s ã�®å†…å®¹ã‚’ã‚·ãƒ¼ãƒˆã�®ãƒ­ã‚¸ãƒƒã‚¯å¼�åˆ—ã�¸é�©ç”¨ã�—ã�¾ã�—ã�Ÿï¼ˆç¶šã�‘ã�¦ä¿�å­˜ã‚’è©¦ã�¿ã�¾ã�™ï¼‰ã€‚",
            log_prefix,
            path_sc,
        )
    return n


def _read_exclude_rules_d_cells_data_only_for_rows(
    wb_path: str, rows: list[int], c_d: int
) -> dict[int, object]:
    """
    D åˆ—ã�Œæ•°å¼�ã�®ã�¨ã��ã€�openpyxl ã�®é€šå¸¸èª­è¾¼ã�§ã�¯ '=...' ã�—ã�‹å�–ã‚Œã�ªã�„ã€‚
    data_only=True ã�§ã‚­ãƒ£ãƒƒã‚·ãƒ¥å€¤ã‚’èª­ã‚€ï¼ˆExcel ã�Œä¸€åº¦ã�§ã‚‚ä¿�å­˜ãƒ»è¨ˆç®—æ¸ˆã�¿ã�®ãƒ–ãƒƒã‚¯ã�§æœ‰åŠ¹ï¼‰ã€‚
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
    ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�ã�®è¡Œå�ŒæœŸãƒ»Dâ†’E ã�® AI è£œå®Œãƒ»ãƒ‡ã‚£ã‚¹ã‚¯å��æ˜ ï¼ˆæ—¢å®šã�¯ xlwings ã�§ Aã€œE å�ŒæœŸâ†’Saveã€‚``EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1`` ã�®ã�¨ã�� openpyxl save ã‚’è©¦è¡Œï¼‰ã€‚

    xlwings ã�§ã‚‚ä¿�å­˜ã�§ã��ã�ªã�„ã�¨ã��ã�¯ ``log/exclude_rules_matrix_vba.tsv`` ã‚’æ®‹ã�—ã€�ãƒžã‚¯ãƒ­
    ``è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹_Aã�‹ã‚‰E_TSVã�‹ã‚‰å��æ˜ `` ã�§ Aã€œE ã‚’å��æ˜ ã�™ã‚‹ã€‚
    ä½µã�›ã�¦å¾“æ�¥ã�©ã�Šã‚Š E åˆ—ã�®ã�¿ã�® ``exclude_rules_e_column_vba.tsv`` ã‚‚å‡ºåŠ›ã�•ã‚Œå¾—ã‚‹ï¼ˆè¡Œåˆ— TSV å„ªå…ˆã�§å��æ˜ å¾Œã�¯å‰Šé™¤ï¼‰ã€‚
    ä¿�å­˜æˆ�åŠŸæ™‚ã�¯ TSV/JSON ã�¯å‰Šé™¤ã�•ã‚Œã‚‹ã€‚

    ``json/exclude_rules_e_column_pending.json`` ã�¯ Python æ¬¡å›žèµ·å‹•æ™‚ã�® E åˆ—å¾©å…ƒç”¨ã€‚
    ã‚·ãƒ¼ãƒˆã�®æ–°è¦�ä½œæˆ�ã�¨ 1 è¡Œç›®è¦‹å‡ºã�—ã�¯ VBAã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹_ã‚·ãƒ¼ãƒˆã‚’ç¢ºä¿�ã€�ã€‚
    """
    if not wb_path:
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_PATH",
            log_prefix,
            "TASK_INPUT_WORKBOOK ã�Œç©ºã�®ã�Ÿã‚�è¨­å®šã‚·ãƒ¼ãƒˆå‡¦ç�†ã‚’ã�—ã�¾ã�›ã‚“ã€‚",
        )
        return
    if not os.path.exists(wb_path):
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_FILE",
            log_prefix,
            "ãƒ–ãƒƒã‚¯ã�Œå­˜åœ¨ã�—ã�¾ã�›ã‚“ã€‚",
            details=f"path={wb_path}",
        )
        return

    _log_exclude_rules_sheet_debug(
        "START",
        log_prefix,
        "è¨­å®šã‚·ãƒ¼ãƒˆä¿�å®ˆé–‹å§‹",
        details=f"path={wb_path} pairs={len(pairs)}",
    )
    global _exclude_rules_effective_read_path
    _exclude_rules_effective_read_path = None

    if _workbook_should_skip_openpyxl_io(wb_path):
        _log_exclude_rules_sheet_debug(
            "SKIP_OPENPYXL_INCOMPATIBLE_BOOK",
            log_prefix,
            f"ãƒ–ãƒƒã‚¯ã�«ã€Œ{OPENPYXL_INCOMPATIBLE_SHEET_MARKER}ã€�ã�Œå�«ã�¾ã‚Œã‚‹ã�Ÿã‚�ã€�openpyxl ã�«ã‚ˆã‚‹è¨­å®šã‚·ãƒ¼ãƒˆä¿�å®ˆã�¯è¡Œã�„ã�¾ã�›ã‚“ã€‚",
            details=f"path={wb_path}",
        )
        logging.warning(
            "%s: ã€Œ%sã€�å�«æœ‰ã�®ã�Ÿã‚�ã€Œ%sã€�ã�® openpyxl ä¿�å®ˆã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿï¼ˆExcelï¼�xlwings ã�§ç·¨é›†ã�—ã�¦ã��ã� ã�•ã�„ï¼‰ã€‚",
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
                "keep_vba=True ã�§ãƒ–ãƒƒã‚¯ã‚’é–‹ã�‘ã�š keep_vba=False ã�§å†�è©¦è¡Œã�—ã�¾ã�™ï¼ˆãƒžã‚¯ãƒ­ã�Œå¤±ã‚�ã‚Œã‚‹å�¯èƒ½æ€§ï¼‰ã€‚",
                exc=e1,
            )
            try:
                wb = load_workbook(wb_path, keep_vba=False, read_only=False, data_only=False)
            except Exception as e2:
                _log_exclude_rules_sheet_debug(
                    "OPEN_FAIL",
                    log_prefix,
                    "ãƒ–ãƒƒã‚¯ã‚’é–‹ã�‘ã�¾ã�›ã‚“ã€‚ã‚·ãƒ¼ãƒˆã�¯ä½œæˆ�ãƒ»ä¿�å­˜ã�•ã‚Œã�¾ã�›ã‚“ã€‚",
                    details=f"path={wb_path}",
                    exc=e2,
                )
                return
        else:
            _log_exclude_rules_sheet_debug(
                "OPEN_FAIL",
                log_prefix,
                "ãƒ–ãƒƒã‚¯ã‚’é–‹ã�‘ã�¾ã�›ã‚“ã€‚ã‚·ãƒ¼ãƒˆã�¯ä½œæˆ�ãƒ»ä¿�å­˜ã�•ã‚Œã�¾ã�›ã‚“ã€‚",
                details=f"path={wb_path}",
                exc=e1,
            )
            return

    _log_exclude_rules_sheet_debug(
        "OPEN_OK",
        log_prefix,
        "ãƒ–ãƒƒã‚¯ã‚’é–‹ã��ã�¾ã�—ã�Ÿã€‚",
        details=f"keep_vba={keep_vba} sheets={len(wb.sheetnames)}",
    )

    try:
        if EXCLUDE_RULES_SHEET_NAME not in wb.sheetnames:
            _log_exclude_rules_sheet_debug(
                "SKIP_NO_SHEET",
                log_prefix,
                "ã‚·ãƒ¼ãƒˆã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚VBA ã�®ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹_ã‚·ãƒ¼ãƒˆã‚’ç¢ºä¿�ã€�ã‚’å®Ÿè¡Œã�™ã‚‹ã�‹ã€�æ®µéšŽ1/2 ã‚’ãƒžã‚¯ãƒ­ã�‹ã‚‰èµ·å‹•ã�—ã�¦ã��ã� ã�•ã�„ã€‚",
                details=f"path={wb_path}",
            )
            logging.error(
                "%s: ã€Œ%sã€�ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚Python ã�§ã�¯ã‚·ãƒ¼ãƒˆã‚’ä½œæˆ�ã�—ã�¾ã�›ã‚“ã€‚",
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
                "1è¡Œç›®ã�«æ¨™æº–è¦‹å‡ºã�—ã‚’æ›¸ã��è¾¼ã�¿ã�¾ã�—ã�Ÿï¼ˆç©ºã‚·ãƒ¼ãƒˆãƒ»åˆ—å��ä¸�ä¸€è‡´ã�®è£œæ­£ï¼‰ã€‚",
                details=f"cols=({c_proc},{c_mach},{c_flag},{c_d},{c_e})",
            )

        # å‰�å›žãƒ–ãƒƒã‚¯ä¿�å­˜ã�«å¤±æ•—ã�—ã�Ÿã�¨ã��é€€é�¿ã�—ã�Ÿ E åˆ—ã‚’ã€�å…ˆã�«ãƒ¯ãƒ¼ã‚¯ã‚·ãƒ¼ãƒˆã�¸æˆ»ã�™ï¼ˆç¶šã��ä¿�å­˜ã�§ãƒ‡ã‚£ã‚¹ã‚¯ã�¸è¼‰ã‚‹ï¼‰
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
                f"å·¥ç¨‹+æ©Ÿæ¢°ã�®è¡Œã‚’ {added} ä»¶è¿½åŠ ã�—ã�¾ã�—ã�Ÿã€‚",
            )
            logging.info(
                "%s: ã€Œ%sã€�ã�«å·¥ç¨‹+æ©Ÿæ¢°ã�®çµ„ã�¿å�ˆã‚�ã�›ã‚’ %s è¡Œè¿½åŠ ã�—ã�¾ã�—ã�Ÿã€‚",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                added,
            )

        # åŠ å·¥è¨ˆç”»ã�‹ã‚‰ãƒšã‚¢ã�Œ1ä»¶ã‚‚å�–ã‚Œã�šã€�ã‚·ãƒ¼ãƒˆã�«ã‚‚ãƒ‡ãƒ¼ã‚¿è¡Œã�Œç„¡ã�„ã�¨ã��ã�¯ä¾‹è¡Œã�®ã�¿ï¼ˆå¾“æ�¥ã�®æ–°è¦�ã‚·ãƒ¼ãƒˆç›¸å½“ï¼‰
        if added == 0 and not existing_keys:
            ws.append(["æ¢±åŒ…", "", "yes", "", ""])
            existing_keys.add(
                (_normalize_process_name_for_rule_match("æ¢±åŒ…"), _normalize_equipment_match_key(""))
            )
            _log_exclude_rules_sheet_debug(
                "EXAMPLE_ROW",
                log_prefix,
                "ãƒ‡ãƒ¼ã‚¿è¡Œã�Œç„¡ã�‹ã�£ã�Ÿã�Ÿã‚�ä¾‹ï¼ˆæ¢±åŒ…=yesï¼‰ã‚’1è¡Œè¿½åŠ ã€‚",
            )
            logging.info(
                "%s: ã€Œ%sã€�ã�«ãƒ‡ãƒ¼ã‚¿è¡Œã�Œç„¡ã�‹ã�£ã�Ÿã�Ÿã‚�ã€�ä¾‹ï¼ˆæ¢±åŒ…=yesï¼‰ã‚’1è¡Œè¿½åŠ ã�—ã�¾ã�—ã�Ÿã€‚",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
            )

        # ç©ºè¡Œè©°ã‚�ã�¯ AI ã‚ˆã‚Šå…ˆã�«è¡Œã�†ï¼ˆå¾Œã�‹ã‚‰è©°ã‚�ã‚‹ã�¨ã€�æ›¸ã��è¾¼ã‚“ã� è¡Œç•ªå�·ã�¨ç”»é�¢ä¸Šã�®è¡Œã�Œã�šã‚Œã‚‹ï¼‰
        n_kept, n_removed_empty = _compact_exclude_rules_data_rows(
            ws, c_proc, c_mach, c_flag, c_d, c_e, log_prefix
        )
        if n_removed_empty:
            _log_exclude_rules_sheet_debug(
                "DATA_COMPACT",
                log_prefix,
                "ç©ºè¡Œã‚’å‰Šé™¤ã�—ã�¦ãƒ‡ãƒ¼ã‚¿è¡Œã‚’è©°ã‚�ã�¾ã�—ã�Ÿï¼ˆä¸¦ã�³é †ã�¯ç¶­æŒ�ï¼‰ã€‚AI è£œå®Œã‚ˆã‚Šå‰�ã€‚",
                details=f"rows={n_kept} removed_empty={n_removed_empty}",
            )

        max_r = int(ws.max_row or 1)
        pending_rows: list[int] = []
        for r in range(2, max_r + 1):
            dv = ws.cell(row=r, column=c_d).value
            ev = ws.cell(row=r, column=c_e).value
            # C åˆ—ã�®æœ‰ç„¡ã�«é–¢ä¿‚ã�ªã��ã€�D ã�«èª¬æ˜Žã�Œã�‚ã‚Š E ã�Œç©ºã�ªã‚‰ Dâ†’E ã‚’è©¦ã�™
            if _cell_is_blank_for_rule(dv):
                continue
            if not _cell_is_blank_for_rule(ev):
                continue
            pending_rows.append(r)

        # D ã�Œæ•°å¼�ã�®ã�¨ã��ã�¯é€šå¸¸èª­è¾¼ã�§ã�¯ '=...' ã� ã�‘å�–ã‚Œã‚‹ã€‚data_only ã�§ã‚­ãƒ£ãƒƒã‚·ãƒ¥è¡¨ç¤ºå€¤ã‚’è£œã�†ã€‚
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
                        "%s: ã€Œ%sã€�%s è¡Œç›®ã�® D åˆ—ã�Œæ•°å¼�ã�§ã€�ã‚­ãƒ£ãƒƒã‚·ãƒ¥å€¤ã‚’èª­ã‚�ã�¾ã�›ã‚“ã�§ã�—ã�Ÿï¼ˆExcel ã�§ä¸€åº¦ä¿�å­˜ã�™ã‚‹ã�‹ D ã‚’å€¤ã�«ã�—ã�¦ã��ã� ã�•ã�„ï¼‰ã€‚",
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
                        "%s: ã€Œ%sã€�%s è¡Œç›®ã�® D åˆ—ã‚’ JSON ã�«ã�§ã��ã�¾ã�›ã‚“ã�§ã�—ã�Ÿï¼ˆAPIã‚­ãƒ¼ãƒ»å¿œç­”ã‚’ç¢ºèª�ï¼‰ã€‚",
                        log_prefix,
                        EXCLUDE_RULES_SHEET_NAME,
                        r,
                    )
                    continue
                jstr = json.dumps(parsed, ensure_ascii=False)
                ws.cell(row=r, column=c_e, value=jstr)
                cell_addr = f"{get_column_letter(c_e)}{r}"
                ai_e_cell_addrs.append(cell_addr)
                preview = jstr if len(jstr) <= 160 else (jstr[:160] + "â€¦")
                logging.info(
                    "%s: ã€Œ%sã€�ãƒ­ã‚¸ãƒƒã‚¯å¼�åˆ—ã€Œ%sã€�ã‚»ãƒ« %s ã�« JSON ã‚’æ›¸ã��è¾¼ã�¿: %s",
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
                f"Dâ†’E ã�® AI è£œå®Œã‚’ {ai_filled} è¡Œå®Ÿæ–½ã€‚",
                details="cells=" + ",".join(ai_e_cell_addrs),
            )
            logging.info(
                "%s: ã€Œ%sã€�ã�§ Dâ†’E ã�® AI è£œå®Œã‚’ %s è¡Œï¼ˆã‚»ãƒ«: %sï¼‰ã€‚",
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
                f'Eåˆ— {_e_addr} ã�«ãƒ†ã‚¹ãƒˆã�§ "1234" ã‚’æ›¸ã��è¾¼ã�¿',
                details=f"row={_er_row}",
            )
            logging.warning(
                '%s: ã€�ãƒ†ã‚¹ãƒˆã€‘%s ã�« "1234" ã‚’æ›¸ã��è¾¼ã�¿ï¼ˆEXCLUDE_RULES_TEST_E1234ï¼‰ã€‚',
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
                "%s: è¨­å®šã‚·ãƒ¼ãƒˆã�® openpyxl ä¿�å­˜ã�«å¤±æ•—ã�—ã�¾ã�—ã�Ÿã€‚"
                " log ã�®è¡Œåˆ— TSV ã‚’ãƒžã‚¯ãƒ­ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹_Aã�‹ã‚‰E_TSVã�‹ã‚‰å��æ˜ ã€�ã€�"
                "ã�¾ã�Ÿã�¯ E åˆ—ã�®ã�¿ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹_Eåˆ—_TSVã�‹ã‚‰å��æ˜ ã€�ã�§å��æ˜ ã�—ã�¦ã��ã� ã�•ã�„ã€‚",
                log_prefix,
            )
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "FATAL",
            log_prefix,
            "è¨­å®šã‚·ãƒ¼ãƒˆå‡¦ç�†ä¸­ã�«æœªæ�•æ�‰ä¾‹å¤–ã�Œç™ºç”Ÿã�—ã�¾ã�—ã�Ÿã€‚",
            exc=ex,
        )
        logging.exception("%s: è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã�®å‡¦ç�†ã�§ä¾‹å¤–", log_prefix)
    finally:
        if wb is not None:
            wb.close()
            _log_exclude_rules_sheet_debug("CLOSED", log_prefix, "ãƒ–ãƒƒã‚¯ã‚’ã‚¯ãƒ­ãƒ¼ã‚ºã�—ã�¾ã�—ã�Ÿã€‚")


def _resolve_exclude_rules_workbook_path_for_read(wb_path: str) -> str:
    """ç›´å‰�ã�®ä¿�å®ˆã�§å®ŸåŠ¹ãƒ‘ã‚¹ã�Œå¤‰ã‚�ã�£ã�Ÿã�¨ã��ï¼ˆé€šå¸¸ã�¯ä¿�å­˜æˆ�åŠŸå¾Œã�®å…ƒãƒ–ãƒƒã‚¯ï¼‰ã�«ã��ã‚Œã‚’ä½¿ã�†ã€‚"""
    p = _exclude_rules_effective_read_path
    if p and os.path.exists(p):
        return p
    return wb_path


def _load_exclude_rules_from_workbook(wb_path: str) -> list[dict]:
    """ã‚·ãƒ¼ãƒˆã�‹ã‚‰ãƒ«ãƒ¼ãƒ«è¡Œã‚’èª­ã�¿ã€�è©•ä¾¡ç”¨ãƒªã‚¹ãƒˆã‚’è¿”ã�™ã€‚"""
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
            "é…�å�°ä¸�è¦�ãƒ«ãƒ¼ãƒ«: ãƒ–ãƒƒã‚¯ã�«ã€Œ%sã€�ã�Œã�‚ã‚‹ã�Ÿã‚� pandas(openpyxl) ã�§ã�®ã€Œ%sã€�èª­è¾¼ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿï¼ˆãƒ«ãƒ¼ãƒ«ã�¯æœªé�©ç”¨ï¼‰ã€‚",
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
    """è¨­å®šã‚·ãƒ¼ãƒˆã�«åŸºã�¥ã��ã€Œé…�å�°ä¸�è¦�ã€�ã‚’è¨­å®šï¼ˆC=yes ã�¾ã�Ÿã�¯ E ã�® JSON ã�ŒçœŸï¼‰ã€‚"""
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
        logging.info("%s: è¨­å®šã€Œ%sã€�ã�«ã‚ˆã‚Šé…�å�°ä¸�è¦�=yes ã‚’ %s è¡Œã�«è¨­å®šã�—ã�¾ã�—ã�Ÿã€‚", log_prefix, EXCLUDE_RULES_SHEET_NAME, n)
    return df


def _sort_stage1_plan_df_by_dispatch_trial_order_asc(plan_df: "pd.DataFrame") -> "pd.DataFrame":
    """
    æ®µéšŽ1å‡ºåŠ›ç›´å‰�: é…�å�°è©¦è¡Œé †ç•ªã�®æ˜‡é †ã�«è¡Œã‚’ä¸¦ã�¹æ›¿ã�ˆã�Ÿ DataFrame ã‚’è¿”ã�™ã€‚
    æ­£ã�®æ•´æ•°ã�§ã�ªã�„ã‚»ãƒ«ã�¯æœ€å¾Œï¼ˆå�Œå¸¯å†…ã�¯å…ƒã�®è¡Œé †ï¼‰ã€‚
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
# æ®µéšŽ1ã‚¨ãƒ³ãƒˆãƒªï¼ˆtask_extract_stage1.py â†’ run_stage1_extractï¼‰
#   åŠ å·¥è¨ˆç”»DATA èª­å�– â†’ é…�å�°ä¸�è¦�è‡ªå‹•å‡¦ç�† â†’ è¨­å®šã‚·ãƒ¼ãƒˆä¿�å®ˆ â†’ plan_input_tasks.xlsx å‡ºåŠ›
# =============================================================================
def run_stage1_extract():
    """
    æ®µéšŽ1: åŠ å·¥è¨ˆç”»DATA ã�‹ã‚‰é…�å�°ç”¨ã‚¿ã‚¹ã‚¯ä¸€è¦§ã‚’æŠ½å‡ºã�— output/plan_input_tasks.xlsx ã�¸å‡ºåŠ›ã€‚
    å�Œä¸€ä¾�é ¼NOã�§å�Œä¸€æ©Ÿæ¢°å��ã�Œè¤‡æ•°è¡Œã�‚ã‚‹ã�¨ã��ã€�å·¥ç¨‹å��ã€Œåˆ†å‰²ã€�è¡Œã�®ç©ºã�®ã€Œé…�å�°ä¸�è¦�ã€�ã�« yes ã‚’è‡ªå‹•è¨­å®šã�™ã‚‹ã€‚
    ãƒžã‚¯ãƒ­ãƒ–ãƒƒã‚¯ã�®ã€Œè¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã€�ã�§å·¥ç¨‹+æ©Ÿæ¢°ã�”ã�¨ã�®é…�å�°ä¸�è¦�ãƒ»æ�¡ä»¶å¼�ï¼ˆAIï¼‰ã‚’ç®¡ç�†ã�™ã‚‹ï¼ˆã‚·ãƒ¼ãƒˆä½œæˆ�ã�¯ VBAï¼‰ã€‚
    """
    if not TASKS_INPUT_WORKBOOK:
        logging.error("TASK_INPUT_WORKBOOK ã�Œæœªè¨­å®šã�§ã�™ã€‚")
        return False
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        logging.error(f"TASK_INPUT_WORKBOOK ã�Œå­˜åœ¨ã�—ã�¾ã�›ã‚“: {TASKS_INPUT_WORKBOOK}")
        return False
    reset_gemini_usage_tracker()
    df_src = load_tasks_df()
    try:
        _pm_pairs = _collect_process_machine_pairs_for_exclude_rules(df_src)
        run_exclude_rules_sheet_maintenance(TASKS_INPUT_WORKBOOK, _pm_pairs, "æ®µéšŽ1")
    except Exception:
        logging.exception("æ®µéšŽ1: è¨­å®š_é…�å�°ä¸�è¦�å·¥ç¨‹ã�®ä¿�å®ˆã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰")
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
        # å·¥ç¨‹å�� + æ©Ÿæ¢°å�� ã‚’â€œå› å­�â€�ã�¨ã�—ã�¦è¡¨ç¤ºç”¨ã�«è¿½åŠ ï¼ˆå¾Œæ®µã�¯è¨ˆç®—ã‚­ãƒ¼ã�«ã‚‚ä½¿ç”¨ï¼‰
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
        logging.warning("æ®µéšŽ1: æŠ½å‡ºå¯¾è±¡ã‚¿ã‚¹ã‚¯ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚")
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
        logging.error("æ®µéšŽ1ã‚’ä¸­æ–­: ãƒžã‚¹ã‚¿ skills ã�®æ¤œè¨¼ã‚¨ãƒ©ãƒ¼ï¼ˆå„ªå…ˆåº¦ã�®æ•°å€¤é‡�è¤‡ã�ªã�©ï¼‰ã€‚")
        raise
    except Exception as e:
        logging.info("æ®µéšŽ1: ãƒžã‚¹ã‚¿ need ã‚’èª­ã‚�ã�šå…ƒåˆ—ã�¯ need ã�ªã�—ã�§åŸ‹ã‚�ã�¾ã�™ (%s)", e)
        req_map, need_rules = {}, []
        equipment_list_stage1 = []
        need_combo_col_index_stage1 = {}
    out_df = _merge_plan_sheet_user_overrides(out_df)
    _refresh_plan_reference_columns(out_df, req_map, need_rules)
    try:
        _apply_auto_exclude_bunkatsu_duplicate_machine(out_df, log_prefix="æ®µéšŽ1")
    except Exception as ex:
        logging.exception("æ®µéšŽ1: åˆ†å‰²è¡Œã�®é…�å�°ä¸�è¦�è‡ªå‹•è¨­å®šã�§ä¾‹å¤–ï¼ˆå‡ºåŠ›ã�¯ç¶šè¡Œï¼‰: %s", ex)
    try:
        out_df = apply_exclude_rules_config_to_plan_df(out_df, TASKS_INPUT_WORKBOOK, "æ®µéšŽ1")
    except Exception as ex:
        logging.warning("æ®µéšŽ1: è¨­å®šã‚·ãƒ¼ãƒˆã�«ã‚ˆã‚‹é…�å�°ä¸�è¦�é�©ç”¨ã�§ä¾‹å¤–ï¼ˆç¶šè¡Œï¼‰: %s", ex)
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
        logging.warning("æ®µéšŽ1: é…�å�°è©¦è¡Œé †ç•ªåˆ—ã�®è¨ˆç®—ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿï¼ˆç¶šè¡Œï¼‰: %s", ex)
    out_df = _sort_stage1_plan_df_by_dispatch_trial_order_asc(out_df)
    out_path = os.path.join(output_dir, STAGE1_OUTPUT_FILENAME)
    out_df.to_excel(out_path, sheet_name="ã‚¿ã‚¹ã‚¯ä¸€è¦§", index=False)
    _apply_excel_date_columns_date_only_display(out_path, "ã‚¿ã‚¹ã‚¯ä¸€è¦§")
    _apply_plan_input_visual_format(out_path, "ã‚¿ã‚¹ã‚¯ä¸€è¦§")
    logging.info(f"æ®µéšŽ1å®Œäº†: '{out_path}' ã‚’å‡ºåŠ›ã�—ã�¾ã�—ã�Ÿã€‚ãƒžã‚¯ãƒ­ã�§ '{PLAN_INPUT_SHEET_NAME}' ã�«å�–ã‚Šè¾¼ã‚“ã�§ã��ã� ã�•ã�„ã€‚")
    _try_write_main_sheet_gemini_usage_summary("æ®µéšŽ1")
    return True


# ç¨¼åƒ�ãƒ«ãƒ¼ãƒ«ï¼ˆãƒ‡ãƒ•ã‚©ãƒ«ãƒˆå€¤ãƒ»2026å¹´3æœˆåŸºæº–ï¼‰
TARGET_YEAR = 2026
TARGET_MONTH = 3
DEFAULT_START_TIME = time(8, 45)
DEFAULT_END_TIME = time(17, 0)
DEFAULT_BREAKS = [
    (time(12, 0), time(12, 50)),
    (time(14, 45), time(15, 0))
]
# çµ‚æ¥­ç›´å‰�ãƒ‡ãƒ•ã‚¡ãƒ¼: ASSIGN_END_OF_DAY_DEFER_MINUTES ã�Œæ­£ã�®ã�¨ã��ã€�team_end_limit ã�¾ã�§ã�®æ®‹ã‚Šã�Œã��ã�®åˆ†æ•°ä»¥ä¸‹ã�§ã€�
# ã�‹ã�¤ remaining_unitsï¼ˆåˆ‡ã‚Šä¸Šã�’ï¼‰ã�Œ ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS ä»¥ä¸‹ã�®ã�¨ã��ã€�ã��ã�®æ—¥ã�®é–‹å§‹ä¸�å�¯ï¼ˆNoneï¼‰ã€‚
# å�Œã�˜ã‚¦ã‚£ãƒ³ãƒ‰ã‚¦ã�§ã€ŒASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS ãƒ­ãƒ¼ãƒ«åˆ†ä»¥ä¸Šã�¯å›žã�›ã�ªã�„ã€�ï¼ˆå�Žå®¹ã�Œé–¾å€¤æœªæº€ï¼‰ã�¨ã��ã�¯
# æ–°è¦�ã�«åŠ å·¥ã‚’å§‹ã‚�ã�ªã�„ï¼ˆ_eod_reject_capacity_units_below_thresholdï¼‰ã€‚
# ASSIGN_END_OF_DAY_DEFER_MINUTES æ—¢å®š 45ï¼ˆåˆ†ï¼‰ã€‚0 ã‚’æ˜Žç¤ºã�™ã‚‹ã�¨ç„¡åŠ¹ï¼ˆå¾“æ�¥ã�©ã�Šã‚Šï¼‰ã€‚
# ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS æ—¢å®š 5ã€‚å��åˆ†å¤§ã��ã�ªå€¤ï¼ˆä¾‹: 999999ï¼‰ã�«ã�™ã‚‹ã�¨å®Ÿè³ªã€Œæ®‹ãƒ­ãƒ¼ãƒ«ã�«ä¾�ã‚‰ã�šçµ‚æ¥­ç›´å‰�ã�¯ä¸�å�¯ã€�ã€‚
# ä¼‘æ†©: å¸¯å†…ã�«è�½ã�¡ã�Ÿé–‹å§‹ã�¯ _defer_team_start_past_prebreak_and_end_of_day ã�§ä¼‘æ†©çµ‚äº†ã�¸ç¹°ã‚Šä¸‹ã�’ã€‚
# ä¼‘æ†©ã‚’ã�¾ã�Ÿã��é€£ç¶šé…�å�°ã�¯ _contiguous_work_minutes_until_next_break_or_limit ã�§å�´ä¸‹ã€‚
# ï¼ˆæ—§ ASSIGN_DEFER_MIN_REMAINING_ROLLS / ASSIGN_PRE_BREAK_DEFER_GAP_MINUTES ã�¯å»ƒæ­¢ãƒ»ç„¡è¦–ï¼‰
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
    """ASSIGN_END_OF_DAY_DEFER_MINUTES ã�Œæ­£ã�®ã�¨ã��ã€�é–‹å§‹ã�Œçµ‚æ¥­ä¸Šé™�ã�®ã��ã�®åˆ†æ•°ä»¥å†…ã�‹ã€‚"""
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
    çµ‚æ¥­ã�¾ã�§ã�‚ã�¨ ASSIGN_END_OF_DAY_DEFER_MINUTES åˆ†ä»¥å†…ã�®ã‚¦ã‚£ãƒ³ãƒ‰ã‚¦å†…ã�§ã€�
    ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS ãƒ­ãƒ¼ãƒ«åˆ†ä»¥ä¸Šã�¯å›žã�›ã�ªã�„ï¼ˆå�Žå®¹ãƒ­ãƒ¼ãƒ«æ•°ã�Œé–¾å€¤æœªæº€ï¼‰ã�¨ã�� Trueï¼ˆæ–°è¦�åŠ å·¥ã‚’å§‹ã‚�ã�ªã�„ï¼�å€™è£œå�´ä¸‹ï¼‰ã€‚
    """
    th = ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS
    if th <= 0:
        return False
    if not _eod_minutes_window_covers_start(team_start, team_end_limit):
        return False
    return int(units_fit_until_close) < int(th)


# =========================================================
# 1. ã‚³ã‚¢è¨ˆç®—ãƒ­ã‚¸ãƒƒã‚¯ (æ—¥æ™‚ãƒ™ãƒ¼ã‚¹)
#    ä¼‘æ†©å¸¯ã‚’æŒŸã‚“ã� ã€Œå®Ÿåƒ�åˆ†ã€�æ�›ç®—ãƒ»çµ‚äº†æ™‚åˆ»ã�®ç¹°ã‚Šä¸Šã�’ã€‚å‰²ä»˜ãƒ«ãƒ¼ãƒ—ã�®ä¸‹å›žã‚Šã€‚
# =========================================================
def merge_time_intervals(intervals):
    """æ™‚åˆ»åŒºé–“ã�®ãƒªã‚¹ãƒˆã‚’ã‚½ãƒ¼ãƒˆã�—ã€�é‡�ã�ªã‚‹åŒºé–“ã‚’çµ�å�ˆã�—ã�¦è¿”ã�™ã€‚"""
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
    start_dt ã�‹ã‚‰æ¬¡ã�®ä¼‘æ†©é–‹å§‹ï¼ˆã�¾ã�Ÿã�¯çµ‚æ¥­ä¸Šé™�ï¼‰ã�¾ã�§ã�®ã€�é€£ç¶šã�—ã�¦å®Ÿåƒ�ã�«ä½¿ã�ˆã‚‹åˆ†æ•°ã€‚
    é–‹å§‹ã�Œä¼‘æ†©å¸¯å†…ã�ªã‚‰ 0ï¼ˆå‘¼ã�³å‡ºã�—å…ƒã�§å�´ä¸‹ï¼‰ã€‚breaks_dt ã�¯ merge æ¸ˆã�¿æƒ³å®šã€‚
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
    ä¼‘æ†©å¸¯å¤–ã�§ã‚‚ã€�æ¬¡ã�®ä¼‘æ†©é–‹å§‹ã�¾ã�§ã�®é€£ç¶šå®Ÿåƒ�ã�Œ min_contiguous_mins æœªæº€ã�ªã‚‰ã€�
    ã��ã�®ä¼‘æ†©åŒºé–“ã�®çµ‚äº†æ™‚åˆ»ã‚’è¿”ã�™ï¼ˆå�ˆå¾Œä¼‘æ†©ç›´å‰�ã�« 1 ãƒ­ãƒ¼ãƒ«åˆ†ã�Œå�Žã�¾ã‚‰ã�ªã�„é–‹å§‹ã� ã�‘é€²ã‚�ã‚‹ï¼‰ã€‚
    çµ‚æ¥­ã�¾ã�§ã�—ã�‹å®Ÿåƒ�ã�Œç¶šã�‹ã�ªã�„å ´å�ˆã�¯ Noneã€‚
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
    - ASSIGN_END_OF_DAY_DEFER_MINUTES > 0 ã�‹ã�¤ (team_end_limit - è©¦è¡Œé–‹å§‹) ã�Œã��ã�®åˆ†æ•°ä»¥ä¸‹ã�§ã€�
      remaining_units åˆ‡ã‚Šä¸Šã�’ã�Œ ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS ä»¥ä¸‹ã�®ã�¨ã��ã€�å½“æ—¥é–‹å§‹ä¸�å�¯ï¼ˆNoneï¼‰ã€‚
    - è©¦è¡Œé–‹å§‹ã�Œä¼‘æ†©å¸¯å†…ã�®ã�¨ã��ã�¯ **ä¼‘æ†©çµ‚äº†æ™‚åˆ»ã�¸ç¹°ã‚Šä¸‹ã�’**ã�—ã€�`refloor_fn` ã�§è¨­å‚™ä¸‹é™�ãƒ»avail ã‚’å†�é�©ç”¨ã�™ã‚‹ã€‚
      ç¹°ã‚Šä¸‹ã�’ã�®ã�‚ã�¨çµ‚æ¥­è¶…é�Žãƒ»EOD ãƒ‡ãƒ•ã‚¡ãƒ¼ã�«è©²å½“ã�™ã‚Œã�° Noneã€‚
    - min_contiguous_work_mins ã�Œæ­£ã�®ã�¨ã��ã€�å¸¯å¤–ã�§ã‚‚ **æ¬¡ã�®ä¼‘æ†©ã�¾ã�§ã�®é€£ç¶šå®Ÿåƒ�**ã�Œã��ã‚Œæœªæº€ã�ªã‚‰
      å½“è©²ä¼‘æ†©ã�®çµ‚äº†ã�¸ç¹°ã‚Šä¸‹ã�’ï¼ˆä¸Šã�¨å�Œæ§˜ã�« refloor ã�—ãƒ«ãƒ¼ãƒ—ï¼‰ã€‚
    """
    _tid = str(task.get("task_id", "") or "").strip()
    _team_txt = ", ".join(str(x) for x in team) if team else "â€”"

    def _trace_block(msg: str, *a) -> None:
        if not _trace_schedule_task_enabled(_tid):
            return
        _log_dispatch_trace_schedule(
            _tid,
            "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ãƒ–ãƒ­ãƒƒã‚¯åˆ¤å®š: " + msg,
            _tid,
            *a,
        )

    ts = refloor_fn(team_start)
    for _ in range(64):
        if ts >= team_end_limit:
            _trace_block(
                "é–‹å§‹ä¸�å�¯(çµ‚æ¥­è¶…é�Ž) machine=%s team=%s rem=%.4f trial_start=%s end_limit=%s",
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
                "ä¼‘æ†©å¸¯å†…ã�®ã�Ÿã‚�çµ‚äº†ã�¸ç¹°ã‚Šä¸‹ã�’ machine=%s team=%s rem=%.4f break_end=%s trial_was=%s",
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
                    "ä¼‘æ†©ç›´å‰�ã�§é€£ç¶šå®Ÿåƒ�ä¸�è¶³ã�®ã�Ÿã‚�ä¼‘æ†©çµ‚äº†ã�¸ç¹°ã‚Šä¸‹ã�’ machine=%s team=%s rem=%.4f need_contig_min=%s trial_was=%s break_end=%s",
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
                "é–‹å§‹ä¸�å�¯(çµ‚æ¥­ç›´å‰�ãƒ»å°�æ®‹ãƒ­ãƒ¼ãƒ«) machine=%s team=%s rem_ceil=%s max_rem=%s trial_start=%s end_limit=%s gap_end_min=%s",
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
        "é–‹å§‹ä¸�å�¯(ä¼‘æ†©ç¹°ã‚Šä¸‹ã�’æ‰“åˆ‡ã‚Š) machine=%s team=%s rem=%.4f trial_start=%s",
        task.get("machine"),
        _team_txt,
        float(task.get("remaining_units") or 0),
        ts,
    )
    return None


def _expand_timeline_events_for_equipment_grid(timeline_events: list) -> list:
    """
    è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ãƒ»ãƒ¡ãƒ³ãƒ�ãƒ¼æ—¥ç¨‹ãƒ»ç¨¼åƒ�çŽ‡ç”¨ã‚¤ãƒ³ãƒ‡ãƒƒã‚¯ã‚¹å�‘ã�‘ã€‚
    1 æœ¬ã�®ã‚¤ãƒ™ãƒ³ãƒˆã�Œæ—¥ã‚’ã�¾ã�Ÿã��å ´å�ˆã€�e["date"] ã� ã�‘å½“æ—¥ã�«è¼‰ã�›ã‚‹ã�¨ç¿Œæœ�ã‚»ã‚°ãƒ¡ãƒ³ãƒˆã�Œæ¬ ã�‘ã‚‹ã�Ÿã‚�ã€�
    start_dtã€œend_dt ã‚’å�„å°±æ¥­æ—¥ DEFAULT_START_TIMEã€œDEFAULT_END_TIME ã�«ã‚¯ãƒªãƒƒãƒ—ã�—ã�Ÿè¤‡è£½ã�¸å±•é–‹ã�™ã‚‹ã€‚
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
    start_dt ã�‹ã‚‰ end_dt ã�¾ã�§ã�®ã€Œä¼‘æ†©ã‚’é™¤ã�„ã�Ÿå®Ÿåƒ�åˆ†æ•°ã€�ã€‚
    breaks_dt â€¦ (åŒºé–“é–‹å§‹, åŒºé–“çµ‚äº†) ã�®åˆ—ï¼ˆdatetime ã�¾ã�Ÿã�¯ timeã€‚å‘¼ã�³å‡ºã�—å…ƒã�®å‹¤æ€ ã‚¤ãƒ™ãƒ³ãƒˆã�¨æ•´å�ˆï¼‰ã€‚
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
    start_dt ã�‹ã‚‰å®Ÿåƒ� duration_minutes åˆ†é€²ã‚�ã�Ÿçµ‚äº† datetime ã‚’æ±‚ã‚�ã‚‹ï¼ˆä¼‘æ†©ã�¯ã‚¹ã‚­ãƒƒãƒ—ï¼‰ã€‚
    end_limit_dt ã‚’è¶…ã�ˆã�ªã�„ã‚ˆã�†æ‰“ã�¡åˆ‡ã‚Šã€‚æˆ»ã‚Šå€¤: (çµ‚äº†æ™‚åˆ», å®Ÿéš›ã�«é€²ã‚�ã�Ÿå®Ÿåƒ�åˆ†, æ®‹ã‚Šæœªæ¶ˆåŒ–åˆ†)
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
    need ã‚·ãƒ¼ãƒˆã€Œä¾�é ¼NOæ�¡ä»¶ã€�æ¬„ã�®è§£é‡ˆã€‚
    ç©ºãƒ»*ãƒ»å…¨ä»¶ â†’ å¸¸ã�«ãƒžãƒƒãƒ�ã€‚
    prefix:ABC / æŽ¥é ­è¾ž:ABC â†’ ä¾�é ¼NO ã�Œã��ã�®æ–‡å­—åˆ—ã�§å§‹ã�¾ã‚‹
    regex:... / æ­£è¦�è¡¨ç�¾:... â†’ æ­£è¦�è¡¨ç�¾ï¼ˆéƒ¨åˆ†ä¸€è‡´ï¼‰
    ã��ã‚Œä»¥å¤–ã�®çŸ­æ–‡ã�¯æŽ¥é ­è¾žã�¨ã�—ã�¦æ‰±ã�†ã€‚å¾“æ�¥ã�®æ—¥æœ¬èªžä¾‹ã€Œä¾�é ¼NOã�ŒJRã�§â€¦ã€�ã�¯ JR ã‚’æ¤œå‡ºã�—ã�Ÿã‚‰æŽ¥é ­è¾žJRæ‰±ã�„ã€‚
    """
    cond = (condition_raw or "").strip()
    tid = str(task_id).strip()
    if not cond or cond in ("*", "å…¨ä»¶", "å…¨ã�¦", "any", "ANY"):
        return True
    low = cond.lower()
    cn = cond.replace("ï¼š", ":")
    if low.startswith("prefix:") or low.startswith("æŽ¥é ­è¾ž:"):
        pref = cn.split(":", 1)[1].strip() if ":" in cn else ""
        return bool(pref) and tid.startswith(pref)
    if low.startswith("regex:") or low.startswith("æ­£è¦�è¡¨ç�¾:"):
        pat = cn.split(":", 1)[1].strip() if ":" in cn else ""
        if not pat:
            return False
        try:
            return re.search(pat, tid) is not None
        except re.error:
            logging.warning(f"need ä¾�é ¼NOæ�¡ä»¶ã�®æ­£è¦�è¡¨ç�¾ã�Œç„¡åŠ¹ã�§ã�™: {pat}")
            return False
    if "ä¾�é ¼" in cond and "JR" in cond.upper():
        return tid.upper().startswith("JR")
    return tid.startswith(cond)


def parse_need_sheet_special_rules(needs_df, label_col, equipment_list, cond_col):
    """ç‰¹åˆ¥æŒ‡å®š1ï½ž99 è¡Œã�‹ã‚‰ã€�è¨­å‚™åˆ¥ã�®å¿…è¦�äººæ•°ä¸Šæ›¸ã��ï¼ˆ1ï½ž99ï¼‰ã‚’æŠ½å‡ºï¼ˆå…ˆã�«å®šç¾©ã�•ã‚Œã�Ÿç•ªå�·ã�Œå„ªå…ˆï¼‰ã€‚"""
    rules = []
    for _, row in needs_df.iterrows():
        lab = str(row.get(label_col, "") or "").strip()
        m = re.match(r"ç‰¹åˆ¥æŒ‡å®š\s*(\d+)", lab)
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
    need ã‚·ãƒ¼ãƒˆã�®ã€Œå·¥ç¨‹å�� + æ©Ÿæ¢°å��ã€�ã�§å¿…è¦�OPäººæ•°ã‚’è§£æ±ºï¼ˆç‰¹åˆ¥æŒ‡å®š1ã€œ99ã�¯ order ã�Œå°�ã�•ã�„ã�»ã�©å„ªå…ˆï¼‰ã€‚

    req_map ã�¯
      - f\"{process}+{machine_name}\"ï¼ˆåŽ³å¯†ã‚­ãƒ¼ï¼‰
      - machine_nameï¼ˆæ©Ÿæ¢°ã� ã�‘ã�®ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ï¼‰
      - processï¼ˆå·¥ç¨‹ã� ã�‘ã�®ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ï¼‰
    ã�®ã�„ã�šã‚Œã�‹ã�§ base ã‚’å¼•ã�‘ã‚‹å‰�æ��ã€‚
    need_rules ã�® overrides ã‚‚å�Œæ§˜ã�«ã‚­ãƒ¼ã‚’æŒ�ã�¤ã€‚
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
    resolve_need_required_op ã�¨å�Œå€¤ã‚’è¿”ã�—ã�¤ã�¤ã€�ãƒ­ã‚°ç”¨ã�«å�‚ç…§å…ƒã�®èª¬æ˜Žæ–‡å­—åˆ—ã‚’ä»˜ã�‘ã‚‹ã€‚
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
        base_src = f"req_map[æ©Ÿæ¢°å��ã�®ã�¿ {m!r}]={base}ï¼ˆè¤‡å�ˆã‚­ãƒ¼ä¸�åœ¨ï¼‰"
    elif p and p in req_map:
        base = req_map[p]
        base_src = f"req_map[å·¥ç¨‹å��ã�®ã�¿ {p!r}]={base}ï¼ˆè¤‡å�ˆãƒ»æ©Ÿæ¢°ã‚­ãƒ¼ä¸�åœ¨ï¼‰"
    else:
        base = 1
        base_src = "req_mapè©²å½“ã�ªã�—â†’æ—¢å®š1"
    for rule in need_rules:
        if not match_need_sheet_condition(rule["condition"], task_id):
            continue
        order = rule.get("order", "?")
        if combo_key and combo_key in rule["overrides"]:
            v = int(rule["overrides"][combo_key])
            return v, f"needç‰¹åˆ¥æŒ‡å®š{order} [{combo_key!r}]={v}"
        if m and m in rule["overrides"]:
            v = int(rule["overrides"][m])
            return v, f"needç‰¹åˆ¥æŒ‡å®š{order} [æ©Ÿæ¢°å��{m!r}]={v}"
        if p and p in rule["overrides"]:
            v = int(rule["overrides"][p])
            return v, f"needç‰¹åˆ¥æŒ‡å®š{order} [å·¥ç¨‹å��{p!r}]={v}"
    return int(base), base_src


def _need_row_label_hints_surplus_add(label_a0: str) -> bool:
    """need ã‚·ãƒ¼ãƒˆ Aåˆ—: åŸºæœ¬å¿…è¦�äººæ•°ã�®ç›´ä¸‹ã�«ã�‚ã‚‹ã€Œé…�å�°çµ�æžœã�§ä½™å‰°ã�Œå‡ºã�Ÿã�¨ã��ã�®è¿½åŠ å¢—å“¡ä¸Šé™�ã€�è¡Œã�‹ã€‚"""
    s = unicodedata.normalize("NFKC", str(label_a0 or "").strip())
    if not s or s.startswith("ç‰¹åˆ¥æŒ‡å®š"):
        return False
    if "ä¾�é ¼" in s and "æ�¡ä»¶" in s:
        return False
    if "è¿½åŠ " in s and ("äººæ•°" in s or "äººå“¡" in s or "å¢—å“¡" in s):
        return True
    if "å¢—å“¡" in s or "ä½™å‰°" in s:
        return True
    if "é…�å�°" in s and ("è¿½åŠ " in s or "å¢—" in s or "ä½™å‰°" in s):
        return True
    return False


def _find_need_surplus_add_row_index(
    needs_raw, base_row: int, col0: int, pm_cols: list
) -> int | None:
    """åŸºæœ¬å¿…è¦�äººæ•°è¡Œã�®æ¬¡è¡Œã‚’å„ªå…ˆã€‚ãƒ©ãƒ™ãƒ«ã�¾ã�Ÿã�¯æ•°å€¤ã�§è¿½åŠ äººæ•°è¡Œã�¨åˆ¤å®šã€‚"""
    r = base_row + 1
    if r >= needs_raw.shape[0]:
        return None
    v0 = needs_raw.iat[r, col0]
    s0 = "" if pd.isna(v0) else str(v0).strip()
    if s0.startswith("ç‰¹åˆ¥æŒ‡å®š"):
        return None
    if _need_row_label_hints_surplus_add(s0):
        return r
    nz = 0
    for col_idx, _, _ in pm_cols:
        if parse_optional_int(needs_raw.iat[r, col_idx]) is not None:
            nz += 1
    if nz > 0 and not unicodedata.normalize("NFKC", s0).startswith("ç‰¹åˆ¥"):
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
    need ã‚·ãƒ¼ãƒˆã€Œé…�å�°æ™‚è¿½åŠ äººæ•°ã€�è¡Œï¼ˆå·¥ç¨‹Ã—æ©Ÿæ¢°åˆ—ï¼‰ã�®å€¤ï¼�å¿…è¦�äººæ•°ã‚’æº€ã�Ÿã�—ã�Ÿã�†ã�ˆã�§
    ã�•ã‚‰ã�«å‰²ã‚Šå½“ã�¦å�¯èƒ½ã�ªäººæ•°ã�®ä¸Šé™�ï¼ˆ0 ã�ªã‚‰å¾“æ�¥ã�©ã�Šã‚Šå¿…è¦�äººæ•°ã�¡ã‚‡ã�†ã�©ã�®ã�¿ï¼‰ã€‚
    need_rules ã�¯ç�¾çŠ¶ã�“ã�®è¡Œã‚’ä¸Šæ›¸ã��ã�—ã�ªã�„ï¼ˆå°†æ�¥æ‹¡å¼µç”¨ã�« task_id ã‚’å�—ã�‘å�–ã‚‹ï¼‰ã€‚
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
    """resolve_need_surplus_extra_max ã�¨å�Œå€¤ï¼‹å�‚ç…§å…ƒèª¬æ˜Žï¼ˆãƒ­ã‚°ç”¨ï¼‰ã€‚"""
    val = resolve_need_surplus_extra_max(
        process, machine_name, task_id, surplus_map, need_rules
    )
    _ = need_rules
    if not surplus_map:
        return val, "surplus_mapç©ºï¼ˆé…�å�°æ™‚è¿½åŠ äººæ•°è¡Œã�ªã�—ï¼‰"
    p = str(process).strip()
    m = str(machine_name).strip()
    combo_key = f"{p}+{m}" if p and m else None
    if combo_key and combo_key in surplus_map:
        raw = surplus_map[combo_key]
        return val, f"surplus_map[{combo_key!r}]={raw}"
    if m and m in surplus_map:
        raw = surplus_map[m]
        return val, f"surplus_map[æ©Ÿæ¢°å��ã�®ã�¿ {m!r}]={raw}ï¼ˆè¤‡å�ˆã‚­ãƒ¼ä¸�åœ¨ï¼‰"
    if p and p in surplus_map:
        raw = surplus_map[p]
        return val, f"surplus_map[å·¥ç¨‹å��ã�®ã�¿ {p!r}]={raw}ï¼ˆè¤‡å�ˆã‚­ãƒ¼ä¸�åœ¨ï¼‰"
    return val, "surpluså½“ã‚­ãƒ¼ã�ªã�—â†’0"


def _surplus_team_time_factor(
    rq_base: int, team_len: int, extra_max_allowed: int
) -> float:
    """
    å¿…è¦�äººæ•°ã‚’è¶…ã�ˆã�¦å…¥ã‚Œã�Ÿãƒ¡ãƒ³ãƒ�ãƒ¼ã�«ã‚ˆã‚‹å�˜ä½�æ™‚é–“ã�¸ã�®ä¿‚æ•°ï¼ˆ1.0ï¼�çŸ­ç¸®ã�ªã�—ï¼‰ã€‚
    è¿½åŠ æž ï¼ˆextra_max_allowedï¼‰ã‚’ä½¿ã�„åˆ‡ã�£ã�Ÿã�¨ã��ã�§ã‚‚ã€�çŸ­ç¸®ã�¯ SURPLUS_TEAM_MAX_SPEEDUP_RATIO ã‚’ä¸Šé™�ã�¨ã�™ã‚‹ç·šå½¢ãƒ¢ãƒ‡ãƒ«ã€‚
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
        return "(-äººæ•°, é–‹å§‹, -å�˜ä½�æ•°, å„ªå…ˆåº¦å�ˆè¨ˆ)"
    if TEAM_ASSIGN_START_SLACK_WAIT_MINUTES <= 0:
        return "(é–‹å§‹, -å�˜ä½�æ•°, å„ªå…ˆåº¦å�ˆè¨ˆ)"
    return (
        f"æœ€æ—©é–‹å§‹ã�‹ã‚‰{TEAM_ASSIGN_START_SLACK_WAIT_MINUTES}åˆ†ä»¥å†…ã�¯"
        "(0,-äººæ•°,é–‹å§‹,-å�˜ä½�æ•°,å„ªå…ˆåº¦)ã€�è¶…é�Žã�¯(1,é–‹å§‹,-äººæ•°,-å�˜ä½�æ•°,å„ªå…ˆåº¦)"
    )


def _team_assignment_sort_tuple(
    team: tuple,
    team_start: datetime,
    units_today: int,
    team_prio_sum: int,
    t_min: datetime | None = None,
) -> tuple:
    """
    ãƒ�ãƒ¼ãƒ å€™è£œã�®å„ªåŠ£ç”¨ã‚¿ãƒ—ãƒ«ï¼ˆè¾žæ›¸å¼�ã�§å°�ã�•ã�„æ–¹ã�ŒæŽ¡ç”¨ï¼‰ã€‚
    - TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF: (-äººæ•°, é–‹å§‹, -å�˜ä½�æ•°, å„ªå…ˆåº¦å�ˆè¨ˆ)
    - ã��ã‚Œä»¥å¤–ã�‹ã�¤ TEAM_ASSIGN_START_SLACK_WAIT_MINUTES>0 ã�‹ã�¤ t_min ã�‚ã‚Š:
        æœ€æ—©é–‹å§‹ã�‹ã‚‰ã‚¹ãƒ©ãƒƒã‚¯ä»¥å†… â†’ (0, -äººæ•°, é–‹å§‹, -å�˜ä½�æ•°, å„ªå…ˆåº¦) â€¦ é�…ã‚Œã�¦ã‚‚äººæ•°ã‚’åŽšã��
        ã‚¹ãƒ©ãƒƒã‚¯è¶… â†’ (1, é–‹å§‹, -äººæ•°, -å�˜ä½�æ•°, å„ªå…ˆåº¦) â€¦ é–‹å§‹ã‚’å„ªå…ˆ
    - ä¸Šè¨˜ä»¥å¤–: (é–‹å§‹, -å�˜ä½�æ•°, å„ªå…ˆåº¦å�ˆè¨ˆ)
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


# skills ã‚»ãƒ«: OP / AS + ä»»æ„�ã�®å„ªå…ˆåº¦æ•´æ•°ï¼ˆä¾‹ OP1, AS 3ï¼‰ã€‚æ•°å€¤ã�Œå°�ã�•ã�„ã�»ã�©å‰²å½“ã�§å…ˆã�«é�¸ã�°ã‚Œã‚‹ã€‚
_SKILL_OP_AS_CELL_RE = re.compile(r"^(OP|AS)(\d*)$", re.IGNORECASE)


def parse_op_as_skill_cell(cell_val):
    """
    master.xlsmã€Œskillsã€�ã�®ã‚»ãƒ«1ã�¤ã‚’è§£é‡ˆã�™ã‚‹ã€‚
    - ã€ŒOPã€�ã�¾ã�Ÿã�¯ã€ŒASã€�ã�®ç›´å¾Œã�«å„ªå…ˆåº¦ç”¨ã�®æ•´æ•°ï¼ˆç©ºç™½ã�¯é™¤åŽ»ã�—ã�¦è§£é‡ˆï¼‰ã€‚ä¾‹: OP, OP1, AS3, AS 12
    - å„ªå…ˆåº¦ã�¯å°�ã�•ã�„ã�»ã�©é«˜å„ªå…ˆï¼ˆå�Œä¸€æ�¡ä»¶ã�®ãƒ�ãƒ¼ãƒ å€™è£œã�‹ã‚‰å…ˆã�«é�¸ã�°ã‚Œã‚‹ï¼‰ã€‚æ•°å­—çœ�ç•¥æ™‚ã�¯ 1ã€‚
    - OP/AS ã�§å§‹ã�¾ã‚‰ã�ªã�„ãƒ»ç©ºã�¯ã‚¹ã‚­ãƒ«ã�ªã�—ã€‚
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
    masterã€Œskillsã€�ã�®å�„åˆ—ï¼ˆå·¥ç¨‹+æ©Ÿæ¢°ã‚­ãƒ¼ç­‰ï¼‰ã�«ã�¤ã�„ã�¦ã€�OP/AS ã�®å‰²å½“å„ªå…ˆåº¦ã�®**æ•°å€¤**ã�Œ
    ãƒ¡ãƒ³ãƒ�ãƒ¼é–“ã�§é‡�è¤‡ã�—ã�¦ã�„ã�ªã�„ã�‹æ¤œè¨¼ã�™ã‚‹ã€‚é‡�è¤‡æ™‚ã�¯ PlanningValidationErrorã€‚
    ï¼ˆOP1 ã�¨ AS1 ã�®ã‚ˆã�†ã�«ãƒ­ãƒ¼ãƒ«ã�Œç•°ã�ªã�£ã�¦ã‚‚å�Œä¸€æ•°å€¤ã�ªã‚‰é‡�è¤‡ã�¨ã�¿ã�ªã�™ï¼‰
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
                errors.append(f'åˆ—ã€Œ{ck}ã€�: å„ªå…ˆåº¦ {pr} ã�Œé‡�è¤‡ â†’ ' + "ã€�".join(entries))
    if errors:
        cap = 50
        tail = errors[:cap]
        msg = (
            "ãƒžã‚¹ã‚¿ã€Œskillsã€�ã�§ã€�å�Œä¸€åˆ—ã�® OP/AS å„ªå…ˆåº¦ã�®æ•°å€¤ã�Œé‡�è¤‡ã�—ã�¦ã�„ã�¾ã�™ã€‚"
            " åˆ—ã�”ã�¨ã�«æ•°å€¤ã�¯1äººã�«ã�¤ã��1ç¨®é¡žã�«ã�—ã�¦ã��ã� ã�•ã�„ã€‚\n"
            + "\n".join(tail)
        )
        if len(errors) > cap:
            msg += f"\nâ€¦ä»– {len(errors) - cap} ä»¶"
        raise PlanningValidationError(msg)


def build_member_assignment_priority_reference(
    skills_dict: dict,
    members: list | None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    çµ�æžœãƒ–ãƒƒã‚¯ç”¨: ãƒžã‚¹ã‚¿ skills ã�®ã€Œå·¥ç¨‹å��+æ©Ÿæ¢°å��ã€�åˆ—ã�”ã�¨ã�«ã€�å‰²å½“ã‚¢ãƒ«ã‚´ãƒªã‚ºãƒ ã�¨å�Œã�˜
    (å„ªå…ˆåº¦å€¤æ˜‡é †, ãƒ¡ãƒ³ãƒ�ãƒ¼å��æ˜‡é †) ã�§ä¸¦ã�¹ã�Ÿå�‚è€ƒè¡¨ã�¨ã€�ãƒ«ãƒ¼ãƒ«èª¬æ˜Žã�®è¡¨ã‚’è¿”ã�™ã€‚
    å½“æ—¥ã�®å‡ºå‹¤ãƒ»è¨­å‚™ç©ºã��ãƒ»å�Œä¸€ä¾�é ¼ã�®å·¥ç¨‹é †ãƒ»ãƒ�ãƒ¼ãƒ äººæ•°ã�¯å��æ˜ ã�—ã�ªã�„ï¼ˆã�‚ã��ã�¾ã�§ãƒžã‚¹ã‚¿ä¸Šã�®é †åº�ï¼‰ã€‚
    """
    mem_list = list(members) if members else list((skills_dict or {}).keys())
    mem_list = [str(m).strip() for m in mem_list if m and str(m).strip()]

    surplus_on = bool(TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF)
    slack_m = TEAM_ASSIGN_START_SLACK_WAIT_MINUTES
    if surplus_on:
        team_rule = (
            "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF=æœ‰åŠ¹: "
            "(-äººæ•°, é–‹å§‹, -å�˜ä½�æ•°, å„ªå…ˆåº¦å�ˆè¨ˆ) ã�®è¾žæ›¸å¼�ï¼ˆäººæ•°æœ€å„ªå…ˆãƒ»å¾“æ�¥ï¼‰ã€‚"
        )
    elif slack_m > 0:
        team_rule = (
            f"æ—¢å®š: ã��ã�®æ—¥ã�®æˆ�ç«‹å€™è£œå…¨ä½“ã�®ã€Œæœ€æ—©é–‹å§‹ã€�ã‚’åŸºæº–ã�«ã€�"
            f"é–‹å§‹ã�Œã��ã�®{slack_m}åˆ†ä»¥å†…ã�®é�…ã‚Œã�ªã‚‰äººæ•°ã‚’åŽšã��å„ªå…ˆï¼ˆ0,-äººæ•°,é–‹å§‹,-å�˜ä½�æ•°,å„ªå…ˆåº¦ï¼‰ã€�"
            f"ã��ã‚Œã‚ˆã‚Šé�…ã�„å€™è£œã�¯é–‹å§‹ã‚’å„ªå…ˆï¼ˆ1,é–‹å§‹,-äººæ•°,-å�˜ä½�æ•°,å„ªå…ˆåº¦ï¼‰ã€‚"
            f"ç’°å¢ƒå¤‰æ•° TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=0 ã�§ç„¡åŠ¹åŒ–ã€‚"
        )
    else:
        team_rule = (
            "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=0: "
            "(é–‹å§‹, -å�˜ä½�æ•°, å„ªå…ˆåº¦å�ˆè¨ˆ) ã�®ã�¿ï¼ˆé–‹å§‹æœ€å„ªå…ˆï¼‰ã€‚"
        )

    legend_rows = [
        {
            "åŒºåˆ†": "ã‚¹ã‚­ãƒ«åˆ—ã�®ä¸¦ã�³",
            "å†…å®¹": "å�„ã€Œå·¥ç¨‹å��+æ©Ÿæ¢°å��ã€�åˆ—ã�«ã�¤ã�„ã�¦ã€�ã‚»ãƒ«ã�Œ OP/ASï¼ˆ+å„ªå…ˆåº¦æ•´æ•°ï¼‰ã�®ãƒ¡ãƒ³ãƒ�ãƒ¼ã�®ã�¿å¯¾è±¡ã€‚"
            " æ•°å€¤ã�Œå°�ã�•ã�„ã�»ã�©é«˜å„ªå…ˆã€‚çœ�ç•¥æ™‚ã�¯å„ªå…ˆåº¦ 1ï¼ˆparse_op_as_skill_cell ã�¨å�Œä¸€ï¼‰ã€‚"
            " å�Œä¸€åˆ—ã�§ã�¯å„ªå…ˆåº¦ã�®æ•°å€¤ã�¯ãƒ¡ãƒ³ãƒ�ãƒ¼é–“ã�§é‡�è¤‡ä¸�å�¯ï¼ˆãƒžã‚¹ã‚¿èª­è¾¼æ™‚ã�«æ¤œè¨¼ï¼‰ã€‚",
        },
        {
            "åŒºåˆ†": "å½“æ—¥ã�¨ã�®å·®",
            "å†…å®¹": "å®Ÿéš›ã�®é…�å�°ã�¯ã€�ã�“ã�®é †ã�®ã�†ã�¡ã��ã�®æ—¥å‡ºå‹¤ã�‹ã�¤ AS/OP è¦�ä»¶ã‚’æº€ã�Ÿã�™è€…ã� ã�‘ã�Œå€™è£œã€‚"
            " è¨­å‚™ã�®ç©ºã��ãƒ»å�Œä¸€ä¾�é ¼NOã�®å·¥ç¨‹é †ãƒ»å¿…è¦�äººæ•°ãƒ»å¢—å“¡æž ãƒ»æŒ‡å��OPã�§å¤‰ã‚�ã‚Šã�¾ã�™ã€‚",
        },
        {
            "åŒºåˆ†": "ãƒ�ãƒ¼ãƒ å€™è£œã�®æ¯”è¼ƒ",
            "å†…å®¹": team_rule,
        },
        {
            "åŒºåˆ†": "æŒ‡å��ãƒ»ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ä¸Šæ›¸ã��",
            "å†…å®¹": "æ‹…å½“OP_æŒ‡å®šãƒ»ãƒ¡ã‚¤ãƒ³ã€Œå†�å„ªå…ˆç‰¹åˆ¥è¨˜è¼‰ã€�ã�® OP æŒ‡å��ã�¯æœ¬è¡¨ã‚ˆã‚Šå„ªå…ˆã�•ã‚Œã�¾ã�™ã€‚",
        },
        {
            "åŒºåˆ†": "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF",
            "å†…å®¹": "1/æœ‰åŠ¹ï¼ˆäººæ•°æœ€å„ªå…ˆãƒ»å¾“æ�¥ï¼‰" if surplus_on else "0/ç„¡åŠ¹ï¼ˆæ—¢å®šï¼‰",
        },
        {
            "åŒºåˆ†": "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES",
            "å†…å®¹": str(slack_m),
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
                    "å·¥ç¨‹å��": proc,
                    "æ©Ÿæ¢°å��": mach,
                    "ã‚¹ã‚­ãƒ«åˆ—ã‚­ãƒ¼": combo,
                    "å„ªå…ˆé †ä½�": "",
                    "ãƒ¡ãƒ³ãƒ�ãƒ¼": "ï¼ˆã�ªã�—ï¼‰",
                    "ãƒ­ãƒ¼ãƒ«": "",
                    "å„ªå…ˆåº¦å€¤_å°�ã�•ã�„ã�»ã�©å…ˆ": "",
                    "skillsã‚»ãƒ«å€¤": "",
                    "å‚™è€ƒ": "ã�“ã�®åˆ—ã�« OP/AS ã�®è³‡æ ¼ã‚»ãƒ«ã�Œã�‚ã‚‹ãƒ¡ãƒ³ãƒ�ãƒ¼ã�Œã�„ã�¾ã�›ã‚“",
                }
            )
            continue
        for i, (pr, m, role, cell_s) in enumerate(ranked, start=1):
            out.append(
                {
                    "å·¥ç¨‹å��": proc,
                    "æ©Ÿæ¢°å��": mach,
                    "ã‚¹ã‚­ãƒ«åˆ—ã‚­ãƒ¼": combo,
                    "å„ªå…ˆé †ä½�": i,
                    "ãƒ¡ãƒ³ãƒ�ãƒ¼": m,
                    "ãƒ­ãƒ¼ãƒ«": role,
                    "å„ªå…ˆåº¦å€¤_å°�ã�•ã�„ã�»ã�©å…ˆ": pr,
                    "skillsã‚»ãƒ«å€¤": cell_s,
                    "å‚™è€ƒ": "",
                }
            )

    df_tbl = pd.DataFrame(out)
    return df_legend, df_tbl


def _normalize_person_name_for_match(s):
    """æ‹…å½“è€…æŒ‡å��ã�®ã�‚ã�„ã�¾ã�„ä¸€è‡´ç”¨ï¼ˆNFKCãƒ»å¯Œç”°/å†¨ç”°ã�®è¡¨è¨˜å¯„ã�›ãƒ»ç©ºç™½é™¤åŽ»ãƒ»æœ«å°¾æ•¬ç§°ã�®ã�¿é™¤åŽ»ï¼‰ã€‚"""
    if s is None:
        return ""
    t = unicodedata.normalize("NFKC", str(s).strip())
    if "å¯Œç”°" in t:
        t = t.replace("å¯Œç”°", "å†¨ç”°")
    t = re.sub(r"[\sã€€]+", "", t)
    t = re.sub(r"(ã�•ã‚“|æ§˜|æ°�)$", "", t)
    return t


def _split_person_sei_mei(s) -> tuple[str, str]:
    """
    æ°�å��ã‚’å§“ãƒ»å��ã�«åˆ†ã�‘ã‚‹ã€‚æœ€åˆ�ã�®å�Šè§’ï¼�å…¨è§’ç©ºç™½ã�®æ‰‹å‰�ã‚’å§“ã€�ä»¥é™�ã‚’å��ã�¨ã�™ã‚‹ã€‚
    ç©ºç™½ã�Œç„¡ã�„å ´å�ˆã�¯ (å…¨ä½“, '')ï¼ˆå��ã�ªã�—æ‰±ã�„ï¼‰ã€‚
    æœ«å°¾ã�® ã�•ã‚“ï¼�æ§˜ï¼�æ°� ã�¯åˆ†å‰²å‰�ã�«é™¤åŽ»ã�™ã‚‹ã€‚
    """
    if s is None:
        return "", ""
    t = unicodedata.normalize("NFKC", str(s).strip())
    if not t or t.lower() in ("nan", "none", "null"):
        return "", ""
    t = re.sub(r"(ã�•ã‚“|æ§˜|æ°�)$", "", t)
    for i, ch in enumerate(t):
        if ch in " \u3000":
            sei = t[:i].strip()
            rest = t[i + 1 :]
            mei = re.sub(r"[\sã€€]+", "", rest.strip())
            return sei, mei
    return t.strip(), ""


def _normalize_sei_for_match(sei: str) -> str:
    """å§“ã�®ã�¿æ­£è¦�åŒ–ã€‚è¡¨è¨˜ã‚†ã‚Œã�¯è¨±å®¹ã�—ã�ªã�„å‰�æ��ã�§ã€�NFKCãƒ»å¯Œç”°/å†¨ç”°å¯„ã�›ãƒ»ç©ºç™½é™¤åŽ»ã€‚"""
    if not sei:
        return ""
    t = unicodedata.normalize("NFKC", str(sei).strip())
    if "å¯Œç”°" in t:
        t = t.replace("å¯Œç”°", "å†¨ç”°")
    t = re.sub(r"[\sã€€]+", "", t)
    return t


def _normalize_mei_for_match(mei: str) -> str:
    """å��ã�®æ­£è¦�åŒ–ï¼ˆã‚†ã‚Œè¨±å®¹ã�®å‰�å‡¦ç�†ï¼‰ã€‚NFKCãƒ»ç©ºç™½é™¤åŽ»ã€‚å§“ç”¨ã�®å¯Œç”°ç½®æ�›ã�¯è¡Œã‚�ã�ªã�„ã€‚"""
    if not mei:
        return ""
    t = unicodedata.normalize("NFKC", str(mei).strip())
    t = re.sub(r"[\sã€€]+", "", t)
    return t


def _has_duplicate_surname_among_members(member_names) -> bool:
    """skills ãƒ¡ãƒ³ãƒ�ãƒ¼ä¸€è¦§ã�«ã€�æ­£è¦�åŒ–å¾Œå�Œä¸€ã�®å§“ã�Œ2äººä»¥ä¸Šã�„ã‚‹ã�‹ã€‚"""
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
    """å�Œä¸€å§“ã�Œãƒ­ã‚¹ã‚¿ãƒ¼ã�§é‡�è¤‡ã�—ã�ªã�„ã�¨ã��ã�®ã�¿ä½¿ã�†å��ã�®ã‚†ã‚Œè¨±å®¹ã€‚"""
    if not r_mei_n and not m_mei_n:
        return True
    if not r_mei_n or not m_mei_n:
        return False
    if r_mei_n == m_mei_n:
        return True
    return r_mei_n in m_mei_n or m_mei_n in r_mei_n


def _resolve_preferred_name_to_capable_member(raw, capable_candidates, roster_member_names=None):
    """
    è‡ªç”±è¨˜è¿°ã�®æŒ‡å��ã‚’ã€�å½“æ—¥ã‚¹ã‚­ãƒ«ä¸Š OP/AS ã�®ãƒ¡ãƒ³ãƒ�ãƒ¼å��ï¼ˆskills ã‚·ãƒ¼ãƒˆã�®è¡Œã‚­ãƒ¼ï¼‰ã�«è§£æ±ºã�™ã‚‹ã€‚
    capable_candidates: ã��ã�®è¨­å‚™ã�§ OP ã�¾ã�Ÿã�¯ AS ã�¨ã�—ã�¦å‰²å½“å�¯èƒ½ã�ªãƒ¡ãƒ³ãƒ�ãƒ¼å��ãƒªã‚¹ãƒˆã€‚
    roster_member_names: skills ã�®å…¨ãƒ¡ãƒ³ãƒ�ãƒ¼å��ï¼ˆçœ�ç•¥æ™‚ã�¯ capable_candidatesï¼‰ã€‚å�Œä¸€å§“ã�®é‡�è¤‡åˆ¤å®šã�«ä½¿ç”¨ã€‚

    å��å‰�ã�®è¡¨è¨˜ã‚†ã‚Œ:
    - å§“ã�¯æ­£è¦�åŒ–å¾Œã�«å®Œå…¨ä¸€è‡´ã�®ã�¿ï¼ˆã‚†ã‚Œè¨±å®¹ã�—ã�ªã�„ã€‚å¯Œç”°/å†¨ç”°ã�®ã�¿å¾“æ�¥ã�©ã�Šã‚Šå¯„ã�›ï¼‰ã€‚
    - roster ã�«å�Œä¸€å§“ã�Œ2äººä»¥ä¸Šã�„ã�ªã�„ã�¨ã��ã� ã�‘ã€�å��ã�¯éƒ¨åˆ†ä¸€è‡´ï¼ˆã�©ã�¡ã‚‰ã�‹ã�Œä»–æ–¹ã‚’å�«ã‚€ï¼‰ã�¾ã�Ÿã�¯å®Œå…¨ä¸€è‡´ã‚’è¨±å®¹ã€‚
    - å�Œä¸€å§“ã�Œãƒ­ã‚¹ã‚¿ãƒ¼ã�«ã�„ã‚‹é–“ã�¯å��ã‚‚å®Œå…¨ä¸€è‡´å¿…é ˆã€‚
    - å§“ã�®ã�¿ã�®å…¥åŠ›ã�§å��ã‚†ã‚Œãƒ¢ãƒ¼ãƒ‰ã�®ã�¨ã��ã€�å§“ã�Œä¸€è‡´ã�™ã‚‹å€™è£œã�Œè¤‡æ•°ã�„ã‚Œã�°è§£æ±ºä¸�å�¯ï¼ˆNoneï¼‰ã€‚
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
    """å½“æ—¥ã‚¹ã‚­ãƒ«ä¸Š OP ã�®ã�¿ã�¸è§£æ±ºï¼ˆå¾“æ�¥ APIï¼‰ã€‚å®Ÿä½“ã�¯ `_resolve_preferred_name_to_capable_member`ã€‚"""
    return _resolve_preferred_name_to_capable_member(
        raw, op_candidates, roster_member_names
    )


def _task_process_matches_global_contains(machine_val: str, contains: str) -> bool:
    """å·¥ç¨‹å��ï¼ˆã‚¿ã‚¹ã‚¯ã�® machineï¼‰ã�«éƒ¨åˆ†ä¸€è‡´ï¼ˆNFKCãƒ»å¤§å°�ç„¡è¦–ï¼‰ã€‚"""
    m = unicodedata.normalize("NFKC", str(machine_val or "").strip()).casefold()
    c = unicodedata.normalize("NFKC", str(contains or "").strip()).casefold()
    if not c:
        return False
    return c in m


def _coerce_global_day_process_operator_rules(raw_val) -> list:
    """Gemini ã�® global_day_process_operator_rules ã‚’æ­£è¦�åŒ–ï¼ˆç©ºãƒ»ä¸�æ­£ã�¯é™¤å¤–ï¼‰ã€‚"""
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
    ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆç”±æ�¥ã�®ã€Œæ—¥ä»˜Ã—å·¥ç¨‹Ã—è¤‡æ•°æŒ‡å��ã€�ã�§ã€�ã��ã�®æ—¥ãƒ»ã��ã�®å·¥ç¨‹ã‚¿ã‚¹ã‚¯ã�«
    **ãƒ�ãƒ¼ãƒ ã�¸å¿…ã�šå�«ã‚�ã‚‹**ãƒ¡ãƒ³ãƒ�ãƒ¼ï¼ˆskills è¡Œã‚­ãƒ¼ï¼‰ã�¨è­¦å‘Šãƒ¡ãƒƒã‚»ãƒ¼ã‚¸ã‚’è¿”ã�™ã€‚
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
                    "ãƒ¡ã‚¤ãƒ³ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«(æ—¥ä»˜Ã—å·¥ç¨‹)æŒ‡å��: "
                    f"ä¾�é ¼NO={tid} æ—¥ä»˜={current_date} å·¥ç¨‹={machine!r} ã�® "
                    f"æŒ‡å��ã€Œ{raw_name}ã€�ã‚’å½“æ—¥ã‚¹ã‚­ãƒ«è©²å½“ãƒ¡ãƒ³ãƒ�ãƒ¼ã�«è§£æ±ºã�§ã��ã�¾ã�›ã‚“"
                )
    return acc, warns


def _merge_global_day_process_and_pref_anchor(
    must_include: list, pref_mem, capable_members: list
) -> list[str]:
    """å¿…é ˆãƒ¡ãƒ³ãƒ�ãƒ¼ã�¨æ‹…å½“OPæŒ‡å��ã‚’1æœ¬åŒ–ï¼ˆcapable ã�«ã�„ã‚‹ã‚‚ã�®ã� ã�‘ï¼‰ã€‚"""
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
# 2. ãƒžã‚¹ã‚¿ãƒ‡ãƒ¼ã‚¿ãƒ»å‡ºå‹¤ç°¿(ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼) ã�¨ AIè§£æž�
#    master.xlsm ã�® skills / need / å�„ãƒ¡ãƒ³ãƒ�ãƒ¼å‹¤æ€ ã‚·ãƒ¼ãƒˆã‚’èª­ã�¿ã€�
#    å‚™è€ƒãƒ»ä¼‘æš‡åŒºåˆ†ã�¯å¿…è¦�ã�«å¿œã�˜ã�¦ Gemini ã�§æ§‹é€ åŒ–ã�™ã‚‹ã€‚
# =========================================================
def load_skills_and_needs():
    """
    çµ±å�ˆãƒ•ã‚¡ã‚¤ãƒ«(MASTER_FILE)ã�‹ã‚‰ã‚¹ã‚­ãƒ«ã�¨ need ã‚’å‹•çš„ã�«èª­ã�¿è¾¼ã�¿ã�¾ã�™ã€‚

    æˆ»ã‚Šå€¤ã�¯7è¦�ç´ ã€‚æœ€å¾Œã�¯ need ã‚·ãƒ¼ãƒˆä¸Šã�®ã€Œå·¥ç¨‹å��+æ©Ÿæ¢°å��ã€�åˆ—ä½�ç½®ï¼ˆå·¦ã�»ã�©å°�ã�•ã�„æ•´æ•°ï¼‰ã�®è¾žæ›¸
    ``need_combo_col_index``ï¼ˆé…�å�°ã‚­ãƒ¥ãƒ¼ã‚½ãƒ¼ãƒˆç”¨ï¼‰ã€‚

    ä»Šå›žã�® need ã�¯ï¼ˆExcelä¸Šã�§ï¼‰
      å·¥ç¨‹å��è¡Œãƒ»æ©Ÿæ¢°å��è¡Œã�®ã�‚ã�¨ã€ŒåŸºæœ¬å¿…è¦�äººæ•°ã€�è¡Œï¼ˆAåˆ—ã�«ã€Œå¿…è¦�äººæ•°ã€�ã‚’å�«ã‚€ï¼‰
      ã��ã�®ç›´ä¸‹: é…�å�°ã�§ä½™å‰°äººå“¡ã�Œã�‚ã‚‹ã�¨ã��ã�«è¿½åŠ ã�§å…¥ã‚Œã‚‰ã‚Œã‚‹äººæ•°ï¼ˆå·¥ç¨‹Ã—æ©Ÿæ¢°ã�”ã�¨ã€‚æœªè¨­å®šã�¯ 0ï¼‰
      ä»¥é™�: ç‰¹åˆ¥æŒ‡å®š1ã€œ99
    ã�¨ã�„ã�†æ§‹é€ ã�®ã�Ÿã‚�ã€�å¿…è¦�OPã�¯ã€Œå·¥ç¨‹å��+æ©Ÿæ¢°å��ã€�ã�§è§£æ±ºã�™ã‚‹ã€‚

    skills äº¤å·®ã‚»ãƒ«ã�¯ OP/AS ã�®å¾Œã�«å„ªå…ˆåº¦æ•´æ•°ï¼ˆä¾‹ OP1, AS3ï¼‰ã€‚æ•°å€¤ã�Œå°�ã�•ã�„ã�»ã�©å½“è©²å·¥ç¨‹ã�¸ã�®å‰²å½“ã�§å„ªå…ˆã€‚
    æ•°å­—çœ�ç•¥ã�® OP/AS ã�¯å„ªå…ˆåº¦ 1ã€‚
    å�Œä¸€åˆ—ï¼ˆå�Œä¸€å·¥ç¨‹Ã—æ©Ÿæ¢°ï¼‰ã�§ã�¯å„ªå…ˆåº¦ã�®æ•°å€¤ã�¯ãƒ¡ãƒ³ãƒ�ãƒ¼é–“ã�§é‡�è¤‡ä¸�å�¯ï¼ˆé‡�è¤‡æ™‚ã�¯ PlanningValidationErrorï¼‰ã€‚
    """
    try:
        # å�Œä¸€ãƒ–ãƒƒã‚¯ã‚’ pd.read_excel ã�§éƒ½åº¦é–‹ã��ã�¨ I/O ã�Œé‡�ã�„ã�Ÿã‚�ã€�ExcelFile ã‚’1å›žã� ã�‘é–‹ã�„ã�¦ã‚·ãƒ¼ãƒˆã‚’ parse ã�™ã‚‹ã€‚
        with pd.ExcelFile(MASTER_FILE) as _master_xls:
            # skills ã�¯æ–°ä»•æ§˜:
            #   1è¡Œç›®: å·¥ç¨‹å��
            #   2è¡Œç›®: æ©Ÿæ¢°å��
            #   A3ä»¥é™�: ãƒ¡ãƒ³ãƒ�ãƒ¼å��
            #   äº¤å·®ã‚»ãƒ«: OP ã�¾ã�Ÿã�¯ AS ã�®å¾Œã�«å‰²å½“å„ªå…ˆåº¦ã�®æ•´æ•°ï¼ˆä¾‹ OP1, AS3ï¼‰ã€‚æ•°å€¤ã�Œå°�ã�•ã�„ã�»ã�©å½“è©²å·¥ç¨‹ã�¸å„ªå…ˆå‰²å½“ã€‚
            #             æ•°å­—çœ�ç•¥ã�® OP/AS ã�¯å„ªå…ˆåº¦ 1ï¼ˆå¾“æ�¥ã�©ã�Šã‚Šæœ€å„ªå…ˆæ‰±ã�„ï¼‰ã€‚
            # ã‚’åŸºæœ¬ã�¨ã�—ã�¤ã�¤ã€�æ—§ä»•æ§˜ï¼ˆ1è¡Œãƒ˜ãƒƒãƒ€ï¼‰ã�«ã‚‚ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯å¯¾å¿œã�™ã‚‹ã€‚
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
                    "skillsã‚·ãƒ¼ãƒˆ: 2æ®µãƒ˜ãƒƒãƒ€å½¢å¼�ã�§èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿï¼ˆå·¥ç¨‹+æ©Ÿæ¢°=%såˆ—, ãƒ¡ãƒ³ãƒ�ãƒ¼=%säººï¼‰ã€‚",
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
                    if c in ("ãƒ¡ãƒ³ãƒ�ãƒ¼", "æ‹…å½“è€…", "æ°�å��", "ä½œæ¥­è€…"):
                        member_col = c
                        break
                if member_col is None and skill_cols:
                    member_col = skill_cols[0]
                    logging.warning(
                        "skillsã‚·ãƒ¼ãƒˆ: ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ—å��ã�Œæ¨™æº–ã�¨ä¸€è‡´ã�—ã�ªã�„ã�Ÿã‚�ã€�å…ˆé ­åˆ— '%s' ã‚’ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ—ã�¨ã�—ã�¦æ‰±ã�„ã�¾ã�™ã€‚",
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
                    "skillsã‚·ãƒ¼ãƒˆ: 1è¡Œãƒ˜ãƒƒãƒ€å½¢å¼�ï¼ˆæ—§äº’æ�›ï¼‰ã�§èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿï¼ˆãƒ¡ãƒ³ãƒ�ãƒ¼=%säººï¼‰ã€‚",
                    len(members),
                )

            if not members:
                logging.error("skillsã‚·ãƒ¼ãƒˆã�‹ã‚‰ãƒ¡ãƒ³ãƒ�ãƒ¼ã‚’èª­ã�¿è¾¼ã‚�ã�¾ã�›ã‚“ã�§ã�—ã�Ÿã€‚")
            else:
                _validate_skills_op_as_priority_numbers_unique(
                    skills_dict, equipment_list
                )

            # need ã�¯ header=None ã�§èª­ã�¿ã€�å…ˆé ­ã�®è¤‡æ•°è¡Œã‚’â€œè¦‹å‡ºã�—è¡Œâ€�ã�¨ã�—ã�¦è§£é‡ˆ
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
            if process_header_row is None and s0 == "å·¥ç¨‹å��":
                process_header_row = r
            elif machine_header_row is None and s0 == "æ©Ÿæ¢°å��":
                machine_header_row = r
            if base_row is None and "å¿…è¦�äººæ•°" in s0 and not s0.startswith("ç‰¹åˆ¥æŒ‡å®š"):
                base_row = r
            if process_header_row is not None and machine_header_row is not None and base_row is not None:
                break

        if process_header_row is None or machine_header_row is None or base_row is None:
            raise ValueError("need ã‚·ãƒ¼ãƒˆã�®ãƒ˜ãƒƒãƒ€ãƒ¼è¡Œï¼ˆå·¥ç¨‹å��/æ©Ÿæ¢°å��/åŸºæœ¬å¿…è¦�äººæ•°ï¼‰ã�Œè¦‹ã�¤ã�‹ã‚Šã�¾ã�›ã‚“ã€‚")

        # ã€Œä¾�é ¼NOæ�¡ä»¶ã€�åˆ—ä½�ç½®ï¼ˆãƒ‡ãƒ•ã‚©ãƒ«ãƒˆã�¯ 1åˆ—ç›®ï¼‰
        cond_col_idx = 1
        for r in range(needs_raw.shape[0]):
            c1 = needs_raw.iat[r, 1] if needs_raw.shape[1] > 1 else None
            c2 = needs_raw.iat[r, 2] if needs_raw.shape[1] > 2 else None
            if pd.isna(c1) or pd.isna(c2):
                continue
            if str(c1).strip() == NEED_COL_CONDITION and str(c2).strip() == NEED_COL_NOTE:
                cond_col_idx = 1
                break

        # å·¥ç¨‹å��Ã—æ©Ÿæ¢°å�� ã�®åˆ—ä¸€è¦§ï¼ˆåˆ—ç•ªå�·ã�¯ Excelä¸Šã�®å®Ÿåˆ—ã‚’ä¿�æŒ�ï¼‰
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
        # å·¥ç¨‹å��+æ©Ÿæ¢°å��ã‚³ãƒ³ãƒœ â†’ need ã‚·ãƒ¼ãƒˆä¸Šã�®åˆ—ã‚¤ãƒ³ãƒ‡ãƒƒã‚¯ã‚¹ï¼ˆå·¦ã�»ã�©å°�ã�•ã�„ï¼�é…�å�°ã‚­ãƒ¥ãƒ¼ã�§å…ˆï¼‰
        need_combo_col_index: dict[str, int] = {}
        # need_rules: [{'order': int, 'condition': str, 'overrides': {combo_key/machine/process: int}}]
        need_rules = []

        # åŸºæœ¬å¿…è¦�äººæ•°
        for col_idx, p_s, m_s in pm_cols:
            n = parse_optional_int(needs_raw.iat[base_row, col_idx])
            if n is None or n < 1:
                n = 1
            combo_key = f"{p_s}+{m_s}"
            need_combo_col_index[combo_key] = col_idx
            req_map[combo_key] = n
            # ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ç”¨ï¼ˆæ©Ÿæ¢°å�� or å·¥ç¨‹å��ã� ã�‘ã�§å¼•ã�‘ã‚‹ã‚ˆã�†ã�«ã�™ã‚‹ï¼‰
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
                "need ã‚·ãƒ¼ãƒˆ: é…�å�°æ™‚è¿½åŠ äººæ•°è¡Œã‚’æ¤œå‡ºï¼ˆExcelè¡Œâ‰ˆ%sï¼‰ã€‚åˆ—ã�”ã�¨ã�®ä¸Šé™�ã‚’èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚",
                surplus_row + 1,
            )
        else:
            logging.info(
                "need ã‚·ãƒ¼ãƒˆ: åŸºæœ¬å¿…è¦�äººæ•°ã�®ç›´ä¸‹ã�«é…�å�°æ™‚è¿½åŠ äººæ•°è¡Œã‚’æ¤œå‡ºã�§ã��ã�¾ã�›ã‚“ã�§ã�—ã�Ÿï¼ˆçœ�ç•¥å�¯ï¼‰ã€‚"
            )

        if TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW:
            logging.info(
                "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW ã�Œæœ‰åŠ¹: é…�å�°æ™‚è¿½åŠ äººæ•°ã�¯èª­ã�¿è¾¼ã‚“ã�§ã‚‚å¸¸ã�« 0 æ‰±ã�„ï¼ˆãƒ�ãƒ¼ãƒ ã�¯åŸºæœ¬å¿…è¦�äººæ•°ã�®ã�¿è©¦è¡Œï¼‰ã€‚"
            )

        logging.info(
            "needäººæ•°ãƒžã‚¹ã‚¿: %s ã�® need ã‚·ãƒ¼ãƒˆã‚’èª­ã�¿è¾¼ã�¿ï¼ˆskills ã�¨å�Œä¸€ ExcelFile ã�§é–‹ã�„ã�Ÿç›´å¾Œã€‚need å°‚ç”¨ãƒ‡ã‚£ã‚¹ã‚¯ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã�¯ç„¡ã�—ãƒ»AI json ã�¨ã�¯ç„¡é–¢ä¿‚ï¼‰ã€‚",
            os.path.abspath(MASTER_FILE),
        )
        for _ci, _ps, _ms in pm_cols:
            _ck = f"{_ps}+{_ms}"
            _bn = req_map.get(_ck)
            _sx = surplus_map.get(_ck, 0) if surplus_map else 0
            logging.info(
                "needåˆ—ã‚µãƒžãƒª combo=%r åŸºæœ¬å¿…è¦�äººæ•°=%s é…�å�°æ™‚è¿½åŠ äººæ•°ä¸Šé™�=%s",
                _ck,
                _bn,
                _sx,
            )

        # ç‰¹åˆ¥æŒ‡å®š
        for r in range(needs_raw.shape[0]):
            v0 = needs_raw.iat[r, col0]
            if pd.isna(v0):
                continue
            lab = str(v0).strip()
            m = re.match(r"ç‰¹åˆ¥æŒ‡å®š\s*(\d+)", lab)
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
                    # ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ç”¨
                    overrides[p_s] = n
                    overrides[m_s] = n

            if overrides:
                need_rules.append({"order": order, "condition": cond, "overrides": overrides})

        need_rules.sort(key=lambda rr: rr["order"])
        logging.info(f"need ç‰¹åˆ¥æŒ‡å®šãƒ«ãƒ¼ãƒ«: {len(need_rules)} ä»¶ï¼ˆå·¥ç¨‹å��+æ©Ÿæ¢°å��ã‚­ãƒ¼ï¼‰ã€‚")

        logging.info(f"ã€Ž{MASTER_FILE}ã€�ã�‹ã‚‰ã‚¹ã‚­ãƒ«ã�¨è¨­å‚™è¦�ä»¶(need)ã‚’èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚")
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
        logging.error(f"ãƒžã‚¹ã‚¿ãƒ•ã‚¡ã‚¤ãƒ«({MASTER_FILE})ã�®ã‚¹ã‚­ãƒ«/needèª­ã�¿è¾¼ã�¿ã‚¨ãƒ©ãƒ¼: {e}")
        return {}, [], [], {}, [], [], {}


def load_team_combination_presets_from_master() -> dict[
    str, list[tuple[int, int | None, tuple[str, ...], int | None]]
]:
    """
    master.xlsmã€Œçµ„ã�¿å�ˆã‚�ã�›è¡¨ã€�ã‚’èª­ã�¿ã€�å·¥ç¨‹+æ©Ÿæ¢°ã‚­ãƒ¼ã�”ã�¨ã�«
    [(çµ„å�ˆã�›å„ªå…ˆåº¦, å¿…è¦�äººæ•°ã�¾ã�Ÿã�¯None, ãƒ¡ãƒ³ãƒ�ãƒ¼ã‚¿ãƒ—ãƒ«, çµ„å�ˆã�›è¡ŒIDã�¾ã�Ÿã�¯None), ...] ã‚’è¿”ã�™ã€‚
    å�Œä¸€ã‚­ãƒ¼å†…ã�¯å„ªå…ˆåº¦æ˜‡é †ã€�å�Œé †ä½�ã�¯ã‚·ãƒ¼ãƒˆä¸Šã�®è¡Œé †ã€‚
    ã€Œå¿…è¦�äººæ•°ã€�åˆ—ã�¯é…�å�°æ™‚ã�« need åŸºæœ¬äººæ•°ã‚ˆã‚Šå„ªå…ˆã�™ã‚‹ï¼ˆãƒ¡ãƒ³ãƒ�ãƒ¼åˆ—äººæ•°ã�¨ä¸€è‡´ã�™ã‚‹ã�“ã�¨ï¼‰ã€‚
    é…�å�°ã�§ã�¯æˆ�ç«‹ã�—ã�Ÿãƒ—ãƒªã‚»ãƒƒãƒˆã‚’ã�™ã�¹ã�¦å€™è£œã�«è¼‰ã�›ã€�çµ„å�ˆã�›æŽ¢ç´¢ã�¨ã�¾ã�¨ã‚�ã�¦ team_start ç­‰ã�§æœ€è‰¯ã‚’é�¸ã�¶
    ï¼ˆã‚·ãƒ¼ãƒˆå„ªå…ˆåº¦ã�¯è©¦è¡Œé †ã�®ã�¿ã€‚å…ˆé ­ãƒ—ãƒªã‚»ãƒƒãƒˆã�®å�³æ±ºã�¯ã�—ã�ªã�„ï¼‰ã€‚
    A åˆ—ã€Œçµ„å�ˆã�›è¡ŒIDã€�ã�Œç„¡ã�„ï¼�ç©ºã�®æ—§ã‚·ãƒ¼ãƒˆã�§ã�¯ ID ã�¯ Noneã€‚
    """
    if not TEAM_ASSIGN_USE_MASTER_COMBO_SHEET:
        return {}
    path = MASTER_FILE
    if not os.path.isfile(path):
        return {}
    try:
        df = pd.read_excel(path, sheet_name=MASTER_SHEET_TEAM_COMBINATIONS, header=0)
    except Exception as e:
        logging.info("çµ„ã�¿å�ˆã‚�ã�›è¡¨ã‚·ãƒ¼ãƒˆã�®èª­è¾¼ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�™: %s", e)
        return {}
    if df is None or df.empty:
        return {}

    def norm_cell(x) -> str:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip()

    colmap = {norm_cell(c): c for c in df.columns if norm_cell(c)}
    id_c = colmap.get("çµ„å�ˆã�›è¡ŒID") or colmap.get("ã‚¤ãƒ³ãƒ‡ãƒƒã‚¯ã‚¹")
    proc_c = colmap.get("å·¥ç¨‹å��")
    mach_c = colmap.get("æ©Ÿæ¢°å��")
    combo_c = colmap.get("å·¥ç¨‹+æ©Ÿæ¢°")
    prio_c = colmap.get("çµ„å�ˆã�›å„ªå…ˆåº¦")
    req_c = colmap.get("å¿…è¦�äººæ•°")

    def mem_col_order(c) -> int:
        m = re.search(r"ãƒ¡ãƒ³ãƒ�ãƒ¼\s*(\d+)", norm_cell(c))
        return int(m.group(1)) if m else 9999

    mem_keys = sorted(
        [c for c in df.columns if norm_cell(str(c)).startswith("ãƒ¡ãƒ³ãƒ�ãƒ¼")],
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
    æŽ¡ç”¨ãƒ�ãƒ¼ãƒ ã�®ãƒ¡ãƒ³ãƒ�ãƒ¼é›†å�ˆï¼ˆNFKCãƒ»trimï¼‰ã�Œçµ„ã�¿å�ˆã‚�ã�›è¡¨ãƒ—ãƒªã‚»ãƒƒãƒˆã�®ã�„ã�šã‚Œã�‹ã�¨ä¸€è‡´ã�™ã‚‹ã�¨ã��ã€�
    ã��ã�®è¡Œã�®çµ„å�ˆã�›è¡ŒIDï¼ˆAåˆ—ï¼‰ã‚’è¿”ã�™ã€‚çµ„å�ˆã�›æŽ¢ç´¢ã�®ã�¿ã�§æ±ºã�¾ã‚Š combo_sheet_row_id ã�Œä»˜ã�„ã�¦ã�„ã�ªã�„
    å±¥æ­´è¡Œã�®è£œå®Œã�«ä½¿ã�†ã€‚è¤‡æ•°ä¸€è‡´æ™‚ã�¯çµ„å�ˆã�›å„ªå…ˆåº¦ï¼ˆæ•°å€¤ã�Œå°�ã�•ã�„æ–¹ï¼‰ã‚’æŽ¡ç”¨ã€‚
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
    """master ãƒ¡ã‚¤ãƒ³ã�®æ™‚åˆ»ã‚»ãƒ«ï¼ˆdatetime / time / æ–‡å­—åˆ—ï¼‰ã‚’ time ã�«ã€‚è§£é‡ˆä¸�èƒ½ã�¯ Noneã€‚"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, time):
        return v
    if isinstance(v, datetime):
        return v.time()
    return parse_time_str(v, None)


def _pick_master_main_sheet_name(sheetnames: list[str]) -> str | None:
    """
    master.xlsm ã�®ã€Œãƒ¡ã‚¤ãƒ³ã€�è¨­å®šã‚·ãƒ¼ãƒˆå��ã‚’è§£æ±ºã�™ã‚‹ï¼ˆVBA MasterGetMainWorksheet ã�¨å�Œè¶£æ—¨ï¼‰ã€‚
    ã€Œã€‡æœˆãƒ¡ã‚¤ãƒ³ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã€�ç­‰ã‚’èª¤æŽ¡ç”¨ã�—ã�ªã�„ã‚ˆã�†ã€Œã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã€�ã‚’å�«ã‚€å��å‰�ã�¯é™¤å¤–ã�—ã€�
    è¤‡æ•°å€™è£œã�¯ã‚·ãƒ¼ãƒˆå��ã�Œæœ€çŸ­ã�®ã‚‚ã�®ã‚’å„ªå…ˆã�™ã‚‹ã€‚
    """
    for prefer in ("ãƒ¡ã‚¤ãƒ³", "ãƒ¡ã‚¤ãƒ³_", "Main"):
        if prefer in sheetnames:
            return prefer
    cand = [sn for sn in sheetnames if "ãƒ¡ã‚¤ãƒ³" in sn and "ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼" not in sn]
    if not cand:
        return None
    return min(cand, key=len)


def _read_master_main_factory_operating_times(master_path: str) -> tuple[time | None, time | None]:
    """
    master.xlsm ã�®ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆ A12ï¼ˆç¨¼åƒ�é–‹å§‹ï¼‰ãƒ»B12ï¼ˆç¨¼åƒ�çµ‚äº†ï¼‰ã‚’èª­ã‚€ã€‚
    ã�„ã�šã‚Œã�‹æ¬ æ��ãƒ»ä¸�æ­£ãƒ»é–‹å§‹>=çµ‚äº†ã�®ã�¨ã��ã�¯ (None, None)ã€‚
    """
    p = (master_path or "").strip()
    if not p or not os.path.isfile(p):
        return None, None
    if _workbook_should_skip_openpyxl_io(p):
        return None, None
    try:
        wb = load_workbook(p, data_only=True, read_only=False)
    except Exception as e:
        logging.warning("å·¥å ´ç¨¼åƒ�æ™‚åˆ»: master ã‚’ openpyxl ã�§é–‹ã�‘ã�¾ã�›ã‚“ã�§ã�—ã�Ÿï¼ˆæ—¢å®šã�®æ—¥å†…æž ã‚’ä½¿ã�„ã�¾ã�™ï¼‰: %s", e)
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
                "å·¥å ´ç¨¼åƒ�æ™‚åˆ»: master ãƒ¡ã‚¤ãƒ³ A12/B12 ã�Œé–‹å§‹>=çµ‚äº† (%s >= %s) ã�®ã�Ÿã‚�æ—¢å®šå€¤ã‚’ä½¿ã�„ã�¾ã�™ã€‚",
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
    master.xlsm ã�®ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆ A15ï¼ˆå®šå¸¸é–‹å§‹ï¼‰ãƒ»B15ï¼ˆå®šå¸¸çµ‚äº†ï¼‰ã‚’èª­ã‚€ã€‚
    ã�„ã�šã‚Œã�‹æ¬ æ��ãƒ»ä¸�æ­£ãƒ»é–‹å§‹>=çµ‚äº†ã�®ã�¨ã��ã�¯ (None, None)ã€‚
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
            "å®šå¸¸æ™‚åˆ»: master ã‚’ openpyxl ã�§é–‹ã�‘ã�¾ã�›ã‚“ã�§ã�—ã�Ÿï¼ˆçµ�æžœã‚·ãƒ¼ãƒˆã�®å®šå¸¸å¤–ç�€è‰²ã‚’ã‚¹ã‚­ãƒƒãƒ—ï¼‰: %s",
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
                "å®šå¸¸æ™‚åˆ»: master ãƒ¡ã‚¤ãƒ³ A15/B15 ã�Œé–‹å§‹>=çµ‚äº† (%s >= %s) ã�®ã�Ÿã‚�ç�€è‰²ãƒ»æ¯”è¼ƒã�«ä½¿ã�„ã�¾ã�›ã‚“ã€‚",
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
    """æ®µéšŽ2ã�®é–“ã� ã�‘ DEFAULT_START_TIME / DEFAULT_END_TIME ã‚’ master ãƒ¡ã‚¤ãƒ³ A12/B12 ã�§ä¸Šæ›¸ã��ã€‚"""
    global DEFAULT_START_TIME, DEFAULT_END_TIME
    orig_s, orig_e = DEFAULT_START_TIME, DEFAULT_END_TIME
    ns, ne = _read_master_main_factory_operating_times(master_path)
    try:
        if ns is not None and ne is not None:
            DEFAULT_START_TIME = ns
            DEFAULT_END_TIME = ne
            logging.info(
                "å·¥å ´ç¨¼åƒ�æž : master.xlsm ãƒ¡ã‚¤ãƒ³ A12/B12 ã‚’æŽ¡ç”¨ â†’ %s ï½ž %sï¼ˆçµ�æžœ_* ã�®æ—¥å†…ã‚°ãƒªãƒƒãƒ‰ãƒ»é…�å�°æž ï¼‰",
                DEFAULT_START_TIME.strftime("%H:%M"),
                DEFAULT_END_TIME.strftime("%H:%M"),
            )
        yield
    finally:
        DEFAULT_START_TIME, DEFAULT_END_TIME = orig_s, orig_e


def infer_mid_break_from_reason(reason_text, start_t, end_t, break1_start=None, break1_end=None):
    """
    å‚™è€ƒã�‹ã‚‰ä¸­æŠœã�‘æ™‚é–“ã‚’æŽ¨å®šã�™ã‚‹ãƒ­ãƒ¼ã‚«ãƒ«è£œæ­£ã€‚
    AIã�Œä¸­æŠœã�‘ã‚’è¿”ã�•ã�ªã�„å ´å�ˆã�®ãƒ•ã‚§ã‚¤ãƒ«ã‚»ãƒ¼ãƒ•ã�¨ã�—ã�¦ä½¿ã�†ã€‚
    master.xlsm ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ç”±æ�¥ã�®ä¼‘æš‡åŒºåˆ†: å‰�ä¼‘=å�ˆå‰�å¹´ä¼‘ãƒ»å�ˆå¾Œã�®ã�¿å‹¤å‹™ã€�å¾Œä¼‘=å�ˆå¾Œå¹´ä¼‘ãƒ»å�ˆå‰�ã�®ã�¿å‹¤å‹™ï¼ˆå‡ºå‹¤ç°¿.txt ã�¨å�Œç¾©ï¼‰ã€‚
    å‰�ä¼‘ãƒ»å¾Œä¼‘ã�®å¢ƒç•Œã�¯ãƒ¡ãƒ³ãƒ�ãƒ¼å‹¤æ€ ã�®ä¼‘æ†©æ™‚é–“1_é–‹å§‹/çµ‚äº†ï¼ˆæœªæŒ‡å®šæ™‚ã�¯ DEFAULT_BREAKS[0]ï¼‰ã�«å�ˆã‚�ã�›ã‚‹ã€‚
    """
    if reason_text is None:
        return None, None
    txt = str(reason_text).strip()
    if not txt or txt.lower() in ("nan", "none", "null", "é€šå¸¸"):
        return None, None

    b1_s = break1_start if break1_start is not None else DEFAULT_BREAKS[0][0]
    b1_e = break1_end if break1_end is not None else DEFAULT_BREAKS[0][1]

    noon_end = time(12, 0)
    afternoon_start = time(13, 0)
    # ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼è¨˜å�·ã�¨ä¸€è‡´ã�•ã�›ã‚‹ï¼ˆã‚·ãƒ•ãƒˆæ™‚åˆ»ã�Œèª¤ã�£ã�¦ã�„ã‚‹å ´å�ˆã�®è£œå®Œç”¨ã€‚æ­£ã�—ã�„è¡Œã�§ã�¯åŒºé–“ã�Œç©ºã�«ã�ªã‚Šè¿½åŠ ã�•ã‚Œã�ªã�„ï¼‰
    if txt == "å‰�ä¼‘":
        # æ­£ã�—ã�„è¡Œã�¯å‡ºå‹¤ã�Œä¼‘æ†©1çµ‚äº†ä»¥é™�ã�§è£œå®Œä¸�è¦�ã€‚å…¨æ—¥ã‚·ãƒ•ãƒˆã�®èª¤å…¥åŠ›æ™‚ã�¯ã��ã�“ã�¾ã�§ã‚’ä¸­æŠœã�‘ï¼ˆå�ˆå‰�å¹´ä¼‘ç›¸å½“ï¼‰
        if start_t and start_t < b1_e:
            return start_t, b1_e
        return None, None
    if txt == "å¾Œä¼‘":
        if end_t and b1_s < end_t:
            return b1_s, end_t
        return None, None

    # 1) æ˜Žç¤ºçš„ã�ªæ™‚åˆ»ç¯„å›²ï¼ˆä¾‹: 11:00-14:00 / 11:00ï½ž14:00ï¼‰
    m = re.search(r"(\d{1,2}[:ï¼š]\d{2})\s*[~ã€œ\-ï¼�ãƒ¼]\s*(\d{1,2}[:ï¼š]\d{2})", txt)
    if m:
        s = parse_time_str(m.group(1).replace("ï¼š", ":"), None)
        e = parse_time_str(m.group(2).replace("ï¼š", ":"), None)
        if s and e and s < e:
            return s, e

    # 2) ã�‚ã�„ã�¾ã�„èªžï¼ˆå�ˆå‰�/å�ˆå¾Œ/çµ‚æ—¥ï¼‰ + ç�¾å ´é›¢è„±ãƒ»ä¼‘æš‡ç³»ã‚­ãƒ¼ãƒ¯ãƒ¼ãƒ‰
    # ã€Œå�ˆå¾Œä¼‘ã�¿ã�§ã�™ã€�ç­‰ã�¯ã€Œå�ˆå¾Œã€�ã‚’å�«ã‚€ã�Œã€�æ—§ãƒ­ã‚¸ãƒƒã‚¯ã�¯ã€ŒæŠœã�‘ã€�ç­‰ã�®ã�¿è¦‹ã�¦ã�Šã‚Šä¸­æŠœã�‘æŽ¨å®šã�«åˆ°é�”ã�—ã�ªã�‹ã�£ã�Ÿ
    leave_keywords = (
        "äº‹å‹™æ‰€", "ä¼šè­°", "æ•™è‚²", "ç ”ä¿®", "å¤–å‡º", "é›¢ã‚Œ", "æŠœã�‘", "ä¸­æŠœã�‘", "æ‰“å�ˆã�›",
        "ä¼‘ã�¿", "ä¼‘æš‡", "æ¬ å‹¤",
    )
    has_leave_hint = any(k in txt for k in leave_keywords)
    if not has_leave_hint:
        return None, None

    if ("çµ‚æ—¥" in txt) or ("1æ—¥" in txt and "é€šå¸¸" not in txt):
        return start_t, end_t
    if ("å�ˆå‰�ä¸­" in txt) or ("å�ˆå‰�" in txt):
        return start_t, noon_end
    if ("å�ˆå¾Œ" in txt):
        return afternoon_start, end_t

    return None, None


# çµ�æžœ_ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼(å‡ºå‹¤ç°¿) ã�®é€€å‹¤è¡¨ç¤ºã€‚VBA å‡ºå‹¤ç°¿ã€Œå¾Œä¼‘ã€�ï¼ˆå�ˆå¾Œå¹´ä¼‘ï¼‰ã�¨å�Œæ§˜ã�«å®Ÿè³ª ä¼‘æ†©æ™‚é–“1_é–‹å§‹ã�§çµ‚äº†ã�¨ã�¿ã�ªã�™ã€‚
_AFTERNOON_OFF_DISPLAY_END = DEFAULT_BREAKS[0][0]


def _reason_is_afternoon_off(reason: str) -> bool:
    """å¾Œä¼‘ï¼ˆå�ˆå¾Œå¹´ä¼‘ãƒ»å�ˆå‰�ã�®ã�¿å‹¤å‹™ï¼‰ã�¾ã�Ÿã�¯å‚™è€ƒã�®å�ˆå¾Œä¼‘ç³»ã€‚"""
    r = str(reason or "")
    return ("å�ˆå¾Œ" in r and ("ä¼‘" in r or "ä¼‘ã�¿" in r)) or ("å¾Œä¼‘" in r)


def _reason_is_morning_off(reason: str) -> bool:
    """å‰�ä¼‘ï¼ˆå�ˆå‰�å¹´ä¼‘ãƒ»å�ˆå¾Œã�®ã�¿å‹¤å‹™ï¼‰ã€‚ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ç”±æ�¥ã�®ç•¥å�·ã�®ã�¿æ˜Žç¤ºæ‰±ã�„ï¼ˆäº‹å‹™æ‰€å‹¤å‹™ã�ªã�©ã�¨æ··å�Œã�—ã�ªã�„ï¼‰ã€‚"""
    return "å‰�ä¼‘" in str(reason or "")


def _calendar_display_clock_out_for_calendar_sheet(entry: dict, day_date: date):
    """
    é…�å�°ã�¯ breaks_dt ã�®å�ˆå¾Œä¸­æŠœã�‘ã�§æ­£ã�—ã��ã�ªã‚‹ä¸€æ–¹ã€�end_dt ã�Œ 17:00 ã�®ã�¾ã�¾ã� ã�¨çµ�æžœã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã�®é€€å‹¤åˆ—ã� ã�‘èª¤ã‚‹ã€‚
    å¾Œä¼‘ï¼ˆå�ˆå¾Œå¹´ä¼‘ï¼‰ã�¾ã�Ÿã�¯å‚™è€ƒã�Œå�ˆå¾Œä¼‘ã�¿ç³»ã�§ã€�å®šæ™‚ã�¾ã�§ç¶šã��å�ˆå¾Œã�®ä¸­æŠœã�‘ã�Œã�‚ã‚‹ã�¨ã��ã� ã�‘é€€å‹¤è¡¨ç¤ºã‚’ä¼‘æ†©æ™‚é–“1_é–‹å§‹ã�«æ�ƒã�ˆã‚‹ï¼ˆend_dt æœ¬ä½“ã�¯å¤‰æ›´ã�—ã�ªã�„ï¼‰ã€‚
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
    å€‹äºº_* ã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒ«ã�®10åˆ†æž ã�Œä¼‘æ†©å¸¯ã�«å…¥ã‚‹ã�¨ã��ã�®æ–‡è¨€ã€‚
    æ˜¼é£Ÿã�ªã�©é€šå¸¸ä¼‘æ†©ã�¯ã€Œä¼‘æ†©ã€�ã€‚å¾Œä¼‘ï¼ˆå�ˆå¾Œå¹´ä¼‘ï¼‰ã�§å®šæ™‚ã�¾ã�§å·¥å ´ã�«ã�„ã�ªã�„å�ˆå¾Œå¸¯ã�¯ã€Œä¼‘æš‡ã€�ã€‚
    å‰�ä¼‘ï¼ˆå�ˆå‰�å¹´ä¼‘ï¼‰ã�§å�ˆå‰�ã�®æ¬ å‹¤åŒºé–“ã�Œä¼‘æ†©å¸¯ã�¨ã�—ã�¦å…¥ã�£ã�¦ã�„ã‚‹å ´å�ˆã�¯ã€Œä¼‘æš‡ã€�ã€‚
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
                return "ä¼‘æš‡"
            if morning_off and bs < DEFAULT_BREAKS[0][0]:
                be_t = b_e.time() if isinstance(b_e, datetime) else b_e
                if be_t <= time(13, 0):
                    return "ä¼‘æš‡"
        return "ä¼‘æ†©"
    return None


def _member_schedule_off_shift_label(
    day_date: date,
    grid_mid_dt: datetime,
    d_start_dt: datetime,
    d_end_dt: datetime,
    reason: str,
) -> str:
    """
    å€‹äºº_* ã‚·ãƒ¼ãƒˆã�§æ‰€å®šå‡ºé€€å‹¤ã�®å¤–å�´ã�®10åˆ†æž ã€‚
    å‰�ä¼‘ã�®å�ˆå‰�ï¼ˆå·¥å ´æ—¥ã�®æ‰€å®šé–‹å§‹ï½žå�ˆå¾Œå‡ºå‹¤ã�¾ã�§ï¼‰ã�¯å¹´ä¼‘ã€�å¾Œä¼‘ã�®å�ˆå¾Œã�¯å¹´ä¼‘ã€‚ã��ã‚Œä»¥å¤–ã�®ã‚·ãƒ•ãƒˆå¤–ã�¯å‹¤å‹™å¤–ã€‚
    """
    r = str(reason or "")
    day_start = datetime.combine(day_date, DEFAULT_START_TIME)
    day_end = datetime.combine(day_date, DEFAULT_END_TIME)
    if grid_mid_dt < d_start_dt:
        if _reason_is_morning_off(r) and grid_mid_dt >= day_start:
            return "å¹´ä¼‘"
        return "å‹¤å‹™å¤–"
    if grid_mid_dt >= d_end_dt:
        if _reason_is_afternoon_off(r) and grid_mid_dt < day_end:
            return "å¹´ä¼‘"
        return "å‹¤å‹™å¤–"
    return "å‹¤å‹™å¤–"


def _member_schedule_full_day_off_label(entry) -> str:
    """
    å…¨æ—¥é�žå‹¤å‹™ï¼ˆis_working=Falseï¼‰ã�®å€‹äººã‚·ãƒ¼ãƒˆåˆ—ã�®è¡¨ç¤ºã€‚
    ä¼‘æš‡åŒºåˆ†ã�Œå¹´ä¼‘ï¼ˆã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ *ï¼‰ã�®ã�¨ã��ã�¯ã€Žå¹´ä¼‘ã€�ã€‚å·¥å ´ä¼‘æ—¥ã�ªã�©ã�¯ã€Žä¼‘ã€�ã€‚
    """
    if not entry:
        return "ä¼‘"
    r = str(entry.get("reason") or "").strip()
    if r == "å¹´ä¼‘" or r.startswith("å¹´ä¼‘ "):
        return "å¹´ä¼‘"
    return "ä¼‘"


def _attendance_remark_text(row) -> str:
    """
    å‹¤æ€ 1è¡Œã�‹ã‚‰ã€Œå‚™è€ƒã€�åˆ—ã�®ãƒ†ã‚­ã‚¹ãƒˆã�®ã�¿å�–å¾—ã�™ã‚‹ã€‚
    å‹¤æ€ AIã�®è§£æž�ãƒªã‚¹ãƒˆã�¸ã�®æŠ•å…¥ã�¯ã�“ã�®åˆ—ã�®ã�¿ã€‚reason æ–‡å­—åˆ—ã�¯ load_attendance ã�§å‚™è€ƒã�¨ä¼‘æš‡åŒºåˆ†ã‚’å�ˆæˆ�ã�™ã‚‹ã€‚
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
    """å‹¤æ€ 1è¡Œã�‹ã‚‰ã€Œä¼‘æš‡åŒºåˆ†ã€�åˆ—ï¼ˆã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ç”±æ�¥ã�® å‰�ä¼‘/å¾Œä¼‘ ç­‰ï¼‰ã€‚"""
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
    """å‹¤æ€ å‚™è€ƒ AI ã�®çœŸå�½å€¤ï¼ˆbool / æ•°å€¤ / æ–‡å­—åˆ—ã�®æ�ºã‚Œã‚’å�¸å�Žï¼‰ã€‚"""
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
    if s in ("true", "1", "yes", "y", "ã�¯ã�„", "çœŸ", "on"):
        return True
    if s in ("false", "0", "no", "n", "ã�„ã�„ã�ˆ", "å�½", "off", ""):
        return False
    return default


def _parse_attendance_overtime_end_optional(v) -> time | None:
    """å‹¤æ€ ã€Œæ®‹æ¥­çµ‚æ¥­ã€�åˆ—ã€‚æœ‰åŠ¹ã�ªæ™‚åˆ»ã�®ã�¿ã€‚ç©ºãƒ»ä¸�æ­£ã�¯ Noneï¼ˆ_excel_scalar_to_time_optional ã�¨å�Œè¶£æ—¨ï¼‰ã€‚"""
    return _excel_scalar_to_time_optional(v)


def load_attendance_and_analyze(members):
    attendance_data = {}
    # â€»ã€Œå‹¤æ€ å‚™è€ƒã€�ã�¯ master å�„ãƒ¡ãƒ³ãƒ�ãƒ¼ã‚·ãƒ¼ãƒˆã�®ã€Œå‚™è€ƒã€�åˆ—ã�®ã�¿ã€‚ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆãƒ»ç‰¹åˆ¥æŒ‡å®š_å‚™è€ƒã�¯åˆ¥APIï¼ˆgenerate_plan å�´ã�§è¿½è¨˜ï¼‰ã€‚
    ai_log = {
        "ï¼ˆæ³¨ï¼‰ã�“ã�®ã‚·ãƒ¼ãƒˆã�®è¦‹æ–¹": "å…ˆé ­2è¡Œã�¯å‹¤æ€ ã€Œå‚™è€ƒã€�ã�®å‡ºé€€å‹¤AIã�®ã�¿ã€‚ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆãƒ»ç‰¹åˆ¥æŒ‡å®šã�¯ä¸‹æ®µã�®JSONã�¨ã€Œ_*_AI_APIã€�è¡Œã€‚",
        "å‹¤æ€ å‚™è€ƒ_AI_API": "ã�ªã�—",
        "å‹¤æ€ å‚™è€ƒ_AI_è©³ç´°": "è§£æž�å¯¾è±¡ã�®å‚™è€ƒè¡Œã�ªã�—",
    }
    
    # 1. ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ¥ã‚·ãƒ¼ãƒˆã�‹ã‚‰ã�®èª­ã�¿è¾¼ã�¿
    all_records = []
    try:
        xls = pd.ExcelFile(MASTER_FILE)
        for sheet_name in xls.sheet_names:
            if "ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼" in sheet_name or sheet_name.lower() in ['skills', 'need', 'tasks']:
                continue 
                
            m_name = sheet_name.strip()
            if m_name not in members:
                continue 
                
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
            df_sheet.columns = df_sheet.columns.str.strip()
            df_sheet['ãƒ¡ãƒ³ãƒ�ãƒ¼'] = m_name 
            all_records.append(df_sheet)
            
        if all_records:
            df = pd.concat(all_records, ignore_index=True)
            df['æ—¥ä»˜'] = pd.to_datetime(df['æ—¥ä»˜'], errors='coerce').dt.date
            df = df.dropna(subset=['æ—¥ä»˜'])
            logging.info(f"ã€Ž{MASTER_FILE}ã€�ã�®å�„ãƒ¡ãƒ³ãƒ�ãƒ¼ã�®å‹¤æ€ ã‚·ãƒ¼ãƒˆã‚’èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚")
            _cols = {str(c).strip() for c in df.columns}
            if ATT_COL_REMARK in _cols and ATT_COL_LEAVE_TYPE in _cols:
                logging.info(
                    "å‹¤æ€ åˆ—: AI å…¥åŠ›ã�¯ã€Œ%sã€�ã�®ã�¿ã€‚å‚™è€ƒã�Œç©ºã�®æ—¥ã�¯ã€Œ%sã€�ï¼ˆå‰�ä¼‘ãƒ»å¾Œä¼‘ãƒ»ä»–æ‹ ç‚¹å‹¤å‹™ã�ªã�©ï¼‰ã‚’ reason ã�«å��æ˜ ã�—ã�¾ã�™ã€‚",
                    ATT_COL_REMARK,
                    ATT_COL_LEAVE_TYPE,
                )
            elif ATT_COL_REMARK not in _cols:
                logging.warning(
                    "å‹¤æ€ ãƒ‡ãƒ¼ã‚¿ã�«ã€Œ%sã€�åˆ—ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚å‚™è€ƒãƒ™ãƒ¼ã‚¹ã�® AI è§£æž�ã�¯ç©ºæ‰±ã�„ã�«ã�ªã‚Šã�¾ã�™ã€‚",
                    ATT_COL_REMARK,
                )
            if ATT_COL_OT_END in _cols:
                logging.info(
                    "å‹¤æ€ åˆ—: ä»»æ„�ã€Œ%sã€�ã�§é€€å‹¤ä¸Šé™�æ™‚åˆ»ã‚’æŒ‡å®šã�§ã��ã�¾ã�™ï¼ˆå…¨æ—¥ä¼‘ã�¿è¡Œã�§ã�¯ç„¡è¦–ï¼‰ã€‚",
                    ATT_COL_OT_END,
                )
        else:
            raise FileNotFoundError("æœ‰åŠ¹ã�ªãƒ¡ãƒ³ãƒ�ãƒ¼åˆ¥å‹¤æ€ ã‚·ãƒ¼ãƒˆã�Œè¦‹ã�¤ã�‹ã‚Šã�¾ã�›ã‚“ã€‚")
            
    except Exception as e:
        logging.warning(f"å‹¤æ€ ã‚·ãƒ¼ãƒˆèª­ã�¿è¾¼ã�¿ã‚¨ãƒ©ãƒ¼: {e} ãƒ‡ãƒ•ã‚©ãƒ«ãƒˆã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã‚’ç”Ÿæˆ�ã�—ã�¾ã�™ã€‚")
        default_dates = generate_default_calendar_dates(TARGET_YEAR, TARGET_MONTH)
        records = []
        for d in default_dates:
            for m in members: records.append({'æ—¥ä»˜': d, 'ãƒ¡ãƒ³ãƒ�ãƒ¼': m, 'å‚™è€ƒ': 'é€šå¸¸'})
        df = pd.DataFrame(records)

    # 2. AI ã�«ã‚ˆã‚‹å‹¤æ€ æ–‡è„ˆã�®è§£æž�ï¼ˆå‚™è€ƒã�Œç©ºã�§ã‚‚ä¼‘æš‡åŒºåˆ†ã�®ã�¿ã�®è¡Œã�¯ AI ã�«æ¸¡ã�—ã€�è¡¨è¨˜æ�ºã‚Œã�¯ãƒ¢ãƒ‡ãƒ«ã�«è§£é‡ˆã�•ã�›ã‚‹ï¼‰
    remarks_to_analyze = []
    for _, row in df.iterrows():
        m = str(row.get('ãƒ¡ãƒ³ãƒ�ãƒ¼', '')).strip()
        if m not in members:
            continue
        rem = _attendance_remark_text(row)
        lt = _attendance_leave_type_text(row)
        d_str = row['æ—¥ä»˜'].strftime("%Y-%m-%d") if pd.notna(row['æ—¥ä»˜']) else ""
        if rem:
            remarks_to_analyze.append(f"{d_str}_{m} ã�®å‚™è€ƒ: {rem}")
        elif lt and lt not in ("é€šå¸¸", ""):
            remarks_to_analyze.append(f"{d_str}_{m} ã�®ä¼‘æš‡åŒºåˆ†ï¼ˆå‚™è€ƒã�¯ç©ºï¼‰: {lt}")

    if remarks_to_analyze:
        remarks_blob = "\n".join(remarks_to_analyze)
        cache_key = hashlib.sha256(
            (remarks_blob + "\n" + ATTENDANCE_REMARK_AI_SCHEMA_ID).encode("utf-8")
        ).hexdigest()
        ai_cache = load_ai_cache()

        # å�Œä¸€å‚™è€ƒã‚»ãƒƒãƒˆã�¯ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã‚’å„ªå…ˆåˆ©ç”¨ã�—ã€�APIã‚³ãƒ¼ãƒ«ã‚’ç¯€ç´„
        cached_data = get_cached_ai_result(ai_cache, cache_key)
        if cached_data is not None:
            ai_parsed = cached_data
            ai_log["å‹¤æ€ å‚™è€ƒ_AI_API"] = "ã�ªã�—(ã‚­ãƒ£ãƒƒã‚·ãƒ¥ä½¿ç”¨)"
            ai_log["å‹¤æ€ å‚™è€ƒ_AI_è©³ç´°"] = "ã‚­ãƒ£ãƒƒã‚·ãƒ¥ãƒ’ãƒƒãƒˆ"
        elif not API_KEY:
            ai_parsed = {}
            ai_log["å‹¤æ€ å‚™è€ƒ_AI_API"] = "ã�ªã�—"
            ai_log["å‹¤æ€ å‚™è€ƒ_AI_è©³ç´°"] = "GEMINI_API_KEYæœªè¨­å®šã�®ã�Ÿã‚�å‹¤æ€ å‚™è€ƒAIã‚’ã‚¹ã‚­ãƒƒãƒ—"
            logging.info("GEMINI_API_KEY æœªè¨­å®šã�®ã�Ÿã‚�å‚™è€ƒAIè§£æž�ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿã€‚")
        else:
            logging.info("â–  AIã�Œè¤‡æ•°æ—¥ã�®ç‰¹è¨˜äº‹é …ã‚’è§£æž�ä¸­...")
            ai_log["å‹¤æ€ å‚™è€ƒ_AI_API"] = "ã�‚ã‚Š"
            
            prompt = f"""
            ä»¥ä¸‹ã�®å�„æ—¥ãƒ»ãƒ¡ãƒ³ãƒ�ãƒ¼ã�®å‚™è€ƒã‚’èª­ã�¿å�–ã‚Šã€�å‡ºé€€å‹¤æ™‚åˆ»ã�®å¤‰æ›´ã‚„ä¸­æŠœã�‘ã€�ä¼‘æ—¥ã�®åˆ¤å®šã‚’è¡Œã�„ã€�JSONå½¢å¼�ã�§å‡ºåŠ›ã�—ã�¦ã��ã� ã�•ã�„ã€‚
            ãƒžãƒ¼ã‚¯ãƒ€ã‚¦ãƒ³è¨˜å�·(``` ç­‰)ã�¯ä¸€åˆ‡å�«ã‚�ã�šã€�ç´”ç²‹ã�ªJSONæ–‡å­—åˆ—ã�®ã�¿ã‚’è¿”ã�—ã�¦ã��ã� ã�•ã�„ã€‚

            ã€�JSONã�®å‡ºåŠ›å½¢å¼�ï¼ˆã‚­ãƒ¼å��ã‚’åŽ³å¯†ã�«å®ˆã‚‹ã�“ã�¨ï¼‰ã€‘
            {{
              "YYYY-MM-DD_ãƒ¡ãƒ³ãƒ�ãƒ¼å��": {{
                "å‡ºå‹¤æ™‚åˆ»": "HH:MM", 
                "é€€å‹¤æ™‚åˆ»": "HH:MM", 
                "ä¸­æŠœã�‘é–‹å§‹": "HH:MM",
                "ä¸­æŠœã�‘çµ‚äº†": "HH:MM",
                "ä½œæ¥­åŠ¹çŽ‡": 1.0,     
                "is_holiday": false,
                "é…�å�°ä¸�å�‚åŠ ": false
              }}
            }}
            ãƒ»ã‚­ãƒ¼å��ã�¯ä¸Šè¨˜ã�®æ—¥æœ¬èªžã‚­ãƒ¼ã‚’ã��ã�®ã�¾ã�¾ä½¿ã�†ï¼ˆè‹±èªžã‚­ãƒ¼ã�«ç½®ã��æ�›ã�ˆã�ªã�„ï¼‰
            ãƒ»å‡ºå‹¤æ™‚åˆ»/é€€å‹¤æ™‚åˆ»: å½“è©²è¡Œã�®ã€Œå‚™è€ƒã€�ã�¾ã�Ÿã�¯ã€Œä¼‘æš‡åŒºåˆ†ï¼ˆå‚™è€ƒã�¯ç©ºï¼‰ã€�ã�®æ–‡è„ˆã�‹ã‚‰æŽ¨æ¸¬ã€‚ä¸�æ˜Žã‚„å¤‰æ›´ã�ªã�—ã�ªã‚‰ null
            ãƒ»ä¸­æŠœã�‘é–‹å§‹/çµ‚äº†: ä¸€æ™‚çš„ã�ªé›¢è„±ï¼ˆä¸­æŠœã�‘ãƒ»äº‹å‹™æ‰€ãƒ»ä¼šè­°ã�ªã�©ï¼‰ã�Œã�‚ã‚‹å ´å�ˆã€�ã��ã�®é–‹å§‹ãƒ»çµ‚äº†ã€‚ã�ªã�„å ´å�ˆã�¯ null
            ãƒ»æ›–æ˜§èªžã�®è§£é‡ˆä¾‹:
              - ã€Œå�ˆå‰�ä¸­ã�¯äº‹å‹™æ‰€ã�§ä½œæ¥­ã€�=> ä¸­æŠœã�‘é–‹å§‹ "08:45", ä¸­æŠœã�‘çµ‚äº† "12:00"
              - ã€Œå�ˆå¾Œã�¯ä¼šè­°ã€�=> ä¸­æŠœã�‘é–‹å§‹ "13:00", ä¸­æŠœã�‘çµ‚äº† "17:00"
            ãƒ»is_holiday: ã��ã�®æ—¥ã�Œä¼šç¤¾ã�«æ�¥ã�ªã�„ãƒ»çµ‚æ—¥ä¼‘æš‡ãƒ»æ¬ å‹¤ã�ªã�© **å‹¤å‹™è‡ªä½“ã�Œã�ªã�„** ã�¨åˆ¤æ–­ã�§ã��ã‚‹å ´å�ˆã�®ã�¿ trueã€‚å�ˆå‰�ä¼‘ãƒ»å�ˆå¾Œä¼‘ã�ªã�©éƒ¨åˆ†çš„ã�ªä¼‘ã�¿ã�¯ falseï¼ˆä¸­æŠœã�‘ã‚„æ™‚åˆ»ã�§è¡¨ç�¾ï¼‰
            ãƒ»é…�å�°ä¸�å�‚åŠ : å‹¤å‹™ã�¯ã�‚ã‚‹ã�Œ **åŠ å·¥ãƒ©ã‚¤ãƒ³ã�¸ã�®é…�å�°ï¼ˆOP/AS ã�®å‰²å½“ï¼‰ã�«è¼‰ã�›ã�¦ã�¯ã�„ã�‘ã�ªã�„** ã�¨èª­ã�¿å�–ã‚Œã‚‹å ´å�ˆã�¯ trueã€‚è¡¨è¨˜ã�¯å•�ã‚�ã�šæ„�å‘³ã�§åˆ¤æ–­ã�™ã‚‹ã�“ã�¨ã€‚
              ä¾‹: ã€Œé…�å�°ä¸�å�¯ã€�ã€Œé…�å�°ï¼®ï¼§ã€�ã€Œãƒ©ã‚¤ãƒ³ã�«ä¹—ã‚‰ã�ªã�„ã€�ã€Œæœˆæ¬¡ç‚¹æ¤œã�®ã�¿ã€�ã€Œç‚¹æ¤œã�§ä¸€æ—¥ã€�ã€Œäº‹å‹™ã�®ã�¿ã€�ã€Œæ•™è‚²ã�§ç�¾å ´ä¸�å�¯ã€�ã€Œæ‰‹é…�ã�ªã�—ã€�ã€Œã‚¢ã‚µã‚¤ãƒ³ä¸�è¦�ã€�ã�ªã�©ã�®æ�ºã‚Œã‚„å©‰æ›²è¡¨ç�¾ã‚‚å�«ã‚€ã€‚
              é€šå¸¸å‹¤å‹™ã�§ç‰¹ã�«åˆ¶é™�ã�Œèª­ã�¿å�–ã‚Œã�ªã�„å ´å�ˆã�¯ false
            ãƒ»ä½œæ¥­åŠ¹çŽ‡: 0.0ã€œ1.0ã�®æ•°å€¤
            
            ã€�ç‰¹è¨˜äº‹é …ãƒªã‚¹ãƒˆã€‘
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
                    ai_log["å‹¤æ€ å‚™è€ƒ_AI_è©³ç´°"] = "è§£æž�æˆ�åŠŸ"
                else:
                    ai_parsed = {}
                    ai_log["å‹¤æ€ å‚™è€ƒ_AI_è©³ç´°"] = "JSONãƒ‘ãƒ¼ã‚¹å¤±æ•—"
            except Exception as e:
                err_text = str(e)
                is_quota_or_rate = ("429" in err_text) or ("RESOURCE_EXHAUSTED" in err_text)
                retry_sec = extract_retry_seconds(err_text) if is_quota_or_rate else None

                if is_quota_or_rate and retry_sec is not None:
                    wait_sec = min(max(retry_sec, 1.0), 90.0)
                    logging.warning(f"AIé€šä¿¡ 429/RESOURCE_EXHAUSTEDã€‚{wait_sec:.1f}ç§’å¾…æ©Ÿã�—ã�¦1å›žã� ã�‘å†�è©¦è¡Œã�—ã�¾ã�™ã€‚")
                    time_module.sleep(wait_sec)
                    try:
                        res = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                        match = re.search(r'\{.*\}', res.text, re.DOTALL)
                        if match:
                            ai_parsed = json.loads(match.group(0))
                            put_cached_ai_result(ai_cache, cache_key, ai_parsed)
                            save_ai_cache(ai_cache)
                            ai_log["å‹¤æ€ å‚™è€ƒ_AI_è©³ç´°"] = "å†�è©¦è¡Œã�§è§£æž�æˆ�åŠŸ"
                        else:
                            ai_parsed = {}
                            ai_log["å‹¤æ€ å‚™è€ƒ_AI_è©³ç´°"] = "å†�è©¦è¡Œå¾ŒJSONãƒ‘ãƒ¼ã‚¹å¤±æ•—"
                    except Exception as e2:
                        ai_parsed = {}
                        logging.warning(f"AIå†�è©¦è¡Œã‚¨ãƒ©ãƒ¼: {e2}")
                        ai_log["å‹¤æ€ å‚™è€ƒ_AI_è©³ç´°"] = f"429å¾Œå†�è©¦è¡Œå¤±æ•—: {e2}"
                else:
                    ai_parsed = {}
                    logging.warning(f"AIé€šä¿¡ã‚¨ãƒ©ãƒ¼: {e}")
                    ai_log["å‹¤æ€ å‚™è€ƒ_AI_è©³ç´°"] = str(e)
    else:
        ai_parsed = {}

    # 3. æ—¥ä»˜ã�”ã�¨ã�®åˆ¶ç´„è¾žæ›¸ã‚’æ§‹ç¯‰
    for _, row in df.iterrows():
        if pd.isna(row['æ—¥ä»˜']): continue
        curr_date = row['æ—¥ä»˜']
        m = str(row.get('ãƒ¡ãƒ³ãƒ�ãƒ¼', '')).strip()
        if m not in members: continue

        if curr_date not in attendance_data:
            attendance_data[curr_date] = {}

        original_reason = _attendance_remark_text(row)
        leave_type = _attendance_leave_type_text(row)

        key = f"{curr_date.strftime('%Y-%m-%d')}_{m}"
        ai_info = ai_parsed.get(key, {})

        is_empty_shift = pd.isna(row.get('å‡ºå‹¤æ™‚é–“')) and pd.isna(row.get('é€€å‹¤æ™‚é–“')) and not ai_info
        is_holiday = _ai_json_bool(ai_info.get("is_holiday"), False) or is_empty_shift
        exclude_from_line = _ai_json_bool(ai_info.get("é…�å�°ä¸�å�‚åŠ "), False)

        ai_eff = ai_info.get("ä½œæ¥­åŠ¹çŽ‡")
        excel_eff = row.get('ä½œæ¥­åŠ¹çŽ‡')
        
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
                and leave_type not in ("é€šå¸¸", "")
                and leave_type not in original_reason
            ):
                reason = f"{leave_type} {original_reason}"
            else:
                reason = original_reason
        elif leave_type and leave_type not in ("é€šå¸¸", ""):
            reason = leave_type
        else:
            reason = 'é€šå¸¸' if not is_empty_shift else 'ä¼‘æ—¥ã‚·ãƒ•ãƒˆ'

        # ãƒžã‚¹ã‚¿ã�«å‡ºå‹¤ãƒ»é€€å‹¤ã�®ä¸¡æ–¹ã�Œå…¥ã�£ã�¦ã�„ã‚‹æ—¥ã�¯ã€�å‹¤æ€ AIã�®å‡ºå‹¤/é€€å‹¤æ™‚åˆ»ã�§ä¸Šæ›¸ã��ã�—ã�ªã�„ï¼ˆä¼‘æš‡åŒºåˆ†ã�®ã�¿ã�®è¡Œã�§èª¤æŽ¨å®šã�•ã‚Œã�†ã‚‹ï¼‰
        excel_s = row.get("å‡ºå‹¤æ™‚é–“")
        excel_e = row.get("é€€å‹¤æ™‚é–“")
        if not pd.isna(excel_s) and not pd.isna(excel_e):
            start_t = parse_time_str(excel_s, DEFAULT_START_TIME)
            end_t = parse_time_str(excel_e, DEFAULT_END_TIME)
        else:
            start_t = parse_time_str(ai_info.get("å‡ºå‹¤æ™‚åˆ»") or excel_s, DEFAULT_START_TIME)
            end_t = parse_time_str(ai_info.get("é€€å‹¤æ™‚åˆ»") or excel_e, DEFAULT_END_TIME)
        base_end_t = end_t

        b1_s = parse_time_str(row.get('ä¼‘æ†©æ™‚é–“1_é–‹å§‹'), DEFAULT_BREAKS[0][0])
        b1_e = parse_time_str(row.get('ä¼‘æ†©æ™‚é–“1_çµ‚äº†'), DEFAULT_BREAKS[0][1])
        b2_s = parse_time_str(row.get('ä¼‘æ†©æ™‚é–“2_é–‹å§‹'), DEFAULT_BREAKS[1][0])
        b2_e = parse_time_str(row.get('ä¼‘æ†©æ™‚é–“2_çµ‚äº†'), DEFAULT_BREAKS[1][1])

        # â˜…è¿½åŠ : AIã�‹ã‚‰ä¸­æŠœã�‘æ™‚é–“ã‚’å�–å¾—
        mid_break_s = parse_time_str(ai_info.get("ä¸­æŠœã�‘é–‹å§‹"), None)
        mid_break_e = parse_time_str(ai_info.get("ä¸­æŠœã�‘çµ‚äº†"), None)
        # AIã�Œä¸­æŠœã�‘ã‚’è¿”ã�•ã�ªã�‹ã�£ã�Ÿå ´å�ˆã�¯ã€�å‚™è€ƒæ–‡è¨€ã�‹ã‚‰ãƒ­ãƒ¼ã‚«ãƒ«æŽ¨å®šã�§è£œå®Œ
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
                "å‹¤æ€  %s %s: æ®‹æ¥­çµ‚æ¥­é�©ç”¨å¾Œã�«é€€å‹¤ã�Œå‡ºå‹¤ä»¥å‰�ã�¨ã�ªã�£ã�Ÿã�Ÿã‚�ã€�æ®‹æ¥­çµ‚æ¥­ã‚’ç„¡è¦–ã�—ã�¦å®šæ™‚é€€å‹¤ã�«æˆ»ã�—ã�¾ã�™ã€‚",
                curr_date,
                m,
            )
            end_t = base_end_t
            end_dt = combine_dt(end_t)
        breaks_dt = []
        
        # é€šå¸¸ã�®ä¼‘æ†©ã‚’è¿½åŠ 
        if b1_s and b1_e: breaks_dt.append((combine_dt(b1_s), combine_dt(b1_e)))
        if b2_s and b2_e: breaks_dt.append((combine_dt(b2_s), combine_dt(b2_e)))
        
        # â˜…è¿½åŠ : ä¸­æŠœã�‘æ™‚é–“ã�Œã�‚ã‚‹å ´å�ˆã�¯ã€�ç‰¹åˆ¥ã�ªã€Œä¼‘æ†©ã€�ã�¨ã�—ã�¦ã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒ«è¨ˆç®—ã�«è¿½åŠ 
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
# å…¨ä¾�é ¼å…±é€š: åŠ å·¥å†…å®¹åˆ—ã�®å·¥ç¨‹é †åº� / å€‹åˆ¥: ECâ†’æ¤œæŸ»ãƒ­ãƒ¼ãƒ«ãƒ‘ã‚¤ãƒ—ãƒ©ã‚¤ãƒ³
# ---------------------------------------------------------------------------
ROLL_PIPELINE_EC_PROCESS = "EC"
ROLL_PIPELINE_EC_MACHINE = "ECæ©Ÿã€€æ¹–å�—"
ROLL_PIPELINE_INSP_PROCESS = "æ¤œæŸ»"
ROLL_PIPELINE_INSP_MACHINE = "ç†±èž�ç�€æ©Ÿã€€æ¹–å�—"
# Â§B-3: å¾Œç¶šã�¯ B-2 ã�®ã€Œæ¤œæŸ»ã€�ã�«ç›¸å½“ã�™ã‚‹å·¥ç¨‹ã�¨ã�—ã�¦å·»è¿”ã�—ï¼ˆå�Œä¸€ä¾�é ¼ã�§ EC å…ˆè¡Œãƒ»ãƒ­ãƒ¼ãƒ«æž ãƒ»ãƒªãƒ¯ã‚¤ãƒ³ãƒ‰ç­‰ã�¯ B-2 ã�¨å�Œè¶£æ—¨ï¼‰
ROLL_PIPELINE_REWIND_PROCESS = "å·»è¿”ã�—"
ROLL_PIPELINE_REWIND_MACHINE = "ECæ©Ÿã€€æ¹–å�—"
ROLL_PIPELINE_INITIAL_BUFFER_ROLLS = 2
# æ¤œæŸ»ã�®å‰²å½“ä¸Šé™� min ã�«ä½¿ã�†ã€‚å�Œä¸€ä¾�é ¼ã�« EC è¡Œã�Œç„¡ã�„ã�¨ã��ã�¯ needãƒ»ã‚¹ã‚­ãƒ«ã�«å¾“ã�„é€šå¸¸é…�å�°ã�™ã‚‹ï¼ˆec_done=0 å›ºå®šã�§æ°¸ä¹…ã‚¹ã‚­ãƒƒãƒ—ã�—ã�ªã�„ï¼‰ã€‚
ROLL_PIPELINE_INSP_UNCAPPED_ROOM = 1.0e18


# å‹¤æ€ ã�«è¼‰ã�£ã�¦ã�„ã‚‹æœ€çµ‚æ—¥ã�¾ã�§ã�§å‰²ä»˜ã�Œçµ‚ã‚�ã‚‰ã�ªã�„ã�¨ã��ã€�æœ€çµ‚æ—¥ã�¨å�Œã�˜ã‚·ãƒ•ãƒˆåž‹ã�§æ—¥ä»˜ã‚’å»¶é•·ã�™ã‚‹ï¼ˆã‚ªãƒ—ã‚·ãƒ§ãƒ³ï¼‰ã€‚
# False ã�®ã�¨ã��æ®µéšŽ2ã�¯ãƒžã‚¹ã‚¿å‹¤æ€ ã�®æ—¥ä»˜ç¯„å›²ã�®ã�¿ã�§å‰²ä»˜ã�—ã€�æ®‹ã‚Šã�¯é…�å�°æ®‹ãƒ»é…�å�°ä¸�å�¯ã�®ã�¾ã�¾ã�¨ã�™ã‚‹ã€‚
STAGE2_EXTEND_ATTENDANCE_CALENDAR = False
SCHEDULE_EXTEND_MAX_EXTRA_DAYS = 366

# ç´�æœŸåŸºæº–æ—¥ã‚’é�Žã�Žã�¦ã‚‚å½“è©²ä¾�é ¼ã�«æ®‹é‡�ã�Œã�‚ã‚‹ã�¨ã��ã€�**ã��ã�®ä¾�é ¼NOã� ã�‘** due_basis ã‚’ +1 ã�—ã€�
# å½“è©²ä¾�é ¼ã�®å‰²å½“ãƒ»ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ã‚’å·»ã��æˆ»ã�—ã�¦**ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼å…ˆé ­ã�‹ã‚‰**å†�ã‚·ãƒŸãƒ¥ãƒ¬ãƒ¼ã‚·ãƒ§ãƒ³ã�™ã‚‹ï¼ˆä»–ä¾�é ¼ã�®å‰²å½“ã�¯ç¶­æŒ�ï¼‰ã€‚
# ãƒžã‚¹ã‚¿å‹¤æ€ ã�®æœ€çµ‚æ—¥ã‚’è¶…ã�ˆã�¦å¾Œã‚�å€’ã�—ã�§ã��ã�ªã�„ä¾�é ¼ã�¯ã€Œé…�å�°æ®‹(å‹¤å‹™ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ä¸�è¶³)ã€�ã�¨ã�™ã‚‹ã€‚å�„å†�è©¦è¡Œå‰�ã�«å‹¤æ€ æ‹¡å¼µåˆ†ã�¯ãƒžã‚¹ã‚¿æ—¥ä»˜ã�¸æˆ»ã�™ã€‚
# æ—¢å®š **False**ï¼ˆé…�å�°è©¦è¡Œé †ã‚’æ­£ã�¨ã�—ã€�è¨ˆç”»åŸºæº–è¶…é�Žã�§ã‚‚ã�“ã�®å·»ã��æˆ»ã�—å†�è©¦è¡Œã�¯è¡Œã‚�ã�ªã�„ï¼‰ã€‚å¾“æ�¥æŒ™å‹•ã�Œå¿…è¦�ã�ªã�¨ã��ã� ã�‘ Trueã€‚
STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING = False
# ç´�æœŸåŸºæº–ã�® +1 æ—¥ã�«ã‚ˆã‚‹å·»ã��æˆ»ã�—å†�ã‚·ãƒŸãƒ¥ã�¯ä¾�é ¼NOã�”ã�¨ã�«æœ€å¤§ã�“ã�®å›žæ•°ï¼ˆ6 å›žç›®ä»¥é™�ã�¯å½“è©²ä¾�é ¼ã�®ã�¿ã‚·ãƒ•ãƒˆã�›ã�šã€�æœªå®Œäº†è¡Œã�«ç´�æœŸè¦‹ç›´ã�—å¿…è¦�ã‚’ä»˜ä¸Žã�—å¾—ã‚‹ï¼‰ã€‚
STAGE2_RETRY_SHIFT_DUE_MAX_ROUNDS = 5

# True ã�®ã�¨ã��ã€�é…�å�°è¨ˆç”»ã‚·ãƒ¼ãƒˆã�®èª­ã�¿è¾¼ã�¿è¡Œé †ï¼ˆå�„ä¾�é ¼NOã�®åˆ�å‡ºè¡Œã�Œæ—©ã�„ã�»ã�©å…ˆï¼‰ã�§ 1 ä¾�é ¼ã� ã�‘ã‚’
# å½“æ—¥å€™è£œã�«æ®‹ã�—ã€�å®Œèµ°ã�—ã�¦ã�‹ã‚‰æ¬¡ä¾�é ¼ã�¸é€²ã‚€ã€‚**ä»–ä¾�é ¼ã�¯ä¸€åˆ‡ã��ã�®æ—¥é…�å�°ã�•ã‚Œã�ªã�„**ã�Ÿã‚�ã€�
# ã‚¢ã‚¯ãƒ†ã‚£ãƒ–ä¾�é ¼ã�®1è¡Œã�§ã‚‚è©°ã�¾ã‚‹ã�¨å…¨ä½“ã�Œé…�å�°ä¸�å�¯ã�«è¦‹ã�ˆã‚‹ï¼ˆãƒ­ã‚°ã€Œä¾�é ¼NOç›´åˆ—é…�å�° ç›´åˆ—å¾Œ=1ã€�ï¼‰ã€‚
# æ—¢å®š Falseã€‚åŽ³å¯†ã�ªä¾�é ¼NOç›´åˆ—ã�Œå¿…è¦�ã�ªã�¨ã��ã� ã�‘ STAGE2_SERIAL_DISPATCH_BY_TASK_ID=1 ã‚’è¨­å®šã�™ã‚‹ã€‚
STAGE2_SERIAL_DISPATCH_BY_TASK_ID = (
    os.environ.get("STAGE2_SERIAL_DISPATCH_BY_TASK_ID", "0")
    .strip()
    .lower()
    in ("1", "true", "yes", "on", "ã�¯ã�„")
)

# True: â‘ æ®‹ã‚¿ã‚¹ã‚¯ã�®ã�†ã�¡é…�å�°è©¦è¡Œé †ã�Œæœ€å°�ã�®1ã‚¿ã‚¹ã‚¯ã� ã�‘ã‚’é�¸ã�³ã€�1ãƒ­ãƒ¼ãƒ«ã�šã�¤å‰²ä»˜ã€‚
# â‘¡åŽŸå��æŠ•å…¥æ—¥ã�¨å�Œä¸€æ—¥ã�«é–‹å§‹ã�™ã‚‹å ´å�ˆã�¯ 13:00 ä»¥é™�ï¼ˆsame_day_raw_start_limit ã‚‚ 13:00ï¼‰ã€‚
# â‘¢â‘£è¨­å‚™ç©ºã��ã‚’ max ã�§ç¹°ã‚Šä¸Šã�’ï¼ˆæ—¥å†…ã€‚ç¿Œæ—¥ã�¯æ—¥ä»˜ãƒ«ãƒ¼ãƒ—ã�§ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ã‚·ãƒ¼ãƒ‰ï¼‰ã€‚
# â‘¤â‘¥â‘¦â‘§äººã�®ç©ºã��ã�§ãƒ�ãƒ¼ãƒ ã‚’æ±ºã‚�ã€�ãƒ­ãƒ¼ãƒ«ã�”ã�¨ã�« avail ã‚’æ›´æ–°ï¼ˆå�Œæ—¥ã�¯å‰�ãƒ­ãƒ¼ãƒ«ã�¨å�Œä¸€ãƒ�ãƒ¼ãƒ ã‚’å„ªå…ˆï¼‰ã€‚
# ç„¡åŠ¹åŒ–: ç’°å¢ƒå¤‰æ•° STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST=0
STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST = os.environ.get(
    "STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST", "1"
).strip().lower() not in ("0", "false", "no", "off", "ã�„ã�„ã�ˆ", "ç„¡åŠ¹")

# Trueï¼ˆæ—¢å®šï¼‰: start_date_req<=å½“æ—¥ ã�‹ã�¤æ®‹ã�‚ã‚Šã�®ã‚¿ã‚¹ã‚¯ã�®ã�†ã�¡ã€�é…�å�°è©¦è¡Œé †ã�®æœ€å°�ã€Œæž ã€�ã� ã�‘ã�Œå‰²ä»˜å¯¾è±¡ã€‚
# ã‚ˆã‚Šå¤§ã��ã�„è©¦è¡Œé †ã�¯ã€�ã‚ˆã‚Šå°�ã�•ã�„è©¦è¡Œé †ã�«æœªå®Œäº†ã�Œæ®‹ã‚‹é™�ã‚Šãƒ–ãƒ­ãƒƒã‚¯ï¼ˆç´�æœŸã�Œè¿‘ã��ã�¦ã‚‚å‰²ã‚Šè¾¼ã�¾ã�ªã�„ï¼‰ã€‚
STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT = os.environ.get(
    "STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT", "1"
).strip().lower() not in ("0", "false", "no", "off", "ã�„ã�„ã�ˆ", "ç„¡åŠ¹")

# Trueï¼ˆæ—¢å®šï¼‰: å‰²ä»˜å€™è£œã‚’ã€Œè¨­å‚™ãƒ»äººã�®å£�æ™‚è¨ˆå� æœ‰åŒºé–“ã€�ã�§äºŒé‡�æ¤œæŸ»ã�—ã€�ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³è¿½è¨˜ã�¨å�ŒæœŸç™»éŒ²ã�™ã‚‹
# ï¼ˆãƒ–ãƒ­ãƒƒã‚¯ãƒ†ãƒ¼ãƒ–ãƒ«ã�¨å�Œè¶£æ—¨ã€‚Excel ã‚»ãƒ«é€�æ¬¡ I/O ã�¯è¡Œã‚�ã�ªã�„ï¼‰ã€‚
# False: å¾“æ�¥ã�©ã�Šã‚Š avail_dt / machine_avail_dt ã�®ã�¿ã€‚
DISPATCH_INTERVAL_MIRROR_ENFORCE = os.environ.get(
    "DISPATCH_INTERVAL_MIRROR_ENFORCE", "1"
).strip().lower() not in ("0", "false", "no", "off", "ã�„ã�„ã�ˆ", "ç„¡åŠ¹")


def _clone_attendance_day_shifted(source_day: dict, old_date: date, new_date: date) -> dict:
    """ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ¥å‹¤æ€ ãƒ–ãƒ­ãƒƒã‚¯ã‚’ new_date ã�«ã‚·ãƒ•ãƒˆã�—ã�Ÿæµ…ã�„ã‚³ãƒ”ãƒ¼ã‚’è¿”ã�™ã€‚"""
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
    """é…�å�°å�¯èƒ½ã�ªãƒ¡ãƒ³ãƒ�ãƒ¼ã�Œ1äººã�§ã‚‚ã�„ã‚‹ç›´è¿‘ã�®æ—¥ã‚’ãƒ†ãƒ³ãƒ—ãƒ¬ã�«æŽ¡ç”¨ï¼ˆæœ€çµ‚æ—¥ã�Œå…¨ä¼‘ã�§ã‚‚æœ‰åŠ¹ã�ªåž‹ã‚’ä½¿ã�†ï¼‰ã€‚"""
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
    """ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ä¸Š1æ—¥å…ˆã‚’ plan_dates ã�«è¿½åŠ ã�—ã€�ãƒ†ãƒ³ãƒ—ãƒ¬æ—¥ã�®ã‚·ãƒ•ãƒˆè¤‡è£½ã�§ attendance ã‚’åŸ‹ã‚�ã‚‹ã€‚å¤±æ•—æ™‚ Falseã€‚"""
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
        "é…�å�°å®Œäº†ã�¾ã�§å‹¤æ€ ã‚’è‡ªå‹•æ‹¡å¼µ: %s ã‚’è¿½åŠ ï¼ˆãƒ†ãƒ³ãƒ—ãƒ¬=%sã€�ãƒ¡ãƒ³ãƒ�ãƒ¼æ•°=%sï¼‰",
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
    plan_dates ã‚’å…ˆé ­ã�‹ã‚‰é †ã�« yieldã€‚æœ«å°¾ã�¾ã�§æ�¥ã�¦ã‚‚æ®‹ã‚¿ã‚¹ã‚¯ã�Œã�‚ã‚Œã�°å‹¤æ€ ã‚’1æ—¥ã�šã�¤æ‹¡å¼µã�—ã�¦ç¶™ç¶šã€‚
    plan_dates / attendance_data ã�¯ã‚¤ãƒ³ãƒ—ãƒ¬ãƒ¼ã‚¹æ›´æ–°ã�•ã‚Œã‚‹ã€‚
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
                "æ®‹ã‚¿ã‚¹ã‚¯ã�Œã�‚ã‚Šã�¾ã�™ã�Œå‹¤æ€ ã�®è‡ªå‹•æ‹¡å¼µã�Œä¸Šé™�ï¼ˆ%s æ—¥ï¼‰ã�«é�”ã�—ã�¾ã�—ã�Ÿã€‚é…�å�°æ®‹ãƒ»é…�å�°ä¸�å�¯ã�Œæ®‹ã‚‹å�¯èƒ½æ€§ã�Œã�‚ã‚Šã�¾ã�™ã€‚",
                SCHEDULE_EXTEND_MAX_EXTRA_DAYS,
            )
            return
        if not _extend_attendance_one_calendar_day(attendance_data, plan_dates):
            logging.warning(
                "å‹¤æ€ ã‚’1æ—¥æ‹¡å¼µã�§ã��ã�¾ã�›ã‚“ã�§ã�—ã�Ÿï¼ˆãƒ†ãƒ³ãƒ—ãƒ¬æ—¥ã�®ãƒ‡ãƒ¼ã‚¿æ¬ è�½ï¼‰ã€‚æ®‹ã‚¿ã‚¹ã‚¯ã�¯æœªå‰²å½“ã�®ã�¾ã�¾ã�§ã�™ã€‚"
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
    """ä¾�é ¼NO â†’ åŠ å·¥å†…å®¹ã�®å·¥ç¨‹å��ãƒªã‚¹ãƒˆï¼ˆè¡¨ã�®ä¸Šã�®æ–¹ã�§æœ€åˆ�ã�«ç�¾ã‚Œã�Ÿé�žç©ºã�®è¡Œã‚’æŽ¡ç”¨ï¼‰ã€‚"""
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
    å·¥ç¨‹å��ï¼ˆé…�å�°è¨ˆç”»ã�®ã€Œå·¥ç¨‹å��ã€�åˆ—ï¼‰ã�Œã€�å…ƒãƒ‡ãƒ¼ã‚¿ã�®ã€ŒåŠ å·¥å†…å®¹ã€�ã‚«ãƒ³ãƒžåŒºåˆ‡ã‚Šãƒˆãƒ¼ã‚¯ãƒ³ã�®ã�„ã�šã‚Œã�‹ã�¨
    æ­£è¦�åŒ–ä¸€è‡´ã�™ã‚‹ã�‹ã€‚ãƒˆãƒ¼ã‚¯ãƒ³ã�Œç„¡ã�„ï¼ˆåŠ å·¥å†…å®¹æœªè¨˜å…¥ã�®ä¾�é ¼ï¼‰ã�¯ç…§å�ˆå¯¾è±¡å¤–ã�¨ã�—ã�¦ Trueã€‚
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
    """é…�å�°è¨ˆç”»ã‚·ãƒ¼ãƒˆã�®ã€Œå„ªå…ˆåº¦ã€�ã€‚å°�ã�•ã�„ã�»ã�©å…ˆã€‚æœªå…¥åŠ›ãƒ»ä¸�æ­£ã�¯ 999ã€‚"""
    p = t.get("priority", 999)
    try:
        return int(p)
    except (TypeError, ValueError):
        return 999


def _task_blocked_by_same_request_dependency(task, task_queue) -> bool:
    """
    å�Œä¸€ä¾�é ¼NOã�®ç•°ã�ªã‚‹å·¥ç¨‹ã‚’å�Œæ™‚åˆ»ã�«å›žã�•ã�ªã�„ï¼ˆé…�å�°ãƒ«ãƒ¼ãƒ« Â§A-1ãƒ»Â§A-2ï¼‰ã€‚
    - ä¸¡è¡Œã�«åŠ å·¥å†…å®¹ç”±æ�¥ã�® rank ã�Œã�‚ã‚‹ã�¨ã��ã�¯ rank ã�®ã�¿ã�§å‰�å¾Œï¼ˆÂ§A-1ï¼‰ã€‚
    - ã�©ã�¡ã‚‰ã�‹ã�« rank ã�Œç„¡ã�„ã�¨ã��ã�¯ã€�é…�å�°è¨ˆç”»ã‚·ãƒ¼ãƒˆã�®è¡Œé † same_request_line_seq ã�§å‰�å¾Œï¼ˆÂ§A-2ï¼‰ã€‚
    Â§B-2 / Â§B-3: ``roll_pipeline_inspection`` ã�¾ã�Ÿã�¯ ``roll_pipeline_rewind`` è¡Œã�Œ
    ``roll_pipeline_ec`` å…ˆè¡Œã�«ã‚ˆã‚Š Â§A-1 ã�§æ­¢ã�¾ã‚‹å ´å�ˆã€�
    ``_roll_pipeline_inspection_assign_room`` > 0 ã�ªã‚‰å½“è©²ãƒšã‚¢ã� ã�‘ãƒ–ãƒ­ãƒƒã‚¯ã�—ã�ªã�„ã€‚
    å‰�é€²é…�å�°ã�§ã�¯ ``_trial_order_flow_eligible_tasks`` ã�Œ EC å®Œèµ°ã�¾ã�§æ¤œæŸ»ã‚’å¤–ã�™ã�Ÿã‚�ã€�
    EC æ®‹ã�Œã�‚ã‚‹é–“ã�¯æœ¬åˆ†å²�ã�«åˆ°é�”ã�—ã�ªã�„ã€‚ãƒªãƒ¯ã‚¤ãƒ³ãƒ‰ç­‰ã�§æ¤œæŸ»ã�Œè¼‰ã‚‹å±€é�¢ã�¨ã�®æ•´å�ˆç”¨ã€‚
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
    ã‚­ãƒ¥ãƒ¼çŠ¶æ…‹ä¸Šã€�ã�“ã�®è¡Œã�¯ã�¾ã� æ—¥æ¬¡é…�å�°ã�§é€²ã‚�ã‚‰ã‚Œã�ªã�„ï¼ˆÂ§A å�Œä¸€ä¾�é ¼ã�®å‰�å·¥ç¨‹æ®‹ã€�ã�¾ã�Ÿã�¯ Â§B-2/Â§B-3 ã�®æž ã‚¼ãƒ­ï¼‰ã€‚
    `_min_pending_dispatch_trial_order_for_date` ã�¨ `_equipment_line_lower_dispatch_trial_still_pending`
    ã�§å�Œã�˜åŸºæº–ã‚’å…±æœ‰ã�™ã‚‹ã€‚ç‰‡æ–¹ã� ã�‘ç›´ã�™ã�¨ã€�å�Œä¸€è¨­å‚™ã‚­ãƒ¼ã�§å…¨ä»¶æœªå‰²å½“ã�Œæ®‹ã‚‹ãƒ‡ãƒƒãƒ‰ãƒ­ãƒƒã‚¯ã�Œèµ·ã��å¾—ã‚‹ã€‚
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
    """ç†±èž�ç�€æ¤œæŸ»è¡Œã�®ã�¿ã�®ç´¯è¨ˆå®Œäº†ãƒ­ãƒ¼ãƒ«ï¼ˆãƒˆãƒ¬ãƒ¼ã‚¹ç”¨ï¼‰ã€‚"""
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
    """Â§B-2 æ¤œæŸ»è¡Œï¼‹Â§B-3 å·»è¿”ã�—è¡Œã�®ã€�å�Œä¸€ä¾�é ¼å†…ã�®å¾Œç¶šãƒ‘ã‚¤ãƒ—ãƒ©ã‚¤ãƒ³ç´¯è¨ˆå®Œäº†ãƒ­ãƒ¼ãƒ«ã€‚"""
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
    """å�Œä¸€ä¾�é ¼NOã�« ECï¼ˆãƒ­ãƒ¼ãƒ«ãƒ‘ã‚¤ãƒ—ãƒ©ã‚¤ãƒ³å…ˆè¡Œï¼‰ã‚¿ã‚¹ã‚¯ã�Œã‚­ãƒ¥ãƒ¼ã�«å�«ã�¾ã‚Œã‚‹ã�‹ã€‚"""
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
    """å�Œä¸€ä¾�é ¼NOã�® EC ãƒ­ãƒ¼ãƒ«ãƒ‘ã‚¤ãƒ—ãƒ©ã‚¤ãƒ³è¡Œã�Œã�™ã�¹ã�¦æ®‹é‡�ã‚¼ãƒ­ï¼ˆå®Œèµ°ï¼‰ã�‹ã€‚"""
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
    # EC å…¨ãƒ­ãƒ¼ãƒ«å®Œäº†å¾Œã�¯ã€ŒEC å…ˆè¡Œãƒ»ãƒ�ãƒƒãƒ•ã‚¡ã€�ã�¯æ—¢ã�«æº€ã�Ÿã�•ã‚Œã�¦ã�„ã‚‹ã€‚ã�“ã�“ã�§ max_insp ã‚’ ec_done ã�«
    # æ�ƒã�ˆã‚‹ã�¨ã€�ã‚·ãƒ¼ãƒˆä¸Šã�®æ¤œæŸ»ï¼ˆãƒ»å·»è¿”ã�—ï¼‰æ®‹ãƒ­ãƒ¼ãƒ«æ•°ã�Œ EC å®Œäº†ãƒ­ãƒ¼ãƒ«æ•°ã‚’ä¸Šå›žã‚‹ãƒ‡ãƒ¼ã‚¿ã�§
    # max_insp - insp_done ã�Œ 0 ã�®ã�¾ã�¾æ®‹ã‚Šã€�æ¤œæŸ»è¡Œã�Œ eligible ã�‹ã‚‰å¤–ã‚Œé…�å�°è©¦è¡Œé †ã�Œæ°¸ä¹…ã�«è©°ã�¾ã‚‹
    # ï¼ˆå†�ç�¾ãƒ­ã‚°: ec_fully_done ã�‹ã�¤ insp_done==max_insp==ec_done ã�§ room=0 â†’ å¾Œç¶šè©¦è¡Œé †ã�Œé…�å�°ä¸�å�¯ï¼‰ã€‚
    if _pipeline_ec_fully_done_for_tid(task_queue, task_id):
        return float(ROLL_PIPELINE_INSP_UNCAPPED_ROOM)
    # EC ç¨¼åƒ�ä¸­: å…ˆè¡Œãƒ�ãƒƒãƒ•ã‚¡ B ã�«ã‚ˆã‚Šæ¤œæŸ»ãƒ­ãƒ¼ãƒ«ä¸Šé™�ã‚’ ec_done ã�‹ã‚‰é�…å»¶ã�•ã�›ã‚‹ï¼ˆB=2 ã�®å¼�ã�¯ã‚³ãƒ¡ãƒ³ãƒˆå�‚ç…§ï¼‰ã€‚
    max_insp = max(0.0, ec_done - float(ROLL_PIPELINE_INITIAL_BUFFER_ROLLS) + 1.0)
    _room = max(0.0, max_insp - insp_done)
    return _room


def _roll_pipeline_inspection_task_row_for_tid(
    task_queue: list, task_id: str
) -> dict | None:
    """å�Œä¸€ä¾�é ¼NOã�® Â§B-2 æ¤œæŸ»è¡Œã�¾ã�Ÿã�¯ Â§B-3 å·»è¿”ã�—è¡Œã‚’1ä»¶è¿”ã�™ã€‚ç„¡ã�‘ã‚Œã�° Noneã€‚"""
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
    """å�Œä¸€ä¾�é ¼ã�® EC ãƒ­ãƒ¼ãƒ«ç¢ºå®šã�”ã�¨ã�®çµ‚äº†æ™‚åˆ»ã‚’æ™‚ç³»åˆ—ã�§è¿”ã�™ï¼ˆassigned_history ã�® end_dtï¼‰ã€‚"""
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
    æ¬¡ã�®æ¤œæŸ»ãƒ­ãƒ¼ãƒ«ã‚’é–‹å§‹ã�—ã�¦ã‚ˆã�„æœ€æ—©æ™‚åˆ»ã€‚
    ç´¯è¨ˆæ¤œæŸ»å®Œäº†ãƒ­ãƒ¼ãƒ«æ•°ã‚’ Kã€�ãƒ�ãƒƒãƒ•ã‚¡ã‚’ Bï¼ˆ=ROLL_PIPELINE_INITIAL_BUFFER_ROLLSï¼‰ã�¨ã�™ã‚‹ã�¨ã€�
    EC å®Œäº†ãƒ­ãƒ¼ãƒ«ã�Œæ™‚ç³»åˆ—ã�§ (K+B) æœ¬ç›®ã�«åˆ°é�”ã�—ã�Ÿæ™‚åˆ»ï¼ˆã��ã�®ãƒ­ãƒ¼ãƒ«ã�® end_dtï¼‰æœªæº€ã�«ã�¯é–‹å§‹ã�—ã�ªã�„ã€‚
    ï¼ˆæ¥­å‹™ãƒ«ãƒ¼ãƒ«: ä»»æ„�ã�®æ™‚ç‚¹ã�§ EC_RollEndCount - KENSA_RollEndCount >= B ã‚’æº€ã�Ÿã�™ã�¾ã�§æ¤œæŸ»ã‚’é€²ã‚�ã�ªã�„ã€�
    ã�®ã€Œãƒ­ãƒ¼ãƒ«çµ‚äº†æ™‚åˆ»åŸºæº–ã€�ã�®å®Ÿè£…ã€‚ï¼‰
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
    """assigned_history ã�® team æ–‡å­—åˆ—ï¼ˆä¸»ãƒ»è£œã‚’ã€Œ,ã€�ã€Œã€�ã€�åŒºåˆ‡ã‚Šï¼‰ã�‹ã‚‰æ‹…å½“è€…å��ã‚’æŠ½å‡ºï¼ˆNFKCï¼‰ã€‚"""
    if team_cell is None:
        return set()
    s = str(team_cell).strip()
    if not s:
        return set()
    out: set[str] = set()
    for part in re.split(r"[,ã€�]", s):
        t = part.strip()
        if t:
            out.add(unicodedata.normalize("NFKC", t))
    return out


def _pipeline_b2_assigned_member_names_nfkc_for_side(
    task_queue: list, task_id: str, *, ec_side: bool
) -> set[str]:
    """å�Œä¸€ä¾�é ¼ã�® EC è¡Œã�¾ã�Ÿã�¯æ¤œæŸ»è¡Œã�® assigned_history ã�«å‡ºã�Ÿæ‹…å½“è€…å��ï¼ˆNFKC é›†å�ˆï¼‰ã€‚"""
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
    """å�Œä¸€ä¾�é ¼NOã�« Â§B-2/Â§B-3 ã�® EC è¡Œã�¨å¾Œç¶šè¡Œï¼ˆæ¤œæŸ»ã�¾ã�Ÿã�¯å·»è¿”ã�—ï¼‰ã�®ä¸¡æ–¹ã�Œã‚­ãƒ¥ãƒ¼ã�«ã�‚ã‚‹ã�‹ã€‚"""
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
    Â§B-2 / Â§B-3 å�Œä¸€ä¾�é ¼ã�§ã�¯ã€�EC è¡Œã�«ä¸€åº¦ã�§ã‚‚å…¥ã�£ã�Ÿè€…ã�¯å¾Œç¶šï¼ˆæ¤œæŸ»ï¼�å·»è¿”ã�—ï¼‰ã�®å€™è£œã�‹ã‚‰å¤–ã�—ã€�
    å¾Œç¶šã�«å…¥ã�£ã�Ÿè€…ã�¯ EC ã�®å€™è£œã�‹ã‚‰å¤–ã�™ã€‚
    ï¼ˆç¤¾å†…ãƒ«ãƒ¼ãƒ«: æ‹…å½“è€…é›†å�ˆã‚’å¿…ã�šåˆ†ã�‘ã‚‹ã€‚`PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS` ã�§ç„¡åŠ¹åŒ–å�¯ï¼‰
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
            _side = "å·»è¿”ã�—"
        else:
            _side = "æ¤œæŸ»"
        _log_dispatch_trace_schedule(
            tid,
            "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ãƒ–ãƒ­ãƒƒã‚¯åˆ¤å®š: B-2æ‹…å½“è€…åˆ†é›¢ side=%s machine=%s "
            "å€™è£œé™¤å¤–=%s æ®‹å€™è£œ=%s(%s)",
            tid,
            _side,
            task.get("machine"),
            ",".join(str(x) for x in removed),
            len(filtered),
            ",".join(str(x) for x in filtered) if filtered else "ã�ªã�—",
        )
    return filtered


def _exclusive_b1_inspection_holder_for_machine(task_queue, occupant_key: str):
    """
    å�Œä¸€ç‰©ç�†æ©Ÿæ¢°ï¼ˆæ©Ÿæ¢°å��ãƒ™ãƒ¼ã‚¹ã�®å� æœ‰ã‚­ãƒ¼ï¼‰ä¸Šã�§ã€�Â§B-2 ç†±èž�ç�€æ¤œæŸ»ã�¾ã�Ÿã�¯ Â§B-3 å·»è¿”ã�—ã�Œ **æ—¢ã�«å‰²ä»˜ã‚’é–‹å§‹** ã�—æ®‹ãƒ­ãƒ¼ãƒ«ã�Œæ®‹ã‚‹è¡Œã�Œã�‚ã‚Œã�°
    ã��ã�®ã‚¿ã‚¹ã‚¯ dict ã‚’1ä»¶è¿”ã�™ï¼ˆã�ªã�‘ã‚Œã�° Noneï¼‰ã€‚

    ãƒ‘ã‚¤ãƒ—ãƒ©ã‚¤ãƒ³æž ã�§å¾Œç¶šã‚’æ•°ãƒ­ãƒ¼ãƒ«ã�šã�¤ã�—ã�‹å…¥ã‚Œã�ªã�„è¨­è¨ˆã�®ã�Ÿã‚�ã€�æž ã‚¼ãƒ­ã�®éš™é–“ã�« **åˆ¥ä¾�é ¼** ã�Œå�Œã�˜è¨­å‚™ã�«å…¥ã‚Šã€�
    çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ã�§ã‚¿ã‚¹ã‚¯è¡¨ç¤ºã�Œé€”ä¸­ã�§åˆ‡ã‚Šæ›¿ã‚�ã‚‹äº‹è±¡ã‚’é˜²ã��ã€‚å� æœ‰ä¸­ã�¯å½“è©²ç‰©ç�†æ©Ÿæ¢°ã�§ã�¯ä»–ã‚¿ã‚¹ã‚¯ã‚’è©¦è¡Œã�—ã�ªã�„ã€‚
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
    """need ã‚·ãƒ¼ãƒˆã�§å·¦ã�«ã�‚ã‚‹ã€Œå·¥ç¨‹å��+æ©Ÿæ¢°å��ã€�åˆ—ã�»ã�©å°�ã�•ã�„å€¤ï¼ˆã‚­ãƒ¥ãƒ¼ã�§å…ˆï¼‰ã€‚"""
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
    generate_plan å†’é ­ã�Šã‚ˆã�³ç´�æœŸã‚·ãƒ•ãƒˆå†�è©¦è¡Œæ™‚ã�® task_queue.sort ç”¨ã‚­ãƒ¼ã€‚

    1. åŠ å·¥é€”ä¸­ï¼ˆin_progressï¼‰ã‚’å…ˆ
    2. ç´�æœŸåŸºæº– due_basis_dateï¼ˆå›žç­”ç´�æœŸâ†’æŒ‡å®šç´�æœŸã€‚æ—©ã�„ã�»ã�©å…ˆï¼‰
    3. Â§B-1 â†’ Â§B-2/Â§B-3 å¸¯ â†’ ã��ã�®ä»–ï¼ˆb_tierï¼‰
    4. Â§B-2/Â§B-3 å¸¯å†…ã�®ã�¿ EC ã‚’æœªç�€æ‰‹ã�®æ¤œæŸ»ï¼�å·»è¿”ã�—ã‚ˆã‚Šå…ˆï¼ˆb2_queue_subï¼‰
    5. need ã‚·ãƒ¼ãƒˆå·¦åˆ—ã�»ã�©å…ˆï¼ˆå·¥ç¨‹å��+æ©Ÿæ¢°å��åˆ—ã�®ä½�ç½®ï¼‰
    6. ä¾�é ¼NOã‚¿ã‚¤ãƒ–ãƒ¬ãƒ¼ã‚¯ï¼ˆ_task_id_same_machine_due_tiebreak_keyï¼‰

    _req_map / _need_rules ã�¯å‘¼ã�³å‡ºã�—äº’æ�›ã�®ã�Ÿã‚�æ®‹ã�™ã€‚
    """
    insp = bool(task.get("roll_pipeline_inspection"))
    rw = bool(task.get("roll_pipeline_rewind"))
    ip = bool(task.get("in_progress"))
    ec = bool(task.get("roll_pipeline_ec"))
    if insp and ip:
        b_tier = 0  # Â§B-1
    elif ec or (insp and not ip) or (rw and not ip):
        b_tier = 1  # Â§B-2 / Â§B-3 å¸¯
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
    Â§B-2 / Â§B-3: å�Œä¸€ task_id ã�® `roll_pipeline_ec` è¡Œã�®ç›´å¾Œã�«ã€�æœªç�€æ‰‹ã�®å¾Œç¶šè¡Œ
    ï¼ˆ`roll_pipeline_inspection` ã�¾ã�Ÿã�¯ `roll_pipeline_rewind`ï¼‰ã‚’è¡Œé †ã�§éš£æŽ¥ã�•ã�›ã‚‹ã€‚
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
            "Â§B-2/Â§B-3 é…�å�°è©¦è¡Œé †: EC ã�¨æœªç�€æ‰‹å¾Œç¶šï¼ˆæ¤œæŸ»ï¼�å·»è¿”ã�—ï¼‰ã‚’éš£æŽ¥ã�—ã�Ÿä¾�é ¼NO: %s",
            ",".join(moved_tids),
        )


def _assign_sequential_dispatch_trial_order(task_queue: list) -> None:
    """
    `task_queue` ã�®ãƒªã‚¹ãƒˆé †ã�«å�ˆã‚�ã�›ã�¦ `dispatch_trial_order` ã‚’ 1..n ã�¸ä»˜ã�‘ç›´ã�™ã€‚
    `_reorder_task_queue_b2_ec_inspection_consecutive` ã�®ç›´å¾Œï¼ˆã�Šã‚ˆã�³ã‚­ãƒ¥ãƒ¼å†�ã‚½ãƒ¼ãƒˆã�®ç›´å¾Œï¼‰ã�«å‘¼ã�³ã€�
    EC ã�¨å¾Œç¶šï¼ˆæ¤œæŸ»ï¼�å·»è¿”ã�—ï¼‰ã�®é€£ç¶šç•ªå�·ã‚’ä¿�è¨¼ã�™ã‚‹ã€‚
    """
    for i, t in enumerate(task_queue, start=1):
        t["dispatch_trial_order"] = i


def _task_queue_all_have_sheet_dispatch_trial_order(task_queue: list) -> bool:
    """é…�å�°è¨ˆç”»ã‚·ãƒ¼ãƒˆã�®ã€Œé…�å�°è©¦è¡Œé †ç•ªã€�ã�Œã‚­ãƒ¥ãƒ¼å…¨è¡Œã�«æ­£ã�®æ•´æ•°ã�§å…¥ã�£ã�¦ã�„ã‚‹ã�‹ã€‚"""
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
    é…�å�°è©¦è¡Œé †ã�®ç¢ºå®šã€‚ã‚·ãƒ¼ãƒˆã�«å…¨è¡Œåˆ†ã�®è©¦è¡Œé †ã�Œã�‚ã‚Œã�°ã��ã‚Œã‚’æŽ¡ç”¨ï¼ˆÂ§B-2/3 ã�®éš£æŽ¥ç¹°ã‚Šä¸Šã�’ã�¯è¡Œã‚�ã�ªã�„ï¼‰ã€‚
    æ¬ æ��ã�Œã�‚ã‚Œã�°å¾“æ�¥ã�©ã�Šã‚Šãƒžã‚¹ã‚¿ãƒ»ç´�æœŸãƒ»need åˆ—é †ã�ªã�©ã�§ã‚½ãƒ¼ãƒˆã�—ã€�EC éš£æŽ¥å¾Œã�« 1..n ã‚’ä»˜ä¸Žã€‚
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
            "é…�å�°è©¦è¡Œé †ç•ª: ã€Œ%sã€�åˆ—ã�®å€¤ã‚’ã��ã�®ã�¾ã�¾ä½¿ç”¨ã�—ã�¾ã�—ã�Ÿï¼ˆå…¨ %s è¡Œï¼‰ã€‚",
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
        "é…�å�°è©¦è¡Œé †ç•ª: ãƒžã‚¹ã‚¿ãƒ»ã‚¿ã‚¹ã‚¯å…¥åŠ›ã�‹ã‚‰è‡ªå‹•è¨ˆç®—ã�— 1..%s ã‚’ä»˜ä¸Žã�—ã�¾ã�—ã�Ÿã€‚",
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
    æ®µéšŽ1å‡ºåŠ› DataFrame ã�®ã€Œé…�å�°è©¦è¡Œé †ç•ªã€�ã‚’ã€�æ®µéšŽ2 å†’é ­ã�¨å�Œã�˜æ‰‹é †ï¼ˆã‚½ãƒ¼ãƒˆãƒ»Â§B-2/3 éš£æŽ¥ãƒ»é€£ç•ªï¼‰ã�§åŸ‹ã‚�ã‚‹ã€‚
    é…�å�°å¯¾è±¡å¤–ã�®è¡Œã�¯ç©ºã�®ã�¾ã�¾ã€‚
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
            # Excel ä¸Šã�¯æ•°å€¤ã‚»ãƒ«ã�«ã�—ã€�ãƒ•ã‚£ãƒ«ã‚¿ãƒ¼ãƒ»ä¸¦ã�¹æ›¿ã�ˆã‚’ã�—ã‚„ã�™ã��ã�™ã‚‹ï¼ˆæ–‡å­—åˆ—ã� ã�¨æ•°å€¤ã�¨åˆ¥ã‚°ãƒ«ãƒ¼ãƒ—ã�«ã�ªã‚‹ï¼‰
            plan_df.iat[iloc, col_idx] = int(dto)
        except (TypeError, ValueError):
            plan_df.iat[iloc, col_idx] = float("nan")


def _equipment_schedule_unified_sub_string_map(timeline_for_eq_grid: list) -> dict:
    """
    å�Œä¸€æ—¥ãƒ»å�Œä¸€è¨­å‚™åˆ—ã‚­ãƒ¼ãƒ»å�Œä¸€ä¾�é ¼NO ã�®åŠ å·¥ã�«ã�¤ã�„ã�¦ã€�è¨­å‚™æ™‚é–“å‰²ã‚»ãƒ«ç”¨ã�®ã€Œè£œã€�è¡¨ç¤ºæ–‡å­—åˆ—ã€‚
    ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ä¸Šã�®å�„ãƒ–ãƒ­ãƒƒã‚¯ã�® `sub` ã�«ç�¾ã‚Œã�Ÿè£œåŠ©è€…å��ã‚’å’Œé›†å�ˆã�—ã€�æ˜‡é †ã�§ ", " é€£çµ�ã�™ã‚‹ã€‚
    ãƒ¡ãƒ³ãƒ�ãƒ¼æ—¥ç¨‹ãƒ»å� æœ‰è¨ˆç®—ã�«ä½¿ã�†ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ã�® `sub` ã�¯å¤‰æ›´ã�—ã�ªã�„ï¼ˆè¡¨ç¤ºå°‚ç”¨ï¼‰ã€‚
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
    çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ã�¨å�Œå½¢å¼�ã�® DataFrameï¼ˆ10 åˆ†æž ãƒ»è¨­å‚™åˆ—ï¼‹é€²åº¦åˆ—ï¼‰ã€‚
    first_eq_schedule_cell_by_task_id ã‚’æ¸¡ã�—ã�Ÿã�¨ã��ã�®ã�¿ã€�åˆ�å‡ºã‚»ãƒ«åº§æ¨™ã‚’è¨˜éŒ²ï¼ˆçµ�æžœãƒ�ã‚¤ãƒ‘ãƒ¼ãƒªãƒ³ã‚¯ç”¨ï¼‰ã€‚
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
        eq_empty_cols[f"{eq}é€²åº¦"] = ""

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

        all_eq_rows.append({"æ—¥æ™‚å¸¯": f"â–  {d.strftime('%Y/%m/%d (%a)')} â– ", **eq_empty_cols})

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
                "æ—¥æ™‚å¸¯": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"
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
                        eq_text = "ä¼‘æ†©"
                    elif not _use_prog:
                        _ek_disp = _timeline_event_kind(active_ev)
                        _tag = {
                            TIMELINE_EVENT_MACHINE_DAILY_STARTUP: "æ—¥æ¬¡å§‹æ¥­æº–å‚™",
                            TIMELINE_EVENT_CHANGEOVER_CLEANUP: "ä¾�é ¼åˆ‡æ›¿å¾Œå§‹æœ«",
                            TIMELINE_EVENT_CHANGEOVER_PREP: "åŠ å·¥å‰�æº–å‚™",
                        }.get(
                            _ek_disp,
                            "ã‚»ãƒƒãƒˆã‚¢ãƒƒãƒ—",
                        )
                        _sub_n = _eq_cell_display_sub(active_ev, d)
                        _sub_text = f" è£œ:{_sub_n}" if _sub_n else ""
                        _tid_d = str(active_ev.get("task_id") or "").strip()
                        eq_text = (
                            f"[{_tid_d}] ä¸»:{active_ev.get('op', '')}{_sub_text} ({_tag})"
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
                        sub_text = f" è£œ:{_sub_s}" if _sub_s else ""
                        eq_text = f"[{active_ev['task_id']}] ä¸»:{active_ev['op']}{sub_text}"
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
                row_data[f"{eq}é€²åº¦"] = progress_text

            all_eq_rows.append(row_data)
            curr_grid = next_grid
        all_eq_rows.append({"æ—¥æ™‚å¸¯": "", **eq_empty_cols})

    df_eq = pd.DataFrame(all_eq_rows)
    _eq_hdr = _equipment_schedule_header_labels(equipment_list)
    _eq_rename = {}
    for _eq, _lab in zip(equipment_list, _eq_hdr):
        if _eq in df_eq.columns:
            _eq_rename[_eq] = _lab
        _pqc = f"{_eq}é€²åº¦"
        if _pqc in df_eq.columns:
            _eq_rename[_pqc] = f"{_lab}é€²åº¦"
    if _eq_rename:
        df_eq = df_eq.rename(columns=_eq_rename)
    return df_eq


def _machine_display_key_for_equipment(eq: str) -> str:
    """skills åˆ—ã‚­ãƒ¼ã€Œå·¥ç¨‹+æ©Ÿæ¢°ã€�ã�‹ã‚‰æ©Ÿæ¢°å��è¡¨ç¤ºã‚­ãƒ¼ã‚’å¾—ã‚‹ï¼ˆé‡�è¤‡æ™‚ã�¯è¤‡å�ˆã‚­ãƒ¼ã�”ã�¨ã�«åˆ¥åˆ—ï¼‰ã€‚"""
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
    æ©Ÿæ¢°å��å�˜ä½�ã�«åˆ—ã‚’ã�¾ã�¨ã‚�ã€�å�„ 10 åˆ†æž ã�§å� æœ‰ä¸­ã�®ä¾�é ¼NOï¼ˆè¤‡æ•°æ™‚ã�¯ã€Œï¼�ã€�ï¼‰ã‚’è¡¨ç¤ºã�™ã‚‹ã€‚
    åˆ—è¦‹å‡ºã�—ã�¯æ©Ÿæ¢°å��ã�®ã�¿ï¼ˆå·¥ç¨‹+æ©Ÿæ¢°ã�®è¤‡å�ˆã‚­ãƒ¼ã�¯ä»˜ã�‘ã�ªã�„ï¼‰ã€‚å�Œä¸€ç‰©ç�†æ©Ÿæ¢°ã�¯å� æœ‰ã‚­ãƒ¼ã�§1åˆ—ã�«é›†ç´„ã�™ã‚‹ã€‚
    """
    timeline_for_eq_grid = _expand_timeline_events_for_equipment_grid(timeline_events)
    events_by_date = defaultdict(list)
    for e in timeline_for_eq_grid:
        events_by_date[e["date"]].append(e)

    # å� æœ‰ã‚­ãƒ¼ï¼ˆæ©Ÿæ¢°å��å�´ãƒ»æ­£è¦�åŒ–ï¼‰ã�”ã�¨ã�«1åˆ—ã€‚è¦‹å‡ºã�—ã�¯ equipment_list åˆ�å‡ºã�®æ©Ÿæ¢°å��è¡¨ç¤ºã�®ã�¿ã€‚
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

        all_rows.append({"æ—¥æ™‚å¸¯": f"â–  {d.strftime('%Y/%m/%d (%a)')} â– ", **empty_tail})

        curr_grid = d_start
        while curr_grid < d_end:
            next_grid = curr_grid + timedelta(minutes=10)
            if next_grid > d_end:
                next_grid = d_end
            mid_t = curr_grid + (next_grid - curr_grid) / 2
            row_data = {
                "æ—¥æ™‚å¸¯": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"
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
                    tids_by_mcol[mcol].add("ï¼ˆä¼‘æ†©ï¼‰")
                else:
                    tid = str(active_ev.get("task_id") or "").strip()
                    if tid:
                        tids_by_mcol[mcol].add(tid)
            for mcol in machine_cols:
                parts = sorted(tids_by_mcol.get(mcol, ()))
                row_data[mcol] = "ï¼�".join(parts) if parts else ""
            all_rows.append(row_data)
            curr_grid = next_grid
        all_rows.append({"æ—¥æ™‚å¸¯": "", **empty_tail})

    return pd.DataFrame(all_rows)


def _build_block_table_dataframe(
    sorted_dates: list,
    equipment_list: list,
    members: list,
    attendance_data: dict,
    timeline_events: list,
) -> "pd.DataFrame":
    """
    è¨­å‚™åˆ—ï¼ˆå� æœ‰ä¸­ã�®ä¾�é ¼NOï¼‰ï¼‹ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ—ï¼ˆå�Œï¼‰ã‚’ 10 åˆ†æž ã�§ä¸¦ã�¹ã�Ÿãƒ–ãƒ­ãƒƒã‚¯å�¯è¦–åŒ–ç”¨ã‚·ãƒ¼ãƒˆã€‚
    """
    timeline_for_eq_grid = _expand_timeline_events_for_equipment_grid(timeline_events)
    events_by_date = defaultdict(list)
    for e in timeline_for_eq_grid:
        events_by_date[e["date"]].append(e)

    _eq_hdr = _equipment_schedule_header_labels(equipment_list)
    eq_disp_to_key: dict[str, str] = {}
    for eq, lab in zip(equipment_list, _eq_hdr):
        eq_disp_to_key[f"è¨­å‚™:{lab}"] = eq

    mem_cols = [f"äºº:{m}" for m in members]
    eq_cols = [f"è¨­å‚™:{lab}" for lab in _eq_hdr]
    all_cols = ["æ—¥æ™‚å¸¯"] + eq_cols + mem_cols
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

        banner = {"æ—¥æ™‚å¸¯": f"â–  {d.strftime('%Y/%m/%d (%a)')} â– "}
        banner.update({c: "" for c in all_cols if c != "æ—¥æ™‚å¸¯"})
        rows_out.append(banner)

        curr_grid = d_start
        while curr_grid < d_end:
            next_grid = curr_grid + timedelta(minutes=10)
            if next_grid > d_end:
                next_grid = d_end
            mid_t = curr_grid + (next_grid - curr_grid) / 2
            row_data: dict = {
                "æ—¥æ™‚å¸¯": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"
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
                    row_data[col_eq] = "ä¼‘æ†©"
                else:
                    tid = str(active_ev.get("task_id") or "").strip()
                    row_data[col_eq] = tid if tid else "å� æœ‰"

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
                        busy_member_task[op].add("ä¼‘æ†©" if tid else "ä¼‘æ†©")
                    elif tid:
                        busy_member_task[op].add(tid)
                for s in str(ev.get("sub") or "").split(","):
                    s = s.strip()
                    if not s:
                        continue
                    if any(
                        b_s <= mid_t < b_e for b_s, b_e in ev.get("breaks") or ()
                    ):
                        busy_member_task[s].add("ä¼‘æ†©")
                    elif tid:
                        busy_member_task[s].add(tid)

            for m in members:
                col_m = f"äºº:{m}"
                parts = sorted(busy_member_task.get(m, ()))
                row_data[col_m] = "ï¼�".join(parts) if parts else ""

            rows_out.append(row_data)
            curr_grid = next_grid

        tail = {"æ—¥æ™‚å¸¯": ""}
        tail.update({c: "" for c in all_cols if c != "æ—¥æ™‚å¸¯"})
        rows_out.append(tail)

    return pd.DataFrame(rows_out, columns=all_cols)


def _day_schedule_task_sort_key(
    task: dict,
    _task_queue: list | None = None,
    need_combo_col_index: dict | None = None,
):
    """
    å�Œä¸€æ—¥å†…ã�®å‰²ä»˜è©¦è¡Œé †ï¼ˆSTAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST=0 ã�®ä¸»ãƒ«ãƒ¼ãƒ—ç”¨ï¼‰ã€‚
    å…ˆé ­ã‚­ãƒ¼ã�¯ _generate_plan_task_queue_sort_key ã�¨å�Œè¶£æ—¨ï¼ˆåŠ å·¥é€”ä¸­ãƒ»ç´�æœŸåŸºæº– due_basis_dateãƒ»Â§B æ®µãƒ»b2_queue_subãƒ»need åˆ—é †ãƒ»ä¾�é ¼NOï¼‰ã€‚
    ç¶šã�‘ã�¦ Â§B-1 ã�®é…�å�°è©¦è¡Œé †ç¹°ã‚Šä¸Šã�’ã€�å·¥ç¨‹ rankã€�dispatch_trial_orderã€�Â§B-2 æ®µå†… EC å…ˆè¡Œã€�å„ªå…ˆåº¦ã€�çµ�æžœç”¨ã‚­ãƒ¼ã€‚
    å�Œä¸€ç‰©ç�†æ©Ÿæ¢°ä¸Šã�®éš™é–“å‰²ã‚Šè¾¼ã�¿ã�¯ _equipment_line_lower_dispatch_trial_still_pending ã�§è©¦è¡Œé †ã‚’å¼·åˆ¶ã�™ã‚‹ã€‚
    STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT=1 ã�®ã�¨ã��ã�¯ _task_blocked_by_global_dispatch_trial_order ã�Œ
    ã‚ˆã‚Šå°�ã�•ã�„è©¦è¡Œé †ã�®æœªå®Œäº†ã‚’è·¨ã�„ã� å‰²ã‚Šè¾¼ã�¿ã‚’åˆ¥é€”ãƒ–ãƒ­ãƒƒã‚¯ã�™ã‚‹ã€‚
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
    å�Œä¸€ç‰©ç�†æ©Ÿæ¢°ï¼ˆmachine å� æœ‰ã‚­ãƒ¼ï¼‰ä¸Šã�§ã€�ã‚ˆã‚Šå°�ã�•ã�„é…�å�°è©¦è¡Œé †ã�®è¡Œã�Œã�¾ã� æ®‹é‡�ã‚’æŒ�ã�¤ã�‹ã€‚
    machine_avail_dt ã�¯ãƒ�ãƒ£ãƒ³ã‚¯é–“ã�®éš™é–“ã�«å¾Œç¶šè©¦è¡Œé †ã�Œå…¥ã‚Šè¾¼ã‚�ã‚‹ã�Ÿã‚�ã€�ã�“ã�“ã�§é †åº�ã‚’å¼·åˆ¶ã�™ã‚‹ã€‚
    è¨­å‚™ã‚’è·¨ã�„ã� è©¦è¡Œé †ã�®å‰�å¾Œã�¯ _task_blocked_by_global_dispatch_trial_order ã�§åˆ¥é€”åˆ¶å¾¡ã�™ã‚‹ã€‚

    ã‚­ãƒ¥ãƒ¼å…ˆé ­ã�«æ®‹é‡�ã�Œã�‚ã‚‹ã� ã�‘ã�§ã�¯ãƒ–ãƒ­ãƒƒã‚¯ã�—ã�ªã�„ã€‚tasks_today ã�¨å�Œæ§˜ã�«
    start_date_req <= current_date ã�®è¡Œã� ã�‘ã‚’ã€Œå…ˆè©¦è¡Œé †ã�®ç«¶å�ˆã€�ã�¨ã�¿ã�ªã�™ã€‚
    ï¼ˆã�¾ã� é–‹å§‹æ—¥ã�«é�”ã�—ã�¦ã�„ã�ªã�„è¡Œã�Œå…¨æ—¥ãƒ–ãƒ­ãƒƒã‚«ãƒ¼ã�«ã�ªã‚Šã€�å¾Œç¶šã�Œã�»ã�¼é…�å�°ä¸�å�¯ã�«ã�ªã‚‹ã�®ã‚’é˜²ã��ã€‚ï¼‰

    ã‚ˆã‚Šå°�ã�•ã�„è©¦è¡Œé †ã�®è¡Œã�Œ **å�Œä¸€ä¾�é ¼ã�®å‰�å·¥ç¨‹å¾…ã�¡ç­‰ã�§ã�¾ã� å‰²ä»˜ä¸�èƒ½**ã�ªã�¨ã��ã�¯ã€Œç«¶å�ˆã�®æ®‹ã€�ã�¨ã�¿ã�ªã�•ã�ªã�„ã€‚
    ï¼ˆå½“è©²è¡Œã�¯ eligible ã�«ã‚‚å…¥ã‚‰ã�ªã�„ã�Ÿã‚�ã€�ã�“ã�“ã�§å¾…ã�Ÿã�›ã‚‹ã�¨å¾Œç¶šè©¦è¡Œé †ã�Œå�Œä¸€è¨­å‚™ã�§æ°¸ä¹…å�œæ­¢ã�—å¾—ã‚‹ã€‚ï¼‰

    ã‚ˆã‚Šå°�ã�•ã�„è©¦è¡Œé †ã�®è¡Œã�Œ **å½“æ—¥ã�®æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã� ã�‘ã�§è¨ˆç”»çª“ã‚’å…¨æ—¥å� æœ‰**ï¼ˆã��ã�®è¨­å‚™ã�¯å½“æ—¥ã‚¹ãƒ­ãƒƒãƒˆã‚¼ãƒ­ï¼‰ã�ªã‚‰
    ã€Œç«¶å�ˆã�®æ®‹ã€�ã�¨ã�¿ã�ªã�•ã�ªã�„ï¼ˆã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è©¦è¡Œé †ã�¨ã�‚ã‚�ã�›ã�¦ä»–è¨­å‚™ã�Œå…¨æ—¥æ­¢ã�¾ã‚‹ã�®ã‚’é˜²ã��ï¼‰ã€‚
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
    start_date_req <= current_date ã�‹ã�¤æ®‹é‡�ã�‚ã‚Šã�®ã‚¿ã‚¹ã‚¯ã�®é…�å�°è©¦è¡Œé †ã�®æœ€å°�å€¤ã€‚
    _equipment_line_lower_dispatch_trial_still_pending ã�¨å�Œæ§˜ã€�ã�¾ã� é–‹å§‹æ—¥ã�«é�”ã�—ã�¦ã�„ã�ªã�„è¡Œã�¯
    ã€Œå…ˆè¡Œè©¦è¡Œé †ã�®ç«¶å�ˆã€�ã�«å�«ã‚�ã�ªã�„ã€‚

    **ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è©¦è¡Œé †ãƒ–ãƒ­ãƒƒã‚¯**ï¼ˆSTAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICTï¼‰ç”¨ã�«ã€�
    ã€Œã�“ã�®æ—¥ã�¾ã� å‰²ä»˜å€™è£œã�«ã�ªã‚Šå¾—ã�ªã�„ã€�è¡Œã�¯æœ€å°�å€¤ã�‹ã‚‰é™¤å¤–ã�™ã‚‹ã€‚ã�•ã‚‚ã�ªã�„ã�¨å�Œä¸€ä¾�é ¼ã�®
    Â§A-1/Â§A-2 å‰�å·¥ç¨‹ï¼ˆè©¦è¡Œé †ã�¯å¾Œã‚�ã� ã�Œè¡Œé †ã�¯å…ˆï¼‰ã�Œå¿…è¦�ã�ªè¡Œã�Œã€�ã‚ˆã‚Šå°�ã�•ã�„è©¦è¡Œé †ã�®è¡Œã�¨
    å¾ªç’°ã�—ã�¦æ°¸ä¹…ã�«å‹•ã�‘ã�ªã�„ã€‚
    - `_task_not_yet_schedulable_due_to_dependency_or_b2_room` ã�Œ True ã�®è¡Œ
    - ï¼ˆdaily_statusãƒ»members ã�Œæ¸¡ã‚‹ã�¨ã��ï¼‰å½“æ—¥æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã� ã�‘ã�§è¨ˆç”»çª“å…¨æ—¥å� æœ‰ã�®è¡Œ
    - ï¼ˆmachine_avail_dt ç­‰ã�Œæ¸¡ã‚‹ã�¨ã��ï¼‰è¨­å‚™å£�æ™‚è¨ˆã�Œè¨ˆç”»çµ‚ç«¯ä»¥ä¸Šã�§å½“æ—¥ã‚¹ãƒ­ãƒƒãƒˆã�ªã�—ã�®è¡Œ

    1 ãƒ­ãƒ¼ãƒ«å‰²å½“ãƒ—ãƒ­ãƒ¼ãƒ–ã�«ã‚ˆã‚‹é™¤å¤–ã�¯è¡Œã‚�ã�ªã�„ï¼ˆ`_effective_min_dispatch_trial_order_from_pool` å�´ã�§å±¤ã�”ã�¨ã�«åˆ¤å®šï¼‰ã€‚
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
    ã‚ˆã‚Šå°�ã�•ã�„é…�å�°è©¦è¡Œé †ã�«ã€�å½“æ—¥å‰²ä»˜å�¯èƒ½ã�ªæœªå®Œäº†ã�Œã�‚ã‚‹ã�¨ã��ã€�å½“è©²ã‚¿ã‚¹ã‚¯ã‚’ãƒ–ãƒ­ãƒƒã‚¯ã�™ã‚‹ã€‚
    min_dispatch_effective: ãƒ—ãƒ¼ãƒ«ï¼‹ãƒ—ãƒ­ãƒ¼ãƒ–ã�§æ±‚ã‚�ã�Ÿå®ŸåŠ¹æœ€å°�è©¦è¡Œé †ï¼ˆæœªæŒ‡å®šæ™‚ã�¯å®‰ä¾¡ãƒ•ã‚£ãƒ«ã‚¿ã�®ã�¿ã�®æœ€å°�ï¼‰ã€‚
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
    """å‹¤æ€ è¾žæ›¸ã�‹ã‚‰ãƒžã‚¹ã‚¿ã�«ç„¡ã�„æ—¥ä»˜ã‚­ãƒ¼ã‚’å‰Šé™¤ã�™ã‚‹ï¼ˆè‡ªå‹•æ‹¡å¼µåˆ†ã�®å·»ã��æˆ»ã�—ï¼‰ã€‚"""
    for dk in list(attendance_data.keys()):
        if dk not in keep_dates:
            del attendance_data[dk]


def _partial_task_id_due_shift_outcome(
    task_queue: list, task_id: str, calendar_last: date
) -> tuple[bool, bool]:
    """
    é…�å�°æ®‹ã�®ä¾�é ¼NOã�«ã�¤ã�„ã�¦ç´�æœŸ+1æ—¥ãƒªãƒˆãƒ©ã‚¤ã�®åˆ†é¡žã€‚
    æˆ»ã‚Šå€¤: (shift_ok, calendar_shortfall)
    - shift_ok: ç´�æœŸåŸºæº–ï¼ˆdue_basis_dateï¼‰ã‚’æŒ�ã�¤è¡Œã�Œã�‚ã‚Šã€�ã��ã‚Œã‚‰ã�™ã�¹ã�¦ã�§ +1 æ—¥ã�Œãƒžã‚¹ã‚¿æœ€çµ‚è¨ˆç”»æ—¥ä»¥ä¸‹
    - calendar_shortfall: ç´�æœŸåŸºæº–ã‚’æŒ�ã�¤è¡Œã�Œã�‚ã‚Šã€�ã�„ã�šã‚Œã�‹ã�§ +1 æ—¥ã�Œãƒžã‚¹ã‚¿æœ€çµ‚è¨ˆç”»æ—¥ã‚’è¶…ã�ˆã‚‹
    åŸºæº–ç´�æœŸã�Œä¸€è¡Œã‚‚ç„¡ã�„ä¾�é ¼ã�¯ (False, False)ï¼ˆé€šå¸¸ã�®é…�å�°æ®‹ã�®ã�¾ã�¾ï¼‰ã€‚
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
    é…�å�°æ®‹ãƒªãƒˆãƒ©ã‚¤ç”¨: **å†…éƒ¨ã�®ç´�æœŸåŸºæº–ï¼ˆdue_basis_dateï¼‰ã� ã�‘**ã‚’ +1 æ—¥ã�™ã‚‹ã€‚
    çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ç”¨ã�® ``due_basis_date_result_sheet`` ã�¯å¤‰æ›´ã�—ã�ªã�„ï¼ˆ+1 å‰�ã�®æ—¥ä»˜ã‚’ä¿�æŒ�ï¼‰ã€‚
    å›žç­”ç´�æœŸãƒ»æŒ‡å®šç´�æœŸã‚‚é…�å�°è¨ˆç”»ã‚·ãƒ¼ãƒˆç”±æ�¥ã�®ã�¾ã�¾ã€‚
    due_urgent ã�¯ã�šã‚‰ã�—ã�Ÿ due_basis_date ã�§å†�è¨ˆç®—ã�™ã‚‹ã€‚
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
    """å�Œä¸€æ—¥å†…ã�®æ—¢å­˜ timeline ã�‹ã‚‰è¨­å‚™ç©ºã��ãƒ»ãƒ¡ãƒ³ãƒ�ãƒ¼ç©ºã��ã�®ä¸‹é™�ã‚’å��æ˜ ã�™ã‚‹ï¼ˆéƒ¨åˆ†å†�é…�å�°ç”¨ï¼‰ã€‚"""
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
    """å�Šé–‹åŒºé–“ãƒ–ãƒ­ãƒƒã‚¯ [start,end) ã�« t ã�Œå…¥ã‚‹é–“ã€�çµ‚ç«¯ã�¸ç¹°ã‚Šä¸Šã�’ã‚‹ã€‚"""
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
    # Excel ã�§ 0 ã‚’ã€Œç©ºã€�ã�¨ã�—ã�¦ã�„ã‚‹åˆ—ã‚„ã€�æ•°å¼�ã�®çµ�æžœ 0 ã�¯å� æœ‰ã�—ã�ªã�„ï¼ˆå¾“æ�¥ True ã� ã�¨å…¨æ—¥å� æœ‰æ‰±ã�„ã�«ã�ªã‚Šå¾—ã‚‹ï¼‰
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
    æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼1ã‚¹ãƒ­ãƒƒãƒˆ [slot_start, slot_end) ã‚’å·¥å ´ç¨¼åƒ�æž ã�«ã‚¯ãƒªãƒƒãƒ—ã�™ã‚‹ã€‚
    æž å¤–ã�®ã�¿ã�®ã‚¹ãƒ­ãƒƒãƒˆã�¯ Noneï¼ˆé…�å�°ã�§ã�¯ç„¡è¦–ï¼‰ã€‚æ®µéšŽ2ã�§ã�¯ master ãƒ¡ã‚¤ãƒ³ A12/B12 ã�§
    DEFAULT_START_TIME / DEFAULT_END_TIME ã�Œä¸Šæ›¸ã��æ¸ˆã�¿ï¼ˆgenerate_plan ã�®ã‚³ãƒ³ãƒ†ã‚­ã‚¹ãƒˆå†…ã�§èª­è¾¼ï¼‰ã€‚
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
    æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼å� æœ‰ã�®å�³ç«¯ã‚’åˆ‡ã‚‹ä¸Šé™�ã€‚å·¥å ´ãƒžã‚¹ã‚¿çµ‚æ¥­ï¼ˆDEFAULT_END_TIMEï¼‰ã�¨ã€�
    å½“æ—¥é…�å�°å¯¾è±¡ãƒ¡ãƒ³ãƒ�ãƒ¼ã�®å‹¤å‹™çµ‚äº†æ™‚åˆ»ã�®æœ€å°�ã�®å°�ã�•ã�„æ–¹ï¼ˆäººã�Œã�„ã�ªã�„æ™‚é–“å¸¯ã�®ã€Œå� æœ‰ã€�ã�§
    è¨­å‚™åºŠã� ã�‘ã�Œçµ‚æ¥­ã‚’è¶…ã�ˆã�ªã�„ã‚ˆã�†ã�«ã�™ã‚‹ï¼‰ã€‚
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
    """å� æœ‰å�Šé–‹åŒºé–“ã‚’ [w0, w1) ã�«ã‚¯ãƒªãƒƒãƒ—ã�—ã�¦ã�‹ã‚‰ãƒžãƒ¼ã‚¸ã�™ã‚‹ã€‚"""
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
    master.xlsmã€Œæ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã€�ã‚’èª­ã�¿ã€�è¨­å‚™åˆ—ã�®é�žç©ºã‚»ãƒ«ï¼�å½“è©² 1 æ™‚é–“ã‚¹ãƒ­ãƒƒãƒˆå� æœ‰ã�¨ã�¿ã�ªã�™ã€‚
    æˆ»ã‚Š: æ—¥ä»˜ -> equipment_list ã�®ã‚­ãƒ¼ -> å�Šé–‹åŒºé–“ [start, end) ã�®ãƒªã‚¹ãƒˆï¼ˆãƒžãƒ¼ã‚¸æ¸ˆã�¿ï¼‰ã€‚
    """
    if not master_path or not os.path.isfile(master_path):
        return {}
    try:
        xls = pd.ExcelFile(master_path)
        if SHEET_MACHINE_CALENDAR not in xls.sheet_names:
            return {}
        raw = pd.read_excel(master_path, sheet_name=SHEET_MACHINE_CALENDAR, header=None)
    except Exception as e:
        logging.warning("æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼: ã‚·ãƒ¼ãƒˆèª­è¾¼ã‚’ã‚¹ã‚­ãƒƒãƒ—ã�—ã�¾ã�—ã�Ÿ (%s)", e)
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
    """å½“æ—¥ã�®ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ã‚·ãƒ¼ãƒ‰å¾Œã€�æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼å� æœ‰ã�§è¨­å‚™ç©ºã��ä¸‹é™�ã‚’ç¹°ã‚Šä¸Šã�’ã‚‹ã€‚"""
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
    """day_blocks ã�‹ã‚‰å� æœ‰ã‚­ãƒ¼ï¼ˆè¡¨è¨˜ã‚†ã‚‰ã�Žè¨±å®¹ï¼‰ã�«ä¸€è‡´ã�™ã‚‹åŒºé–“ãƒªã‚¹ãƒˆã‚’å¾—ã‚‹ã€‚"""
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
    å½“æ—¥ã�®æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼å� æœ‰ã�Œè¨ˆç”»çª“ [å§‹æ¥­, min(çµ‚æ¥­,ç¨¼åƒ�ãƒ¡ãƒ³ãƒ�ãƒ¼çµ‚äº†) ) å…¨ä½“ã‚’å¡žã�Žã€�
    ã��ã�®è¨­å‚™ã�§ã�¯å½“æ—¥ 1 æœ¬ã‚‚åŠ å·¥ã‚’å…¥ã‚Œã‚‰ã‚Œã�ªã�„ã�¨ã�� Trueã€‚
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
    å½“è©²ã‚¿ã‚¹ã‚¯ã�®å� æœ‰è¨­å‚™ã�Œã€�å½“æ—¥ã�®æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã� ã�‘ã�§è¨ˆç”»çª“ã‚’å…¨æ—¥å¡žã�Œã‚Œã�¦ã�„ã‚‹ã€‚
    ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è©¦è¡Œé †ãƒ–ãƒ­ãƒƒã‚¯ç”¨ã�®ã€Œæœ€å°�è©¦è¡Œé †ã€�ã�‹ã‚‰å¤–ã�™ï¼ˆä»–è¨­å‚™ã�®é…�å�°ãƒ‡ãƒƒãƒ‰ãƒ­ãƒƒã‚¯é˜²æ­¢ï¼‰ã€‚
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
    machine_avail_dtï¼ˆã‚·ãƒ¼ãƒ‰ãƒ»æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼åºŠãƒ»å½“æ—¥ç¢ºå®šãƒ­ãƒ¼ãƒ«å��æ˜ å¾Œï¼‰ã�§ã€�
    å� æœ‰è¨­å‚™ã�®ç©ºã��ä¸‹é™�ã�Œè¨ˆç”»çª“çµ‚ç«¯ä»¥ä¸Šã�ªã‚‰å½“æ—¥ã�¯å½“è¨­å‚™ã�«ã‚¹ãƒ­ãƒƒãƒˆã�ªã�—ã€‚
    `machine_handoff` ç­‰ã�Œæ¸¡ã‚‹ã�¨ã��ã�¯ `_resolve_machine_changeover_floor_segments` ã�«ã‚ˆã‚Š
    `_assign_one_roll_trial_order_flow` ã�¨å�Œã�˜ **å®ŸåŠ¹åŠ å·¥é–‹å§‹ä¸‹é™�** ã�§åˆ¤å®šã�™ã‚‹
    ï¼ˆç”Ÿã�® machine_avail ã� ã�‘ã�§ã�¯ãƒ�ãƒ£ãƒ³ã‚¸ã‚ªãƒ¼ãƒ�ãƒ¼å¾Œã�®ä¸‹é™�ã�Œæ¬ ã�‘ã€�å€™è£œã‚„ min_dto ã�Œç‹‚ã�†ã�®ã‚’é˜²ã��ï¼‰ã€‚
    ã�¾ã�Ÿç©ºã��ä¸‹é™�ã�Œçµ‚ç«¯ã‚ˆã‚Šå‰�ã�§ã‚‚ã€�è¨ˆç”»çª“ã�§ã�® **æ®‹ã‚Šé€£ç¶šã�Œ 1 ãƒ­ãƒ¼ãƒ«åˆ†ã�«è¶³ã‚Šã�ªã�„**
    ã�¨åˆ¤æ–­ã�§ã��ã‚‹å ´å�ˆã�¯ Trueï¼ˆå®Ÿåƒ�ä¸�è¶³ãƒ‡ãƒƒãƒ‰ãƒ­ãƒƒã‚¯é˜²æ­¢ï¼‰ã€‚
    ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼åŒºé–“ç…§å�ˆã�®ã‚­ãƒ¼å�–ã‚Šã�“ã�¼ã�—ã‚’é˜²ã��ã€‚
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
    # eff_time_per_unit â‰ˆ base / avg_eff / t_eff Ã— ä½™åŠ›ä¿‚æ•°ã€‚avg_eff ã�¯ãƒ�ãƒ¼ãƒ æ¬¡ç¬¬ã�§ä¸‹ã�Œã‚‹ã€‚
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
    """ãƒ­ãƒ¼ãƒ«ç¢ºå®šç›´å¾Œ: çµ‚äº†æ™‚åˆ»ã�Œã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼å� æœ‰ã‚¹ãƒ­ãƒƒãƒˆå†…ã�ªã‚‰çµ‚ç«¯ã�¾ã�§ç¹°ã‚Šä¸Šã�’ã€‚"""
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
    master.xlsm ã�®ä»»æ„�ã‚·ãƒ¼ãƒˆ:
      - ã€Œè¨­å®š_ä¾�é ¼åˆ‡æ›¿å‰�å¾Œæ™‚é–“ã€�â€¦ å·¥ç¨‹å��ãƒ»æ©Ÿæ¢°å��ãƒ»æº–å‚™åˆ†ãƒ»å¾Œå§‹æœ«åˆ†ï¼ˆ1 è¡Œç›®è¦‹å‡ºã�—ã€�2 è¡Œç›®ä»¥é™�ãƒ‡ãƒ¼ã‚¿ï¼‰
      - ã€Œè¨­å®š_æ©Ÿæ¢°_æ—¥æ¬¡å§‹æ¥­æº–å‚™ã€�â€¦ æ©Ÿæ¢°å��ãƒ»æ—¥æ¬¡å§‹æ¥­æº–å‚™åˆ†

    ä¾�é ¼NOï¼ˆã‚¿ã‚¹ã‚¯ï¼‰ã�Œå�Œä¸€ç‰©ç�†æ©Ÿæ¢°ä¸Šã�§åˆ‡ã‚Šæ›¿ã‚�ã‚‹ã�¨ã��ã€�ç›´å‰�ãƒ–ãƒ­ãƒƒã‚¯ã�®å¾Œå§‹æœ«â†’å½“è©²ãƒ–ãƒ­ãƒƒã‚¯ã�®æº–å‚™ã‚’
    è¨­å‚™ç©ºã��ä¸‹é™�ã�«åŠ ç®—ã�™ã‚‹ã€‚å�Œä¸€ä¾�é ¼NOã�®é€£ç¶šãƒ­ãƒ¼ãƒ«ã�®é–“ã�«ã�¯åŠ ç®—ã�—ã�ªã�„ã€‚
    æ—¥æ¬¡å§‹æ¥­æº–å‚™ã�¯ã€�å�Œä¸€ã‚«ãƒ¬ãƒ³ãƒ€æ—¥ã�§å½“è©²æ©Ÿæ¢°ã�®å…ˆé ­ãƒ­ãƒ¼ãƒ«ã�«ã�®ã�¿åŠ ç®—ã�™ã‚‹ã€‚

    æˆ»ã‚Š: (è¨­å‚™è¡Œã‚­ãƒ¼ã€Œå·¥ç¨‹+æ©Ÿæ¢°ã€�ã�Šã‚ˆã�³æ­£è¦�åŒ–ã‚­ãƒ¼ -> (æº–å‚™åˆ†, å¾Œå§‹æœ«åˆ†),
          æ©Ÿæ¢°å��ã�Šã‚ˆã�³æ­£è¦�åŒ–ã‚­ãƒ¼ -> å§‹æ¥­æº–å‚™åˆ†)
    """
    changeover: dict[str, tuple[int, int]] = {}
    startup: dict[str, int] = {}
    if not master_path or not os.path.isfile(master_path):
        return changeover, startup
    try:
        xls = pd.ExcelFile(master_path)
    except Exception as e:
        logging.warning("æ©Ÿæ¢°æº–å‚™/åˆ‡æ›¿è¨­å®š: ãƒ–ãƒƒã‚¯ã‚’é–‹ã�‘ã�¾ã�›ã‚“ (%s)", e)
        return changeover, startup

    if SHEET_MACHINE_CHANGEOVER in xls.sheet_names:
        try:
            df = pd.read_excel(
                master_path, sheet_name=SHEET_MACHINE_CHANGEOVER, header=0
            )
            df.columns = [str(c).strip() for c in df.columns]
            c_proc = _df_pick_column(df, "å·¥ç¨‹å��", "å·¥ç¨‹")
            c_mac = _df_pick_column(df, "æ©Ÿæ¢°å��", "æ©Ÿæ¢°")
            c_prep = _df_pick_column(
                df,
                "æº–å‚™æ™‚é–“_åˆ†",
                "æº–å‚™åˆ†",
                "åŠ å·¥å‰�æº–å‚™_åˆ†",
                "åŠ å·¥é–‹å§‹å‰�æº–å‚™_åˆ†",
            )
            c_clean = _df_pick_column(
                df,
                "å¾Œå§‹æœ«æ™‚é–“_åˆ†",
                "å¾Œå§‹æœ«åˆ†",
                "åŠ å·¥å¾Œå¾Œå§‹æœ«_åˆ†",
                "åŠ å·¥çµ‚äº†å¾Œå¾Œå§‹æœ«_åˆ†",
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
                        "ãƒžã‚¹ã‚¿ã€Œ%sã€�: å·¥ç¨‹+æ©Ÿæ¢° %s è¡Œã�®æº–å‚™/å¾Œå§‹æœ«ï¼ˆåˆ†ï¼‰ã‚’èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚",
                        SHEET_MACHINE_CHANGEOVER,
                        n_ent,
                    )
        except Exception as e:
            logging.warning(
                "ãƒžã‚¹ã‚¿ã€Œ%sã€�èª­è¾¼å¤±æ•—ï¼ˆç„¡è¦–ï¼‰: %s", SHEET_MACHINE_CHANGEOVER, e
            )

    if SHEET_MACHINE_DAILY_STARTUP in xls.sheet_names:
        try:
            df2 = pd.read_excel(
                master_path, sheet_name=SHEET_MACHINE_DAILY_STARTUP, header=0
            )
            df2.columns = [str(c).strip() for c in df2.columns]
            c_mn = _df_pick_column(df2, "æ©Ÿæ¢°å��", "æ©Ÿæ¢°")
            c_su = _df_pick_column(
                df2, "æ—¥æ¬¡å§‹æ¥­æº–å‚™_åˆ†", "å§‹æ¥­æº–å‚™_åˆ†", "æ—¥å§‹æ¥­æº–å‚™_åˆ†"
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
                        "ãƒžã‚¹ã‚¿ã€Œ%sã€�: æ©Ÿæ¢° %s ä»¶ã�®æ—¥æ¬¡å§‹æ¥­æº–å‚™ï¼ˆåˆ†ï¼‰ã‚’èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚",
                        SHEET_MACHINE_DAILY_STARTUP,
                        len({k for k in startup if "+" not in str(k)}),
                    )
        except Exception as e:
            logging.warning(
                "ãƒžã‚¹ã‚¿ã€Œ%sã€�èª­è¾¼å¤±æ•—ï¼ˆç„¡è¦–ï¼‰: %s", SHEET_MACHINE_DAILY_STARTUP, e
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
    å½“æ—¥ eligible ã�®ã�†ã�¡ã€�å½“è©²å·¥ç¨‹+æ©Ÿæ¢°ã�§ OP ã‚¹ã‚­ãƒ«ã‚’æŒ�ã�¤è€…ã�®ã�†ã�¡å„ªå…ˆåº¦ã�Œæœ€å°�ã�®1å��ã€‚
    æº–å‚™ãƒ»æ—¥æ¬¡å§‹æ¥­ã�®ä¼‘æ†©ã‚¹ã‚­ãƒƒãƒ—ã�«ç”¨ã�„ã‚‹ï¼ˆavail_dt ã�¯è¦‹ã�ªã�„ï¼‰ã€‚
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
    """ã‚¹ã‚­ãƒ« OP ã�Œæ‹¾ã�ˆã�ªã�„ã�¨ã��ã�®ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ï¼ˆå£�æ™‚è¨ˆã�«åˆ†ã‚’è¶³ã�™ï¼�å®šå¸¸é–‹å§‹åŸºæº–ã�®æ—¥æ¬¡å§‹æ¥­ã�¯çµ‚äº†æ™‚åˆ»ã�§ maxï¼‰ã€‚"""
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
    å‰�ãƒ­ãƒ¼ãƒ«åŠ å·¥çµ‚äº† prev_machining_end_dt ã�‹ã‚‰ã€�æ—¥æ¬¡å§‹æ¥­ï¼ˆå½“æ—¥å…ˆé ­ã�®ã�¿ï¼‰ãƒ»å�Œæ—¥ä¾�é ¼åˆ‡æ›¿ã�®å¾Œå§‹æœ«ãƒ»æº–å‚™ã‚’
    çµ„ã�¿ç«‹ã�¦ã€�(åŠ å·¥é–‹å§‹æœ€æ—©æ™‚åˆ», ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ç”¨ã‚»ã‚°ãƒ¡ãƒ³ãƒˆé››å½¢) ã‚’è¿”ã�™ã€‚
    æ—¥æ¬¡å§‹æ¥­ã�¯ master ãƒ¡ã‚¤ãƒ³ A15ï¼ˆå®šå¸¸é–‹å§‹ï¼‰ã�Œèª­ã‚�ã‚Œã�° [é–‹å§‹, é–‹å§‹+Nåˆ†) ã�®å£�æ™‚è¨ˆï¼ˆå‹¤æ€  forward ã�—ã�ªã�„ï¼‰ã€‚
    A15 ã�Œèª­ã‚�ã�ªã�„ã�¨ã��ã�®ã�¿ã€�å¾“æ�¥ã�©ã�Šã‚Šä»£è¡¨ã‚¹ã‚­ãƒ« OP ã�®å‹¤å‹™ãƒ»ä¼‘æ†©ã�«æ²¿ã�£ã�¦ forward ã�™ã‚‹ã€‚
    å�Œä¸€å� æœ‰ã‚­ãƒ¼ã�§ç›´å‰�åŠ å·¥ã�¨å�Œä¸€ä¾�é ¼NOã�®ã�¨ã��ã�¯åŠ å·¥å‰�æº–å‚™ã‚’ä»˜ã�‘ã�ªã�„ï¼ˆé€£ç¶šãƒ­ãƒ¼ãƒ«ï¼‰ã€‚
    æ—¥æ¬¡å§‹æ¥­ã‚»ã‚°ãƒ¡ãƒ³ãƒˆã�® op ã�¯ç©ºï¼ˆã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ã�§ã�¯äººã‚’è¼‰ã�›ã�šè¨­å‚™ã�®ã�¿ï¼‰ã€‚æº–å‚™ãƒ»å¾Œå§‹æœ«ã�® op ã�¯ forward ç”¨ã�®ä»£è¡¨ï¼�ç›´å‰�ä¸»ã€‚
    ã‚»ã‚°ãƒ¡ãƒ³ãƒˆ dict ã�¯ start_dt, end_dt, op, event_kind, machine, machine_occupancy_key ã‚’æŒ�ã�¤ã€‚
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
    è¨­å‚™ã�®å£�æ™‚è¨ˆã�«ã�Šã�‘ã‚‹ã€Œå½“è©²ãƒ­ãƒ¼ãƒ«ã�®åŠ å·¥é–‹å§‹ã€�ä»¥å‰�ã�®ä¸‹é™�ã€‚
    daily_statusãƒ»skills_dictãƒ»current_date ã�Œæ�ƒã�†ã�¨ã��ã�¯ã€�skills é�©å�ˆ OP ã�®å‹¤å‹™ãƒ»ä¼‘æ†©ã�«æ²¿ã�£ã�¦
    æ—¥æ¬¡å§‹æ¥­ãƒ»å¾Œå§‹æœ«ãƒ»æº–å‚™ã‚’ forward ã�—ã�Ÿæœ€æ—©åŠ å·¥é–‹å§‹ã€‚æ�ƒã‚�ã�ªã�„ã�¨ã��ã�¯åˆ†ã�®å£�æ™‚è¨ˆåŠ ç®—ã�«ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ã€‚
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
    è¨­å‚™ã�®åŠ å·¥é–‹å§‹ä¸‹é™�ã�¨ã€�ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³è¿½è¨˜ç”¨ã‚»ãƒƒãƒˆã‚¢ãƒƒãƒ—åŒºé–“ã€‚
    æˆ»ã‚Šå€¤ (floor_dt, segments, abort)ã€‚abort ã�Œ True ã�®ã�¨ã��ã�¯å½“è©²ãƒ­ãƒ¼ãƒ«å‰²å½“ã‚’å…¨ä½“ã�¨ã�—ã�¦æ£„å�´ã�™ã‚‹ã€‚
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
    """ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ç”¨ã�®ä¸»ï¼�è£œã€‚æ—¥æ¬¡å§‹æ¥­ã�¯äººã�ªã�—ã€‚æº–å‚™ã�¯ç›´å¾Œãƒ­ãƒ¼ãƒ«ã€�å¾Œå§‹æœ«ã�¯ handoff ã�®ç›´å‰�ãƒ­ãƒ¼ãƒ«ã€‚"""
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
    """ã‚»ãƒƒãƒˆã‚¢ãƒƒãƒ—ç³»ã‚»ã‚°ãƒ¡ãƒ³ãƒˆã‚’ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ãƒ»ãƒŸãƒ©ãƒ¼ãƒ»æ‹…å½“è€… avail ã�«å��æ˜ ã€‚"""
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
    å½“è©²æ—¥ã�®çµ‚äº†æ™‚ç‚¹ã�§ã€�ç´�æœŸåŸºæº–æ—¥ï¼ˆå½“æ—¥å�«ã‚€ï¼‰ä»¥å‰�ã�ªã�®ã�«æ®‹é‡�ã�Œæ®‹ã‚‹ä¾�é ¼NOã€‚
    ã€Œç´�æœŸæ—¥å†…ã�«å®Œé�‚ã�§ã��ã�ªã�‹ã�£ã�Ÿã€�= å¾Œã‚�å€’ã�—å†�è©¦è¡Œã�®å€™è£œã€‚
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
    ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ã�‹ã‚‰ã€�å�„ machine_occupancy_key ã�«ã�¤ã�„ã�¦
    è¨ˆç”»æ—¥ current_date ä»¥å‰�ã�® **åŠ å·¥ (machining)** ã‚¤ãƒ™ãƒ³ãƒˆã�®æœ€çµ‚çµ‚äº†ã‚’å¾©å…ƒã�™ã‚‹ã€‚
    ã‚»ãƒƒãƒˆã‚¢ãƒƒãƒ—ç³» event_kind ã�¯ last_tid / å¾Œå§‹æœ«åˆ¤å®šã�«å�«ã‚�ã�ªã�„ã€‚
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
    """åŽŸå��æŠ•å…¥æ—¥ã‚’èµ·ç‚¹ã�«ã€�ã��ã�®æ—¥ã�®åŠ å·¥é–‹å§‹ã�®ä¸‹é™�æ™‚åˆ»ï¼ˆå�Œæ—¥ã�¯ 13:00 ä»¥é™�ã‚’å�«ã‚€ï¼‰ã€‚"""
    floor = datetime.combine(current_date, DEFAULT_START_TIME)
    # Â§B-2 æ¤œæŸ» / Â§B-3 å·»è¿”ã�—ã�¯ EC å®Œäº†ã‚’å¾…ã�£ã�¦é–‹å§‹ã�§ã��ã‚‹ã�Ÿã‚�ã€�
    # åŽŸå��æŠ•å…¥æ—¥ï¼ˆ=å�Œæ—¥13:00ä»¥é™�ï¼‰ã�®åˆ¶ç´„ã‚’ã��ã�®ã�¾ã�¾é�©ç”¨ã�™ã‚‹ã�¨å¾Œç¶šã�Œä¸�å¿…è¦�ã�«å¾Œã‚�ã�¸å€’ã‚Œã‚‹ã€‚
    # ECå®Œäº†æ™‚åˆ»ä¸‹é™�ï¼ˆ_roll_pipeline_b2_inspection_ec_completion_floor_dtï¼‰ã�§æ•´å�ˆã‚’å�–ã‚‹ã€‚
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
        # min_dto ã�‹ã‚‰å…¨æ—¥ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼å� æœ‰ã�¯é™¤å¤–æ¸ˆã�¿ã�§ã‚‚ã€�å�Œæ—¥è©¦è¡Œé †ã�®ã€Œãƒ–ãƒ­ãƒƒã‚¯ã€�ã�¯ my_o>m ã�®ã�¿ã�®ã�Ÿã‚�
        # è©¦è¡Œé †=min ã�®å� æœ‰è¡Œã�Œæ®‹ã‚Šã€�ä»–è©¦è¡Œé †ã�Œæ°¸ä¹…å�œæ­¢ã�—å¾—ã‚‹ã€‚å½“æ—¥ã‚¹ãƒ­ãƒƒãƒˆã‚¼ãƒ­ã�®è¡Œã�¯å€™è£œå¤–ã�«ã�™ã‚‹ã€‚
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
    çµ„ã�¿å�ˆã‚�ã�›è¡¨ãƒ—ãƒªã‚»ãƒƒãƒˆ1è¡Œã�®äººæ•°ç¯„å›² (lo, hi)ã€‚need ã�®åŸºæœ¬äººæ•°ã‚ˆã‚Šã‚·ãƒ¼ãƒˆå�´ã‚’å„ªå…ˆã�™ã‚‹ã€‚
    - å¿…è¦�äººæ•°åˆ—ã�Œæ­£ã�®ã�¨ã��ã�¯ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ—ã�®äººæ•°ã�¨ä¸€è‡´ã�™ã‚‹ã�“ã�¨ã€‚
    - hi ã�¯ need ã�®ä¸Šé™�ã�¨å®Ÿäººæ•°ã�®å¤§ã��ã�„æ–¹ï¼ˆãƒ—ãƒªã‚»ãƒƒãƒˆã�Œ need ã‚ˆã‚Šå°‘äººæ•°ã�§ã‚‚æŽ¡ç”¨å�¯èƒ½ï¼‰ã€‚
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
    """åŠ å·¥è¨ˆç”»ã�®å¿…è¦�äººæ•°åˆ—ã�Œæ­£ã�®æ•´æ•°ã�ªã‚‰ã��ã�®å€¤ã€‚ç„¡åŠ¹ã�ªã‚‰ Noneã€‚"""
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
    """ãƒ¬ã‚¬ã‚·ãƒ¼æ—¥æ¬¡é…�å�°ãƒ«ãƒ¼ãƒ—ç”¨: å�˜ä¸€ãƒ�ãƒ¼ãƒ ã�Œæˆ�ç«‹ã�™ã‚Œã�° team_candidates ã�« 1 ä»¶è¿½åŠ ã�—ã�¦ Trueã€‚"""
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
    1ãƒ­ãƒ¼ãƒ«åˆ†ã�®æœ€è‰¯ãƒ�ãƒ¼ãƒ ã‚’æ±ºå®šã�™ã‚‹ã€‚è¨­å‚™ç©ºã��ãƒ»æ—¥é–‹å§‹ä¸‹é™�ã‚’ team_start ã�«ç¹”ã‚Šè¾¼ã‚€ã€‚
    preferred_team ã�Œä¸Žã�ˆã‚‰ã‚Œã€�ã�‹ã�¤ã€Œå�Œä¸€æ—¥å†…ã�®ç›´å‰�ãƒ­ãƒ¼ãƒ«ã€�ã�¨ã�—ã�¦æˆ�ç«‹ã�™ã‚Œã�°ã€�
    çµ„å�ˆã�›æŽ¢ç´¢ã‚ˆã‚Šå„ªå…ˆã�—ã�¦æŽ¡ç”¨ã�™ã‚‹ï¼ˆç¿Œæ—¥ã�«ã�¯æŒ�ã�¡è¶Šã�•ã�ªã�„ï¼‰ã€‚
    æˆ»ã‚Šå€¤: team(tuple), start_dt, end_dt, breaks, eff, op, eff_time_per_unit, extra_max, rq_base, need_src_line, extra_src_line, machine, machine_name, eq_line, req_num, max_team_size
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
            need_src_line = (need_src_line + "ï¼›") if need_src_line else ""
            need_src_line += f"è¨ˆç”»ã‚·ãƒ¼ãƒˆå¿…è¦�äººæ•°{plan_ro}ã�¯æœªä½¿ç”¨ï¼ˆneedåŸºæº–={req_num}ï¼‰"
    else:
        if plan_ro is not None:
            req_num = plan_ro
            need_src_line = f"è¨ˆç”»ã‚·ãƒ¼ãƒˆã€Œå¿…è¦�OP(ä¸Šæ›¸)ã€�={req_num}"
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
            (need_src_line + " â†’ ") if need_src_line else ""
        ) + "ãƒ¡ã‚¤ãƒ³ä¸Šæ›¸ignore_need_minimumã�§req=1"

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
            "ãƒ¡ã‚¤ãƒ³ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«(æ—¥ä»˜Ã—å·¥ç¨‹): task=%s date=%s å·¥ç¨‹=%r ãƒ�ãƒ¼ãƒ å¿…é ˆ=%s",
            task.get("task_id"),
            current_date,
            machine,
            ",".join(_gdp_must),
        )
    if fixed_team_anchor:
        _nfix = len(fixed_team_anchor)
        if _nfix > req_num:
            need_src_line = (need_src_line + " â†’ ") if need_src_line else ""
            need_src_line += f"ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«(æ—¥ä»˜Ã—å·¥ç¨‹)æŒ‡å��ã�§æœ€ä½Ž{_nfix}äºº"
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
            (extra_src_line + " â†’ ") if extra_src_line else ""
        ) + "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROWã�§0"
    extra_max = (
        extra_max_sheet if TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS else 0
    )
    if (
        extra_max_sheet > 0
        and not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
        and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
    ):
        extra_src_line = (
            (extra_src_line + " â†’ ") if extra_src_line else ""
        ) + "ãƒ¡ã‚¤ãƒ³ã�¯åŸºæœ¬äººæ•°ã�®ã�¿ï¼ˆä½™åŠ›æž ã�¯å…¨é…�å�°å¾Œã�«æœªå‰²å½“Ã—ã‚¹ã‚­ãƒ«ã�§è¿½è¨˜ï¼‰"
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
            "needäººæ•°(è©¦è¡Œé †å„ªå…ˆãƒ•ãƒ­ãƒ¼) order=%s task=%s å·¥ç¨‹/æ©Ÿæ¢°=%s/%s "
            "req_num=%s [%s] extra_max=%s [%s] max_teamå€™è£œ=%s capable=%säºº",
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
            "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] " + msg,
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
                "å€™è£œå�´ä¸‹: ãƒ�ãƒ¼ãƒ äººæ•°å¤– team=%s size=%s req=%s max=%s",
                ",".join(str(x) for x in team),
                len(team),
                lo,
                hi,
            )
            return None
        op_list = [m for m in team if skill_role_priority(m)[0] == "OP"]
        if not op_list:
            _trace_assign(
                "å€™è£œå�´ä¸‹: OPä¸�åœ¨ team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        if not all(m in daily_status for m in team):
            _trace_assign(
                "å€™è£œå�´ä¸‹: å½“æ—¥å‹¤æ€ ã‚­ãƒ¼ã�ªã�— team=%s",
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
                "å€™è£œå�´ä¸‹: é–‹å§‹>=çµ‚æ¥­ team=%s start=%s end_limit=%s",
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
                "å€™è£œå�´ä¸‹: ä¼‘æ†©å¸¯å†…ãƒ»çµ‚æ¥­ç›´å‰�(å°�æ®‹)ã�§å½“æ—¥ä¸�å�¯ team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        team_start = team_start_d
        if team_start >= team_end_limit:
            _trace_assign(
                "å€™è£œå�´ä¸‹: ãƒ‡ãƒ•ã‚¡ãƒ¼å¾Œã�«é–‹å§‹>=çµ‚æ¥­ team=%s start=%s end_limit=%s",
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
                "å€™è£œå�´ä¸‹: å®Ÿåƒ�ä¸�è¶³ team=%s start=%s avail_mins=%s need_mins=%.2f",
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
                "å€™è£œå�´ä¸‹: çµ‚æ¥­ç›´å‰�ã�§å½“æ—¥å�Žå®¹ãƒ­ãƒ¼ãƒ«æ•°ã�Œé–¾å€¤æœªæº€ team=%s cap=%s th=%s start=%s",
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
                "å€™è£œå�´ä¸‹: ä¼‘æ†©ã�¾ã�Ÿã�Žã�®ã�Ÿã‚�é€£ç¶šå®Ÿåƒ�ä¸�è¶³ team=%s contiguous_min=%s need_mins=%s start=%s",
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
                "åŒºé–“ãƒŸãƒ©ãƒ¼å�´ä¸‹: team=%s start=%s end=%s eq=%s",
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

    # ç‰¹åˆ¥æŒ‡å®š: å�Œä¸€æ—¥ãƒ»é€£ç¶šãƒ­ãƒ¼ãƒ«ã�¯å‰�å›žãƒ�ãƒ¼ãƒ ã‚’å„ªå…ˆï¼ˆç¿Œæ—¥ã�¸ã�¯æŒ�ã�¡è¶Šã�•ã�ªã�„ï¼‰ã€‚
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
    # çµ„ã�¿å�ˆã‚�ã�›è¡¨ãƒ—ãƒªã‚»ãƒƒãƒˆã�¯ã€Œæˆ�ç«‹ã�—ã�Ÿã‚‰å�³ returnã€�ã�›ã�šã€�çµ„å�ˆã�›æŽ¢ç´¢ã�¨ã�¾ã�¨ã‚�ã�¦
    # team_start / ã‚¹ãƒ©ãƒƒã‚¯ä»˜ã��ã‚¿ãƒ—ãƒ«ã�§æœ€è‰¯ã‚’é�¸ã�¶ï¼ˆã‚·ãƒ¼ãƒˆä¸Šã�®å„ªå…ˆåº¦é †ã�¯è©¦è¡Œé †ã�®ã�¿ï¼‰ã€‚
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
                    "å€™è£œå›ºå®š: æ‹…å½“OPæŒ‡å��=%s ã�®ã�Ÿã‚� 1äººãƒ�ãƒ¼ãƒ ã�¯å½“äººã�®ã�¿è©¦è¡Œ",
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
                "æ®µéšŽ2: ä¾�é ¼NO=%s æ—¥ä»˜=%s å·¥ç¨‹/æ©Ÿæ¢°=%s/%s ã�§ãƒ�ãƒ¼ãƒ å€™è£œã�Œ0ä»¶ã€‚"
                "ã‚¹ã‚­ãƒ«é�©å�ˆ(OP/AS)ã�¯ %s äººã�„ã�¾ã�™ã�Œã€�è¨­å‚™ã�®åŠ å·¥é–‹å§‹ä¸‹é™�=%s ã�Œ"
                "å½“æ—¥ã�®æ‹…å½“å€™è£œã�®é€€å‹¤(%s)ä»¥é™�ã�®ã�Ÿã‚�ã�“ã�®æ—¥ã�¯å‰²å½“ã�§ã��ã�¾ã�›ã‚“ã€‚"
                "masterã€Œæ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã€�ã�§å½“è©²æ—¥ãƒ»å½“è©²æ©Ÿæ¢°åˆ—ã�«ä¸�è¦�ã�ªè¨˜å…¥ã�Œã�ªã�„ã�‹ã€�"
                "ã�¾ã�Ÿã�¯å‰�å·¥ç¨‹ã�®å� æœ‰ã�§è¨­å‚™ä¸‹é™�ã�Œçµ‚æ¥­ã�¾ã�§ç¹°ã‚Šä¸Šã�Œã�£ã�¦ã�„ã�ªã�„ã�‹ç¢ºèª�ã�—ã�¦ã��ã� ã�•ã�„"
                "ï¼ˆé…�å�°ãƒ«ãƒ¼ãƒ« 3.2.1 æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ãƒ»ãƒˆãƒ©ãƒ–ãƒ«ã‚·ãƒ¥ãƒ¼ãƒˆï¼‰ã€‚"
                "å�‚è€ƒ: changeoverå‰�ã�®è¨­å‚™ç©ºã��ä¸‹é™�=%s å� æœ‰ã‚­ãƒ¼=%s",
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
    ç�¾åœ¨ã�® avail_dt / machine_avail_dt / machine_handoff ã�®ã‚¹ãƒŠãƒƒãƒ—ã‚·ãƒ§ãƒƒãƒˆã�§
    `_assign_one_roll_trial_order_flow` ã�Œ None ã�«ã�ªã‚‹ã�ªã‚‰ Trueã€‚
    æ©Ÿæ¢°æž ã�¯å��åˆ†ã�§ã‚‚äººãƒ»ä¼‘æ†©ãƒ»ãƒŸãƒ©ãƒ¼ç­‰ã�§è©°ã�¾ã‚Šã€�ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«è©¦è¡Œé †ã� ã�‘ã�Œå…ˆé ­è¡Œã�«å¼µã‚Šä»˜ã��ã�®ã‚’é˜²ã��ã€‚
    å‰¯ä½œç”¨ã�ªã�—ï¼ˆneed äººæ•°ãƒ­ã‚°ç”¨ set ã�¯æ¯Žå›žç©ºï¼‰ã€‚
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
            "trial_order_assign_probe ä¾‹å¤–ã�®ã�Ÿã‚�å½“è©²è¡Œã�¯é™¤å¤–ã�—ã�ªã�„: task=%s err=%s",
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
    """`_min_pending_dispatch_trial_order_for_date` ã�¨å�Œä¸€ã�®å®‰ä¾¡ãƒ•ã‚£ãƒ«ã‚¿ã‚’é€šé�Žã�—ã�Ÿã‚¿ã‚¹ã‚¯ã�®ãƒªã‚¹ãƒˆã€‚"""
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
    pool ã‚’æ˜‡é † dto ã�§è¦‹ã�¦ã€�**ã��ã�® dto ã�«å±žã�™ã‚‹è¡Œã�®ã�†ã�¡ 1 ä»¶ã�§ã‚‚** 1 ãƒ­ãƒ¼ãƒ«å‰²å½“ãƒ—ãƒ­ãƒ¼ãƒ–ã�Œé€šã‚Œã�°
    ã��ã�® dto ã‚’ã€Œå®ŸåŠ¹ã�®æœ€å°�è©¦è¡Œé †ã€�ã�¨ã�™ã‚‹ã€‚
    å…ˆé ­ dto å±¤ã�Œå…¨æ»…ï¼ˆæ©Ÿæ¢°ã�¯ç©ºã�„ã�¦ã�„ã‚‹ã�Œäººã�§ç©�ã‚�ã�ªã�„ç­‰ï¼‰ã�®ã�¨ã��ã€�æ¬¡ã�® dto ã�«é€²ã�¿ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«å�œæ­¢ã‚’é˜²ã��ã€‚
    ãƒ—ãƒ­ãƒ¼ãƒ–ç„¡ã�—ã�®ã�¨ã��ã�¯ pool ã�®æœ€å°� dto ã‚’è¿”ã�™ã€‚
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
    â‘ å½“æ—¥å€™è£œã‚’é…�å�°è©¦è¡Œé †ã�®æ˜‡é †ã�«ä¸¦ã�¹ã‚‹ï¼ˆ1 ãƒ‘ã‚¹åˆ†ï¼‰ã€‚
    **å®Œå…¨äºŒç›¸ï¼ˆÂ§B-2 / Â§B-3ï¼‰**: **ãƒ•ã‚§ãƒ¼ã‚º1**ã�§ **å¾Œç¶šãƒ‘ã‚¤ãƒ—ãƒ©ã‚¤ãƒ³è¡Œ**ï¼ˆç†±èž�ç�€æ¤œæŸ»ãƒ»å·»è¿”ã�—ï¼‰**ã‚’é™¤ã��**å€™è£œï¼ˆECãƒ»ä»–ä¾�é ¼ãƒ»ä»–å·¥ç¨‹ï¼‰ã‚’è©¦è¡Œé †ã�©ã�Šã‚Š
    **`_drain_rolls_for_task`** ã�—ã€�**ãƒ•ã‚§ãƒ¼ã‚º2**ã�¯ Â§B-2 æ¤œæŸ»ï¼�Â§B-3 å·»è¿”ã�—è¡Œã�®ã�¿ï¼ˆ**å�Œä¸€ä¾�é ¼ã�® EC ã�Œå…¨æ—¥ã�§å®Œèµ°ã�—ã�Ÿå¾Œ**ã�«é™�ã‚Šå€™è£œåŒ–ã€‚
    EC æ®‹ã�Œã�‚ã‚‹æ—¥ã�¯ `_trial_order_flow_eligible_tasks` ã�§å¾Œç¶šã‚’å¤–ã�—ã€�ç¿Œç¨¼åƒ�æ—¥ä»¥é™�ã‚‚ EC ã�®ã�¿å‰�é€²ã�™ã‚‹ã€‚
    ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼é€šç®—ã�§ EC å®Œèµ°å¾Œã€�`_run_b2_inspection_rewind_pass` ã�Œæ—¥ä»˜å…ˆé ­ã�‹ã‚‰å¾Œç¶šã� ã�‘å†�èµ°æŸ»ã�™ã‚‹ï¼‰ã€‚
    EC ã�¨å¾Œç¶šã‚’ **å�Œä¸€æ‹…å½“è€…ã�§** äº¤äº’ã�«è©°ã‚�ã‚‹ã�¨ EC ã�Œãƒ–ãƒ­ãƒƒã‚¯ã�•ã‚Œã‚‹ã�Ÿã‚�ã€�å¾“æ�¥ã�¯ãƒ•ã‚§ãƒ¼ã‚º1ã‚’å…ˆã�«è©°ã‚�ã�Ÿã€‚
    ã�Ÿã� ã�—å¾Œç¶šã�Œå€™è£œåŒ–ã�—ã�Ÿæ™‚ç‚¹ã�§ **æ¤œæŸ»ã�¨å�Œã�˜ç‰©ç�†æ©Ÿæ¢°**ã�®ãƒ•ã‚§ãƒ¼ã‚º1ã‚„ **å�Œä¸€ä¾�é ¼ã�® EC** ã�Œå…¨æ—¥å…ˆã�«é€²ã‚€ã�¨ã€�
    æ¤œæŸ»ã�¯ `start_ge_end_initial`ï¼ˆè¨­å‚™ç©ºã��ã�Œçµ‚æ¥­ã‚ˆã‚Šå¾Œï¼‰ã�§å…¨æ—¥å¤±æ•—ã�™ã‚‹ã€‚Â§B-2/Â§B-3 å¾Œç¶šã�Œã�‚ã‚‹ã�¨ã��ã�¯
    ã€Œå�Œä¸€ä¾�é ¼ECãƒ»æ¤œæŸ»æ©Ÿã�¨æ©Ÿæ¢°å…±æœ‰ã�™ã‚‹ãƒ•ã‚§ãƒ¼ã‚º1ãƒ»å¾Œç¶šã€�ã‚’ **é…�å�°è©¦è¡Œé †**ã�§ãƒžãƒ¼ã‚¸ã�—ã€�
    å�Œé †ã�§ã�¯ **å¾Œç¶šã‚’ EC ã‚ˆã‚Šå…ˆã�«**ã€�**ã��ã�®ä»–ã�®ãƒ•ã‚§ãƒ¼ã‚º1** ã�¨ã�‚ã‚�ã�›ã�¦ **é…�å�°è©¦è¡Œé †**ã�§æ•´åˆ—ã�—
    **æœ€å¤§1ãƒ­ãƒ¼ãƒ«ã�šã�¤**ã� ã�‘å‘¨å›žã�™ã‚‹ï¼ˆãƒžãƒ¼ã‚¸ãƒ»rest ã�¨ã‚‚ä¸€æ‹¬ãƒ‰ãƒ¬ã‚¤ãƒ³ã�—ã�ªã�„ã€‚æ¤œæŸ»OPã�Œä»–å·¥ç¨‹ã�«
    å�Œæ—¥å�–ã‚Šåˆ‡ã‚‰ã‚Œ start_ge_end_initial ã�«ã�ªã‚‹ã�®ã‚’é˜²ã��ï¼‰ã€‚
    ãƒªãƒ¯ã‚¤ãƒ³ãƒ‰å�´ã�®å¾Œç¶šè¡Œã�¯å�„ãƒ­ãƒ¼ãƒ«ã�«ã�¤ã�„ã�¦ `_roll_pipeline_inspection_assign_room` ã�Šã‚ˆã�³
    `_roll_pipeline_b2_inspection_ec_completion_floor_dt`ï¼ˆEC ãƒ­ãƒ¼ãƒ«çµ‚äº†æ™‚åˆ»ä¸‹é™�ï¼‰ã�§æ•´å�ˆã�™ã‚‹ã€‚
    è©¦è¡Œé †æœ€å°�ã�®è¡Œã� ã�‘ã�Œå½“æ—¥å…¥ã‚‰ã�ªã�„å ´å�ˆã�§ã‚‚ã€�**å�Œã�˜ãƒ•ã‚§ãƒ¼ã‚ºå†…ã�§æ¬¡ã�®è©¦è¡Œé †ã�¸é€²ã�¿**ä»–è¨­å‚™ã‚’åŸ‹ã‚�ã‚‹ã€‚
    æ©Ÿæ¢°ãƒ»äººã�®ç©ºã��ã�¯ãƒ­ãƒ¼ãƒ«ã�”ã�¨ã�«æ›´æ–°ã�™ã‚‹ï¼ˆâ‘¦â‘§ï¼‰ã€‚
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
                    "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ãƒ­ãƒ¼ãƒ«ç¢ºå®š ãƒ¡ã‚¤ãƒ³ day=%s machine=%s machine_name=%s "
                    "start=%s end=%s æŽ¡ç”¨äººæ•°=%s req_num=%s ãƒ¡ã‚¤ãƒ³æŽ¢ç´¢extra_max=%s "
                    "ä½™å‰°äººæ•°é�©ç”¨(ãƒ¡ã‚¤ãƒ³)=%s team=%s",
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
        # å�Œã�˜é…�å�°è©¦è¡Œé †ã�§ã�¯å¾Œç¶šï¼ˆæ¤œæŸ»ãƒ»å·»è¿”ã�—ï¼‰ã‚’ EC ã‚ˆã‚Šå…ˆã�«å›žã�—ã€�ç†±èž�ç�€ã�®å£�æ™‚è¨ˆã‚’
        # å�Œæ—¥æ—©ã�„æ®µéšŽã�§å�–ã‚Šã�«è¡Œã��ï¼ˆÂ§B-2 æ‹…å½“è€…åˆ†é›¢ã�§ EC ã�¨æ¤œæŸ»ã�¯åˆ¥ãƒ¡ãƒ³ãƒ�ãƒ¼æƒ³å®šï¼‰ã€‚
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
    Â§B-2 / Â§B-3: EC å�´ã‚’å…ˆã�«å…¨æ—¥ã�§é€²ã‚�ã�Ÿå¾Œã€�æ¤œæŸ»ï¼�å·»è¿”ã�—å�´ã�®ã�¿ã‚’æ—¥ä»˜å…ˆé ­ã�‹ã‚‰å†�èµ°æŸ»ã�—ã�¦é…�å�°ã�™ã‚‹ã€‚
    timeline_events ã‚’äººãƒ»è¨­å‚™ã�®ãƒ–ãƒ­ãƒƒã‚¯ãƒ†ãƒ¼ãƒ–ãƒ«ã�¨ã�—ã�¦ä½¿ã�„ã€�æ—¥è·¨ã�Žã�®å� æœ‰ã‚’ä¿�æŒ�ã�™ã‚‹ã€‚
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
    needã€Œé…�å�°æ™‚è¿½åŠ äººæ•°ï¼�ä½™åŠ›æ™‚è¿½åŠ äººæ•°ã€�è¡Œã�®ä¸Šé™�ã�¾ã�§ã€�ãƒ¡ã‚¤ãƒ³å‰²ä»˜ã�§æŽ¡ç”¨ã�—ã��ã‚Œã�ªã�‹ã�£ã�Ÿæž ã‚’è¿½è¨˜ã�™ã‚‹ã€‚
    å�„ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ãƒ–ãƒ­ãƒƒã‚¯ã�«ã�¤ã�„ã�¦ã€�ã��ã�®æ™‚é–“å¸¯ã�«ä»–ãƒ–ãƒ­ãƒƒã‚¯ã�¸æœªå�‚åŠ ï¼ˆåŒºé–“é‡�ã�ªã‚Šã�ªã�—ï¼‰ã�§
    eligible ã�‹ã�¤ OP/AS ã‚¹ã‚­ãƒ«ã�®è€…ã‚’ã‚µãƒ–ã�«è¿½åŠ ã�™ã‚‹ã€‚
    æ—¥æ¬¡å§‹æ¥­ãƒ»ä¾�é ¼åˆ‡æ›¿å¾Œå§‹æœ«ãƒ»åŠ å·¥å‰�æº–å‚™ï¼ˆevent_kind ã�ŒåŠ å·¥ä»¥å¤–ï¼‰ã�¯æœ¬å‡¦ç�†ã�®å¯¾è±¡å¤–ï¼ˆä½™å‰°ã‚µãƒ–ã�¯åŠ å·¥ã�«ã�®ã�¿è¿½è¨˜ï¼‰ã€‚
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
                "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ä½™åŠ›è¿½è¨˜(ãƒ¡ã‚¤ãƒ³å®Œäº†å¾Œ) day=%s machine=%s machine_name=%s "
                "start=%s end=%s è¿½è¨˜äººæ•°=%s è¿½è¨˜å‰�äººæ•°=%s è¿½è¨˜å¾Œäººæ•°=%s req_num=%s "
                "needè¿½åŠ æž (ã‚·ãƒ¼ãƒˆ)=%s å±¥æ­´é»„(ä½™å‰°äººæ•°è¶…é�Ž)=%s è¿½è¨˜ãƒ¡ãƒ³ãƒ�ãƒ¼=%s",
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
# 3. ãƒ¡ã‚¤ãƒ³è¨ˆç”»ç”Ÿæˆ� (æ—¥æ¯Žãƒ«ãƒ¼ãƒ—ãƒ»æŒ�ã�¡è¶Šã�—å¯¾å¿œ)
#    æ®µéšŽ2ã�®æœ¬ä½“ã€‚plan_simulation_stage2 ã�‹ã‚‰ã�®ã�¿å‘¼ã�°ã‚Œã‚‹æƒ³å®šã€‚
#    é…�å�°è¨ˆç”»ã‚·ãƒ¼ãƒˆèª­è¾¼ â†’ ã‚¿ã‚¹ã‚¯ã‚­ãƒ¥ãƒ¼ â†’ æ—¥ä»˜ã�”ã�¨ã�«è¨­å‚™ãƒ»OPå‰²ä»˜ â†’ çµ�æžœãƒ–ãƒƒã‚¯å‡ºåŠ›ã€‚
# =========================================================
def generate_plan():
    """
    æ®µéšŽ2ã�®ãƒ¡ã‚¤ãƒ³å‡¦ç�†ã€‚æˆ»ã‚Šå€¤ã�ªã�—ï¼ˆãƒ­ã‚°ãƒ»Excel å‡ºåŠ›ã�§å®Œçµ�ï¼‰ã€‚

    å‰�æ��: ç’°å¢ƒå¤‰æ•° TASK_INPUT_WORKBOOKã€�ã‚«ãƒ¬ãƒ³ãƒˆãƒ‡ã‚£ãƒ¬ã‚¯ãƒˆãƒªã�Œã‚¹ã‚¯ãƒªãƒ—ãƒˆãƒ•ã‚©ãƒ«ãƒ€ã€‚
    å‡ºåŠ›: ``output_dir`` ç›´ä¸‹ã�® ``production_plan_multi_day_*.xlsx`` / ``member_schedule_*.xlsx``ï¼ˆæœ€æ–°1çµ„ã�®ã�¿ï¼‰ã€�ã�Šã‚ˆã�³ log/execution_log.txtã€‚
    """
    master_abs = os.path.abspath(os.path.join(os.getcwd(), MASTER_FILE))
    with _override_default_factory_hours_from_master(master_abs):
        _generate_plan_impl()


def _generate_plan_impl():
    # é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ï¼ˆè¨­å®šã‚·ãƒ¼ãƒˆ A3 ä»¥é™�ã�®ã�¿ï¼‰ã�¯ã€�ãƒ¡ãƒ³ãƒ�ãƒ¼0äººç­‰ã�§æ—©æœŸ return ã�—ã�¦ã‚‚
    # execution_log ã�«æ®‹ã‚‹ã‚ˆã�† skills èª­è¾¼ã‚ˆã‚Šå‰�ã�§ç¢ºå®šãƒ»ãƒ­ã‚°ã�™ã‚‹ã€‚
    global TRACE_SCHEDULE_TASK_IDS, DEBUG_DISPATCH_ONLY_TASK_IDS
    _wb_trace = (os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK)
    _ids_from_sheet = _read_trace_schedule_task_ids_from_config_sheet(_wb_trace)
    TRACE_SCHEDULE_TASK_IDS = frozenset(
        str(x).strip() for x in _ids_from_sheet if str(x).strip()
    )
    if _ids_from_sheet:
        _preview = _ids_from_sheet[:25]
        _suffix = " â€¦" if len(_ids_from_sheet) > 25 else ""
        logging.info(
            "è¨­å®šã‚·ãƒ¼ãƒˆã€Œ%sã€�A3 ä»¥é™�: ãƒˆãƒ¬ãƒ¼ã‚¹ç”¨ä¾�é ¼NOã‚’ %s ä»¶èª­ã�¿è¾¼ã�¿ï¼ˆ%s%sï¼‰",
            APP_CONFIG_SHEET_NAME,
            len(_ids_from_sheet),
            ", ".join(_preview),
            _suffix,
        )
    else:
        logging.info(
            "è¨­å®šã‚·ãƒ¼ãƒˆã€Œ%sã€�A3 ä»¥é™�: ãƒˆãƒ¬ãƒ¼ã‚¹ç”¨ä¾�é ¼NOã�¯ç„¡ã�—ï¼ˆç©ºã�¾ã�Ÿã�¯ã‚·ãƒ¼ãƒˆç„¡ã�—ï¼‰",
            APP_CONFIG_SHEET_NAME,
        )
    if TRACE_SCHEDULE_TASK_IDS:
        logging.info(
            "é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹: æœ‰åŠ¹ task_id = %sï¼ˆè¨­å®šã‚·ãƒ¼ãƒˆ A3 ä»¥é™�ï¼‰",
            ", ".join(sorted(TRACE_SCHEDULE_TASK_IDS)),
        )
    else:
        logging.info(
            "é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹: å¯¾è±¡ã�ªã�—ï¼ˆ[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ â€¦] ãƒ­ã‚°ã�¯å‡ºã�¾ã�›ã‚“ï¼‰"
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
            "ãƒ‡ãƒ�ãƒƒã‚°é…�å�°: ã€Œ%sã€�B3ä»¥é™�ã�«ã‚ˆã‚Šé…�å�°å¯¾è±¡ã‚’ %s ä»¶ã�®ä¾�é ¼NOã�«é™�å®šã�—ã�¾ã�™: %s",
            APP_CONFIG_SHEET_NAME,
            len(DEBUG_DISPATCH_ONLY_TASK_IDS),
            ", ".join(sorted(DEBUG_DISPATCH_ONLY_TASK_IDS)),
        )
        _show_stage2_debug_dispatch_mode_dialog(sorted(DEBUG_DISPATCH_ONLY_TASK_IDS))
    if TRACE_TEAM_ASSIGN_TASK_ID:
        logging.info(
            "ç’°å¢ƒå¤‰æ•° TRACE_TEAM_ASSIGN_TASK_ID=%r â†’ ãƒ�ãƒ¼ãƒ å‰²å½“ãƒˆãƒ¬ãƒ¼ã‚¹æœ‰åŠ¹",
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
            "çµ„ã�¿å�ˆã‚�ã�›è¡¨: å·¥ç¨‹+æ©Ÿæ¢°ã‚­ãƒ¼ %s ç¨®é¡žãƒ»ç·¨æˆ�è¡Œ %s ã‚’é…�å�°ãƒ—ãƒªã‚»ãƒƒãƒˆã�¨ã�—ã�¦èª­ã�¿è¾¼ã�¿ã�¾ã�—ã�Ÿã€‚",
            len(team_combo_presets),
            _nrules,
        )
    elif TEAM_ASSIGN_USE_MASTER_COMBO_SHEET:
        logging.info(
            "çµ„ã�¿å�ˆã‚�ã�›è¡¨: ãƒ—ãƒªã‚»ãƒƒãƒˆç„¡ã�—ï¼ˆã‚·ãƒ¼ãƒˆæ¬ å¦‚ãƒ»ç©ºãƒ»ã�¾ã�Ÿã�¯èª­è¾¼å¤±æ•—ï¼‰ã€‚å¾“æ�¥ã�®ãƒ�ãƒ¼ãƒ æŽ¢ç´¢ã�®ã�¿ã€‚"
        )
    if not members:
        master_abs = os.path.abspath(MASTER_FILE)
        logging.error(
            "æ®µéšŽ2ã‚’ä¸­æ–­ã�—ã�¾ã�—ã�Ÿ: ãƒ¡ãƒ³ãƒ�ãƒ¼ã�Œ0äººã�§ã�™ï¼ˆãƒžã‚¹ã‚¿ã�® skills ã�Œç©ºã€�ã�¾ã�Ÿã�¯èª­ã�¿è¾¼ã�¿å¤±æ•—ï¼‰ã€‚"
            " æœŸå¾…ãƒ‘ã‚¹: %s ï¼ˆã‚«ãƒ¬ãƒ³ãƒˆ: %sï¼‰ã€‚ãƒ†ã‚¹ãƒˆã‚³ãƒ¼ãƒ‰ç›´ä¸‹ã�« master.xlsm ã‚’ç½®ã��ã€�"
            "planning_core ã�®ã‚«ãƒ¬ãƒ³ãƒˆã�Œã��ã�®ãƒ•ã‚©ãƒ«ãƒ€ã�«ã�ªã‚‹ã‚ˆã�† python\\ é…�ç½®ã‚’ç¢ºèª�ã�—ã�¦ã��ã� ã�•ã�„ã€‚"
            " ã�“ã�®çŠ¶æ…‹ã�§ã�¯ production_plan / member_schedule ã�¯å‡ºåŠ›ã�•ã‚Œã�¾ã�›ã‚“ã€‚",
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
            "æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼: èª­è¾¼ä¾‹å¤–ã�®ã�Ÿã‚�å� æœ‰ã�ªã�—ã�¨ã�—ã�¦ç¶šè¡Œã�—ã�¾ã�™ (%s)", e
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
            "æ©Ÿæ¢°æº–å‚™/ä¾�é ¼åˆ‡æ›¿ãƒ»æ—¥æ¬¡å§‹æ¥­è¨­å®š: èª­è¾¼ä¾‹å¤–ã�®ã�Ÿã‚�ç„¡è¦–ã�—ã�¾ã�™ (%s)", e
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
                "æ—¥æ¬¡å§‹æ¥­æº–å‚™: å®šå¸¸é–‹å§‹ master ãƒ¡ã‚¤ãƒ³ A15=%s ã‚’æŽ¡ç”¨ï¼ˆ[é–‹å§‹, é–‹å§‹+åˆ†) ã‚’å£�æ™‚è¨ˆã�§å� æœ‰ã€‚A15 ç„¡åŠ¹æ™‚ã�¯å¾“æ�¥ã�®å‹¤æ€  forwardï¼‰",
                _rs_a15.strftime("%H:%M"),
            )
    except Exception as e:
        logging.warning("å®šå¸¸é–‹å§‹(A15) èª­è¾¼å¤±æ•—: æ—¥æ¬¡å§‹æ¥­ã�¯å¾“æ�¥ã�®å‹¤æ€  forward ã�«ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯ (%s)", e)
        _STAGE2_REGULAR_SHIFT_START = None
    if _MACHINE_CALENDAR_BLOCKS_BY_DATE:
        _n_iv = sum(
            len(ivs)
            for _dm in _MACHINE_CALENDAR_BLOCKS_BY_DATE.values()
            for ivs in _dm.values()
        )
        logging.info(
            "æ©Ÿæ¢°ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼: %s æ—¥åˆ†ãƒ»è¨­å‚™å� æœ‰ãƒ–ãƒ­ãƒƒã‚¯è¨ˆ %s ã‚’é…�å�°ã�«å��æ˜ ã�—ã�¾ã�™ã€‚",
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
            "needé…�å�°æ™‚è¿½åŠ äººæ•°: ãƒ¡ã‚¤ãƒ³å‰²ä»˜ã�¯åŸºæœ¬å¿…è¦�äººæ•°ã�®ã�¿ã€‚"
            "ä½™åŠ›ã�¯å…¨ã‚·ãƒŸãƒ¥ãƒ¬ãƒ¼ã‚·ãƒ§ãƒ³å¾Œã€�æ™‚é–“é‡�ã�ªã‚Šã�®ã�ªã�„æœªå‰²å½“ã�‹ã�¤ã‚¹ã‚­ãƒ«é�©å�ˆè€…ã‚’ã‚µãƒ–ã�«è¿½è¨˜ã�—ã�¾ã�™ã€‚"
            "ï¼ˆãƒ¡ã‚¤ãƒ³ã�§å¢—å“¡æŽ¢ç´¢ã�™ã‚‹å¾“æ�¥æŒ™å‹•: TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS=1ï¼‰"
        )

    # æ®µéšŽ2ã�®åŸºæº–æ—¥æ™‚ã�¯ã€Œãƒžã‚¯ãƒ­å®Ÿè¡Œæ™‚åˆ»ã€�ã�§ã�¯ã�ªã��ã€Œãƒ‡ãƒ¼ã‚¿æŠ½å‡ºæ—¥ã€�ã‚’ä½¿ç”¨
    data_extract_dt = _extract_data_extraction_datetime()
    base_now_dt = data_extract_dt if data_extract_dt is not None else datetime.now()
    run_date = base_now_dt.date()
    data_extract_dt_str = (
        base_now_dt.strftime("%Y/%m/%d %H:%M:%S") if data_extract_dt is not None else "â€”"
    )
    logging.info(
        "è¨ˆç”»åŸºæº–æ—¥æ™‚: %sï¼ˆ%sï¼‰",
        base_now_dt.strftime("%Y/%m/%d %H:%M:%S"),
        "ãƒ‡ãƒ¼ã‚¿æŠ½å‡ºæ—¥" if data_extract_dt is not None else "ç�¾åœ¨æ™‚åˆ»ãƒ•ã‚©ãƒ¼ãƒ«ãƒ�ãƒƒã‚¯",
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
            "ãƒ¡ã‚¤ãƒ³ãƒ»ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆ: å·¥å ´ä¼‘æ¥­æ‰±ã�„ã�®æ—¥ä»˜ â†’ %s",
            ", ".join(str(x) for x in sorted(_factory_closure_dates)),
        )
    ai_log_data["ãƒ¡ã‚¤ãƒ³_ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«_å·¥å ´ä¼‘æ¥­æ—¥(è§£æž�)"] = (
        ", ".join(str(x) for x in sorted(_factory_closure_dates))
        if _factory_closure_dates
        else "ï¼ˆã�ªã�—ï¼‰"
    )
    _sn = str(global_priority_override.get("scheduler_notes_ja") or "").strip()
    if _sn:
        ai_log_data["ãƒ¡ã‚¤ãƒ³_ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«_æœªé�©ç”¨ãƒ¡ãƒ¢(AI)"] = _sn[:2000]

    sorted_dates = sorted(list(attendance_data.keys()))
    # çµ�æžœã‚·ãƒ¼ãƒˆã�¯ã€ŒåŸºæº–æ—¥ï¼ˆãƒ‡ãƒ¼ã‚¿æŠ½å‡ºæ—¥ï¼‰ã€�ä»¥é™�ã�®ã�¿è¡¨ç¤ºãƒ»è¨ˆç”»å¯¾è±¡ã�¨ã�™ã‚‹
    sorted_dates = [d for d in sorted_dates if d >= run_date]
    if not sorted_dates:
        logging.error("å½“æ—¥ä»¥é™�ã�®å‡¦ç�†å¯¾è±¡æ—¥ä»˜ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚")
        _try_write_main_sheet_gemini_usage_summary("æ®µéšŽ2")
        return

    # ã‚¿ã‚¹ã‚¯å…¥åŠ›: ãƒ–ãƒƒã‚¯å†…ã€Œé…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã€�ï¼ˆæ®µéšŽ1ã�§å‡ºåŠ›â†’å�–ã‚Šè¾¼ã�¿å¾Œã�«ç·¨é›†ï¼‰
    try:
        tasks_df = load_planning_tasks_df()
    except Exception as e:
        logging.error(f"é…�å�°è¨ˆç”»ã‚¿ã‚¹ã‚¯ã‚·ãƒ¼ãƒˆèª­ã�¿è¾¼ã�¿ã‚¨ãƒ©ãƒ¼: {e}")
        _try_write_main_sheet_gemini_usage_summary("æ®µéšŽ2")
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
            "ãƒ‡ãƒ�ãƒƒã‚°é…�å�°: ã€Œ%sã€�ã�®è¡Œã‚’ %s â†’ %s ã�«çµžã‚Šè¾¼ã�¿ã�¾ã�—ã�Ÿã€‚",
            PLAN_INPUT_SHEET_NAME,
            _n_tasks_before,
            _n_tasks_after,
        )
        if _n_tasks_after == 0:
            logging.error(
                "ãƒ‡ãƒ�ãƒƒã‚°é…�å�°: B3ä»¥é™�ã�®ä¾�é ¼NOã�«ä¸€è‡´ã�™ã‚‹è¡Œã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚æ®µéšŽ2ã‚’ä¸­æ–­ã�—ã�¾ã�™ã€‚"
            )
            _try_write_main_sheet_gemini_usage_summary("æ®µéšŽ2")
            return

    if global_priority_raw.strip():
        snip = global_priority_raw[:2500]
        if len(global_priority_raw) > 2500:
            snip += "â€¦"
        ai_log_data["ãƒ¡ã‚¤ãƒ³_å†�å„ªå…ˆç‰¹åˆ¥è¨˜è¼‰(åŽŸæ–‡)"] = snip
    else:
        ai_log_data["ãƒ¡ã‚¤ãƒ³_å†�å„ªå…ˆç‰¹åˆ¥è¨˜è¼‰(åŽŸæ–‡)"] = (
            "ï¼ˆç©ºã€�ã�¾ã�Ÿã�¯ãƒ¡ã‚¤ãƒ³ã‚·ãƒ¼ãƒˆã�«ã€Œã‚°ãƒ­ãƒ¼ãƒ�ãƒ«ã‚³ãƒ¡ãƒ³ãƒˆã€�è¦‹å‡ºã�—ã�Œè¦‹ã�¤ã�‹ã‚Šã�¾ã�›ã‚“ï¼‰"
        )
    ai_log_data["ãƒ¡ã‚¤ãƒ³_å†�å„ªå…ˆç‰¹åˆ¥è¨˜è¼‰(AI)"] = json.dumps(
        global_priority_override, ensure_ascii=False
    )
    if global_priority_override.get("ignore_skill_requirements"):
        logging.warning(
            "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: ã‚¹ã‚­ãƒ«è¦�ä»¶ã‚’ç„¡è¦–ã�—ã�¦é…�å�°ã�—ã�¾ã�™ã€‚%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("ignore_need_minimum"):
        logging.warning(
            "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: ãƒ�ãƒ¼ãƒ äººæ•°ã‚’1å��ã�«å›ºå®šã�—ã�¾ã�™ï¼ˆneedãƒ»è¡Œã�®å¿…è¦�OPä¸Šæ›¸ã��ã‚ˆã‚Šå„ªå…ˆï¼‰ã€‚%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("abolish_all_scheduling_limits"):
        logging.warning(
            "ãƒ¡ã‚¤ãƒ³å†�å„ªå…ˆç‰¹è¨˜: è¨­å‚™å°‚æœ‰ãƒ»åŽŸå��å�Œæ—¥é–‹å§‹ãƒ»æŒ‡å®šé–‹å§‹æ™‚åˆ»ãƒ»ãƒžã‚¯ãƒ­å®Ÿè¡Œæ™‚åˆ»ä¸‹é™�ã‚’é�©ç”¨ã�—ã�¾ã�›ã‚“ã€‚%s",
            global_priority_override.get("interpretation_ja", ""),
        )

    # ã€Œå½“æ—¥ã€�åˆ¤å®šã�¨æœ€æ—©é–‹å§‹æ™‚åˆ»ã�«ã�¯åŸºæº–æ—¥æ™‚ï¼ˆãƒ‡ãƒ¼ã‚¿æŠ½å‡ºæ—¥ï¼‰ã‚’ä½¿ã�†
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
    # é–‹å§‹æ—¥ã�Œé�žç¨¼åƒ�æ—¥ã�®å ´å�ˆã�¯ã€�ç›´å‰�ã�®ç¨¼åƒ�æ—¥ã�¸è£œæ­£ï¼ˆä¾‹: 4/4, 4/5 ã�Œé�žç¨¼åƒ�ã�ªã‚‰ 4/3 ã�¸ï¼‰
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
                        "DEBUG[task=%s] start_date_req ã‚’é�žç¨¼åƒ�æ—¥è£œæ­£: %s -> %s",
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
            f"æœ‰åŠ¹ã�ªã‚¿ã‚¹ã‚¯ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚ã€Œ{PLAN_INPUT_SHEET_NAME}ã€�ã�®ã€Œä¾�é ¼NOã€�ã€Œå·¥ç¨‹å��ã€�ã€Œæ�›ç®—æ•°é‡�ã€�ã€�"
            "ã�¾ã�Ÿã�¯å®Œäº†åŒºåˆ†ãƒ»å®Ÿå‡ºæ�¥é«˜æ�›ç®—ã�«ã‚ˆã‚Šæ®‹é‡�ã�Œç„¡ã�„è¡Œã�®ã�¿ã�®å�¯èƒ½æ€§ã�Œã�‚ã‚Šã�¾ã�™ã€‚"
        )

    # é…�å�°è©¦è¡Œé †: ã‚·ãƒ¼ãƒˆåˆ—ã�Œæ�ƒã�£ã�¦ã�„ã‚Œã�°ã��ã‚Œã‚’æŽ¡ç”¨ã€‚æ¬ æ��æ™‚ã�¯ Â§B å¸¯ãƒ»ç´�æœŸãƒ»need åˆ—é †ã�§ã‚½ãƒ¼ãƒˆã�— EC éš£æŽ¥å¾Œã�« 1..n
    _apply_dispatch_trial_order_for_generate_plan(
        task_queue, req_map, need_rules, need_combo_col_index
    )
    if DEBUG_TASK_ID:
        dbg_items = [t for t in task_queue if str(t.get("task_id", "")).strip() == DEBUG_TASK_ID]
        if dbg_items:
            t0 = dbg_items[0]
            logging.info(
                "DEBUG[task=%s] queueåŸºæº–: start_date_req=%s due_basis=%s answer_due=%s specified_due=%s specified_due_ov=%s due_source=%s priority=%s in_progress=%s remark=%s",
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
            logging.info("DEBUG[task=%s] task_queueã�«å­˜åœ¨ã�—ã�¾ã�›ã‚“ï¼ˆå®Œäº†/æ®‹é‡�0/ä¾�é ¼NOä¸�ä¸€è‡´ã�®å�¯èƒ½æ€§ï¼‰ã€‚", DEBUG_TASK_ID)
    timeline_events = []

    # ---------------------------------------------------------
    # æ—¥æ¯Žã�®ã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒªãƒ³ã‚°ãƒ«ãƒ¼ãƒ—
    # STAGE2_EXTEND_ATTENDANCE_CALENDAR ã�Œ True ã�®ã�¨ã��ã�®ã�¿ã€�æ®‹ã‚¿ã‚¹ã‚¯ã�Œã�‚ã‚Œã�°å‹¤æ€ ã‚’æ—¥ä»˜è¤‡è£½ã�§æ‹¡å¼µã€‚
    # STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING ã�Œ True ã�®ã�¨ã��ã�®ã�¿: ç´�æœŸåŸºæº–ã‚’é�Žã�Žã�¦ã‚‚æ®‹ã�Œã�‚ã‚‹ä¾�é ¼ã�«ã�¤ã�„ã�¦
    # due_basis +1ãƒ»å½“è©²ä¾�é ¼ã�®å‰²å½“æˆ»ã�—ãƒ»å…ˆé ­ã�‹ã‚‰å†�å®Ÿè¡Œã€‚å�„å†�è©¦è¡Œå‰�ã�«å‹¤æ€ æ‹¡å¼µåˆ†ã�¯ãƒžã‚¹ã‚¿æ—¥ä»˜ã�¸å·»ã��æˆ»ã�™ã€‚
    # æ—¢å®š False ã�®ã�Ÿã‚�é€šå¸¸ã�¯ 1 ãƒ‘ã‚¹ï¼ˆã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼é€šã�— 1 å›žï¼‰ã�®ã�¿ã€‚
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
            "DISPATCH_INTERVAL_MIRROR_ENFORCE: è¨­å‚™ãƒ»äººã�®å� æœ‰ã‚’åŒºé–“ãƒŸãƒ©ãƒ¼ã�§è¿½è·¡ã�—ã�¾ã�™"
            "ï¼ˆç„¡åŠ¹åŒ–ã�¯ è¨­å®š_ç’°å¢ƒå¤‰æ•° ç­‰ã�§ DISPATCH_INTERVAL_MIRROR_ENFORCE=0ï¼‰ã€‚"
        )

    if STAGE2_SERIAL_DISPATCH_BY_TASK_ID:
        logging.info(
            "ä¾�é ¼NOç›´åˆ—é…�å�°: æœ‰åŠ¹ï¼ˆSTAGE2_SERIAL_DISPATCH_BY_TASK_IDï¼‰ã€‚"
            " å�„æ—¥ã�¯ã‚¢ã‚¯ãƒ†ã‚£ãƒ–ã�ªä¾�é ¼NOã�®è¡Œã� ã�‘ã�Œå€™è£œã�®ã�Ÿã‚�ã€�å½“è©²ä¾�é ¼ã�Œè©°ã�¾ã‚‹ã�¨ä»–ä¾�é ¼ã�¯ä¸€åˆ‡é€²ã�¿ã�¾ã�›ã‚“ã€‚"
        )
    else:
        logging.info(
            "ä¾�é ¼NOç›´åˆ—é…�å�°: ç„¡åŠ¹ã€‚start_date ã‚’æº€ã�Ÿã�™å…¨è¡Œã�Œå½“æ—¥å€™è£œã�«ã�ªã‚Šã€�é…�å�°è©¦è¡Œé †ãƒ»è¨­å‚™ãƒ«ãƒ¼ãƒ«ã�§é †åº�ä»˜ã�‘ã�—ã�¾ã�™ã€‚"
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
            # è¨­å‚™ã�”ã�¨ã�®ç©ºã��æ™‚åˆ»ï¼ˆå�Œä¸€è¨­å‚™ã�®å�Œæ™‚ä¸¦è¡Œå‰²å½“ã‚’é˜²æ­¢ï¼‰
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
                logging.info("DEBUG[day=%s] ç¨¼åƒ�ãƒ¡ãƒ³ãƒ�ãƒ¼0ã�®ã�Ÿã‚�å‰²ä»˜ã‚¹ã‚­ãƒƒãƒ—", current_date)
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
                    "ä¾�é ¼NOç›´åˆ—é…�å�° day=%s ã‚¢ã‚¯ãƒ†ã‚£ãƒ–ä¾�é ¼NO=%s ç›´åˆ—ãƒªã‚¹ãƒˆä½�ç½®=%s/%s "
                    "å½“æ—¥å€™è£œè¡Œæ•°(ç›´åˆ—å‰�)=%s ç›´åˆ—å¾Œ=%s ã‚­ãƒ¥ãƒ¼æ®‹è¡Œ(å…¨æ—¥)=%s",
                    current_date,
                    _active_serial_tid if _active_serial_tid is not None else "â€”",
                    _serial_pos if _serial_pos else "â€”",
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
                    "DEBUG[day=%s] å‰²ä»˜å¯¾è±¡ã‚¿ã‚¹ã‚¯0ä»¶ pending_total=%s earliest_start_date_req=%s",
                    current_date,
                    pending_total,
                    earliest_wait,
                )
            elif DEBUG_TASK_ID:
                has_dbg_today = any(str(t.get("task_id", "")).strip() == DEBUG_TASK_ID for t in tasks_today)
                if current_date.isoformat() == "2026-04-03" or has_dbg_today:
                    logging.info(
                        "DEBUG[day=%s] avail_members=%s tasks_today=%s (task=%s å�«ã‚€=%s)",
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
                                    "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ã‚¹ã‚­ãƒƒãƒ—: å�Œä¸€ä¾�é ¼NOã�®å…ˆè¡Œå·¥ç¨‹å¾…ã�¡ day=%s machine=%s rem=%.4f",
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
                                    "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ã‚¹ã‚­ãƒƒãƒ—: Â§B-2/Â§B-3 å¾Œç¶šãƒ­ãƒ¼ãƒ«æž ã‚¼ãƒ­ day=%s machine=%s "
                                    "ecç´¯è¨ˆå®Œäº†R=%.4f å¾Œç¶šç´¯è¨ˆå®Œäº†R=%.4f rem_follower=%.4f",
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
                                        "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ã‚¹ã‚­ãƒƒãƒ—: å�Œä¸€è¨­å‚™ã�®æ¤œæŸ»å� æœ‰ä¸­ day=%s "
                                        "å� æœ‰è€…ä¾�é ¼NO=%s å� æœ‰è€…è©¦è¡Œé †=%s",
                                        task.get("task_id"),
                                        current_date,
                                        _b1_holder.get("task_id"),
                                        _b1_holder.get("dispatch_trial_order"),
                                    )
                                continue
                        if DEBUG_TASK_ID and str(task.get("task_id", "")).strip() == DEBUG_TASK_ID:
                            logging.info(
                                "DEBUG[task=%s] day=%s é–‹å§‹åˆ¤å®š: start_date_req=%s remaining_units=%s machine=%s",
                                DEBUG_TASK_ID,
                                current_date,
                                task.get("start_date_req"),
                                task.get("remaining_units"),
                                task.get("machine"),
                            )
                        if task.get("has_done_deadline_override"):
                            logging.info(
                                "DEBUG[å®Œäº†æ—¥æŒ‡å®š] ä¾�é ¼NO=%s æ—¥ä»˜=%s start_date_req=%s due_basis=%s æŒ‡å®šç´�æœŸ(ä¸Šæ›¸ã��)=%s é€²æ�—=%s/%s",
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
                                    "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ã‚¹ã‚­ãƒƒãƒ—: ã‚ˆã‚Šå°�ã�•ã�„é…�å�°è©¦è¡Œé †ã�«æœªå®Œäº†ã�‚ã‚Š "
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
                                    "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ã‚¹ã‚­ãƒƒãƒ—: å�Œä¸€è¨­å‚™ã�§é…�å�°è©¦è¡Œé †ã�Œå…ˆã�®è¡Œã�Œæœªå®Œäº† "
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
                                    (need_src_line + "ï¼›") if need_src_line else ""
                                )
                                need_src_line += (
                                    f"è¨ˆç”»ã‚·ãƒ¼ãƒˆå¿…è¦�äººæ•°{plan_ro}ã�¯æœªä½¿ç”¨ï¼ˆneedåŸºæº–={req_num}ï¼‰"
                                )
                        else:
                            if plan_ro is not None:
                                req_num = plan_ro
                                need_src_line = f"è¨ˆç”»ã‚·ãƒ¼ãƒˆã€Œå¿…è¦�OP(ä¸Šæ›¸)ã€�={req_num}"
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
                                (need_src_line + " â†’ ")
                                if need_src_line
                                else ""
                            ) + "ãƒ¡ã‚¤ãƒ³ä¸Šæ›¸ignore_need_minimumã�§req=1"
    
                        # ãƒ¡ãƒ³ãƒ�ãƒ¼Ã—è¨­å‚™ã‚¹ã‚­ãƒ«ï¼ˆparse_op_as_skill_cell: å°�ã�•ã�„å„ªå…ˆåº¦ã�»ã�©å…ˆã�«ãƒ�ãƒ¼ãƒ å€™è£œã�¸æŽ¡ç”¨ï¼‰
                        # skills èª­è¾¼æ™‚ã�«ã€Œæ©Ÿæ¢°å��ã€�å�˜ç‹¬ã‚­ãƒ¼ã�¸ã‚¨ã‚¤ãƒªã‚¢ã‚¹ã�™ã‚‹ã�Ÿã‚�ã€�å·¥ç¨‹å��+æ©Ÿæ¢°å��ã�Œä¸¡æ–¹ã�‚ã‚‹è¡Œã�§ã�¯
                        # è¤‡å�ˆã‚­ãƒ¼ã€Œå·¥ç¨‹å��+æ©Ÿæ¢°å��ã€�ã�®ã�¿ã‚’è¦‹ã‚‹ï¼ˆåˆ¥å·¥ç¨‹ã�®å�Œå��æ©Ÿæ¢°ã�® OP ã�Œæµ�ã‚Œè¾¼ã�¾ã�ªã�„ã‚ˆã�†ã�«ã�™ã‚‹ï¼‰ã€‚
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
                                "DEBUG[å®Œäº†æ—¥æŒ‡å®š] ä¾�é ¼NO=%s è¨­å‚™=%s req_num=%s capable_members=%s machine_free=%s",
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
                                "æ‹…å½“OPæŒ‡å��: å½“æ—¥ã�®OPå€™è£œã�«ä¸€è‡´ã�›ã�šåˆ¶ç´„ã�ªã�— task=%s raw=%r",
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
                                "ãƒ¡ã‚¤ãƒ³ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«(æ—¥ä»˜Ã—å·¥ç¨‹): task=%s date=%s å·¥ç¨‹=%r ãƒ�ãƒ¼ãƒ å¿…é ˆ=%s",
                                task.get("task_id"),
                                current_date,
                                machine,
                                ",".join(_gdp_must),
                            )
                        if fixed_team_anchor:
                            _nfix = len(fixed_team_anchor)
                            if _nfix > req_num:
                                need_src_line = (
                                    (need_src_line + " â†’ ")
                                    if need_src_line
                                    else ""
                                )
                                need_src_line += (
                                    f"ã‚°ãƒ­ãƒ¼ãƒ�ãƒ«(æ—¥ä»˜Ã—å·¥ç¨‹)æŒ‡å��ã�§æœ€ä½Ž{_nfix}äºº"
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
                                (extra_src_line + " â†’ ")
                                if extra_src_line
                                else ""
                            ) + "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROWã�§0"
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
                                (extra_src_line + " â†’ ")
                                if extra_src_line
                                else ""
                            ) + "ãƒ¡ã‚¤ãƒ³ã�¯åŸºæœ¬äººæ•°ã�®ã�¿ï¼ˆä½™åŠ›æž ã�¯å…¨é…�å�°å¾Œã�«æœªå‰²å½“Ã—ã‚¹ã‚­ãƒ«ã�§è¿½è¨˜ï¼‰"
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
                                "needäººæ•°(é…�å�°è©¦è¡Œé †åˆ�å›ž) order=%s task=%s å·¥ç¨‹/æ©Ÿæ¢°=%s/%s "
                                "req_num=%s [%s] extra_max=%s [%s] max_teamå€™è£œ=%s capable=%säºº",
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
                                "TRACEé…�å�°[%s] %s å·¥ç¨‹/æ©Ÿæ¢°=%s / %s req_num=%s extra_max=%s â†’ max_team=%s "
                                "capable(n=%s)=%s ignore_need1=%s ignore_skill=%s abolish=%s æ‹…å½“OPæŒ‡å®š=%râ†’%s",
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
                        # ãƒ—ãƒªã‚»ãƒƒãƒˆã�¯æˆ�ç«‹åˆ†ã‚’ã�™ã�¹ã�¦å€™è£œã�«è¼‰ã�›ã€�ä¸‹ã�®çµ„å�ˆã�›æŽ¢ç´¢ã�¨ã�¾ã�¨ã‚�ã�¦æœ€è‰¯ã‚’é�¸ã�¶ã€‚
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
                                        "æ‹…å½“OPæŒ‡å��: ãƒ�ãƒ¼ãƒ äººæ•°ã‚’æº€ã�Ÿã�›ã�ªã�„ã�Ÿã‚�æŒ‡å��ã‚’ç„¡è¦– task=%s size=%s raw=%r",
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
                                    # å�Œä¸€è¨­å‚™ã�¯1æ™‚ç‚¹ã�§1ã‚¿ã‚¹ã‚¯ã�®ã�¿ï¼ˆè¨­å‚™ç©ºã��ï¼‹æ—¥æ¬¡å§‹æ¥­/ä¾�é ¼åˆ‡æ›¿ã�®æº–å‚™ãƒ»å¾Œå§‹æœ«ï¼‰
                                    machine_free_dt = _mach_floor_legacy
                                    if team_start < machine_free_dt:
                                        team_start = machine_free_dt
                                    # åŽŸå��æŠ•å…¥æ—¥ã�¨å�Œæ—¥ã�®é–‹å§‹ã�¯ 13:00 ä»¥é™�ï¼ˆè©¦è¡Œé †å„ªå…ˆãƒ•ãƒ­ãƒ¼ã�¨ä¸€è‡´ï¼‰
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
                                    # å½“æ—¥ã�¯ã€Œãƒžã‚¯ãƒ­å®Ÿè¡Œã�—ã�Ÿæ™‚åˆ»ã€�ã‚ˆã‚Šå‰�ã�«é–‹å§‹ã�§ã��ã�ªã�„
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
                                        "TRACEé…�å�°[%s] %s tsize=%s â†’ ã�“ã�®äººæ•°ã�§æˆ�ç«‹ã�™ã‚‹ãƒ�ãƒ¼ãƒ ã�ªã�—",
                                        tid,
                                        current_date,
                                        tsize,
                                    )
                                else:
                                    sm = min(sub, key=_team_cand_key)
                                    logging.info(
                                        "TRACEé…�å�°[%s] %s tsize=%s äººæ•°å†…æœ€è‰¯: members=%s "
                                        "start=%s units_today=%s prio_sum=%s eff_t/unit=%.6f "
                                        "æ¯”è¼ƒãƒ«ãƒ¼ãƒ«=%s â€»å…¨æ—¥æœ€æ—©é–‹å§‹=%s ã‚’åŸºæº–ã�«è¾žæ›¸å¼�ã�§å°�ã�•ã�„æ–¹ã�ŒæŽ¡ç”¨",
                                        tid,
                                        current_date,
                                        tsize,
                                        sm["team"],
                                        sm["team_start"],
                                        sm["units_today"],
                                        sm["prio_sum"],
                                        sm["eff_time_per_unit"],
                                        _tk,
                                        t_min.isoformat(sep=" ") if t_min else "â€”",
                                    )
    
                        if trace_assign and best_team is not None:
                            logging.info(
                                "TRACEé…�å�°[%s] %s â˜…æŽ¡ç”¨ n=%s members=%s start=%s units_today=%s prio_sum=%s",
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
                                        "TRACEé…�å�°[%s] %s 1äººæŽ¡ç”¨ï¼ˆTEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFFï¼‰: "
                                        "ã‚ˆã‚Šå¤§ã��ã�„äººæ•°ã�§æœ‰åŠ¹ã�ªãƒ�ãƒ¼ãƒ ã�ªã�—ï¼ˆOPä¸�è¶³ãƒ»0å�˜ä½�ãƒ»é–‹å§‹>=çµ‚äº†ç­‰ï¼‰ã€‚",
                                        task["task_id"],
                                        current_date,
                                    )
                                else:
                                    logging.info(
                                        "TRACEé…�å�°[%s] %s 1äººæŽ¡ç”¨: äººæ•°ã‚’å¢—ã‚„ã�™ã�¨é–‹å§‹ã�Œé�…ã‚Œã€�"
                                        "ã‚¹ãƒ©ãƒƒã‚¯å¤–ã�§ã�¯é–‹å§‹å„ªå…ˆã�§1äººã�Œé�¸ã�°ã‚Œã�Ÿå�¯èƒ½æ€§ã€‚"
                                        "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=%sã€�ã�¾ã�Ÿã�¯å¾“æ�¥ã�®äººæ•°æœ€å„ªå…ˆã�¯ç’°å¢ƒå¤‰æ•°å�‚ç…§ã€‚",
                                        task["task_id"],
                                        current_date,
                                        TEAM_ASSIGN_START_SLACK_WAIT_MINUTES,
                                    )
    
                        if best_team:
                            if len(best_team) > req_num:
                                logging.info(
                                    "é…�å�°æŽ¡ç”¨äººæ•°>req_num task=%s day=%s order=%s å·¥ç¨‹/æ©Ÿæ¢°=%s/%s "
                                    "æŽ¡ç”¨=%säºº req_num=%s extra_max=%s max_team=%s [%s] [%s]",
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
                                        "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ã‚¹ã‚­ãƒƒãƒ—: ãƒ�ãƒ¼ãƒ æŽ¡ç”¨å¾Œã�®å®ŸåŠ¹ãƒ¦ãƒ‹ãƒƒãƒˆ0 "
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
                            
                            # ã€Œãƒžã‚¯ãƒ­å®Ÿè¡Œæ™‚ç‚¹ã€�ã�®å®Œäº†çŽ‡ï¼ˆäºˆå®šã�®é€²æ�—ã�§ã�¯ã�ªã��ã€�å®ŸåŠ å·¥æ•°ãƒ™ãƒ¼ã‚¹ï¼‰
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
                                    "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³è¿½è¨˜ chunk day=%s machine=%s "
                                    "done_units=%s already_done=%s total_u=%s rem_after=%.4f "
                                    "start=%s end=%s eff_t/unit=%.4f rp_room(å½“æ™‚)=%s",
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
                                    "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ãƒ­ãƒ¼ãƒ«ç¢ºå®š ãƒ¡ã‚¤ãƒ³ day=%s machine=%s machine_name=%s "
                                    "start=%s end=%s æŽ¡ç”¨äººæ•°=%s req_num=%s ãƒ¡ã‚¤ãƒ³æŽ¢ç´¢extra_max=%s "
                                    "ä½™å‰°äººæ•°é�©ç”¨(ãƒ¡ã‚¤ãƒ³)=%s team=%s",
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
                                    "DEBUG[å®Œäº†æ—¥æŒ‡å®š] ä¾�é ¼NO=%s æ—¥ä»˜=%s ã�¯å‰²å½“ä¸�å�¯ï¼ˆè¦�å“¡/è¨­å‚™ç©ºã��æ�¡ä»¶ã�§ãƒ�ãƒ¼ãƒ ä¸�æˆ�ç«‹ï¼‰ã€‚remaining_units=%s",
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
                            "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] æ—¥æ¬¡çµ‚äº†æ™‚ç‚¹ã�®æ®‹ day=%s machine=%s "
                            "machine_name=%s rem=%.4f roll_b2_follower=%s è©¦è¡Œé †=%s",
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
                                "ç´�æœŸè¶…é�Žãƒªãƒˆãƒ©ã‚¤: è¨ˆç”»åŸºæº–+1æ—¥ã�—ã�¦å½“è©²ä¾�é ¼ã�®ã�¿å†�é…�å�°ï¼ˆæ¤œå‡ºæ—¥=%s ä¾�é ¼NO=%s å½“è©²ä¾�é ¼ã�®ç´¯è¨ˆè©¦è¡Œ=%sï¼‰",
                                current_date.isoformat(),
                                ",".join(sorted(allowed_shift_tids)),
                                _trials_detail,
                            )
                            _full_calendar_without_deadline_restart = False
                            break
                        else:
                            # ä¾�é ¼ã�”ã�¨ä¸Šé™�ã�§ã‚·ãƒ•ãƒˆã�§ã��ã�ªã�„ã� ã�‘ã�®ã�¨ã��ã�¯æ—¥ä»˜ãƒ«ãƒ¼ãƒ—ã‚’ç¶™ç¶šã�™ã‚‹ï¼ˆbreak ã�™ã‚‹ã�¨æœªå‡¦ç�†æ—¥ã�Œæ®‹ã‚Šé…�å�°ä¸�å�¯ã�Œå¤§é‡�ç™ºç”Ÿã�™ã‚‹ï¼‰ã€‚
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
                                    "ç´�æœŸå¾Œã‚�å€’ã�—å†�é…�å�°: æ¬¡ã�®ä¾�é ¼NOã�¯ä¾�é ¼ã�”ã�¨ã�®ä¸Šé™�ï¼ˆå�„ %s å›žï¼‰ã�®ã�Ÿã‚�ã�“ã�®æ¤œå‡ºã�§ã�¯ +1 ã�—ã�¾ã�›ã‚“ã€‚"
                                    " ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ã�¯ç¶™ç¶šã�—ã�¾ã�™ï¼ˆæœªå®Œäº†ã�¯çµ‚äº†æ™‚ã�«ç´�æœŸè¦‹ç›´ã�—å¿…è¦�ã‚’ä»˜ã�‘å¾—ã�¾ã�™ï¼‰: %s",
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
                    "Â§B-2/Â§B-3 ãƒªãƒ¯ã‚¤ãƒ³ãƒ‰: EC å®Œèµ°å¾Œã�«æ¤œæŸ»ï¼�å·»è¿”ã�—ã�®ã�¿æ—¥ä»˜å…ˆé ­ã�‹ã‚‰å†�é…�å�°ã�—ã�¾ã�—ã�Ÿï¼ˆtimeline_events ã‚’å� æœ‰ãƒ†ãƒ¼ãƒ–ãƒ«ã�¨ã�—ã�¦åˆ©ç”¨ï¼‰ã€‚"
                )
            break

    if TRACE_SCHEDULE_TASK_IDS:
        for _tt in TRACE_SCHEDULE_TASK_IDS:
            for _t in task_queue:
                if str(_t.get("task_id", "")).strip() != _tt:
                    continue
                _log_dispatch_trace_schedule(
                    _tt,
                    "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ã‚·ãƒŸãƒ¥ãƒ¬ãƒ¼ã‚·ãƒ§ãƒ³çµ‚äº†æ™‚ machine=%s machine_name=%s "
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
                    "[é…�å�°ãƒˆãƒ¬ãƒ¼ã‚¹ task=%s] ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³æœ€çµ‚å¡Š(å·¥ç¨‹åˆ—ã�”ã�¨) machine=%s "
                    "already_done+units_done=%s+%s=%s total_units=%s end_dt=%s",
                    _tt,
                    _mk,
                    _ad,
                    _ud,
                    _ad + _ud,
                    _ev.get("total_units"),
                    _ev.get("end_dt"),
                )

    # ãƒ¡ã‚¤ãƒ³å‰²ä»˜ã�¾ã�§ã�®ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ï¼ˆneed ä½™åŠ›è¿½è¨˜å‰�ï¼‰ã€‚TEMP_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ç”¨ã€‚
    timeline_before_need_surplus = copy.deepcopy(timeline_events)

    # needã€Œé…�å�°æ™‚è¿½åŠ äººæ•°ã€�: ãƒ¡ã‚¤ãƒ³å‰²ä»˜å¾Œã�«ã€�æœªå�‚åŠ Ã—ã‚¹ã‚­ãƒ«é�©å�ˆè€…ã‚’ã‚µãƒ–ã�¸è¿½è¨˜ï¼ˆæ—¢å®šï¼‰
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
                "needä½™åŠ›: ãƒ¡ã‚¤ãƒ³å‰²ä»˜å®Œäº†å¾Œã�«ã‚µãƒ– %s å��ã‚’è¿½è¨˜ï¼ˆæœªå‰²å½“Ã—ã‚¹ã‚­ãƒ«ãƒ»æ™‚é–“é‡�ã�ªã‚Šã�ªã�—ï¼‰",
                _n_sur,
            )

    if _dispatch_interval_mirror is not None:
        _dispatch_interval_mirror.rebuild_from_timeline(timeline_events)

    # ã‚¿ã‚¤ãƒ ãƒ©ã‚¤ãƒ³ã‚’æ—¥ä»˜åˆ¥ã�«ã‚¤ãƒ³ãƒ‡ãƒƒã‚¯ã‚¹åŒ–ã�—ã€�ã‚µãƒ–ãƒ¡ãƒ³ãƒ�ãƒ¼ä¸€è¦§ã‚’äº‹å‰�è§£æž�ï¼ˆä»¥é™�ã�®å‡ºåŠ›ãƒ«ãƒ¼ãƒ—ã‚’é«˜é€ŸåŒ–ï¼‰
    for e in timeline_events:
        e["subs_list"] = [s.strip() for s in e["sub"].split(",")] if e.get("sub") else []

    events_by_date = defaultdict(list)
    for e in timeline_events:
        events_by_date[e["date"]].append(e)

    # =========================================================
    # 4. Excelå‡ºåŠ› (ãƒ¡ã‚¤ãƒ³è¨ˆç”»)
    # =========================================================
    _remove_prior_stage2_workbooks_and_prune_empty_dirs(output_dir)
    # å�Œä¸€ç§’å†…ã�®å†�å®Ÿè¡Œã�§ãƒ•ã‚¡ã‚¤ãƒ«å��ã�Œè¡�çª�ã�—ã�ªã�„ã‚ˆã�†ãƒžã‚¤ã‚¯ãƒ­ç§’ã�¾ã�§å�«ã‚�ã‚‹
    _stage2_out_stamp = base_now_dt.strftime("%Y%m%d_%H%M%S_%f")
    output_filename = os.path.join(
        output_dir, f"production_plan_multi_day_{_stage2_out_stamp}.xlsx"
    )
    # ã‚¿ã‚¹ã‚¯ID â†’ çµ�æžœ_è¨­å‚™æ¯Žã�®æ™‚é–“å‰²ã�§å½“è©²ã‚¿ã‚¹ã‚¯ã�Œæœ€åˆ�ã�«ç�¾ã‚Œã‚‹ã‚»ãƒ«ï¼ˆä¾‹ B12ï¼‰ã€‚çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®ãƒªãƒ³ã‚¯ç”¨ã€‚
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

    # çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ç”¨: ã‚·ãƒŸãƒ¥ãƒ¬ãƒ¼ã‚·ãƒ§ãƒ³ä¸Šã�®å½“è©²ã‚¿ã‚¹ã‚¯ã�®æœ€æ—©é–‹å§‹ãƒ»æœ€é�…çµ‚äº†ï¼ˆtimeline_events é›†ç´„ï¼‰
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

    # çµ�æžœ_ã‚¿ã‚¹ã‚¯ä¸€è¦§ã�®ã€Œå›žç­”ç´�æœŸã€�ã€ŒæŒ‡å®šç´�æœŸã€�ã�¯é…�å�°è¨ˆç”»_ã‚¿ã‚¹ã‚¯å…¥åŠ›ã�®å½“è©²è¡Œã‚»ãƒ«ã�®ã�¿ã€‚
    # ã€ŒåŽŸå��æŠ•å…¥æ—¥ã€�ã�¯ä¸Šæ›¸ã��åˆ—ã�«æ—¥ä»˜ã�Œã�‚ã‚‹ã�¨ã��ã��ã�®å€¤ã€�ç„¡ã�„ã�¨ã��åˆ—ã€ŒåŽŸå��æŠ•å…¥æ—¥ã€�ï¼ˆè¨ˆç”»åŸºæº–ç´�æœŸã�¨æ··å�Œã�—ã�ªã�„ï¼‰
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
    
    # ã‚¹ãƒ†ãƒ¼ã‚¿ã‚¹ï¼ˆé…�å�°ã�®å�¯å�¦ãƒ»æ®‹ï¼‰ï¼šå®Œäº†ç›¸å½“=é…�å�°å�¯ï¼�æœªå‰²å½“=é…�å�°ä¸�å�¯ï¼�ä¸€éƒ¨ã�®ã�¿=é…�å�°æ®‹
    # è¨ˆç”»åŸºæº–+1 ã�®å†�è©¦è¡Œã�Œä¾�é ¼NOã�”ã�¨ã�®ä¸Šé™�ã�«é�”ã�—ã�Ÿä¾�é ¼ã�®æœªå®Œäº†è¡Œã�«ã�¯ï¼ˆç´�æœŸè¦‹ç›´ã�—å¿…è¦�ï¼‰ã‚’ä»˜ä¸Žã�™ã‚‹ã€‚
    sorted_tasks_for_result = sorted(task_queue, key=_result_task_sheet_sort_key)
    for t in sorted_tasks_for_result:
        rem_u = float(t.get("remaining_units") or 0)
        hist = bool(t.get("assigned_history"))
        if rem_u <= 1e-9:
            status = "é…�å�°å�¯"
        elif hist and t.get("_partial_retry_calendar_blocked"):
            status = "é…�å�°æ®‹(å‹¤å‹™ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ä¸�è¶³)"
        elif not hist and rem_u > 1e-9:
            status = "é…�å�°ä¸�å�¯"
        else:
            status = "é…�å�°æ®‹"
        _tid_res = str(t.get("task_id", "") or "").strip()
        if (
            _tid_res in _due_shift_exhausted_requests
            and rem_u > 1e-9
            and "ç´�æœŸè¦‹ç›´ã�—å¿…è¦�" not in status
        ):
            status = f"{status}ï¼ˆç´�æœŸè¦‹ç›´ã�—å¿…è¦�ï¼‰"
        
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
        # åˆ—é †: A=ã‚¹ãƒ†ãƒ¼ã‚¿ã‚¹ â†’ ã‚¿ã‚¹ã‚¯ID/å·¥ç¨‹/æ©Ÿæ¢°/å„ªå…ˆåº¦ â†’ å±¥æ­´1..n â†’ ã��ã�®ä»– â†’ æœ€å¾Œã�«ç‰¹åˆ¥æŒ‡å®š_AI
        row_status = {"ã‚¹ãƒ†ãƒ¼ã‚¿ã‚¹": status}
        _dto = t.get("dispatch_trial_order")
        row_core = {
            "ã‚¿ã‚¹ã‚¯ID": t['task_id'],
            "å·¥ç¨‹å��": t['machine'],
            "æ©Ÿæ¢°å��": t.get("machine_name", ""),
            "å„ªå…ˆåº¦": t.get("priority", 999),
            RESULT_TASK_COL_DISPATCH_TRIAL_ORDER: _dto if _dto is not None else "",
        }
        row_history = {}
        for i in range(max_history_len):
            if i < len(t['assigned_history']):
                h = t['assigned_history'][i]
                row_history[f"å±¥æ­´{i+1}"] = _format_result_task_history_cell(t, h)
            else:
                row_history[f"å±¥æ­´{i+1}"] = ""

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
            "å¿…è¦�OP(ä¸Šæ›¸)": rov if rov is not None else "",
            "ã‚¿ã‚¹ã‚¯åŠ¹çŽ‡": parse_float_safe(t.get("task_eff_factor"), 1.0),
            "åŠ å·¥é€”ä¸­": "ã�¯ã�„" if t.get("in_progress") else "ã�„ã�„ã�ˆ",
            "ç‰¹åˆ¥æŒ‡å®šã�‚ã‚Š": "ã�¯ã�„" if t.get("has_special_remark") else "ã�„ã�„ã�ˆ",
            "æ‹…å½“OPæŒ‡å��": (t.get("preferred_operator_raw") or "")[:120],
            "å›žç­”ç´�æœŸ": ans_s,
            "æŒ‡å®šç´�æœŸ": spec_s,
            "è¨ˆç”»åŸºæº–ç´�æœŸ": basis_s,
            TASK_COL_RAW_INPUT_DATE: kenhan_s,
            "ç´�æœŸç·Šæ€¥": "ã�¯ã�„" if t.get("due_urgent") else "ã�„ã�„ã�ˆ",
            "åŠ å·¥é–‹å§‹æ—¥": start_req_s,
            "é…�å®Œ_åŠ å·¥é–‹å§‹": plan_assign_start_s,
            "é…�å®Œ_åŠ å·¥çµ‚äº†": plan_assign_end_s,
            RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16: _plan_end_ans_spec16,
            "ç·�åŠ å·¥é‡�": f"{total_r}R ({t['total_qty_m']}m)",
            "æ®‹åŠ å·¥é‡�": f"{rem_r}R ({int(t['remaining_units'] * t['unit_m'])}m)",
            "å®Œäº†çŽ‡(å®Ÿè¡Œæ™‚ç‚¹)": f"{pct_macro}%",
        }
        row_ai_last = {"ç‰¹åˆ¥æŒ‡å®š_AI": (t.get("task_special_ai_note") or "")[:300]}
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
                    clock_out_s = "ä¼‘"
                cal_rows.append({
                    "æ—¥ä»˜": d,
                    "ãƒ¡ãƒ³ãƒ�ãƒ¼": m,
                    "å‡ºå‹¤": data['start_dt'].strftime("%H:%M") if data['is_working'] else "ä¼‘",
                    "é€€å‹¤": clock_out_s,
                    "åŠ¹çŽ‡": data['efficiency'],
                    "å‚™è€ƒ": data['reason'],
                })

    utilization_data = []
    for d in sorted_dates:
        row_data = {"å¹´æœˆæ—¥": d.strftime("%Y/%m/%d (%a)")}
        # ã��ã�®æ—¥ã�®ã‚¤ãƒ™ãƒ³ãƒˆã�‹ã‚‰ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ¥ä½œæ¥­åˆ†ã‚’ä¸€æ‹¬é›†è¨ˆï¼ˆå…¨ãƒ¡ãƒ³ãƒ�ãƒ¼Ã—å…¨ã‚¤ãƒ™ãƒ³ãƒˆã�®äºŒé‡�ãƒ«ãƒ¼ãƒ—ã‚’é�¿ã�‘ã‚‹ï¼‰
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
                row_data[m] = f"{ratio:.1f}% ({worked_mins}/{total_avail_mins}åˆ†)"
            else:
                row_data[m] = "ä¼‘"
        utilization_data.append(row_data)
        
    df_utilization = pd.DataFrame(utilization_data)

    df_mprio_legend, df_mprio_tbl = build_member_assignment_priority_reference(
        skills_dict, members
    )
    if df_mprio_tbl.empty:
        df_mprio_tbl = pd.DataFrame(
            [
                {
                    "å·¥ç¨‹å��": "",
                    "æ©Ÿæ¢°å��": "",
                    "ã‚¹ã‚­ãƒ«åˆ—ã‚­ãƒ¼": "",
                    "å„ªå…ˆé †ä½�": "",
                    "ãƒ¡ãƒ³ãƒ�ãƒ¼": "",
                    "ãƒ­ãƒ¼ãƒ«": "",
                    "å„ªå…ˆåº¦å€¤_å°�ã�•ã�„ã�»ã�©å…ˆ": "",
                    "skillsã‚»ãƒ«å€¤": "",
                    "å‚™è€ƒ": "ãƒžã‚¹ã‚¿ skills ã�«ã€Œå·¥ç¨‹å��+æ©Ÿæ¢°å��ã€�å½¢å¼�ã�®åˆ—ã�Œè¦‹ã�¤ã�‹ã‚‰ã�ªã�„ã�‹ã€�ãƒ‡ãƒ¼ã‚¿ã�Œã�‚ã‚Šã�¾ã�›ã‚“ã€‚",
                }
            ]
        )

    _usage_txt = build_gemini_usage_summary_text()
    if _usage_txt:
        ai_log_data["Gemini_ãƒˆãƒ¼ã‚¯ãƒ³ãƒ»æ–™é‡‘ã‚µãƒžãƒª"] = _usage_txt[:50000]

    _master_abs_for_result_fmt = os.path.abspath(os.path.join(os.getcwd(), MASTER_FILE))
    _reg_shift_start, _reg_shift_end = _read_master_main_regular_shift_times(
        _master_abs_for_result_fmt
    )
    if _reg_shift_start is not None and _reg_shift_end is not None:
        logging.info(
            "å®šå¸¸æž : master ãƒ¡ã‚¤ãƒ³ A15/B15 â†’ %s ï½ž %sï¼ˆçµ�æžœã�®å®šå¸¸å¤–ã€Œæ—¥æ™‚å¸¯ã€�ç�€è‰²ï¼‰",
            _reg_shift_start.strftime("%H:%M"),
            _reg_shift_end.strftime("%H:%M"),
        )

    logging.info(
        "æ®µéšŽ2: çµ�æžœãƒ–ãƒƒã‚¯ã‚’ä½œæˆ�ã�—ã�¾ã�™ â†’ %s",
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
            pd.DataFrame(cal_rows).to_excel(writer, sheet_name='çµ�æžœ_ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼(å‡ºå‹¤ç°¿)', index=False)
            df_utilization.to_excel(writer, sheet_name='çµ�æžœ_ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ¥ä½œæ¥­å‰²å�ˆ', index=False)
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
                    "åˆ—å��": task_column_order_dedup,
                    "è¡¨ç¤º": vis_list_dedup,
                }
            ).to_excel(writer, sheet_name=COLUMN_CONFIG_SHEET_NAME, index=False)
            df_tasks.to_excel(writer, sheet_name=RESULT_TASK_SHEET_NAME, index=False)
            pd.DataFrame(list(ai_log_data.items()), columns=["é …ç›®", "å†…å®¹"]).to_excel(writer, sheet_name='çµ�æžœ_AIãƒ­ã‚°', index=False)

            _mprio_sheet = RESULT_MEMBER_PRIORITY_SHEET_NAME
            df_mprio_legend.to_excel(writer, sheet_name=_mprio_sheet, index=False)
            _mprio_gap = len(df_mprio_legend) + 2
            df_mprio_tbl.to_excel(
                writer, sheet_name=_mprio_sheet, index=False, startrow=_mprio_gap
            )

            logging.info(
                "æ®µéšŽ2: è¨­å‚™ã‚¬ãƒ³ãƒˆã‚’ç”Ÿæˆ�ã�—ã�¦ã�„ã�¾ã�™ï¼ˆãƒ‡ãƒ¼ã‚¿é‡�ã�«ã‚ˆã‚Šæ•°åˆ†ã�‹ã�‹ã‚‹ã�“ã�¨ã�Œã�‚ã‚Šã�¾ã�™ï¼‰"
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

            # æœªã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒ«è¡Œï¼ˆé…�å�°ä¸�å�¯ãƒ»é…�å�°æ®‹ï¼‰ã‚’ç›®ç«‹ã�Ÿã�›ã‚‹
            status_col_idx = None
            for col_idx, col_name in enumerate(df_tasks.columns, 1):
                if str(col_name) == "ã‚¹ãƒ†ãƒ¼ã‚¿ã‚¹":
                    status_col_idx = col_idx
                    break
            if status_col_idx is not None:
                unscheduled_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                for r in range(2, worksheet_tasks.max_row + 1):
                    st_val = worksheet_tasks.cell(row=r, column=status_col_idx).value
                    st = str(st_val).strip() if st_val is not None else ""
                    if st in ("é…�å�°ä¸�å�¯", "é…�å�°æ®‹"):
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
            "æ®µéšŽ2: çµ�æžœãƒ–ãƒƒã‚¯ã�®ä½œæˆ�ãƒ»ä¿�å­˜ã�«å¤±æ•—ã�—ã�¾ã�—ã�Ÿ: %sï¼ˆ%sï¼‰ã€‚"
            "output å†…ã�® production_plan_multi_day_*.xlsx ã‚’ Excel ã�§é–‹ã�„ã�¦ã�„ã�ªã�„ã�‹ç¢ºèª�ã�—ã�¦ã��ã� ã�•ã�„ã€‚",
            output_filename,
            e,
        )
        raise

    try:
        _apply_excel_date_columns_date_only_display(
            output_filename, "çµ�æžœ_ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼(å‡ºå‹¤ç°¿)", frozenset({"æ—¥ä»˜"})
        )
    except Exception as e:
        logging.warning(f"çµ�æžœ_ã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼(å‡ºå‹¤ç°¿)ã�®æ—¥ä»˜åˆ—è¡¨ç¤ºæ•´å½¢: {e}")

    _stage2_try_copy_column_config_shapes_from_input(
        output_filename,
        (os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK),
    )

    logging.info(f"å®Œäº†: '{output_filename}' ã‚’ç”Ÿæˆ�ã�—ã�¾ã�—ã�Ÿã€‚")

    # =========================================================
    # 5. â˜…è¿½åŠ : ãƒ¡ãƒ³ãƒ�ãƒ¼æ¯Žã�®è¡Œå‹•ã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒ« (åˆ¥ãƒ•ã‚¡ã‚¤ãƒ«) å‡ºåŠ›
    # =========================================================
    member_output_filename = os.path.join(
        output_dir, f"member_schedule_{_stage2_out_stamp}.xlsx"
    )
    
    # æ™‚é–“å¸¯ã�¯å…¨ãƒ¡ãƒ³ãƒ�ãƒ¼å…±é€šã�§1å›žã� ã�‘ç”Ÿæˆ�ï¼ˆãƒ¡ãƒ³ãƒ�ãƒ¼æ•°åˆ†ã�®é‡�è¤‡è¨ˆç®—ã‚’é�¿ã�‘ã‚‹ï¼‰
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
        "æ®µéšŽ2: ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ¥ã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒ«ã‚’ä½œæˆ�ã�—ã�¾ã�™ â†’ %s",
        os.path.basename(member_output_filename),
    )
    try:
        with pd.ExcelWriter(member_output_filename, engine="openpyxl") as member_writer:
            for m in members:
                # å�„è¡Œã�®è¾žæ›¸ã‚’åˆ�æœŸåŒ–
                m_schedule = {t_label: {"æ™‚é–“å¸¯": t_label} for t_label in time_labels}
            
                # å�„æ—¥ä»˜ã�®ã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒ«ã‚’åˆ—ã�¨ã�—ã�¦åŸ‹ã‚�ã�¦ã�„ã��
                for d in sorted_dates:
                    d_str = d.strftime("%m/%d (%a)")
                
                    # å…¨æ—¥é�žå‹¤å‹™: å¹´ä¼‘ï¼ˆã‚«ãƒ¬ãƒ³ãƒ€ãƒ¼ *ï¼‰ã�¯ã€Žå¹´ä¼‘ã€�ã€�å·¥å ´ä¼‘æ—¥ã�ªã�©ã�¯ã€Žä¼‘ã€�
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
                    
                        # åˆ¤å®šç”¨ã�®ä¸­é–“æ™‚åˆ»ã‚’è¨ˆç®—
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
                            # è©²å½“ã�™ã‚‹ã‚¿ã‚¹ã‚¯ã‚’æŽ¢ã�™ï¼ˆsubs_list ã�¯äº‹å‰�è§£æž�æ¸ˆã�¿ï¼‰
                            active_ev = next((e for e in events_today if e['start_dt'] <= grid_mid_dt < e['end_dt'] and (e['op'] == m or m in e.get('subs_list', []))), None)
                            if active_ev:
                                role = "ä¸»" if active_ev['op'] == m else "è£œ"
                                text = f"[{active_ev['task_id']}] {active_ev['machine']}({role})"
                            else:
                                text = "" # ä½•ã‚‚å‰²ã‚Šå½“ã�¦ã‚‰ã‚Œã�¦ã�„ã�ªã�„ç©ºã��æ™‚é–“
                    
                        m_schedule[t_label][d_str] = text
                    
                # ãƒ‡ãƒ¼ã‚¿ãƒ•ãƒ¬ãƒ¼ãƒ åŒ–ã�—ã�¦ã‚·ãƒ¼ãƒˆã�«æ›¸ã��è¾¼ã�¿
                df_m = pd.DataFrame(list(m_schedule.values()))
                cols = ["æ™‚é–“å¸¯"] + [d.strftime("%m/%d (%a)") for d in sorted_dates]
                df_m = df_m[[c for c in cols if c in df_m.columns]]
                df_m.to_excel(member_writer, sheet_name=m, index=False)
            
                # --- æ—¢å®šãƒ•ã‚©ãƒ³ãƒˆãƒ»ç½«ç·šãƒ»è¦‹å‡ºã�—èƒŒæ™¯ï¼ˆåˆ—å¹…ã�¯ VBA å�–ã‚Šè¾¼ã�¿æ™‚ã�® AutoFitï¼‰ ---
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
            "æ®µéšŽ2: ãƒ¡ãƒ³ãƒ�ãƒ¼åˆ¥ã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒ«ã�®ä¿�å­˜ã�«å¤±æ•—ã�—ã�¾ã�—ã�Ÿ: %sï¼ˆ%sï¼‰ã€‚"
            "member_schedule_*.xlsx ã‚’ Excel ã�§é–‹ã�„ã�¦ã�„ã�ªã�„ã�‹ç¢ºèª�ã�—ã�¦ã��ã� ã�•ã�„ã€‚",
            member_output_filename,
            e,
        )
        raise

    logging.info(f"å®Œäº†: å€‹äººåˆ¥ã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒ«ã‚’ '{member_output_filename}' ã�«å‡ºåŠ›ã�—ã�¾ã�—ã�Ÿã€‚")
    _try_write_main_sheet_gemini_usage_summary("æ®µéšŽ2")