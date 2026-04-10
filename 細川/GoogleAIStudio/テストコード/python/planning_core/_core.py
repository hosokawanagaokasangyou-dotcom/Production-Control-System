"""planning_core 螳溯｣�譛ｬ菴難ｼ医ヱ繝�繧ｱ繝ｼ繧ｸ蜀��ｼ峨Ａ`import planning_core`` 縺ｧ bootstrap 縺悟�医↓螳溯｡後＆繧後ｋ縺薙→縲�"""
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

# AI 蛯呵�繝ｻ驟榊床荳崎�ｽ繝ｭ繧ｸ繝�繧ｯ D竊脱 縺ｮ TTL 繧ｭ繝｣繝�繧ｷ繝･�ｼ域立 output/ 縺九ｉ json/ 縺ｸ遘ｻ陦鯉ｼ�
_ai_remarks_cache_name = "ai_remarks_cache.json"
_ai_cache_legacy = os.path.join(output_dir, _ai_remarks_cache_name)
_ai_cache_new = os.path.join(json_data_dir, _ai_remarks_cache_name)
if os.path.isfile(_ai_cache_legacy) and not os.path.isfile(_ai_cache_new):
    try:
        shutil.move(_ai_cache_legacy, _ai_cache_new)
    except OSError:
        pass
ai_cache_path = _ai_cache_new
# 縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九阪す繝ｼ繝井ｽ懈�舌�ｻ菫晏ｭ倥�ｮ謌仙凄繝�繝舌ャ繧ｰ�ｼ�execution_log 縺ｨ菴ｵ逕ｨ�ｼ�
exclude_rules_sheet_debug_log_path = os.path.join(log_dir, "exclude_rules_sheet_debug.txt")
# 菫晏ｭ伜､ｱ謨玲凾縺ｫ E 蛻暦ｼ医Ο繧ｸ繝�繧ｯ蠑擾ｼ峨□縺代ｒ騾驕ｿ縺励∵ｬ｡蝗� run_exclude_rules_sheet_maintenance 縺ｧ閾ｪ蜍暮←逕ｨ縺吶ｋ�ｼ�json 繝輔か繝ｫ繝�ｼ�
EXCLUDE_RULES_E_SIDECAR_FILENAME = "exclude_rules_e_column_pending.json"
# openpyxl 菫晏ｭ伜､ｱ謨玲凾縺ｫ VBA 縺� E 蛻励∈譖ｸ縺崎ｾｼ繧縺溘ａ縺ｮ UTF-8 TSV�ｼ�Base64�ｼ峨�
EXCLUDE_RULES_E_VBA_TSV_FILENAME = "exclude_rules_e_column_vba.tsv"
# openpyxl 菫晏ｭ伜､ｱ謨玲凾縺ｫ VBA 縺� A縲廢 繧剃ｸ諡ｬ蜿肴丐縺吶ｋ UTF-8 TSV�ｼ郁｡後＃縺ｨ縺ｫ 5 繧ｻ繝ｫ蛻� Base64�ｼ峨�
EXCLUDE_RULES_MATRIX_VBA_FILENAME = "exclude_rules_matrix_vba.tsv"
# VBA 縺後Γ繧､繝ｳ P 蛻励∈譖ｸ縺崎ｾｼ繧縺溘ａ縺ｮ UTF-8 繝�繧ｭ繧ｹ繝茨ｼ�Excel 髢九＞縺溘∪縺ｾ save 縺ｧ縺阪↑縺�蝠城｡後�ｮ蝗樣∩�ｼ�
GEMINI_USAGE_SUMMARY_FOR_MAIN_FILE = "gemini_usage_summary_for_main.txt"
# 蜈ｨ螳溯｡後ｒ騾壹＠縺� Gemini 蛻ｩ逕ｨ繝ｻ謗ｨ螳壽侭驥代�ｮ邏ｯ險茨ｼ�API 蠢懃ｭ斐＃縺ｨ縺ｫ譖ｴ譁ｰ縲ゆｿ晏ｭ伜�医�ｯ API_Payment 繝輔か繝ｫ繝�ｼ�
GEMINI_USAGE_CUMULATIVE_JSON_FILE = "gemini_usage_cumulative.json"
# 譛滄俣蛻･繝舌こ繝�繝医ｒ繝輔Λ繝�繝亥喧縺励◆ CSV�ｼ�Excel 縺ｮ謚倥ｌ邱壹�ｻ譽偵げ繝ｩ繝慕畑�ｼ�
GEMINI_USAGE_BUCKETS_CSV_FILE = "gemini_usage_buckets_for_chart.csv"
# 繝｡繧､繝ｳ繧ｷ繝ｼ繝医�ｻGemini 譌･谺｡謗ｨ遘ｻ�ｼ�xlwings: Q縲彝�ｼ晄侭驥代∪縺溘�ｯ蜻ｼ蜃ｺ縺励ヾ縲弋�ｼ晏粋險医ヨ繝ｼ繧ｯ繝ｳ�ｼ�
GEMINI_USAGE_CHART_COL_DATE = 17  # Q
GEMINI_USAGE_CHART_COL_VALUE = 18  # R
GEMINI_USAGE_CHART_COL_TOK_DATE = 19  # S�ｼ医げ繝ｩ繝慕畑縺ｫ譌･莉倥ｒ隍�陬ｽ�ｼ�
GEMINI_USAGE_CHART_COL_TOK_VALUE = 20  # T�ｼ�total_tokens 逶ｸ蠖難ｼ�
GEMINI_USAGE_CHART_HEADER_ROW = 16
GEMINI_USAGE_CHART_ANCHOR_CELL = "T16"
GEMINI_USAGE_CHART_TOKENS_ANCHOR_CELL = "AA16"
GEMINI_USAGE_CHART_MAX_DAYS = 14
GEMINI_USAGE_CHART_CLEAR_ROWS = 36
# xlwings 縺ｧ雋ｼ繧区釜繧檎ｷ壹�ｮ蜷榊燕�ｼ亥�榊ｮ溯｡梧凾縺ｫ蜑企勁縺励※縺九ｉ菴懊ｊ逶ｴ縺呻ｼ�
GEMINI_USAGE_XLW_CHART_NAME = "_GeminiApiDailyTrend"
GEMINI_USAGE_XLW_CHART_TOKENS_NAME = "_GeminiApiDailyTokens"
# 繝�繧ｹ繝�: EXCLUDE_RULES_TEST_E1234=1 縺ｧ EXCLUDE_RULES_SHEET_NAME�ｼ医瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九搾ｼ峨�ｮ E 蛻励↓ "1234" 繧呈嶌縺擾ｼ井ｿ晏ｭ倡ｵ瑚ｷｯ縺ｮ遒ｺ隱咲畑�ｼ峨�
# TASK_INPUT_WORKBOOK 縺ｯ縲悟刈蟾･險育判DATA縲阪す繝ｼ繝井ｻ倥″繝悶ャ繧ｯ�ｼ井ｾ�: 逕溽肇邂｡逅�_AI驟榊床繝�繧ｹ繝�.xlsm�ｼ峨ｒ謖�螳壹☆繧九％縺ｨ縲�
# 陦後�ｯ EXCLUDE_RULES_TEST_E1234_ROW�ｼ域里螳� 9縲�2 譛ｪ貅縺ｯ 9 縺ｫ荳ｸ繧√ｋ�ｼ峨�

# =========================================================
# 縲占ｨｭ螳壹羨PI繧ｭ繝ｼ / 蝓ｺ譛ｬ繝ｫ繝ｼ繝ｫ / 繝輔ぃ繧､繝ｫ蜷�
# =========================================================
# Gemini API 繧ｭ繝ｼ縺ｯ TASK_INPUT_WORKBOOK 遒ｺ螳壼ｾ後∽ｸ玖ｨ倥瑚ｨｭ螳壹坑1 縺ｮ JSON 縺九ｉ隗｣豎ｺ�ｼ亥ｹｳ譁�縺ｾ縺溘�ｯ format_version 2 縺ｮ證怜捷蛹厄ｼ峨�
# 譛ｪ險ｭ螳壽凾縺ｮ縺ｿ遘ｻ陦檎畑縺ｫ迺ｰ蠅�螟画焚 GEMINI_API_KEY 繧貞盾辣ｧ縲�

GEMINI_MODEL_FLASH = "gemini-2.5-flash"
# 謗ｨ螳壽侭驥�: USD / 1M tokens�ｼ亥�･蜉�, 蜃ｺ蜉幢ｼ峨ょ�ｬ蠑上�ｮ譛譁ｰ蜊倅ｾ｡縺ｫ蜷医ｏ縺帙※譖ｴ譁ｰ縺吶ｋ縺薙→縲�
# 迺ｰ蠅�螟画焚 GEMINI_PRICE_USD_IN_PER_M / GEMINI_PRICE_USD_OUT_PER_M 縺ｧ荳頑嶌縺榊庄�ｼ�Flash 蜷代￠�ｼ峨�
_GEMINI_FLASH_IN_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_IN_PER_M", "0.075") or 0.075
)
_GEMINI_FLASH_OUT_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_OUT_PER_M", "0.30") or 0.30
)
GEMINI_JPY_PER_USD = float(os.environ.get("GEMINI_JPY_PER_USD", "150") or 150)

# ---------------------------------------------------------------------------
# 莉･髯阪�ｮ螳壽焚繝悶Ο繝�繧ｯ縺ｯ縲窪xcel 蛻苓ｦ句�ｺ縺励阪→ 1:1 縺ｧ蟇ｾ蠢懊＆縺帙ｋ縲�
# 蛻怜錐繧貞､峨∴繧句�ｴ蜷医�ｯ VBA繝ｻ繝槭け繝ｭ蛛ｴ繧ｷ繝ｼ繝医→蜷梧凾縺ｫ逶ｴ縺吶％縺ｨ縲�
# ---------------------------------------------------------------------------

MASTER_FILE = "master.xlsm" # skills縺ｨattendance(縺翫ｈ縺ｳtasks)繧堤ｵｱ蜷医＠縺溘ヵ繧｡繧､繝ｫ
# VBA縲稽aster_讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ繧剃ｽ懈�舌阪す繝ｼ繝茨ｼ�1 譎る俣繧ｹ繝ｭ繝�繝亥頃譛峨ｒ谿ｵ髫�2縺ｮ machine_avail_dt 縺ｫ蜿肴丐�ｼ�
SHEET_MACHINE_CALENDAR = "讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ"
# ``generate_plan`` 髢句ｧ区凾縺ｫ蜀崎ｨｭ螳壹Ｅate -> 險ｭ蛯吶く繝ｼ -> [ (start, end), ... ] 蜊企幕蛹ｺ髢� [start, end)
_MACHINE_CALENDAR_BLOCKS_BY_DATE: dict[
    date, dict[str, list[tuple[datetime, datetime]]]
] = {}

# master.xlsm: 萓晞�ｼNO 縺悟､峨ｏ繧句燕蠕後�ｮ蟾･遞凝玲ｩ滓｢ｰ縺斐→縺ｮ貅門ｙ繝ｻ蠕悟ｧ区忰�ｼ亥���ｼ会ｼ乗ｩ滓｢ｰ縺斐→縺ｮ譌･谺｡蟋区･ｭ貅門ｙ�ｼ亥���ｼ�
SHEET_MACHINE_CHANGEOVER = "險ｭ螳喟萓晞�ｼ蛻�譖ｿ蜑榊ｾ梧凾髢�"
SHEET_MACHINE_DAILY_STARTUP = "險ｭ螳喟讖滓｢ｰ_譌･谺｡蟋区･ｭ貅門ｙ"
# ``generate_plan`` 髢句ｧ区凾縺ｫ蜀崎ｨｭ螳夲ｼ医す繝ｼ繝育┌縺励�ｻ遨ｺ縺ｯ遨ｺ霎樊嶌�ｼ晏ｾ捺擂縺ｩ縺翫ｊ�ｼ�
_STAGE2_MACHINE_CHANGEOVER_BY_EQ: dict[str, tuple[int, int]] = {}
_STAGE2_MACHINE_DAILY_STARTUP_MIN_BY_MACHINE: dict[str, int] = {}
# master 繝｡繧､繝ｳ A15�ｼ亥ｮ壼ｸｸ髢句ｧ具ｼ峨よ律谺｡蟋区･ｭ貅門ｙ繧貞共諤� forward 縺ｧ縺ｯ縺ｪ縺� [髢句ｧ�, 髢句ｧ�+N蛻�) 縺ｮ螢∵凾險医↓霈峨○繧九�
_STAGE2_REGULAR_SHIFT_START: time | None = None
# timeline_events 縺ｮ event_kind�ｼ育怐逡･譎ゅ�ｯ蜉�蟾･縺ｨ縺ｿ縺ｪ縺呻ｼ�
TIMELINE_EVENT_MACHINING = "machining"
TIMELINE_EVENT_MACHINE_DAILY_STARTUP = "machine_daily_startup"
TIMELINE_EVENT_CHANGEOVER_CLEANUP = "changeover_cleanup"
TIMELINE_EVENT_CHANGEOVER_PREP = "changeover_prep"
# VBA縲稽aster_邨�縺ｿ蜷医ｏ縺幄｡ｨ繧呈峩譁ｰ縲阪〒菴懊ｋ繧ｷ繝ｼ繝茨ｼ亥ｷ･遞�+讖滓｢ｰ繧ｭ繝ｼ縺ｨ繝｡繝ｳ繝舌�ｼ邱ｨ謌撰ｼ�
MASTER_SHEET_TEAM_COMBINATIONS = "邨�縺ｿ蜷医ｏ縺幄｡ｨ"
# 繝｡繝ｳ繝舌�ｼ蛻･蜍､諤�繧ｷ繝ｼ繝�: master.xlsm 縺ｧ縺ｯ縲御ｼ第嚊蛹ｺ蛻�縲阪→縲悟ｙ閠�縲阪′蛻･蛻励�
# 蜍､諤�AI縺ｮ蜈･蜉帙�ｯ蛯呵�縺ｮ縺ｿ縲ゅ◆縺�縺� reason�ｼ郁｡ｨ遉ｺ繝ｻ荳ｭ謚懊￠陬懈ｭ｣繝ｻ蛟倶ｺｺ繧ｷ繝ｼ繝医�ｮ莨第�ｩ/莨第嚊譁�險�ｼ峨�ｯ縲∝ｙ閠�縺檎ｩｺ縺ｮ縺ｨ縺堺ｼ第嚊蛹ｺ蛻�繧貞ｼ輔″邯吶＄縲�
# master 繧ｫ繝ｬ繝ｳ繝繝ｼ�ｼ丞�ｺ蜍､邁ｿ.txt 貅匁侠: 蜑堺ｼ�=蜊亥燕蟷ｴ莨代�ｻ莨第�ｩ譎る俣1_邨ゆｺ��ｽ槫ｮ壼ｸｸ邨ゆｺ��ｼ亥壕蠕御ｼ第�ｩ14:45�ｽ�15:00�ｼ会ｼ丞ｾ御ｼ�=螳壼ｸｸ髢句ｧ具ｽ樔ｼ第�ｩ譎る俣1_髢句ｧ九�ｻ蜊亥ｾ悟ｹｴ莨托ｼ丞嵜=莉匁侠轤ｹ蜍､蜍吶�
# 蛯呵�蛻励�ｻ莨第嚊蛹ｺ蛻�縺ｯ蜍､諤� AI 縺ｧ讒矩�蛹厄ｼ磯�榊床荳榊盾蜉�繝ｻis_holiday繝ｻ荳ｭ謚懊￠遲会ｼ峨ょｙ閠�縺檎ｩｺ縺ｧ繧ゆｼ第嚊蛹ｺ蛻�縺ｮ縺ｿ縺ｮ陦後�ｯ AI 縺ｫ貂｡縺吶�
ATT_COL_LEAVE_TYPE = "莨第嚊蛹ｺ蛻�"
ATT_COL_REMARK = "蛯呵�"
# 繝｡繝ｳ繝舌�ｼ蜍､諤�繧ｷ繝ｼ繝茨ｼ�master.xlsm�ｼ�: 螳壽凾縺ｮ縲碁蜍､譎る俣縲阪→蛻�縺代※騾蜍､荳企剞繧呈欠螳夲ｼ井ｻｻ諢丞�暦ｼ�
ATT_COL_OT_END = "谿区･ｭ邨よ･ｭ"
# 蜍､諤�蛯呵� AI 縺ｮ JSON 繧ｹ繧ｭ繝ｼ繝槭ｒ螟峨∴縺溘ｉ譖ｴ譁ｰ縺励√く繝｣繝�繧ｷ繝･繧ｭ繝ｼ繧堤┌蜉ｹ蛹悶☆繧�
ATTENDANCE_REMARK_AI_SCHEMA_ID = "v2_haitai_fuka"
# need 繧ｷ繝ｼ繝�: 縲悟渕譛ｬ蠢�隕∽ｺｺ謨ｰ縲崎｡鯉ｼ�A蛻励↓縲悟ｿ�隕∽ｺｺ謨ｰ縲阪ｒ蜷ｫ繧�ｼ会ｼ� 縺昴�ｮ逶ｴ荳九�ｮ縲碁�榊床譎りｿｽ蜉�莠ｺ謨ｰ�ｼ丈ｽ吝鴨譎りｿｽ蜉�莠ｺ謨ｰ縲咲ｭ�
# �ｼ�Excel 荳翫�ｯ讎ゅ�ｭ 5 陦檎岼莉倩ｿ代ゆｽ吝臆譎ゅ↓蠅励ｄ縺帙ｋ莠ｺ謨ｰ荳企剞繝ｻ蟾･遞凝玲ｩ滓｢ｰ蛻暦ｼ�
# �ｼ� 陦後檎音蛻･謖�螳�1縲搾ｽ槭檎音蛻･謖�螳�99縲搾ｼ亥ｿ�隕∽ｺｺ謨ｰ縺ｮ荳頑嶌縺阪�ｻ1�ｽ�99�ｼ�
NEED_COL_CONDITION = "萓晞�ｼNO譚｡莉ｶ"
NEED_COL_NOTE = "蛯呵�"
# need縲碁�榊床譎りｿｽ蜉�莠ｺ謨ｰ縲阪ｒ貅譫�菴ｿ縺｣縺ｦ繧ゅ∝腰菴阪≠縺溘ｊ蜉�蟾･譎る俣縺檎洒縺上↑繧九�ｮ縺ｯ譛螟ｧ縺ｧ縺薙�ｮ蜑ｲ蜷茨ｼ井ｾ�: 0.05 竕� 5%�ｼ�
SURPLUS_TEAM_MAX_SPEEDUP_RATIO = 0.05
# 繧ｿ繧ｹ繧ｯ縺ｯ tasks.xlsx 繧剃ｽｿ繧上★縲〃BA 縺九ｉ貂｡縺� TASK_INPUT_WORKBOOK 縺ｮ縲悟刈蟾･險育判DATA縲阪�ｮ縺ｿ
TASKS_INPUT_WORKBOOK = os.environ.get("TASK_INPUT_WORKBOOK", "").strip()
TASKS_SHEET_NAME = "蜉�蟾･險育判DATA"

# 縺薙�ｮ繧ｷ繝ｼ繝亥錐繧貞性繧繝悶ャ繧ｯ縺ｯ openpyxl 縺瑚ｪｭ縺ｿ譖ｸ縺阪↓螟ｱ謨励☆繧九％縺ｨ縺後≠繧九◆繧√〕oad_workbook 繧定ｩｦ陦後＠縺ｪ縺�
OPENPYXL_INCOMPATIBLE_SHEET_MARKER = "驟榊床_驟榊床荳崎ｦ∝ｷ･遞�"


def _ooxml_workbook_sheet_names(wb_path: str) -> list[str] | None:
    """ZIP 蜀� xl/workbook.xml 縺九ｉ繧ｷ繝ｼ繝亥錐荳隕ｧ繧貞叙繧具ｼ�openpyxl 繧剃ｽｿ繧上↑縺��ｼ峨�"""
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
    """蠖楢ｩｲ繝代せ縺� OOXML 縺ｧ繧ｷ繝ｼ繝医碁�榊床_驟榊床荳崎ｦ∝ｷ･遞九阪ｒ蜷ｫ繧縺ｨ縺� True�ｼ�openpyxl 蛻ｩ逕ｨ繧帝∩縺代ｋ�ｼ峨�"""
    p = (wb_path or "").strip()
    if not p:
        return False
    names = _ooxml_workbook_sheet_names(p)
    if not names:
        return False
    return OPENPYXL_INCOMPATIBLE_SHEET_MARKER in names


# 繝槭け繝ｭ繝悶ャ繧ｯ縲瑚ｨｭ螳壹坑1: 遉ｾ蜀�蜈ｱ譛我ｸ翫�ｮ Gemini 隱崎ｨｼ JSON 縺ｮ繝代せ
APP_CONFIG_SHEET_NAME = "險ｭ螳�"
# 證怜捷蛹冶ｪ崎ｨｼ JSON�ｼ�format_version 2�ｼ峨�ｮ蠕ｩ蜿ｷ縺ｯ蟶ｸ縺ｫ縺薙�ｮ螳壽焚縺ｮ縺ｿ�ｼ育､ｾ蜀�謇矩��縺ｮ繝代せ繝輔Ξ繝ｼ繧ｺ縺ｨ荳閾ｴ縺輔○繧九ゅΟ繧ｰ繝ｻUI 縺ｫ蜃ｺ縺輔↑縺��ｼ峨�
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
    """縲瑚ｨｭ螳壹阪す繝ｼ繝� B1 縺九ｉ隱崎ｨｼ JSON 繝輔ぃ繧､繝ｫ繝代せ繧定ｪｭ繧縲ら┌縺代ｌ縺ｰ None縲�"""
    if not wb_path or not os.path.isfile(wb_path):
        return None
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.debug(
            "Gemini: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ openpyxl 縺ｧ縲�%s縲�!B1 繧定ｪｭ縺ｿ縺ｾ縺帙ｓ縲�",
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
            "Gemini: 繝槭け繝ｭ繝悶ャ繧ｯ縲�%s縲阪�ｮ縲�%s縲�!B1 繧定ｪｭ繧√∪縺帙ｓ: %s",
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
    繝槭け繝ｭ繝悶ャ繧ｯ縲瑚ｨｭ螳壹阪す繝ｼ繝医�ｮ謖�螳壼�暦ｼ�1=A, 2=B�ｼ�3 陦檎岼莉･髯阪°繧我ｾ晞�ｼNO繧定ｪｭ繧縲�
    遨ｺ繧ｻ繝ｫ縺ｯ繧ｹ繧ｭ繝�繝励る｣邯� 30 繧ｻ繝ｫ遨ｺ縺ｧ謇薙■蛻�繧翫よ怙螟ｧ 500 陦後ゅき繝ｳ繝槫玄蛻�繧翫〒隍�謨ｰ蜿ｯ縲�
    """
    out: list[str] = []
    if not wb_path or not os.path.isfile(wb_path):
        return out
    if _workbook_should_skip_openpyxl_io(wb_path):
        msg = (
            f"{log_label}: 繝悶ャ繧ｯ縺ｫ縲鶏OPENPYXL_INCOMPATIBLE_SHEET_MARKER}縲阪′縺ゅｋ縺溘ａ"
            f"縲鶏APP_CONFIG_SHEET_NAME}縲�!{column_letter_desc}3 莉･髯阪�ｯ openpyxl 縺ｧ隱ｭ繧√∪縺帙ｓ縲�"
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
            "%s: 縲�%s縲�!%s3 莉･髯阪�ｮ萓晞�ｼNO繧定ｪｭ繧√∪縺帙ｓ�ｼ育┌隕厄ｼ�: %s",
            log_label,
            APP_CONFIG_SHEET_NAME,
            column_letter_desc,
            ex,
        )
        return []
    return out


def _read_trace_schedule_task_ids_from_config_sheet(wb_path: str) -> list[str]:
    """
    繝槭け繝ｭ繝悶ャ繧ｯ縲瑚ｨｭ螳壹阪す繝ｼ繝� A 蛻励�ｮ 3 陦檎岼莉･髯阪ｒ縲�驟榊床繝医Ξ繝ｼ繧ｹ蟇ｾ雎｡縺ｮ萓晞�ｼNO縺ｨ縺励※隱ｭ繧縲�
    遨ｺ繧ｻ繝ｫ縺ｯ繧ｹ繧ｭ繝�繝励る｣邯� 30 繧ｻ繝ｫ遨ｺ縺ｪ繧画遠縺｡蛻�繧翫よ怙螟ｧ 500 陦後∪縺ｧ襍ｰ譟ｻ縲�
    """
    return _read_task_ids_from_config_sheet_column(
        wb_path,
        1,
        "驟榊床繝医Ξ繝ｼ繧ｹ",
        "A",
        openpyxl_skip_hint="驟榊床繝医Ξ繝ｼ繧ｹ縺ｯ縲瑚ｨｭ螳壹阪す繝ｼ繝� A 蛻励ｒ openpyxl 縺ｧ隱ｭ繧√↑縺�縺溘ａ辟｡蜉ｹ縺ｧ縺吶�",
    )


def _read_debug_dispatch_task_ids_from_config_sheet(wb_path: str) -> list[str]:
    """
    繝槭け繝ｭ繝悶ャ繧ｯ縲瑚ｨｭ螳壹阪す繝ｼ繝� B 蛻励�ｮ 3 陦檎岼莉･髯阪ｒ縲∵ｮｵ髫�2繝�繝舌ャ繧ｰ驟榊床縺ｮ蟇ｾ雎｡萓晞�ｼNO縺ｨ縺励※隱ｭ繧縲�
    1 莉ｶ繧ら┌縺�蝣ｴ蜷医�ｯ谿ｵ髫�2縺ｯ騾壼ｸｸ繝｢繝ｼ繝会ｼ亥�ｨ莉ｶ驟榊床�ｼ峨らｩｺ繧ｻ繝ｫ繝ｻ謇薙■蛻�繧顔ｭ峨�ｯ A 蛻励ヨ繝ｬ繝ｼ繧ｹ縺ｨ蜷後§縲�
    """
    return _read_task_ids_from_config_sheet_column(
        wb_path,
        2,
        "繝�繝舌ャ繧ｰ驟榊床",
        "B",
        openpyxl_skip_hint="繝�繝舌ャ繧ｰ驟榊床縺ｯ縲瑚ｨｭ螳壹阪す繝ｼ繝� B 蛻励ｒ openpyxl 縺ｧ隱ｭ繧√↑縺�縺溘ａ辟｡蜉ｹ�ｼ亥�ｨ莉ｶ驟榊床�ｼ峨〒縺吶�",
    )


def _show_stage2_debug_dispatch_mode_dialog(task_ids_sorted: list[str]) -> None:
    """險ｭ螳壹す繝ｼ繝� B3莉･髯阪′遨ｺ縺ｧ縺ｪ縺�縺ｨ縺阪□縺大他縺ｶ縲８indows 縺ｧ縺ｯ MessageBox縲√◎繧御ｻ･螟悶�ｯ WARNING 繝ｭ繧ｰ縲�"""
    if not task_ids_sorted:
        return
    preview_lines = task_ids_sorted[:30]
    preview = "\n".join(preview_lines)
    if len(task_ids_sorted) > 30:
        preview += "\n窶ｦ"
    body = (
        "繝�繝舌ャ繧ｰ繝｢繝ｼ繝峨〒螳溯｡後＠縺ｾ縺吶�\n\n"
        "縲瑚ｨｭ螳壹阪す繝ｼ繝� B3莉･髯阪↓蜈･蜉帙＠縺滉ｾ晞�ｼNO縺ｮ縺ｿ繧帝�榊床蟇ｾ雎｡縺ｨ縺励∪縺吶�\n\n"
        "蟇ｾ雎｡萓晞�ｼNO:\n"
        + preview
    )
    title = "谿ｵ髫�2�ｼ磯�榊床�ｼ俄� 繝�繝舌ャ繧ｰ繝｢繝ｼ繝�"
    if sys.platform != "win32":
        logging.warning("%s\n%s", title, body)
        return
    try:
        ctypes.windll.user32.MessageBoxW(0, body, title, 0x00000040)
    except Exception as ex:
        logging.warning(
            "繝�繝舌ャ繧ｰ驟榊床: 繝｡繝�繧ｻ繝ｼ繧ｸ繝懊ャ繧ｯ繧ｹ繧定｡ｨ遉ｺ縺ｧ縺阪∪縺帙ｓ (%s)縲�%s", ex, body
        )


def _extract_gemini_api_key_from_plain_dict(data: dict, json_path: str) -> str | None:
    key = data.get("gemini_api_key")
    if key is None or (isinstance(key, str) and not key.strip()):
        key = data.get("GEMINI_API_KEY")
    if key is None:
        logging.warning(
            "Gemini: 隱崎ｨｼ繝�繝ｼ繧ｿ縺ｫ gemini_api_key�ｼ医∪縺溘�ｯ GEMINI_API_KEY�ｼ峨′縺ゅｊ縺ｾ縺帙ｓ�ｼ�%s�ｼ峨�",
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
            "Gemini: 證怜捷蛹冶ｪ崎ｨｼ JSON 縺ｫ縺ｯ cryptography 縺悟ｿ�隕√〒縺呻ｼ�pip install cryptography�ｼ峨�"
        )
        return None
    token_s = (data.get("fernet_ciphertext") or "").strip()
    if not token_s:
        logging.warning(
            "Gemini: 證怜捷蛹冶ｪ崎ｨｼ JSON 縺ｫ fernet_ciphertext 縺後≠繧翫∪縺帙ｓ�ｼ�%s�ｼ峨�",
            json_path,
        )
        return None
    salt_b64 = (data.get("salt_b64") or "").strip()
    if not salt_b64:
        logging.warning(
            "Gemini: 證怜捷蛹冶ｪ崎ｨｼ JSON 縺ｫ salt_b64 縺後≠繧翫∪縺帙ｓ�ｼ�%s�ｼ峨�",
            json_path,
        )
        return None
    try:
        salt = base64.standard_b64decode(salt_b64)
    except Exception as ex:
        logging.warning("Gemini: salt_b64 縺ｮ隗｣驥医↓螟ｱ謨励＠縺ｾ縺励◆�ｼ�%s�ｼ�: %s", json_path, ex)
        return None
    iterations = int(data.get("iterations") or _GEMINI_CREDENTIALS_PBKDF2_ITERATIONS_DEFAULT)
    kdf_name = (data.get("kdf") or "pbkdf2_sha256").strip()
    if kdf_name != "pbkdf2_sha256":
        logging.warning("Gemini: 譛ｪ蟇ｾ蠢懊�ｮ kdf�ｼ�%s�ｼ�: %s", kdf_name, json_path)
        return None
    try:
        fkey = _derive_fernet_key_from_passphrase(passphrase, salt, iterations)
        plain = Fernet(fkey).decrypt(token_s.encode("ascii"))
    except Exception:
        logging.debug("Gemini: 證怜捷蛹冶ｪ崎ｨｼ縺ｮ蠕ｩ蜿ｷ蜃ｦ逅�縺ｫ螟ｱ謨励＠縺ｾ縺励◆�ｼ�%s�ｼ峨�", json_path)
        return None
    try:
        inner = json.loads(plain.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as ex:
        logging.warning(
            "Gemini: 蠕ｩ蜿ｷ蠕後�ｮ JSON 縺御ｸ肴ｭ｣縺ｧ縺呻ｼ�%s�ｼ�: %s",
            json_path,
            ex,
        )
        return None
    if not isinstance(inner, dict):
        logging.warning("Gemini: 蠕ｩ蜿ｷ蠕後�ｮ JSON 縺ｯ繧ｪ繝悶ず繧ｧ繧ｯ繝医〒縺ゅｋ蠢�隕√′縺ゅｊ縺ｾ縺呻ｼ�%s�ｼ峨�", json_path)
        return None
    return inner


def _load_gemini_api_key_from_credentials_json(
    json_path: str, workbook_path: str | None = None
) -> tuple[str | None, bool]:
    """謌ｻ繧雁､: (api_key 縺ｾ縺溘�ｯ None, 證怜捷蛹門ｽ｢蠑上□縺｣縺溘°)縲よ囓蜿ｷ蛹匁凾縺ｯ _GEMINI_CREDENTIALS_PASSPHRASE_FIXED 縺ｮ縺ｿ縺ｧ蠕ｩ蜿ｷ縲�"""
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
    except OSError as ex:
        logging.warning("Gemini: 隱崎ｨｼ JSON 繧帝幕縺代∪縺帙ｓ: %s (%s)", json_path, ex)
        return None, False
    except json.JSONDecodeError as ex:
        logging.warning("Gemini: 隱崎ｨｼ JSON 縺ｮ蠖｢蠑上′荳肴ｭ｣縺ｧ縺�: %s (%s)", json_path, ex)
        return None, False
    if not isinstance(data, dict):
        logging.warning("Gemini: 隱崎ｨｼ JSON 縺ｯ繧ｪ繝悶ず繧ｧ繧ｯ繝亥ｽ｢蠑上〒縺ゅｋ蠢�隕√′縺ゅｊ縺ｾ縺�: %s", json_path)
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
            logging.info("Gemini API 繧ｭ繝ｼ: 證怜捷蛹冶ｪ崎ｨｼ JSON 縺九ｉ隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�")
        else:
            logging.info(
                "Gemini API 繧ｭ繝ｼ: 繝槭け繝ｭ繝悶ャ繧ｯ縲�%s縲坑1 縺ｮ繝代せ縺九ｉ隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�",
                APP_CONFIG_SHEET_NAME,
            )
elif _cred_path:
    logging.warning(
        "Gemini: 縲�%s縲坑1 縺ｧ謖�螳壹＆繧後◆隱崎ｨｼ JSON 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲�",
        APP_CONFIG_SHEET_NAME,
    )

# B1 縺梧囓蜿ｷ蛹� JSON 縺ｪ縺ｮ縺ｫ繧ｭ繝ｼ縺悟叙繧後↑縺��ｼ亥ｹｳ譁� JSON 縺ｧ繧ｭ繝ｼ谺�關ｽ縺ｨ縺ｮ蛹ｺ蛻･�ｼ峨ょ次蝗�縺ｮ迚ｹ螳壹�ｯ繝ｭ繧ｰ縺ｫ譖ｸ縺九★豎守畑繝｡繝�繧ｻ繝ｼ繧ｸ縺ｮ縺ｿ縲�
_encrypted_json_missing_key = (
    bool(_cred_path)
    and os.path.isfile(_cred_path)
    and _used_encrypted_credentials
    and not API_KEY
)
if _encrypted_json_missing_key:
    logging.error(
        "Gemini: 縲�%s縲坑1 縺ｮ隱崎ｨｼ繝輔ぃ繧､繝ｫ縺九ｉ API 繧ｭ繝ｼ繧貞茜逕ｨ縺ｧ縺阪∪縺帙ｓ縲�"
        " 遉ｾ蜀�謇矩��縺ｫ蠕薙＞隱崎ｨｼ繧貞�崎ｨｭ螳壹☆繧九°縲∫ｮ｡逅�閠�縺ｫ蝠上＞蜷医ｏ縺帙※縺上□縺輔＞縲�",
        APP_CONFIG_SHEET_NAME,
    )

if not API_KEY:
    logging.warning(
        "Gemini API 繧ｭ繝ｼ縺梧悴險ｭ螳壹〒縺吶ゅ�槭け繝ｭ繝悶ャ繧ｯ縺ｫ縲�%s縲阪す繝ｼ繝医ｒ逕ｨ諢上＠ B1 縺ｫ隱崎ｨｼ JSON 縺ｮ繝輔Ν繝代せ繧呈嶌縺�縺ｦ縺上□縺輔＞縲�"
        " 蛯呵�縺ｮ AI 隗｣譫千ｭ峨�ｯ繧ｹ繧ｭ繝�繝励＆繧後∪縺吶�"
        " 縺ｲ縺ｪ蝙�: gemini_credentials.example.json / encrypt_gemini_credentials.py�ｼ域囓蜿ｷ蛹厄ｼ峨�",
        APP_CONFIG_SHEET_NAME,
    )

RESULT_SHEET_GANTT_NAME = "邨先棡_險ｭ蛯吶ぎ繝ｳ繝�"

# 繧ｿ繧ｹ繧ｯ蛻怜錐�ｼ医�槭け繝ｭ螳溯｡後ヶ繝�繧ｯ縲悟刈蟾･險育判DATA縲搾ｼ�
TASK_COL_TASK_ID = "萓晞�ｼNO"
TASK_COL_MACHINE = "蟾･遞句錐"
TASK_COL_MACHINE_NAME = "讖滓｢ｰ蜷�"
TASK_COL_QTY = "謠帷ｮ玲焚驥�"
TASK_COL_ORDER_QTY = "蜿玲ｳｨ謨ｰ"
TASK_COL_SPEED = "蜉�蟾･騾溷ｺｦ"
TASK_COL_PRODUCT = "陬ｽ蜩∝錐"
TASK_COL_ANSWER_DUE = "蝗樒ｭ皮ｴ肴悄"
TASK_COL_SPECIFIED_DUE = "謖�螳夂ｴ肴悄"
TASK_COL_RAW_INPUT_DATE = "蜴溷渚謚募�･譌･"
# 蜉�蟾･險育判DATA 逕ｱ譚･縲る�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙〒縺ｯ蜴溷渚謚募�･譌･縺ｮ蜿ｳ髫｣�ｼ�SOURCE_BASE_COLUMNS 鬆��ｼ峨�
TASK_COL_STOCK_LOCATION = "蝨ｨ蠎ｫ蝣ｴ謇"
# 蜷御ｸ萓晞�ｼNO縺ｮ蟾･遞矩���ｼ医き繝ｳ繝槫玄蛻�繧翫�ｮ蟾･遞句錐�ｼ峨ょ刈蟾･險育判DATA�ｼ城�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙�
TASK_COL_PROCESS_CONTENT = "蜉�蟾･蜀�螳ｹ"
# 謚募�･蜿ｯ閭ｽ譌･縺ｮ逶ｮ螳峨�ｯ縲悟屓遲皮ｴ肴悄縲阪∵悴蜈･蜉帶凾縺ｯ縲梧欠螳夂ｴ肴悄縲搾ｼ亥燕譌･蝓ｺ貅悶�ｻ蠖捺律/驕�繧後�ｯ譛蜆ｪ蜈茨ｼ峨ゅ悟刈蟾･髢句ｧ区律縲榊�励�ｯ蜿ら�ｧ縺励↑縺�縲�
# 螳御ｺ�蛻､螳壹�ｻ騾ｲ謐暦ｼ亥刈蟾･險育判DATA�ｼ�
TASK_COL_COMPLETION_FLAG = "蜉�蟾･螳御ｺ�蛹ｺ蛻�"
TASK_COL_ACTUAL_DONE = "螳溷刈蟾･謨ｰ"   # 譌ｧ莠呈鋤�ｼ育峩謗･縺ｮ蜉�蟾･貂域焚驥擾ｼ�
TASK_COL_ACTUAL_OUTPUT = "螳溷�ｺ譚･鬮�"  # 螳梧�仙刀謨ｰ驥擾ｼ域鋤邂励↓菴ｿ縺��ｼ�
TASK_COL_DATA_EXTRACTION_DT = "繝�繝ｼ繧ｿ謚ｽ蜃ｺ譌･"
AI_CACHE_TTL_SECONDS = 6 * 60 * 60  # 6譎る俣
# json/ai_remarks_cache.json 蜀�縺ｮ繧ｭ繝ｼ謗･鬆ｭ霎橸ｼ郁ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九�ｻ驟榊床荳崎�ｽ繝ｭ繧ｸ繝�繧ｯ D竊脱�ｼ�
AI_CACHE_KEY_PREFIX_EXCLUDE_RULE_DE = "exclude_rule_de_v1"

# 繝槭け繝ｭ繝悶ャ繧ｯ縲悟刈蟾･螳溽ｸｾDATA縲搾ｼ�Power Query 遲峨〒蜿悶ｊ霎ｼ縺ｿ諠ｳ螳夲ｼ�
ACTUALS_SHEET_NAME = "蜉�蟾･螳溽ｸｾDATA"
ACT_COL_TASK_ID = "萓晞�ｼNO"
ACT_COL_PROCESS = "蟾･遞句錐"
ACT_COL_OPERATOR = "諡�蠖楢�"
ACT_COL_START_DT = "髢句ｧ区律譎�"
ACT_COL_END_DT = "邨ゆｺ�譌･譎�"
ACT_COL_START_ALT = "螳溽ｸｾ髢句ｧ�"
ACT_COL_END_ALT = "螳溽ｸｾ邨ゆｺ�"
ACT_COL_DAY = "譌･莉�"
ACT_COL_TIME_START = "髢句ｧ区凾蛻ｻ"
ACT_COL_TIME_END = "邨ゆｺ�譎ょ綾"
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

# --- 2谿ｵ髫主�ｦ逅�: 谿ｵ髫�1謚ｽ蜃ｺ 竊� 繝悶ャ繧ｯ縲碁�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙咲ｷｨ髮� 竊� 谿ｵ髫�2險育判 ---
STAGE1_OUTPUT_FILENAME = "plan_input_tasks.xlsx"
PLAN_INPUT_SHEET_NAME = os.environ.get("TASK_PLAN_SHEET", "").strip() or "驟榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉�"
PLAN_COL_SPEED_OVERRIDE = "蜉�蟾･騾溷ｺｦ_荳頑嶌縺�"
# 遨ｺ逋ｽ縺ｮ縺ｨ縺阪�ｯ蛻励悟次蜿肴兜蜈･譌･縲搾ｼ亥刈蟾･險育判DATA 逕ｱ譚･�ｼ峨ｒ縺昴�ｮ縺ｾ縺ｾ菴ｿ縺�縲よ律莉倥≠繧翫�ｮ縺ｨ縺阪�ｯ驟榊床縺ｮ蜴溷渚蛻ｶ邏�繝ｻ邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ陦ｨ遉ｺ縺ｮ荳｡譁ｹ縺ｧ縺薙�ｮ譌･莉倥ｒ謗｡逕ｨ縲�
PLAN_COL_RAW_INPUT_DATE_OVERRIDE = "蜴溷渚謚募�･譌･_荳頑嶌縺�"
PLAN_COL_PREFERRED_OP = "諡�蠖徹P_謖�螳�"
PLAN_COL_SPECIAL_REMARK = "迚ｹ蛻･謖�螳喟蛯呵�"
# 蜿ら�ｧ蛻励鯉ｼ亥���ｼ蛾�榊床荳崎ｦ√阪�ｯ鄂ｮ縺九↑縺��ｼ亥��繝�繝ｼ繧ｿ縺ｫ逶ｸ蠖薙☆繧九�槭せ繧ｿ蛻励′辟｡縺�縺溘ａ�ｼ峨�
# 繧ｻ繝ｫ蛟､縺ｮ萓具ｼ磯�榊床縺九ｉ螟悶☆�ｼ�: Excel 縺ｮ TRUE / 謨ｰ蛟､ 1 / 譁�蟄怜�励後�ｯ縺�縲阪軽es縲阪荊rue縲阪娯雷縲阪後�縲阪娯酪縲咲ｭ峨�
# 遨ｺ繝ｻFALSE繝ｻ0繝ｻ縲後＞縺�縺医咲ｭ峨�ｯ驟榊床蟇ｾ雎｡縲りｩｳ邏ｰ縺ｯ _plan_row_exclude_from_assignment縲�
PLAN_COL_EXCLUDE_FROM_ASSIGNMENT = "驟榊床荳崎ｦ�"
PLAN_COL_AI_PARSE = "AI迚ｹ蛻･謖�螳喟隗｣譫�"
PLAN_COL_PROCESS_FACTOR = "蜉�蟾･蟾･遞九�ｮ豎ｺ螳壹�励Ο繧ｻ繧ｹ縺ｮ蝗�蟄�"
# 1繝ｭ繝ｼ繝ｫ縺ゅ◆繧翫�ｮ髟ｷ縺包ｼ�m�ｼ峨る�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙↓縺ｮ縺ｿ蟄伜惠�ｼ亥刈蟾･險育判DATA 縺ｫ縺ｯ辟｡縺��ｼ峨り｣ｽ蜩∝錐蛻励�ｮ蜿ｳ髫｣縺ｫ驟咲ｽｮ縲�
PLAN_COL_ROLL_UNIT_LENGTH = "繝ｭ繝ｼ繝ｫ蜊倅ｽ埼聞縺�"
DEBUG_TASK_ID = os.environ.get("DEBUG_TASK_ID", "Y3-26").strip()
# 萓�: set TRACE_TEAM_ASSIGN_TASK_ID=W3-14 窶ｦ 驟榊床繝ｫ繝ｼ繝励〒縲御ｺｺ謨ｰ蛻･縺ｮ譛濶ｯ蛟呵｣懊阪→謗｡逕ｨ逅�逕ｱ繧� INFO 繝ｭ繧ｰ縺ｫ蜃ｺ縺�
TRACE_TEAM_ASSIGN_TASK_ID = os.environ.get("TRACE_TEAM_ASSIGN_TASK_ID", "").strip()
# 驟榊床繝医Ξ繝ｼ繧ｹ蟇ｾ雎｡縺ｯ繝槭け繝ｭ繝悶ャ繧ｯ縲瑚ｨｭ螳壹阪す繝ｼ繝� A 蛻� 3 陦檎岼莉･髯阪�ｮ縺ｿ�ｼ�generate_plan 蜀帝�ｭ縺ｧ遒ｺ螳夲ｼ峨ら腸蠅�螟画焚縺ｯ菴ｿ繧上↑縺�縲�
TRACE_SCHEDULE_TASK_IDS: frozenset[str] = frozenset()
# 谿ｵ髫�2繝�繝舌ャ繧ｰ驟榊床: 縲瑚ｨｭ螳壹坑 蛻� 3 陦檎岼莉･髯阪↓萓晞�ｼNO縺後≠繧九→縺阪�ｮ縺ｿ縲√◎縺ｮ萓晞�ｼ縺ｮ陦後□縺鷹�榊床�ｼ�generate_plan 蜀帝�ｭ縺ｧ遒ｺ螳夲ｼ峨らｩｺ縺ｪ繧牙�ｨ莉ｶ縲�
DEBUG_DISPATCH_ONLY_TASK_IDS: frozenset[str] = frozenset()
# 邏肴悄雜�驕弱Μ繝医Λ繧､縺ｮ螟門�ｴ繝ｩ繧ｦ繝ｳ繝会ｼ�0=蛻晏屓繧ｫ繝ｬ繝ｳ繝繝ｼ騾壹＠縲∽ｻ･髯阪�ｯ while 蜈磯�ｭ縺ｧ譖ｴ譁ｰ�ｼ峨る�榊床繝医Ξ繝ｼ繧ｹ蜃ｺ蜉帙�ｮ繝輔ぃ繧､繝ｫ蜷阪�ｻ謗･鬆ｭ霎槭↓菴ｿ逕ｨ縲�
DISPATCH_TRACE_OUTER_ROUND: int = 0


def _trace_schedule_task_enabled(task_id) -> bool:
    if not TRACE_SCHEDULE_TASK_IDS:
        return False
    return str(task_id or "").strip() in TRACE_SCHEDULE_TASK_IDS


def _sanitize_dispatch_trace_filename_part(task_id: str) -> str:
    """萓晞�ｼNO繧� log 繝輔ぃ繧､繝ｫ蜷阪↓菴ｿ縺�縺溘ａ縺ｮ邁｡譏薙し繝九ち繧､繧ｺ�ｼ�Windows 遖∵ｭ｢譁�蟄励ｒ驕ｿ縺代ｋ�ｼ峨�"""
    s = "".join(
        c if (c.isalnum() or c in "-_.") else "_"
        for c in str(task_id or "").strip()
    )
    return s[:120] if s else "task"


def _reset_dispatch_trace_per_task_logfiles() -> None:
    """
    谿ｵ髫�2螳溯｡後�ｮ蜀帝�ｭ縺ｧ1蝗槭〕og 蜀�縺ｮ dispatch_trace_*.txt 繧偵☆縺ｹ縺ｦ蜑企勁縺吶ｋ�ｼ磯℃蜴ｻ螳溯｡後�ｮ谿矩ｪｸ繧呈ｮ九＆縺ｪ縺��ｼ峨�
    蜷�螟門�ｴ繝ｩ繧ｦ繝ｳ繝臥畑繝輔ぃ繧､繝ｫ縺ｯ generate_plan 縺ｮ while 蜈磯�ｭ縺ｧ _dispatch_trace_begin_outer_round 縺後�倥ャ繝莉倥″譁ｰ隕丈ｽ懈�舌☆繧九�
    execution_log.txt 縺ｨ縺ｯ蛻･繝輔ぃ繧､繝ｫ縲ょ��螳ｹ縺ｯ [驟榊床繝医Ξ繝ｼ繧ｹ task=窶ｦ] 陦後ｒ _log_dispatch_trace_schedule 縺ｧ霑ｽ險�
    �ｼ域律谺｡谿九�ｻ繝ｭ繝ｼ繝ｫ遒ｺ螳壹�ｮ菴吝臆譛臥┌繝ｻ菴吝鴨霑ｽ險倥�ｻ邨ゆｺ�譎ゅし繝槭Μ遲会ｼ峨�
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
    """邏肴悄雜�驕弱Μ繝医Λ繧､縺ｮ螟門�ｴ繝ｩ繧ｦ繝ｳ繝臥分蜿ｷ繧堤｢ｺ螳壹＠縲∝ｽ薙Λ繧ｦ繝ｳ繝臥畑 dispatch_trace_*_rNN.txt 縺ｮ繝倥ャ繝繧�1蝗槭□縺第嶌縺上�"""
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
                    "# 驟榊床繝医Ξ繝ｼ繧ｹ�ｼ井ｾ晞�ｼNO縺斐→繝ｻ螟門�ｴ繝ｩ繧ｦ繝ｳ繝牙挨�ｼ峨ょ酔荳陦後�ｯ log/execution_log.txt 縺ｫ繧ょ�ｺ蜉帙＆繧後∪縺吶�\n"
                    f"# task_id={t}  outer_round={DISPATCH_TRACE_OUTER_ROUND}  "
                    "# �ｼ�0=蛻晏屓繧ｫ繝ｬ繝ｳ繝繝ｼ騾壹＠縲∽ｻ･髯阪�ｯ邏肴悄雜�驕弱Μ繝医Λ繧､縺斐→縺ｫ +1�ｼ噂n\n"
                )
        except OSError as ex:
            logging.warning("dispatch_trace 繝ｭ繧ｰ縺ｮ蛻晄悄蛹悶↓螟ｱ謨�: %s (%s)", path, ex)


def _log_dispatch_trace_schedule(task_id, msg: str, *args) -> None:
    """[驟榊床繝医Ξ繝ｼ繧ｹ task=窶ｦ] 繧� execution_log 縺ｫ蜃ｺ縺励▽縺､縲∝ｯｾ雎｡萓晞�ｼNO蟆ら畑繝輔ぃ繧､繝ｫ縺ｫ繧りｿｽ險倥☆繧九�"""
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
            logging.warning("dispatch_trace 蛛ｴ繝輔ぃ繧､繝ｫ縺ｸ縺ｮ霑ｽ險倥↓螟ｱ謨�: %s (%s)", path, ex)
        except Exception:
            pass


# True: 蠕捺擂縺ｮ縲御ｺｺ謨ｰ譛蜆ｪ蜈医阪ち繝励Ν (-莠ｺ謨ｰ, 髢句ｧ�, -蜊倅ｽ肴焚, 蜆ｪ蜈亥ｺｦ蜷郁ｨ�)縲�False 縺ｮ縺ｨ縺堺ｸ玖ｨ倥せ繝ｩ繝�繧ｯ蛻�縺ｨ邨�縺ｿ蜷医ｏ縺�
TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF = os.environ.get(
    "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF", "0"
).strip().lower() not in ("0", "false", "no", "off", "縺�縺�縺�")


def _team_assign_start_slack_wait_minutes() -> int:
    """蜈ｨ譌･蛟呵｣懊�ｮ譛譌ｩ髢句ｧ九°繧峨％縺ｮ蛻�莉･蜀�縺ｮ驕�繧後↑繧峨�髢句ｧ九ｈ繧贋ｺｺ謨ｰ繧貞━蜈茨ｼ亥���ｼ峨�0 縺ｧ辟｡蜉ｹ縲�"""
    raw = os.environ.get("TEAM_ASSIGN_START_SLACK_WAIT_MINUTES", "60").strip()
    try:
        v = int(raw)
    except ValueError:
        v = 60
    return max(0, v)


TEAM_ASSIGN_START_SLACK_WAIT_MINUTES = _team_assign_start_slack_wait_minutes()

# True 縺ｮ縺ｨ縺� need 繧ｷ繝ｼ繝医碁�榊床譎りｿｽ蜉�莠ｺ謨ｰ縲崎｡後ｒ辟｡隕悶＠縲√メ繝ｼ繝�莠ｺ謨ｰ縺ｯ蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ�ｼ�req_num�ｼ峨�ｮ縺ｿ隧ｦ陦後＠縲√Γ繧､繝ｳ蠕瑚ｿｽ險倥ｂ縺励↑縺�縲�
TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW = (
    os.environ.get("TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW", "0")
    .strip()
    .lower()
    in ("1", "true", "yes", "on", "縺ｯ縺�")
)

# True: 蠕捺擂縺ｩ縺翫ｊ繝｡繧､繝ｳ蜑ｲ莉倥�ｮ邨�縺ｿ蜷医ｏ縺帶爾邏｢縺ｧ req_num縲徨eq_num+霑ｽ蜉�莠ｺ謨ｰ荳企剞縺ｾ縺ｧ隧ｦ縺吶�
# False�ｼ域里螳夲ｼ�: 繝｡繧､繝ｳ縺ｯ req_num 縺ｮ縺ｿ縲りｿｽ蜉�莠ｺ謨ｰ荳企剞縺ｯ蜈ｨ繧ｷ繝溘Η繝ｬ繝ｼ繧ｷ繝ｧ繝ｳ螳御ｺ�蠕後∝ｽ楢ｩｲ繝悶Ο繝�繧ｯ譎る俣縺ｫ
#     莉悶ち繧ｹ繧ｯ縺ｸ譛ｪ蜑ｲ蠖難ｼ域凾髢馴㍾縺ｪ繧翫↑縺暦ｼ峨°縺､ skills 驕ｩ蜷医�ｮ閠�繧偵し繝悶→縺励※霑ｽ險假ｼ�append_surplus_staff_after_main_dispatch�ｼ峨�
TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS = (
    os.environ.get("TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS", "")
    .strip()
    .lower()
    in ("1", "true", "yes", "on", "縺ｯ縺�")
)

# True�ｼ域里螳夲ｼ�: 繝｡繧､繝ｳ驟榊床縺ｮ蠢�隕∽ｺｺ謨ｰ縺ｯ need�ｼ亥渕譛ｬ蠢�隕∽ｺｺ謨ｰ�ｼ狗音蛻･謖�螳夲ｼ峨�ｮ縺ｿ縲�
# False 縺ｮ縺ｨ縺阪�ｯ迚ｹ蛻･謖�螳壼ｙ閠� AI 縺ｮ required_op 縺ｮ縺ｿ險育判蛛ｴ縺九ｉ蜿ら�ｧ縺怜ｾ励ｋ�ｼ医す繝ｼ繝亥�励悟ｿ�隕∽ｺｺ謨ｰ縲阪�ｯ蟒�豁｢貂医∩�ｼ峨�
TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY = (
    os.environ.get("TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "縺�縺�縺�")
)
# True�ｼ域里螳夲ｼ�: master縲檎ｵ�縺ｿ蜷医ｏ縺幄｡ｨ縲阪↓隧ｲ蠖楢｡後′縺ゅｋ蟾･遞�+讖滓｢ｰ縺ｯ縲∫ｵ�蜷医○蜆ｪ蜈亥ｺｦ縺ｮ譏�鬆�縺ｧ
# 譛蛻昴↓謌千ｫ九＠縺溘Γ繝ｳ繝舌�ｼ邱ｨ謌舌ｒ謗｡逕ｨ縲ゅ☆縺ｹ縺ｦ荳榊庄縺ｪ繧牙ｾ捺擂縺ｮ itertools 邨�蜷医○謗｢邏｢縲�
TEAM_ASSIGN_USE_MASTER_COMBO_SHEET = (
    os.environ.get("TEAM_ASSIGN_USE_MASTER_COMBO_SHEET", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "縺�縺�縺�")
)

# ﾂｧB-2 辭ｱ陞咲捩讀懈渊繧貞酔荳險ｭ蛯呻ｼ亥ｷ･遞句�励く繝ｼ�ｼ峨〒縲碁幕蟋区ｸ医∩1莉ｶ縺ｫ谿九Ο繝ｼ繝ｫ縺後≠繧矩俣縺ｯ莉紋ｾ晞�ｼ縺ｮ讀懈渊繧定ｩｦ縺輔↑縺�縲阪°縲�
# 0 / false / no / off 縺ｧ辟｡蜉ｹ縺ｫ縺吶ｋ縺ｨ險ｭ蛯呎凾髢灘牡荳翫〒萓晞�ｼ縺梧ｷｷ蝨ｨ縺怜ｾ励ｋ縺後∝頃譛峨↓繧医ｋ髟ｷ譛溘ヶ繝ｭ繝�繧ｯ�ｼ井ｾ�: W3-14 蝙具ｼ峨ｒ驕ｿ縺代ｉ繧後ｋ縲�
PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE = (
    os.environ.get("PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "縺�縺�縺�", "辟｡蜉ｹ")
)

# ﾂｧB-2 / ﾂｧB-3 蜷御ｸ萓晞�ｼ縺ｧ EC 縺ｨ蠕檎ｶ夲ｼ域､懈渊�ｼ丞ｷｻ霑斐＠�ｼ峨�ｮ諡�蠖楢�髮�蜷医ｒ謗剃ｻ悶☆繧九°縲�
# 0 / false / no / off / 縺�縺�縺� / 辟｡蜉ｹ 縺ｧ辟｡蜉ｹ蛹悶☆繧九→縲∝ｱ･豁ｴ繝吶�ｼ繧ｹ縺ｮ逶ｸ莠帝勁螟悶ｒ陦後ｏ縺壼酔荳莠ｺ迚ｩ縺御ｸ｡蛛ｴ縺ｮ蛟呵｣懊↓谿九ｊ蠕励ｋ縲�
PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS = (
    os.environ.get("PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS", "1")
    .strip()
    .lower()
    not in ("0", "false", "no", "off", "縺�縺�縺�", "辟｡蜉ｹ")
)

# 繝槭け繝ｭ繝悶ャ繧ｯ縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九�: 譌｢螳壹〒縺ｯ openpyxl save 繧定ｩｦ縺輔★ xlwings 蜷梧悄竊担ave�ｼ�Excel 蜊�譛画凾縺ｯ openpyxl 縺悟ｮ溯ｳｪ螟ｱ謨励☆繧九◆繧��ｼ峨ょ､ｱ謨玲凾縺ｯ TSV竊歎BA 蜿肴丐縲�
# 繧ｳ繝槭Φ繝臥ｭ峨〒 openpyxl 繧定ｩｦ縺吝�ｴ蜷医�ｯ EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1縲�
EXCLUDE_RULES_SHEET_NAME = "險ｭ螳喟驟榊床荳崎ｦ∝ｷ･遞�"
EXCLUDE_RULES_SKIP_OPENPYXL_SAVE = os.environ.get(
    "EXCLUDE_RULES_TRY_OPENPYXL_SAVE", ""
).strip().lower() not in ("1", "true", "yes", "on")
EXCLUDE_RULE_COL_PROCESS = "蟾･遞句錐"
EXCLUDE_RULE_COL_MACHINE = "讖滓｢ｰ蜷�"
EXCLUDE_RULE_COL_FLAG = "驟榊床荳崎ｦ�"
EXCLUDE_RULE_COL_LOGIC_JA = "驟榊床荳崎�ｽ繝ｭ繧ｸ繝�繧ｯ"
EXCLUDE_RULE_COL_LOGIC_JSON = "繝ｭ繧ｸ繝�繧ｯ蠑�"
# 蜈�繝悶ャ繧ｯ縺後Ο繝�繧ｯ縺輔ｌ蛻･蜷堺ｿ晏ｭ倥＠縺溷�ｴ蜷医∝酔荳繝励Ο繧ｻ繧ｹ蜀�縺ｮ繝ｫ繝ｼ繝ｫ隱ｭ霎ｼ縺ｯ縺薙�ｮ繝代せ繧貞━蜈�
_exclude_rules_effective_read_path: str | None = None
# 逶ｴ蠕後�ｮ apply_exclude_rules�ｼ亥酔荳繝励Ο繧ｻ繧ｹ�ｼ臥畑: VBA 蜿肴丐蜑阪〒繧� E 蛻嶺ｻ倥″繝ｫ繝ｼ繝ｫ繧剃ｽｿ縺�
_exclude_rules_rules_snapshot: list | None = None
_exclude_rules_snapshot_wb: str | None = None
# 繝ｫ繝ｼ繝ｫ JSON 縺ｮ conditions 縺ｧ蜿ら�ｧ蜿ｯ閭ｽ縺ｪ蛻暦ｼ�AI 繝励Ο繝ｳ繝励ヨ縺ｨ隧穂ｾ｡蝎ｨ繧剃ｸ閾ｴ縺輔○繧具ｼ�
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

# 險育判邨先棡繝悶ャ繧ｯ縲檎ｵ先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縲阪�ｮ蛻鈴��繝ｻ陦ｨ遉ｺ�ｼ医�槭け繝ｭ螳溯｡後ヶ繝�繧ｯ縺ｮ蜷悟錐繧ｷ繝ｼ繝医〒荳頑嶌縺榊庄�ｼ�
RESULT_TASK_SHEET_NAME = "邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ"
RESULT_EQUIPMENT_SCHEDULE_SHEET_NAME = "邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ"
# 菴吝鴨霑ｽ險伜燕縺ｮ繧ｿ繧､繝�繝ｩ繧､繝ｳ繧貞庄隕門喧�ｼ育ｵ先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ縺ｨ蜷後§ 10 蛻�譫�繝ｻ蛻玲ｧ矩��ｼ�
TEMP_EQUIPMENT_SCHEDULE_SHEET_NAME = "TEMP_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ"
# 險ｭ蛯吶�ｻ莠ｺ縺ｮ蜊�譛会ｼ医ヶ繝ｭ繝�繧ｯ�ｼ峨ｒ 10 蛻�譫�縺ｧ荳隕ｧ�ｼ郁ｪｿ譟ｻ繝ｻ讀懆ｨｼ逕ｨ�ｼ�
BLOCK_TABLE_SHEET_NAME = "繝悶Ο繝�繧ｯ繝�繝ｼ繝悶Ν"
# 蟾･遞句錐+讖滓｢ｰ縺ｮ隍�蜷亥�励〒縺ｯ縺ｪ縺上∵ｩ滓｢ｰ蜷榊腰菴阪〒蜷�譫�縺ｮ萓晞�ｼNO繧呈滑謠｡縺励ｄ縺吶￥縺吶ｋ
RESULT_EQUIPMENT_BY_MACHINE_SHEET_NAME = "邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ_讖滓｢ｰ蜷肴ｯ�"
# master 繝｡繧､繝ｳ A15/B15 縺ｮ螳壼ｸｸ螟悶�ｮ縲梧律譎ょｸｯ縲崎ｦ句�ｺ縺礼捩濶ｲ�ｼ育ｵ先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ繝ｻ邨先棡_險ｭ蛯吶ぎ繝ｳ繝茨ｼ�
RESULT_OUTSIDE_REGULAR_TIME_FILL = "FCE4D6"
# 邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ_讖滓｢ｰ蜷肴ｯ�: 驟榊床貂医∩萓晞�ｼNO繧ｻ繝ｫ�ｼ域ｩ滓｢ｰ蛻暦ｼ峨�ｮ阮�縺�繧ｰ繝ｪ繝ｼ繝ｳ
# 邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ�ｼ医♀繧医�ｳ TEMP�ｼ�: 蜉�蟾･蜑肴ｺ門ｙ繝ｻ萓晞�ｼ蛻�譖ｿ蠕悟ｧ区忰縺ｮ險ｭ蛯吶そ繝ｫ繧ょ酔邉ｻ濶ｲ
RESULT_DISPATCHED_REQUEST_FILL = "C6EFCE"
# 邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ: master縲梧ｩ滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ縲榊頃譛峨→驥阪↑繧玖ｨｭ蛯吶そ繝ｫ�ｼ�10蛻�譫��ｼ�
RESULT_MACHINE_CALENDAR_BLOCK_FILL = "D4B3E8"
# 邨先棡_險ｭ蛯吶ぎ繝ｳ繝�: 讖滓｢ｰ蜷阪げ繝ｫ繝ｼ繝暦ｼ域ｩ滓｢ｰ蜷榊�励�ｮ蜷御ｸ蜷咲ｧｰ�ｼ峨＃縺ｨ縺ｫ B縲廢 蛻励ｒ蛹ｺ蛻･縺吶ｋ豺｡濶ｲ�ｼ磯��縺ｫ蜑ｲ蠖薙�ｻ蠕ｪ迺ｰ�ｼ�
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
# 驟榊床繧ｷ繝溘Η繝ｬ繝ｼ繧ｷ繝ｧ繝ｳ髢句ｧ句燕�ｼ亥�晏屓 task_queue.sort 蠕鯉ｼ峨�ｮ繧ｭ繝･繝ｼ鬆�縲�1 蟋九∪繧翫�ｻ蜈ｨ譌･遞九〒荳榊､�
RESULT_TASK_COL_DISPATCH_TRIAL_ORDER = "驟榊床隧ｦ陦碁��逡ｪ"
# 驟榊ｮ契蜉�蟾･邨ゆｺ�縺後悟屓遲皮ｴ肴悄+16:00縲阪∪縺溘�ｯ縲梧欠螳夂ｴ肴悄+16:00縲搾ｼ亥屓遲斐′遨ｺ縺ｮ縺ｨ縺搾ｼ我ｻ･蜑阪°繧定｡ｨ遉ｺ
RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16 = "驟榊ｮ契蝗樒ｭ疲欠螳�16譎ゅ∪縺ｧ"
# 繝槭せ繧ｿ skills 縺ｮ蟾･遞�+讖滓｢ｰ蛻励＃縺ｨ縺ｮ OP/AS 蜑ｲ蠖灘盾閠�鬆��ｼ亥━蜈亥ｺｦ蛟､繝ｻ豌丞錐鬆��ｼ峨→繝√�ｼ繝�謗｡逕ｨ繝ｫ繝ｼ繝ｫ縺ｮ隱ｬ譏�
RESULT_MEMBER_PRIORITY_SHEET_NAME = "邨先棡_莠ｺ蜩｡驟榊床蜆ｪ蜈磯��"
COLUMN_CONFIG_SHEET_NAME = "蛻苓ｨｭ螳喟邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ"
COLUMN_CONFIG_HEADER_COL = "蛻怜錐"
COLUMN_CONFIG_VISIBLE_COL = "陦ｨ遉ｺ"
# 谿ｵ髫�2縺ｮ邨先棡 xlsx 逕滓�仙ｾ後∝�･蜉帙ヶ繝�繧ｯ縺ｮ蛻苓ｨｭ螳壹す繝ｼ繝井ｸ翫�ｮ蝗ｳ蠖｢�ｼ医ヵ繧ｩ繝ｼ繝�繝懊ち繝ｳ遲会ｼ峨ｒ xlwings 縺ｧ隍�陬ｽ縺吶ｋ�ｼ域里螳� ON縲ら┌蜉ｹ蛹悶�ｯ STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT=0�ｼ�
STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT = os.environ.get(
    "STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT", "1"
).strip().lower() in ("1", "true", "yes", "on")
# 邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ譌･莉倡ｳｻ�ｼ�yyyy/mm/dd 譁�蟄怜�暦ｼ峨↓莉倥￠繧九ヵ繧ｩ繝ｳ繝郁牡縲ょｱ･豁ｴ蛻励�ｮ縲先律莉倥代→謠�縺医ｋ
RESULT_TASK_DATE_STYLE_HEADERS = frozenset(
    {
        "蝗樒ｭ皮ｴ肴悄",
        "謖�螳夂ｴ肴悄",
        "險育判蝓ｺ貅也ｴ肴悄",
        TASK_COL_RAW_INPUT_DATE,
        "蜉�蟾･髢句ｧ区律",
        "驟榊ｮ契蜉�蟾･髢句ｧ�",
        "驟榊ｮ契蜉�蟾･邨ゆｺ�",
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
# 遏帷崟讀懷�ｺ縺ｧ繝ｪ繧ｻ繝�繝亥ｯｾ雎｡縺ｫ縺吶ｋ蛻暦ｼ郁ｦ句�ｺ縺苓｡後�ｮ譁�險縺ｨ荳閾ｴ縺吶ｋ縺薙→�ｼ�
PLAN_CONFLICT_STYLABLE_COLS = tuple(PLAN_OVERRIDE_COLUMNS)
# 谿ｵ髫�1蜀肴歓蜃ｺ譎ゅ∵里蟄倥碁�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙阪°繧臥ｶ呎価縺吶ｋ蛻暦ｼ�AI縺ｮ隗｣譫千ｵ先棡蛻励�ｯ豈主屓遨ｺ縺ｫ謌ｻ縺呻ｼ�
PLAN_STAGE1_MERGE_COLUMNS = tuple(c for c in PLAN_OVERRIDE_COLUMNS if c != PLAN_COL_AI_PARSE)
# 荳頑嶌縺堺ｻ･螟悶〒縲∝�肴歓蜃ｺ譎ゅ↓譌ｧ繧ｷ繝ｼ繝医°繧牙ｼ輔″邯吶＄蛻暦ｼ医そ繝ｫ縺檎ｩｺ縺ｧ縺ｪ縺�縺ｨ縺阪�ｮ縺ｿ�ｼ�
# 驟榊床隧ｦ陦碁��逡ｪ縺ｯ豈主屓遨ｺ繧ｯ繝ｪ繧｢縺ｮ縺�縺� fill_plan_dispatch_trial_order_column_stage1 縺ｧ莉倥￠逶ｴ縺吶◆繧∝ｯｾ雎｡螟悶�
PLAN_STAGE1_MERGE_EXTRA_COLUMNS = (PLAN_COL_ROLL_UNIT_LENGTH,)
# openpyxl 菫晏ｭ倥′繝悶ャ繧ｯ繝ｭ繝�繧ｯ縺ｧ螟ｱ謨励＠縺溘→縺阪〃BA 縺碁幕縺�縺ｦ縺�繧九ヶ繝�繧ｯ縺ｸ譖ｸ蠑城←逕ｨ縺吶ｋ縺溘ａ縺ｮ謖�遉ｺ繝輔ぃ繧､繝ｫ
PLANNING_CONFLICT_SIDECAR = "planning_conflict_highlight.tsv"
# 驟榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙∈縲後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫舌阪ｒ譖ｸ縺丞�暦ｼ郁｡ｨ縺ｮ蜿ｳ遶ｯ繧医ｊ螟門�ｴ縲�1陦檎岼縺九ｉ邵ｦ縺ｫ繝ｩ繝吶Ν�ｼ丞､�ｼ�
# 笘� 蜿ら�ｧ陦ｨ遉ｺ縺ｮ縺ｿ: load_planning_tasks_df 遲峨�ｯ譛ｬ蛻励ｒ荳蛻�隱ｭ縺ｾ縺ｪ縺�縲る�榊床驕ｩ逕ｨ縺ｯ蟶ｸ縺ｫ繝｡繧､繝ｳ縲後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医�1邨瑚ｷｯ縺ｮ縺溘ａ莠碁㍾驕ｩ逕ｨ縺ｫ縺ｪ繧峨↑縺�縲�
PLAN_SHEET_GLOBAL_PARSE_LABEL_COL = 50  # AX
PLAN_SHEET_GLOBAL_PARSE_VALUE_COL = 51  # AY
PLAN_SHEET_GLOBAL_PARSE_MAX_ROWS = 42


def plan_reference_column_name(override_col: str) -> str:
    """荳頑嶌縺榊�励�ｮ蟾ｦ髫｣縺ｫ鄂ｮ縺丞盾辣ｧ蛻励�ｮ隕句�ｺ縺暦ｼ医そ繝ｫ蛟､縺ｯ諡ｬ蠑ｧ莉倥″縺ｧ蜈�繝�繝ｼ繧ｿ繧定｡ｨ遉ｺ�ｼ峨�"""
    return f"�ｼ亥���ｼ閲override_col}"


def plan_input_sheet_column_order():
    """
    驟榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙�ｮ蛻鈴���ｼ域ｮｵ髫�1蜃ｺ蜉帙�ｻ谿ｵ髫�2隱ｭ霎ｼ縺ｧ蜈ｱ騾夲ｼ峨�

    0. 驟榊床隧ｦ陦碁��逡ｪ�ｼ域ｮｵ髫�1謚ｽ蜃ｺ逶ｴ蠕後↓遨ｺ繧ｯ繝ｪ繧｢竊呈ｮｵ髫�2縺ｨ蜷瑚ｶ｣譌ｨ縺ｫ莉倅ｸ弱よｮｵ髫�2縺ｯ蜈ｨ陦後↓蛟､縺後≠繧九→縺阪％縺ｮ鬆�繧貞━蜈茨ｼ�
    1. 驟榊床荳崎ｦ��ｼ亥盾辣ｧ蛻励↑縺暦ｼ�
    2. 蜉�蟾･險育判DATA 逕ｱ譚･�ｼ�SOURCE_BASE_COLUMNS�ｼ俄ｦ 萓晞�ｼNO縲懷ｮ溷�ｺ譚･鬮倥∪縺ｧ�ｼ郁｣ｽ蜩∝錐縺ｮ逶ｴ蠕後↓繝ｭ繝ｼ繝ｫ蜊倅ｽ埼聞縺輔∝次蜿肴兜蜈･譌･縺ｮ逶ｴ蠕後↓蝨ｨ蠎ｫ蝣ｴ謇�ｼ�
    3. 蜉�蟾･蟾･遞九�ｮ豎ｺ螳壹�励Ο繧ｻ繧ｹ縺ｮ蝗�蟄�
    4. 荳頑嶌縺榊�冷ｦ 蜷�蛻励�ｮ逶ｴ蜑阪↓縲鯉ｼ亥���ｼ俄ｦ縲榊盾辣ｧ蛻励�AI迚ｹ蛻･謖�螳喟隗｣譫舌�ｮ縺ｿ蜿ら�ｧ蛻励↑縺励�
       �ｼ域律莉倡ｳｻ荳頑嶌縺阪↓ 蜴溷渚謚募�･譌･_荳頑嶌縺� 繧貞性繧縲らｩｺ逋ｽ譎ゅ�ｯ蛻励悟次蜿肴兜蜈･譌･縲阪ｒ驟榊床縺ｫ菴ｿ逕ｨ�ｼ�

    global_speed_rules 遲峨〒螟峨ｏ繧句ｮ溷柑騾溷ｺｦ縺ｯ繧ｷ繝ｼ繝亥�励〒縺ｯ謖√◆縺壹�驟榊床蜀�驛ｨ縺ｮ縺ｿ縺ｧ蜿肴丐縺吶ｋ縲�
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
    """蜿ら�ｧ陦ｨ遉ｺ逕ｨ: 遨ｺ縺ｯ�ｼ遺包ｼ峨∵律莉倥�ｻ縺昴�ｮ莉悶�ｯ�ｼ亥､�ｼ峨�"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "�ｼ遺包ｼ�"
    if isinstance(val, datetime):
        d = val.date() if hasattr(val, "date") else val
        if isinstance(d, date):
            return f"�ｼ�{d.year}/{d.month}/{d.day}�ｼ�"
    if isinstance(val, date):
        return f"�ｼ�{val.year}/{val.month}/{val.day}�ｼ�"
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none"):
        return "�ｼ遺包ｼ�"
    return f"�ｼ�{s}�ｼ�"


def _reference_text_for_override_row(row, override_col: str, req_map: dict, need_rules: list) -> str:
    """1陦悟��縺ｮ荳頑嶌縺榊�励↓蟇ｾ蠢懊☆繧句盾辣ｧ譁�險�ｼ域峡蠑ｧ莉倥″�ｼ峨�"""
    _ = (req_map, need_rules)  # 譌ｧ縲鯉ｼ亥���ｼ牙ｿ�隕∽ｺｺ謨ｰ縲榊盾辣ｧ縺ｧ菴ｿ逕ｨ縲ょ�怜ｻ�豁｢縺ｫ繧医ｊ譛ｪ菴ｿ逕ｨ縺�縺悟他縺ｳ蜃ｺ縺嶺ｺ呈鋤縺ｮ縺溘ａ谿九☆縲�
    if override_col == PLAN_COL_SPEED_OVERRIDE:
        v = row.get(TASK_COL_SPEED)
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "�ｼ遺包ｼ�"
        try:
            x = float(v)
            if abs(x - round(x)) < 1e-9:
                return f"�ｼ�{int(round(x))}�ｼ�"
            return f"�ｼ�{x}�ｼ�"
        except (TypeError, ValueError):
            return _format_paren_ref_scalar(v)
    if override_col in (PLAN_COL_PREFERRED_OP, PLAN_COL_SPECIAL_REMARK):
        return "�ｼ遺包ｼ�"
    if override_col == PLAN_COL_RAW_INPUT_DATE_OVERRIDE:
        return _format_paren_ref_scalar(
            parse_optional_date(_planning_df_cell_scalar(row, TASK_COL_RAW_INPUT_DATE))
        )
    return "�ｼ遺包ｼ�"


def _refresh_plan_reference_columns(df, req_map: dict, need_rules: list):
    """蜉�蟾･險育判DATA�ｼ熟eed 縺ｫ蝓ｺ縺･縺阪鯉ｼ亥���ｼ俄ｦ縲榊�励ｒ蜀崎ｨ育ｮ暦ｼ医�槭�ｼ繧ｸ蠕後↓蠢�縺壼他縺ｶ�ｼ峨�"""
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


def _apply_plan_input_visual_format(path: str, sheet_name: str = "繧ｿ繧ｹ繧ｯ荳隕ｧ"):
    """荳頑嶌縺榊�･蜉帛�励↓阮�縺�鮟�濶ｲ繧剃ｻ倅ｸ趣ｼ亥盾辣ｧ蛻励�ｯ譛ｪ逹濶ｲ縲�AI隗｣譫仙�励�ｯ髯､螟厄ｼ峨�"""
    # 隕句�ｺ縺玲枚蟄励�ｮ陦ｨ險倥ｆ繧後〒蛻怜錐讀懃ｴ｢縺ｫ螟ｱ謨励＠縺後■縺ｪ縺溘ａ縲∵ｮｵ髫�1縺ｮ蛻鈴���ｼ�plan_input_sheet_column_order�ｼ峨�ｮ
    # 1-based 蛻礼分蜿ｷ縺ｧ蝪励ｋ�ｼ�to_excel 縺ｮ蛻鈴��縺ｨ荳閾ｴ縺輔○繧具ｼ峨�
    fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    order = plan_input_sheet_column_order()
    col_1based = {name: i + 1 for i, name in enumerate(order)}
    if _workbook_should_skip_openpyxl_io(path):
        logging.info(
            "驟榊床險育判縺ｮ隕冶ｦ壽紛蠖｢: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ openpyxl 縺ｧ縺ｮ逹濶ｲ繧偵せ繧ｭ繝�繝励＠縺ｾ縺励◆縲�",
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
    Excel 縺後ヶ繝�繧ｯ繧帝幕縺�縺溘∪縺ｾ縺ｮ縺ｨ縺堺ｿ晏ｭ倥〒縺阪↑縺�蝣ｴ蜷医↓縲〃BA 逕ｨ縺ｮ TSV 繧� log 縺ｫ譖ｸ縺上�
    蠖｢蠑�: V1 / 繧ｷ繝ｼ繝亥錐 / 繝�繝ｼ繧ｿ陦梧焚 / 繧ｯ繝ｪ繧｢蛻励ｒ繧ｿ繝也ｵ仙粋 / 莉･髯� 陦檎分蜿ｷ\\t蛻怜錐
    """
    path = _planning_conflict_sidecar_path()
    clear_cols = "\t".join(PLAN_CONFLICT_STYLABLE_COLS)
    lines = ["V1", sheet_name, str(int(num_data_rows)), clear_cols]
    for r in sorted(conflicts_by_row.keys()):
        for name in sorted(conflicts_by_row[r]):
            lines.append(f"{int(r)}\t{name}")
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")

# 谿ｵ髫�1蜃ｺ蜉帙�ｻ繝悶ャ繧ｯ蜀�縺ｮ譌･莉伜�励ｒ Excel 荳翫梧律莉倥�ｮ縺ｿ縲�(譎ょ綾縺ｪ縺苓｡ｨ遉ｺ) 縺ｫ謨ｴ縺医ｋ
STAGE1_SHEET_DATEONLY_HEADERS = frozenset(
    {
        TASK_COL_ANSWER_DUE,
        TASK_COL_SPECIFIED_DUE,
        TASK_COL_RAW_INPUT_DATE,
        PLAN_COL_RAW_INPUT_DATE_OVERRIDE,
    }
)


def _result_font(**kwargs):
    """邨先棡繝悶ャ繧ｯ逕ｨ Font�ｼ亥他縺ｳ蜃ｺ縺怜�ｴ縺� name/size 遲峨ｒ謖�螳夲ｼ峨�"""
    return Font(**kwargs)


def _output_book_font(bold=False):
    return _result_font(bold=bold)


def _apply_output_font_to_result_sheet(ws):
    """邨先棡_* 縺ｮ縺�縺｡繧ｬ繝ｳ繝井ｻ･螟門髄縺�: 譌｢螳壹ヵ繧ｩ繝ｳ繝医�ｻ1陦檎岼螟ｪ蟄励�ｮ縺ｿ�ｼ亥�怜ｹ�縺ｯ VBA AutoFit�ｼ峨�"""
    base = _output_book_font(bold=False)
    hdr = _output_book_font(bold=True)
    mr, mc = ws.max_row or 1, ws.max_column or 1
    for row in ws.iter_rows(min_row=1, max_row=mr, min_col=1, max_col=mc):
        for cell in row:
            cell.font = base
    for cell in ws[1]:
        cell.font = hdr


def _apply_excel_date_columns_date_only_display(path, sheet_name, header_names=None):
    """openpyxl: 謖�螳壹�倥ャ繝繝ｼ蛻励ｒ yyyy/mm/dd 縺ｮ譌･莉倩｡ｨ遉ｺ縺ｫ縺吶ｋ�ｼ域凾蛻ｻ繧定｡ｨ遉ｺ縺励↑縺��ｼ峨�"""
    from openpyxl import load_workbook

    headers = header_names or STAGE1_SHEET_DATEONLY_HEADERS
    if _workbook_should_skip_openpyxl_io(path):
        logging.info(
            "譌･莉伜�苓｡ｨ遉ｺ謨ｴ蠖｢: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ openpyxl 縺ｧ縺ｮ蜃ｦ逅�繧偵せ繧ｭ繝�繝励＠縺ｾ縺励◆縲�",
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
    `蜉�蟾･險育判DATA` 繧ｷ繝ｼ繝医�ｮ `繝�繝ｼ繧ｿ謚ｽ蜃ｺ譌･` 縺九ｉ datetime 繧貞叙蠕励☆繧九�
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
    `蜉�蟾･險育判DATA` 繧ｷ繝ｼ繝医�ｮ `繝�繝ｼ繧ｿ謚ｽ蜃ｺ譌･` 縺九ｉ縲後ョ繝ｼ繧ｿ蜷ｸ蜃ｺ縺玲律譎ゅ阪ｒ蜿門ｾ励＠縺ｦ譁�蟄怜�怜喧縺吶ｋ縲�
    """
    try:
        dt = _extract_data_extraction_datetime()
        if dt is None:
            return "窶�"
        return dt.strftime("%Y/%m/%d %H:%M:%S")
    except Exception:
        return "窶�"


def _weekday_jp(d):
    return "譛育↓豌ｴ譛ｨ驥大悄譌･"[d.weekday()]


# 繧ｬ繝ｳ繝医�ｮ菴懈･ｭ繝舌�ｼ�ｼ壹＞縺壹ｌ繧よ�弱ｋ縺�蝨ｰ濶ｲ�ｼ矩ｻ呈枚蟄励′隱ｭ繧√ｋ繝医�ｼ繝ｳ�ｼ医Δ繝弱け繝ｭ蜊ｰ蛻ｷ縺ｧ繧よｿ�豺｡縺ｧ隴伜挨縺励ｄ縺吶＞�ｼ�
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

# 螳溽ｸｾ繝舌�ｼ逕ｨ�ｼ郁ｨ育判縺ｨ荳ｦ縺ｹ縺ｦ繧ゅΔ繝弱け繝ｭ縺ｧ蛹ｺ蛻･縺励ｄ縺吶＞繝医�ｼ繝ｳ�ｼ�
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

# 險ｭ蛯吶ぎ繝ｳ繝�: 譌･谺｡蟋区･ｭ貅門ｙ�ｼ�machine_daily_startup�ｼ峨�ｮ蟶ｯ濶ｲ�ｼ磯ｻ�濶ｲ邉ｻ�ｼ�
_GANTT_DAILY_STARTUP_FILL = "FFEB9C"


def _gantt_bar_fill_for_task_id(task_id):
    """萓晞�ｼNO縺斐→縺ｫ荳願ｨ倥ヱ繝ｬ繝�繝医°繧�1濶ｲ�ｼ�RRGGBB�ｼ峨よｿ�濶ｲ�ｼ狗區譁�蟄励�ｮ邨�縺ｿ蜷医ｏ縺帙�ｯ菴ｿ繧上↑縺�縲�"""
    h = hashlib.md5(str(task_id).encode("utf-8")).hexdigest()
    i = int(h[:8], 16) % len(_GANTT_BAR_FILLS_PRINT_SAFE)
    return _GANTT_BAR_FILLS_PRINT_SAFE[i]


def _gantt_bar_fill_actual_for_task_id(task_id):
    h = hashlib.md5(str(task_id).encode("utf-8")).hexdigest()
    i = int(h[:8], 16) % len(_GANTT_BAR_FILLS_ACTUAL)
    return _GANTT_BAR_FILLS_ACTUAL[i]


# 繧ｬ繝ｳ繝域凾蛻ｻ繧ｻ繝ｫ�ｼ育ｵ仙粋蟶ｯ縺ｮ蜈磯�ｭ繧ｻ繝ｫ�ｼ�: 豈弱そ繧ｰ繝｡繝ｳ繝� new 縺励↑縺�
_GANTT_TIMELINE_CELL_ALIGNMENT = Alignment(
    horizontal="left",
    vertical="center",
    wrap_text=False,
    shrink_to_fit=False,
    indent=1,
)
# 繧ｿ繧ｹ繧ｯ蟶ｯ縺ｮ濶ｲ縺ｯ繝代Ξ繝�繝域怏髯舌↑縺ｮ縺ｧ PatternFill 繧� hex 蜊倅ｽ阪〒蜈ｱ譛会ｼ�openpyxl 縺ｮ繧ｹ繧ｿ繧､繝ｫ螻暮幕繧ｳ繧ｹ繝亥炎貂幢ｼ�
_GANTT_TASK_PATTERN_FILL_BY_HEX: dict[str, PatternFill] = {}


def _gantt_cached_pattern_fill(hex_rrggbb: str) -> PatternFill:
    fi = _GANTT_TASK_PATTERN_FILL_BY_HEX.get(hex_rrggbb)
    if fi is None:
        fi = PatternFill(fill_type="solid", start_color=hex_rrggbb, end_color=hex_rrggbb)
        _GANTT_TASK_PATTERN_FILL_BY_HEX[hex_rrggbb] = fi
    return fi


def _gantt_slot_state_tuple(evlist, slot_mid, task_fill_fn=None):
    """繧ｹ繝ｭ繝�繝井ｸｭ螟ｮ譎ょ綾縺ｫ縺翫￠繧�1繝槭せ蛻�縺ｮ迥ｶ諷九�('idle',) | ('break',) | ('daily_startup', fill_hex) | ('task', tid, fill_hex, pct)"""
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
        # 縲後�槭け繝ｭ螳溯｡梧凾轤ｹ縲阪�ｮ螳御ｺ�邇�繧貞━蜈茨ｼ�pct_macro 繧� timeline_event 縺ｫ謖√◆縺帙ｋ�ｼ�
        if active.get("pct_macro") is not None:
            pct = int(round(parse_float_safe(active.get("pct_macro"), 0.0)))
            pct = max(0, min(100, pct))
        else:
            # 繝輔ぉ繧､繝ｫ繧ｻ繝ｼ繝包ｼ亥ｾ捺擂縺ｮ謫ｬ莨ｼ騾ｲ謐苓ｨ育ｮ暦ｼ�
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
    """邨仙粋繧ｻ繧ｰ繝｡繝ｳ繝亥｢�逡悟愛螳夲ｼ域ｯ弱せ繝ｭ繝�繝� tuple 繧貞牡繧雁ｽ薙※縺ｪ縺��ｼ峨�"""
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
    譎る俣霆ｸ繧貞｡励ｊ蛻�縺代◆縺�縺医〒縲∝酔荳迥ｶ諷九′騾｣邯壹☆繧九そ繝ｫ繧呈ｨｪ邨仙粋縺怜ｸｯ迥ｶ縺ｮ繝舌�ｼ縺ｫ縺吶ｋ縲�
    �ｼ育ｴｰ繝槭せ蜊倅ｽ薙�ｮ蝪励ｊ縺ｧ縺ｯ縺ｪ縺�15蛻�蛻ｻ縺ｿ�ｼ句酔荳迥ｶ諷九�ｮ繧ｻ繝ｫ邨仙粋縺ｧ縲∝ｸｯ迥ｶ縺ｮ繝舌�ｼ縺ｨ縺励※陦ｨ迴ｾ縺吶ｋ�ｼ�
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
                    c.value = "(譌･谺｡蟋区･ｭ貅門ｙ)"
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
    """蜊企幕蛹ｺ髢� [a_start, a_end) 縺ｨ [b_start, b_end) 縺碁㍾縺ｪ繧九°�ｼ亥酔荳譌･蜀��ｼ峨�"""

    def _sec(t: time) -> int:
        return t.hour * 3600 + t.minute * 60 + t.second

    return _sec(a_start) < _sec(b_end) and _sec(a_end) > _sec(b_start)


def _parse_equipment_schedule_time_band_cell(val) -> tuple[time | None, time | None]:
    """邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ縲梧律譎ょｸｯ縲阪そ繝ｫ�ｼ井ｾ� 08:45-09:00�ｼ峨ｒ隗｣驥医�"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None, None
    s = str(val).strip()
    if not s or "笆�" in s:
        return None, None
    for sep in ("-", "�ｼ�", "~", "縲�"):
        if sep in s:
            left, right = s.split(sep, 1)
            left = left.strip().replace("�ｼ�", ":")
            right = right.strip().replace("�ｼ�", ":")
            t0 = parse_time_str(left, None)
            t1 = parse_time_str(right, None)
            if t0 is not None and t1 is not None and t0 < t1:
                return t0, t1
            return None, None
    return None, None


def _apply_equipment_schedule_outside_regular_fill(
    ws, reg_start: time, reg_end: time
) -> None:
    """縲梧律譎ょｸｯ縲榊�励〒螳壼ｸｸ [reg_start, reg_end) 縺ｨ驥阪↑繧峨↑縺�陦後�ｮ繧ｻ繝ｫ縺ｫ逹濶ｲ縲�"""
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_OUTSIDE_REGULAR_TIME_FILL,
        end_color=RESULT_OUTSIDE_REGULAR_TIME_FILL,
    )
    col_idx = None
    for i, c in enumerate(ws[1], start=1):
        if c.value is not None and str(c.value).strip() == "譌･譎ょｸｯ":
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
    險ｭ蛯吝�暦ｼ磯ｲ蠎ｦ蛻励ｒ髯､縺擾ｼ峨〒縲∬｡ｨ遉ｺ縺ｫ縲梧律谺｡蟋区･ｭ貅門ｙ縲阪悟刈蟾･蜑肴ｺ門ｙ縲阪御ｾ晞�ｼ蛻�譖ｿ蠕悟ｧ区忰縲阪′蜷ｫ縺ｾ繧後ｋ繧ｻ繝ｫ繧定埋邱代↓縺吶ｋ縲�
    邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ / TEMP_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ 縺ｮ equip 繧ｻ繝ｫ逕ｨ�ｼ域律譎ょｸｯ蛻励�ｯ螟画峩縺励↑縺��ｼ峨�
    """
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_DISPATCHED_REQUEST_FILL,
        end_color=RESULT_DISPATCHED_REQUEST_FILL,
    )
    markers = ("(譌･谺｡蟋区･ｭ貅門ｙ)", "(蜉�蟾･蜑肴ｺ門ｙ)", "(萓晞�ｼ蛻�譖ｿ蠕悟ｧ区忰)")
    col_tb = None
    equip_cols: list[int] = []
    for i, c in enumerate(ws[1], start=1):
        if c.value is None:
            continue
        h = str(c.value).strip()
        if h == "譌･譎ょｸｯ":
            col_tb = i
            continue
        if h.endswith("騾ｲ蠎ｦ"):
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
    """譌･莉倩ｦ句�ｺ縺苓｡後娯蔓 YYYY/MM/DD 窶ｦ 笆�縲阪°繧画律莉倥ｒ蜿悶ｋ縲�"""
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
    """蠖捺律繝ｻ蠖楢ｩｲ險ｭ蛯吝�励く繝ｼ縺ｫ蟇ｾ蠢懊☆繧区ｩ滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ蜊�譛牙玄髢難ｼ亥ｷ･蝣ｴ遞ｼ蜒肴棧縺ｧ繧ｯ繝ｪ繝�繝玲ｸ医∩�ｼ峨�"""
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
    邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ: 讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ蜊�譛峨→驥阪↑繧玖ｨｭ蛯吶そ繝ｫ�ｼ磯ｲ蠎ｦ蛻嶺ｻ･螟厄ｼ峨ｒ邏ｫ濶ｲ縺ｧ蝪励ｋ縲�
    10 蛻�譫�縺ｮ蜊企幕蛹ｺ髢� [slot_start, slot_end) 縺ｨ蜊�譛� [bs, be) 縺碁㍾縺ｪ繧後�ｰ蟇ｾ雎｡縲�
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
        if c.value is not None and str(c.value).strip() == "譌･譎ょｸｯ":
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
    邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ_讖滓｢ｰ蜷肴ｯ弱�ｮ讖滓｢ｰ蜷榊�励〒縲∽ｾ晞�ｼNO縺悟�･縺｣縺ｦ縺�繧九そ繝ｫ縺ｫ阮�邱代ｒ莉倅ｸ弱☆繧九�
    縲鯉ｼ井ｼ第�ｩ�ｼ峨阪�ｮ縺ｿ縺ｮ繧ｻ繝ｫ縺ｯ蟇ｾ雎｡螟悶りｦ句�ｺ縺苓｡後�ｻ譌･譎ょｸｯ蛻励�ｯ螟画峩縺励↑縺�縲�
    """
    fill = PatternFill(
        fill_type="solid",
        start_color=RESULT_DISPATCHED_REQUEST_FILL,
        end_color=RESULT_DISPATCHED_REQUEST_FILL,
    )
    col_tb = None
    for i, c in enumerate(ws[1], start=1):
        if c.value is not None and str(c.value).strip() == "譌･譎ょｸｯ":
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
            if not s or s == "�ｼ井ｼ第�ｩ�ｼ�":
                continue
            cell.fill = fill


def _equipment_gantt_fills_by_machine_name(equipment_list) -> dict[str, PatternFill]:
    """
    邨先棡_險ｭ蛯吶ぎ繝ｳ繝医�ｮ蝗ｺ螳壼�暦ｼ�B縲廢縲、 縺ｯ譌･莉倡ｸｦ邨仙粋�ｼ臥畑縲Ｆquipment_list 蜀�縺ｮ讖滓｢ｰ蜷搾ｼ�+ 辟｡縺玲凾縺ｯ陦悟�ｨ菴薙ｒ讖滓｢ｰ蜷搾ｼ峨�ｮ蜃ｺ迴ｾ鬆�縺ｧ
    豺｡濶ｲ繧貞牡繧雁ｽ薙※縲∝酔荳讖滓｢ｰ蜷阪�ｯ蟶ｸ縺ｫ蜷後§ PatternFill 繧貞�ｱ譛峨☆繧九�
    """
    order: list[str] = []
    seen: set[str] = set()
    for eq in equipment_list or []:
        _, mn = _split_equipment_line_process_machine(eq)
        key = (mn or "").strip() or "窶�"
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
    邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ縺ｨ蜷御ｸ繝�繝ｼ繧ｿ貅撰ｼ�timeline_events�ｼ峨↓蝓ｺ縺･縺阪�
    險ｭ蛯凖玲ｨｪ霆ｸ譎る俣縺ｮ繧ｬ繝ｳ繝√Ε繝ｼ繝磯｢ｨ繧ｷ繝ｼ繝医ｒ霑ｽ蜉�縺吶ｋ縲�
    讓ｪ霆ｸ縺ｯ15蛻�蛻ｻ縺ｿ縲る｣邯壹☆繧句酔荳繧ｿ繧ｹ繧ｯ�ｼ丈ｼ第�ｩ�ｼ冗ｩｺ縺阪�ｯ繧ｻ繝ｫ邨仙粋縺励※蟶ｯ迥ｶ縺ｫ陦ｨ遉ｺ縺吶ｋ縲�
    actual_timeline_events 縺後≠繧後�ｰ險ｭ蛯吶＃縺ｨ縺ｫ縲悟ｮ溽ｸｾ縲崎｡後ｒ險育判陦後�ｮ荳九∈霑ｽ蜉�縺吶ｋ縲�
    """
    wb = writer.book
    try:
        insert_at = wb.sheetnames.index("邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ") + 1
    except ValueError:
        insert_at = len(wb.sheetnames)
    ws = wb.create_sheet("邨先棡_險ｭ蛯吶ぎ繝ｳ繝�", insert_at)
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

    # 讓ｪ霆ｸ(10蛻�蛻ｻ縺ｿ)縺ｯ譌･莉倥〒蜈ｱ騾壹�ｮ縺溘ａ縲《lot_times 繧貞�医↓遒ｺ螳�
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
    n_fixed = 5  # A=譌･莉假ｼ域律繝悶Ο繝�繧ｯ蜀�縺ｧ邵ｦ邨仙粋�ｼ�/ B縲廢=讖滓｢ｰ蜷阪�ｻ蟾･遞句錐繝ｻ諡�蠖楢�繝ｻ繧ｿ繧ｹ繧ｯ讎りｦ�
    last_col = n_fixed + n_slots
    fills_by_mach = _equipment_gantt_fills_by_machine_name(equipment_list)
    fb_gantt = "F5F5F5"
    fill_gantt_fallback = PatternFill(fill_type="solid", start_color=fb_gantt, end_color=fb_gantt)

    # 繧ｿ繧､繝医Ν�ｼ�譌･譎ゑｼ医�壹�ｼ繧ｸ荳企Κ�ｼ�
    create_ts = base_dt.strftime("%Y/%m/%d %H:%M:%S")
    master_path = os.path.join(os.getcwd(), MASTER_FILE) if MASTER_FILE else ""

    def _fmt_mtime(p):
        try:
            if p and os.path.exists(p):
                return datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y/%m/%d %H:%M:%S")
        except Exception:
            pass
        return "窶�"

    master_mtime = _fmt_mtime(master_path)

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    tcell = ws.cell(row=row, column=1, value="貉門漉蟾･蝣ｴ 蜉�蟾･險育判")
    tcell.font = title_font
    tcell.fill = title_fill
    # 邨仙粋繧ｻ繝ｫ縺ｧ繧ょｷｦ遶ｯ縺九ｉ陦ｨ遉ｺ�ｼ育ｸｮ蟆上�ｻ謚倥ｊ霑斐＠縺ｪ縺暦ｼ�
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
        f"菴懈�舌{create_ts}"
        f"縲繝ｻ縲繝�繝ｼ繧ｿ蜷ｸ蜃ｺ縺励{data_extract_dt_str or '窶�'}"
        f"縲繝ｻ縲繝槭せ繧ｿ�ｼ�master.xlsm�ｼ峨{master_mtime}"
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
    fixed_hdr = ["譌･莉�", "讖滓｢ｰ蜷�", "蟾･遞句錐", "諡�蠖楢�", "繧ｿ繧ｹ繧ｯ讎りｦ�"]
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
    # 蜈磯�ｭ繝�繝ｼ繧ｿ陦後�ｮ蟾ｦ荳奇ｼ晄凾蛻ｻ蛻怜�磯�ｭ�ｼ�F4�ｼ峨〒遯捺棧蝗ｺ螳夲ｼ郁｡�1縲�3繝ｻ蛻輸縲廢繧貞崋螳夲ｼ�
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
            mk_key = (mach_nm or "").strip() or "窶�"
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
                task_sum = " ".join(tids) if tids else "窶�"
                member_disp = _gantt_row_member_names(evlist)
            else:
                task_sum = "窶�"
                member_disp = "窶�"

            c1 = ws.cell(row=row, column=2, value=mach_nm if mach_nm else "窶�")
            c2 = ws.cell(row=row, column=3, value=proc_nm if proc_nm else "窶�")
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
                    task_sum_a = " ".join(tids_a) if tids_a else "窶�"
                    member_disp_a = _gantt_row_member_names(evlist_a)
                else:
                    task_sum_a = "窶�"
                    member_disp_a = "窶�"

                lab_fill_a = fills_by_mach.get(mk_key) or fill_gantt_fallback

                if mach_nm:
                    act_mach = f"{mach_nm}�ｼ亥ｮ溽ｸｾ�ｼ�"
                elif proc_nm:
                    act_mach = "�ｼ亥ｮ溽ｸｾ�ｼ�"
                else:
                    act_mach = "窶�"
                ca1 = ws.cell(row=row, column=2, value=act_mach)
                ca2 = ws.cell(row=row, column=3, value=proc_nm if proc_nm else "窶�")
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
                value=f"縲須d.strftime('%Y/%m/%d')}縲�",
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

    # 蜃｡萓九�ｯ鬮倥＆遒ｺ菫昴�ｮ縺溘ａ逵∫払�ｼ医Δ繝弱け繝ｭ蜊ｰ蛻ｷ縺ｯ濶ｲ縺ｮ豼�豺｡/繧ｻ繝ｫ縺ｮ譫�縺ｧ隴伜挨�ｼ�
    # 蛻怜ｹ�繝ｻ謚倥ｊ霑斐＠縺ｯ VBA 蜿悶ｊ霎ｼ縺ｿ譎ゑｼ育ｵ先棡_險ｭ蛯吶ぎ繝ｳ繝�_蛻怜ｹ�繧定ｨｭ螳夲ｼ峨〒險ｭ螳�

    try:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToHeight = False
        ws.page_setup.fitToWidth = 1
        # A3�ｼ�openpyxl 荳翫〒 paperSize=8 縺� A3 逶ｸ蠖難ｼ�
        ws.page_setup.paperSize = 8
        # 菴咏區繧堤強繧√※讓ｪ1繝壹�ｼ繧ｸ縺ｫ蜿弱∪繧翫ｄ縺吶￥縺吶ｋ�ｼ亥腰菴�: 繧､繝ｳ繝��ｼ�
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.2
        ws.page_margins.bottom = 0.2
        # 繧ｿ繧､繝医Ν繝ｻ陦ｨ繧偵�壹�ｼ繧ｸ蟾ｦ蝓ｺ貅悶↓�ｼ医Ξ繝昴�ｼ繝磯｢ｨ�ｼ�
        ws.print_options.horizontalCentered = False
        ws.print_options.verticalCentered = False
        ws.print_options.gridLines = False
    except Exception:
        pass


def row_has_completion_keyword(row):
    """蜉�蟾･螳御ｺ�蛹ｺ蛻�縺ｫ縲悟ｮ御ｺ�縲阪�ｮ譁�蟄励′蜷ｫ縺ｾ繧後ｋ蝣ｴ蜷医�ｯ繧ｿ繧ｹ繧ｯ螳御ｺ�縺ｨ縺ｿ縺ｪ縺吶�"""
    v = row.get(TASK_COL_COMPLETION_FLAG)
    if v is None or pd.isna(v):
        return False
    return "螳御ｺ�" in str(v)


def _plan_row_exclude_from_assignment(row) -> bool:
    """
    縲碁�榊床荳崎ｦ√榊�励′繧ｪ繝ｳ縺ｪ繧峨√◎縺ｮ陦後�ｯ驟榊床繧ｭ繝･繝ｼ縺ｸ蜈･繧後★縲∫音蛻･謖�螳喟蛯呵�縺ｮ AI 隗｣譫占｡後°繧峨ｂ髯､縺上�

    驟榊床縺九ｉ螟悶☆�ｼ育悄�ｼ�: 隲也炊蛟､ True縲∵焚蛟､ 1縲∵枚蟄怜�暦ｼ�NFKC 蠕後�ｻ蟆乗枚蟄暦ｼ�
      true / 1 / yes / on / y / t / 縺ｯ縺� / 笳� / 縲� / 笳�
    驟榊床蟇ｾ雎｡�ｼ亥⊃�ｼ�: 遨ｺ縲¨one縲：alse縲�0縲］o / off / false / 縺�縺�縺� / 蜷ｦ 遲�
    荳願ｨ倅ｻ･螟悶�ｮ譁�蟄怜�励�ｯ蛛ｽ�ｼ磯�榊床縺吶ｋ�ｼ峨ゅメ繧ｧ繝�繧ｯ繝懊ャ繧ｯ繧ｹ騾｣蜍輔そ繝ｫ縺ｯ騾壼ｸｸ TRUE/FALSE 縺ｾ縺溘�ｯ 1/0縲�
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
    if not s or s in ("nan", "none", "false", "0", "no", "off", "縺�縺�縺�", "蜷ｦ"):
        return False
    if s in ("true", "1", "yes", "on", "縺ｯ縺�", "y", "t", "笳�", "縲�", "笳�"):
        return True
    return False


def _coerce_plan_exclude_column_value_for_storage(v):
    """
    縲碁�榊床荳崎ｦ√榊�励∈譖ｸ縺崎ｾｼ繧蛟､繧偵ヾtringDtype 蛻励〒繧ゆｻ｣蜈･繧ｨ繝ｩ繝ｼ縺ｫ縺ｪ繧峨↑縺�蠖｢縺ｫ縺昴ｍ縺医ｋ縲�
    Excel 蜿悶ｊ霎ｼ縺ｿ縺ｮ True / 1 / False / 0 縺ｨ譁�蟄怜�励ｒ菫晄戟縺励＼plan_row_exclude_from_assignment 縺ｨ謨ｴ蜷医☆繧九�
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
    蜉�蟾･貂域焚驥擾ｼ亥ｷ･遞区兜蜈･驥乗鋤邂暦ｼ峨ｒ霑斐☆縲�

    蝓ｺ譛ｬ蠑�:
      螳溷�ｺ譚･鬮� ﾃｷ (蜿玲ｳｨ謨ｰ ﾃｷ 謠帷ｮ玲焚驥�)
    = 螳溷�ｺ譚･鬮� * 謠帷ｮ玲焚驥� / 蜿玲ｳｨ謨ｰ

    蜿玲ｳｨ謨ｰ縺檎┌縺�/荳肴ｭ｣縺ｪ蝣ｴ蜷医�ｯ縲∵立蛻励悟ｮ溷刈蟾･謨ｰ縲阪ｒ莠呈鋤繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ縺ｨ縺励※菴ｿ縺�縲�
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
    iterrows() 1陦悟��縺九ｉ蛻怜､繧貞叙繧九ょ酔荳隕句�ｺ縺励�ｮ驥崎､�蛻励′縺ゅｋ縺ｨ row.get 縺ｯ Series 縺ｫ縺ｪ繧翫�
    str竊稚o_datetime 縺ｧ隱､縺｣縺滓律莉倥↓縺ｪ繧九％縺ｨ縺後≠繧九◆繧√∝�磯�ｭ縺ｮ髱樊ｬ�謳阪せ繧ｫ繝ｩ繝ｼ繧定ｿ斐☆縲�
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
                    # 譛滄剞蛻�繧後お繝ｳ繝医Μ繧帝勁蜴ｻ�ｼ�6譎る俣�ｼ�
                    now_ts = time_module.time()
                    cleaned = {}
                    expired_count = 0
                    for k, v in data.items():
                        # 譁ｰ蠖｢蠑�: {"ts": epoch_seconds, "data": {...}}
                        if isinstance(v, dict) and "ts" in v and "data" in v:
                            ts = parse_float_safe(v.get("ts"), 0.0)
                            if ts > 0 and (now_ts - ts) <= AI_CACHE_TTL_SECONDS:
                                cleaned[k] = v
                            else:
                                expired_count += 1
                        # 譌ｧ蠖｢蠑�: 蛟､縺檎峩謗･AI邨先棡dict�ｼ井ｺ呈鋤縺ｧ隱ｭ縺ｿ蜿悶ｊ縲∝叉譎ゅ↓譁ｰ蠖｢蠑上∈蜀堺ｿ晏ｭ倥＆繧後ｋ�ｼ�
                        else:
                            cleaned[k] = {"ts": now_ts, "data": v}
                    if expired_count > 0:
                        logging.info(f"AI繧ｭ繝｣繝�繧ｷ繝･譛滄剞蛻�繧後ｒ蜑企勁: {expired_count}莉ｶ")
                    return cleaned
    except Exception as e:
        logging.warning(f"AI繧ｭ繝｣繝�繧ｷ繝･隱ｭ縺ｿ霎ｼ縺ｿ螟ｱ謨�: {e}")
    return {}

def save_ai_cache(cache_obj):
    try:
        with open(ai_cache_path, "w", encoding="utf-8") as f:
            json.dump(cache_obj, f, ensure_ascii=False)
    except Exception as e:
        logging.warning(f"AI繧ｭ繝｣繝�繧ｷ繝･菫晏ｭ伜､ｱ謨�: {e}")

def get_cached_ai_result(cache_obj, cache_key, content_key=None):
    """
    content_key: 繧ｪ繝励す繝ｧ繝ｳ縲ゆｿ晏ｭ俶凾縺ｨ蜷御ｸ縺ｮ譁�蟄怜�励〒縺ｪ縺�繝偵ャ繝医�ｯ辟｡蜉ｹ蛹悶☆繧具ｼ育音蛻･謖�螳壹�ｻ辣ｧ蜷育畑縺ｮ莠梧ｬ｡繝√ぉ繝�繧ｯ�ｼ峨�
    譌ｧ繧ｨ繝ｳ繝医Μ縺ｫ content_key 縺檎┌縺�蝣ｴ蜷医�ｯ SHA256 繧ｭ繝ｼ荳閾ｴ縺ｮ縺ｿ縺ｧ蠕捺擂縺ｩ縺翫ｊ繝偵ャ繝医→縺ｿ縺ｪ縺吶�
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
                "AI繧ｭ繝｣繝�繧ｷ繝･: 繧ｭ繝ｼ縺ｯ荳閾ｴ縺励∪縺吶′ content_key 縺檎樟陦悟�･蜉帙→逡ｰ縺ｪ繧九◆繧∫┌蜉ｹ蛹悶＠縺ｾ縺吶�"
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
    # 萓�: "Please retry in 57.089735313s."
    m = re.search(r"retry in ([0-9]+(?:\.[0-9]+)?)s", err_text, re.IGNORECASE)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    # 萓�: "'retryDelay': '57s'"
    m = re.search(r"retryDelay'\s*:\s*'([0-9]+)s'", err_text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return None


def infer_unit_m_from_product_name(product_name, fallback_unit):
    """
    陬ｽ蜩∝錐譁�蟄怜�励°繧牙刈蟾･蜊倅ｽ�(m)繧呈耳螳壹☆繧区圻螳壹Ν繝ｼ繝ｫ縲�
    萓�: 15020-JX5R- 770X300F-A   R -> 300
    窶ｻ 繝舌Μ繧ｨ繝ｼ繧ｷ繝ｧ繝ｳ縺悟､壹＞蜑肴署縺ｮ縺溘ａ縲√％縺薙ｒ驛ｽ蠎ｦ隱ｿ謨ｴ縺ｧ縺阪ｋ繧医≧髢｢謨ｰ蛹悶＠縺ｦ縺�繧九�
    """
    if product_name is None or pd.isna(product_name):
        return fallback_unit
    s = str(product_name)
    # "770X300..." 縺ｮ繧医≧縺ｪ繝代ち繝ｼ繝ｳ縺九ｉ X 縺ｮ蠕後�ｮ謨ｰ蛟､繧呈鏡縺��ｼ域怙蠕後↓隕九▽縺九▲縺盜繧貞━蜈茨ｼ�
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
    繧ｿ繧ｹ繧ｯ蜈･蜉帙ｒ蜿門ｾ励☆繧具ｼ�tasks.xlsx 縺ｯ菴ｿ逕ｨ縺励↑縺��ｼ峨�
    蠢�鬆�: 迺ｰ蠅�螟画焚 TASK_INPUT_WORKBOOK 縺ｫ繝槭け繝ｭ螳溯｡後ヶ繝�繧ｯ縺ｮ繝輔Ν繝代せ�ｼ�VBA 縺瑚ｨｭ螳夲ｼ�
         繧ｷ繝ｼ繝医悟刈蟾･險育判DATA縲阪ｒ隱ｭ縺ｿ霎ｼ繧�ｼ域兜蜈･逶ｮ螳峨�ｯ縲悟屓遲皮ｴ肴悄縲阪∵悴蜈･蜉帶凾縺ｯ縲梧欠螳夂ｴ肴悄縲搾ｼ峨�
    """
    if not TASKS_INPUT_WORKBOOK:
        raise FileNotFoundError(
            "TASK_INPUT_WORKBOOK 縺梧悴險ｭ螳壹〒縺吶７BA 縺ｮ RunPython 縺ｧ繝槭け繝ｭ螳溯｡後ヶ繝�繧ｯ縺ｮ繝代せ繧呈ｸ｡縺励※縺上□縺輔＞縲�"
        )
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        raise FileNotFoundError(f"TASK_INPUT_WORKBOOK 縺悟ｭ伜惠縺励∪縺帙ｓ: {TASKS_INPUT_WORKBOOK}")
    df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=TASKS_SHEET_NAME)
    df.columns = df.columns.str.strip()
    logging.info(f"繧ｿ繧ｹ繧ｯ蜈･蜉�: '{TASKS_INPUT_WORKBOOK}' 縺ｮ '{TASKS_SHEET_NAME}' 繧定ｪｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�")
    return df


def _nfkc_column_aliases(canonical_name):
    """隕句�ｺ縺励�ｮ陦ｨ險倥ｆ繧鯉ｼ亥�ｨ隗定ｨ伜捷繝ｻ莠呈鋤譁�蟄暦ｼ峨ｒ蜷ｸ蜿弱☆繧九◆繧√�ｮ豈碑ｼ�繧ｭ繝ｼ縲�"""
    return unicodedata.normalize("NFKC", str(canonical_name).strip())


def _align_dataframe_headers_to_canonical(df, canonical_names):
    """蛻怜錐繧� NFKC 荳閾ｴ縺ｧ canonical 縺ｫ蟇�縺帙ｋ�ｼ�Excel 蛛ｴ縺悟�ｨ隗� '_' 遲峨〒繧りｪｭ繧√ｋ繧医≧縺ｫ�ｼ峨�"""
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
    蟾･遞句錐�ｼ郁ｨｭ蛯吝錐�ｼ峨�ｮ辣ｧ蜷育畑繧ｭ繝ｼ縲�
    NFKC繝ｻ蜑榊ｾ檎ｩｺ逋ｽ繝ｻ騾｣邯夂ｩｺ逋ｽ繝ｻNBSP/蜈ｨ隗偵せ繝壹�ｼ繧ｹ繝ｻ繧ｼ繝ｭ蟷�譁�蟄励ｒ豁｣隕丞喧縺吶ｋ縲�
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    t = unicodedata.normalize("NFKC", str(val))
    t = t.replace("\u00a0", " ").replace("\u3000", " ")
    t = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _equipment_line_key_to_physical_occupancy_key(eq_line: str) -> str:
    """險ｭ蛯吝�励く繝ｼ�ｼ亥ｷ･遞�+讖滓｢ｰ 遲会ｼ峨°繧峨∫黄逅�讖滓｢ｰ縺ｮ蜊�譛峨↓逕ｨ縺�繧九く繝ｼ�ｼ域ｩ滓｢ｰ蜷榊�ｴ繝ｻ豁｣隕丞喧�ｼ峨ｒ蠕励ｋ縲�"""
    s = str(eq_line or "").strip()
    if not s:
        return ""
    nk = _normalize_equipment_match_key(s)
    if "+" in nk:
        return _normalize_equipment_match_key(nk.split("+", 1)[1])
    return nk


def _physical_machine_occupancy_key_for_task(task: dict) -> str:
    """
    險ｭ蛯吶�ｮ螢∵凾險亥頃譛会ｼ�machine_avail_dt繝ｻ髢馴囈繝溘Λ繝ｼ�ｼ峨↓逕ｨ縺�繧九く繝ｼ縲�
    讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ蛻励�ｯ equipment_line_key 縺ｮ縲悟ｷ･遞�+讖滓｢ｰ縲阪→荳閾ｴ縺吶ｋ縺溘ａ縲�
    豁｣隕丞喧蠕後↓縲�+縲阪ｒ蜷ｫ繧縺ｨ縺阪�ｯ **machine_name 繧医ｊ蜈医↓** 縺昴％縺九ｉ迚ｩ逅�讖滓｢ｰ蜷阪ｒ謗｡逕ｨ縺吶ｋ縲�
    �ｼ�machine_name 縺ｫ蟾･遞句錐縺ｮ縺ｿ縺ｪ縺ｩ縺悟�･繧翫∝ｺ翫く繝ｼ縲檎�ｱ陞咲捩讖� 貉門漉縲阪→縺壹ｌ縺ｦ蛟呵｣懷､悶＠貍上ｌ縺吶ｋ縺ｮ繧帝亟縺撰ｼ�
    蜊倅ｸ蜷阪�ｮ縺ｨ縺阪�ｯ蠕捺擂縺ｩ縺翫ｊ machine_name 繧貞━蜈医＠縲∫┌縺代ｌ縺ｰ equipment_line_key / machine 縺九ｉ謗ｨ螳壹☆繧九�
    machine_name 縺ｫ縲悟ｷ･遞�+讖滓｢ｰ縲阪→蜈･縺｣縺ｦ縺�繧句�ｴ蜷医〒繧ゅ∝頃譛峨�ｯ迚ｩ逅�讖滓｢ｰ蜷搾ｼ�+ 縺ｮ蜿ｳ蛛ｴ�ｼ峨↓蟇�縺帙ｋ縲�
    蜈ｨ隗偵鯉ｼ九阪�ｮ縺ｿ縺ｮ蛻励�ｯ NFKC 蠕後↓蜊願ｧ偵�+縲阪↓縺ｪ繧九◆繧√∝��蜑ｲ蛻､螳壹�ｯ豁｣隕丞喧蠕後↓陦後≧縲�
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
    machine_avail_dt繝ｻ讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ蠎翫→謨ｴ蜷医☆繧句頃譛峨く繝ｼ�ｼ亥次蜑�: 迚ｩ逅�讖滓｢ｰ蜷搾ｼ峨�
    task 縺九ｉ蜿悶ｌ縺ｪ縺�縺ｨ縺阪�ｯ eq_line�ｼ亥ｷ･遞�+讖滓｢ｰ�ｼ峨°繧画ｩ滓｢ｰ蜷榊�ｴ繧呈耳螳壹＠縲∵怙蠕後�ｮ謇区ｮｵ縺ｧ eq_line縲�
    縲娯ｦ or eq_line縲阪↓繧医ｋ蟾･遞�+讖滓｢ｰ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ縺ｯ讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ迚ｩ逅�繧ｭ繝ｼ縺ｨ荳堺ｸ閾ｴ縺ｫ縺ｪ繧雁ｾ励ｋ縺溘ａ遖∵ｭ｢縲�
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
    """豁｣隕丞喧繧ｭ繝ｼ 竊� master 繧ｹ繧ｭ繝ｫ繧ｷ繝ｼ繝井ｸ翫�ｮ蛻怜錐�ｼ�canonical 陦ｨ險假ｼ峨�"""
    lookup = {}
    for eq in equipment_list:
        k = _normalize_equipment_match_key(eq)
        if k and k not in lookup:
            lookup[k] = eq
    # 蟾･遞句錐縺ｮ縺ｿ縺ｮ辣ｧ蜷茨ｼ亥刈蟾･螳溽ｸｾDATA遲会ｼ�: 蜷御ｸ蟾･遞九�ｮ蜈磯�ｭ蛻暦ｼ亥ｷ･遞�+讖滓｢ｰ�ｼ峨∈蟇�縺帙ｋ
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
    邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ繝ｻ邨先棡_險ｭ蛯吶ぎ繝ｳ繝医�ｮ陦鯉ｼ丞�苓ｦ句�ｺ縺礼畑縲�
    蜀�驛ｨ繧ｭ繝ｼ縺後悟ｷ･遞�+讖滓｢ｰ縲阪�ｮ縺ｨ縺阪�ｯ讖滓｢ｰ蜷阪ｒ陦ｨ遉ｺ縺励∵ｩ滓｢ｰ蜷阪�ｮ驥崎､�譎ゅ�ｮ縺ｿ蟾･遞九ｒ諡ｬ蠑ｧ縺ｧ陬懊≧縲�
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
                out.append(f"{r}�ｼ�{p}�ｼ�" if p else r)
            else:
                out.append(r)
        else:
            out.append(r)
    return out


def _split_equipment_line_process_machine(eq_line: str) -> tuple[str, str]:
    """
    險ｭ蛯吶�槭せ繧ｿ縺ｮ蛻励く繝ｼ縲悟ｷ･遞�+讖滓｢ｰ縲阪ｒ (蟾･遞句錐, 讖滓｢ｰ蜷�) 縺ｫ蛻�蜑ｲ縺吶ｋ縲�
    '+' 縺檎┌縺�縺ｨ縺阪�ｯ讖滓｢ｰ蜷阪�ｮ縺ｿ縺ｨ縺ｿ縺ｪ縺励∝ｷ･遞句錐縺ｯ遨ｺ譁�蟄励�
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
    險ｭ蛯吶ぎ繝ｳ繝医�ｮ諡�蠖楢�繧ｻ繝ｫ逕ｨ縲ょ濠隗抵ｼ丞�ｨ隗堤ｩｺ逋ｽ縺後≠繧後�ｰ謇句燕繧貞ｧ薙→縺ｿ縺ｪ縺励∫┌縺�縺ｨ縺阪�ｯ蜈ｨ菴薙ｒ陦ｨ遉ｺ
    �ｼ域ｰ丞錐縺�1繝医�ｼ繧ｯ繝ｳ縺ｮ縺ｿ縺ｮ縺ｨ縺阪�ｯ蟋薙�ｮ蛻�繧雁�ｺ縺嶺ｸ榊庄縺ｮ縺溘ａ縺昴�ｮ縺ｾ縺ｾ�ｼ峨�NFKC繝ｻ蟇檎伐/蜀ｨ逕ｰ蟇�縺帙�ｯ蟋鍋畑縺ｨ蜷後§縲�
    """
    sei, mei = _split_person_sei_mei(raw)
    if not sei:
        return ""
    n = _normalize_sei_for_match(sei)
    return n if n else sei


def _gantt_row_member_names(evlist) -> str:
    """險ｭ蛯吶ぎ繝ｳ繝郁｡檎畑: 荳ｻ諡�蠖�(op)縺ｨ繧ｵ繝�(sub)繧貞�ｺ迴ｾ鬆�縺ｧ驥崎､�髯､蜴ｻ縺励∝ｧ薙�ｮ縺ｿ繧偵き繝ｳ繝�+蜊願ｧ偵せ繝壹�ｼ繧ｹ蛹ｺ蛻�繧翫�"""
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
        for seg in re.split(r"[,縲‐", sub_raw):
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
    return ", ".join(labels) if labels else "窶�"


def _resolve_equipment_line_key_for_task(task: dict, equipment_list: list | None) -> str:
    """
    險ｭ蛯呎凾髢灘牡繝ｻ險ｭ蛯吝ｰよ怏遨ｺ縺阪�ｮ蛻励く繝ｼ�ｼ�skills / need 縺ｨ蜷後§縲悟ｷ･遞�+讖滓｢ｰ縲阪ｒ蝓ｺ譛ｬ縺ｨ縺吶ｋ�ｼ峨�
    讖滓｢ｰ蜷阪′遨ｺ縺ｧ繝槭せ繧ｿ縺ｫ蠖楢ｩｲ蟾･遞九�ｮ蛻励′1縺､縺�縺代↑繧峨◎縺ｮ隍�蜷医く繝ｼ縺ｸ蟇�縺帙ｋ縲�
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
    驟榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙ｒ DataFrame 蛹悶＠縺溽峩蠕後�ｮ蜈ｱ騾壼�ｦ逅�縲�
    谿ｵ髫�2縺ｮ ``load_planning_tasks_df`` 縺ｨ蜷後§�ｼ郁ｨｭ螳壹す繝ｼ繝医�ｻ蛻�蜑ｲ陦後�ｻ驟榊床荳崎ｦ√Ν繝ｼ繝ｫ�ｼ峨�
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
        logging.exception("%s: 險ｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九�ｮ菫晏ｮ医〒萓句､厄ｼ育ｶ夊｡鯉ｼ�", log_prefix)
    try:
        _apply_auto_exclude_bunkatsu_duplicate_machine(df, log_prefix=log_prefix)
    except Exception as ex:
        logging.warning(
            "%s: 蛻�蜑ｲ陦後�ｮ驟榊床荳崎ｦ∬�ｪ蜍戊ｨｭ螳壹〒萓句､厄ｼ育ｶ夊｡鯉ｼ�: %s",
            log_prefix,
            ex,
        )
    try:
        apply_exclude_rules_config_to_plan_df(df, wb_path, log_prefix)
    except Exception as ex:
        logging.warning(
            "%s: 險ｭ螳壹す繝ｼ繝医↓繧医ｋ驟榊床荳崎ｦ�驕ｩ逕ｨ縺ｧ萓句､厄ｼ育ｶ夊｡鯉ｼ�: %s",
            log_prefix,
            ex,
        )

def load_planning_tasks_df():
    """
    2谿ｵ髫守岼逕ｨ: 繝槭け繝ｭ繝悶ャ繧ｯ荳翫�ｮ縲碁�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙阪す繝ｼ繝医ｒ隱ｭ縺ｿ霎ｼ繧縲�

    縲梧球蠖徹P_謖�螳壹榊�励∪縺溘�ｯ迚ｹ蛻･謖�螳壼ｙ閠�縺ｮ AI 蜃ｺ蜉� preferred_operator 縺ｧ荳ｻ諡�蠖� OP 繧呈欠蜷阪〒縺阪ｋ�ｼ�skills 縺ｮ繝｡繝ｳ繝舌�ｼ蜷阪→縺ゅ＞縺ｾ縺�荳閾ｴ�ｼ峨�
    繝｡繧､繝ｳ縲悟�榊━蜈育音蛻･險倩ｼ峨阪�ｮ task_preferred_operators 縺ｯ generate_plan 蛛ｴ縺ｧ譛蜆ｪ蜈医�槭�ｼ繧ｸ縺輔ｌ繧九�
    縲碁�榊床荳崎ｦ√阪′繧ｪ繝ｳ�ｼ�TRUE/1/縺ｯ縺� 遲会ｼ峨�ｮ陦後�ｯ驟榊床蟇ｾ雎｡螟悶�
    隱ｭ縺ｿ霎ｼ縺ｿ蠕後∝酔荳萓晞�ｼNO繝ｻ驥崎､�讖滓｢ｰ蜷阪′縺ゅｋ繧ｰ繝ｫ繝ｼ繝励�ｮ蟾･遞九悟��蜑ｲ縲崎｡後∈遨ｺ縺ｪ繧峨碁�榊床荳崎ｦ√�=yes�ｼ域ｮｵ髫�1縺ｨ蜷後§�ｼ峨�
    縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九阪〒蟾･遞�+讖滓｢ｰ縺ｮ邨�繧貞酔譛溘＠縲，/D/E 縺ｫ蝓ｺ縺･縺埼�榊床荳崎ｦ√ｒ蜿肴丐縺吶ｋ�ｼ医す繝ｼ繝井ｽ懈�舌�ｯ VBA�ｼ峨�
    """
    if not TASKS_INPUT_WORKBOOK:
        raise FileNotFoundError(
            "TASK_INPUT_WORKBOOK 縺梧悴險ｭ螳壹〒縺吶７BA 縺ｮ RunPython 縺ｧ繝槭け繝ｭ螳溯｡後ヶ繝�繧ｯ縺ｮ繝代せ繧呈ｸ｡縺励※縺上□縺輔＞縲�"
        )
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        raise FileNotFoundError(f"TASK_INPUT_WORKBOOK 縺悟ｭ伜惠縺励∪縺帙ｓ: {TASKS_INPUT_WORKBOOK}")
    df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=PLAN_INPUT_SHEET_NAME)
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(
        df, plan_input_sheet_column_order()
    )
    for c in plan_input_sheet_column_order():
        if c not in df.columns:
            df[c] = ""
    _apply_planning_sheet_post_load_mutations(df, TASKS_INPUT_WORKBOOK, "驟榊床繧ｷ繝ｼ繝郁ｪｭ霎ｼ")
    logging.info(
        f"險育判繧ｿ繧ｹ繧ｯ蜈･蜉�: '{TASKS_INPUT_WORKBOOK}' 縺ｮ '{PLAN_INPUT_SHEET_NAME}' 繧定ｪｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�"
    )
    return df


def _main_sheet_cell_is_global_comment_label(val) -> bool:
    """繝｡繧､繝ｳ繧ｷ繝ｼ繝井ｸ翫後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医崎ｦ句�ｺ縺励そ繝ｫ縺具ｼ郁｡ｨ險倥ｆ繧瑚ｨｱ螳ｹ�ｼ峨�"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    s = unicodedata.normalize("NFKC", str(val).strip())
    if not s:
        return False
    if _nfkc_column_aliases(s) == _nfkc_column_aliases("繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝�"):
        return True
    if "繧ｰ繝ｭ繝ｼ繝舌Ν" in s and "繧ｳ繝｡繝ｳ繝�" in s:
        return True
    return False


def load_main_sheet_global_priority_override_text() -> str:
    """
    TASK_INPUT_WORKBOOK 縺ｮ繝｡繧､繝ｳ繧ｷ繝ｼ繝医〒縲後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医阪→譖ｸ縺九ｌ縺溘そ繝ｫ縺ｮ **逶ｴ荳�** 繧定ｪｭ繧縲�
    繧ｷ繝ｼ繝亥錐: 縲後Γ繧､繝ｳ縲阪後Γ繧､繝ｳ_縲阪勲ain縲阪�ｮ縺�縺壹ｌ縺九√∪縺溘�ｯ蜷榊燕縺ｫ縲後Γ繧､繝ｳ縲阪ｒ蜷ｫ繧�ｼ�VBA GetMainWorksheet 縺ｨ蜷瑚ｶ｣譌ｨ�ｼ峨�

    蜀�螳ｹ縺ｯ **Gemini 縺ｧ荳諡ｬ隗｣驥�**�ｼ�`analyze_global_priority_override_comment`�ｼ峨ょｷ･蝣ｴ莨第･ｭ譌･繝ｻ蜀榊━蜈医ヵ繝ｩ繧ｰ繝ｻ譛ｪ螳溯｣�謖�遉ｺ縺ｮ繝｡繝｢繧� JSON 蛹悶☆繧九�
    API 繧ｭ繝ｼ縺檎┌縺�蝣ｴ蜷医�ｮ縺ｿ縲∝ｷ･蝣ｴ莨第･ｭ譌･縺ｯ繝ｫ繝ｼ繝ｫ繝吶�ｼ繧ｹ縺ｮ `parse_factory_closure_dates_from_global_comment` 縺ｧ陬懷ｮ後☆繧九�
    """
    wb_path = TASKS_INPUT_WORKBOOK.strip() if TASKS_INPUT_WORKBOOK else ""
    if not wb_path or not os.path.exists(wb_path):
        return ""
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "繝｡繧､繝ｳ蜀榊━蜈育音險�: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ openpyxl 縺ｧ繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医ｒ隱ｭ縺ｿ縺ｾ縺帙ｓ縲�",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return ""
    try:
        wb = load_workbook(wb_path, data_only=True, read_only=False)
    except Exception as e:
        logging.warning("繝｡繧､繝ｳ蜀榊━蜈育音險�: 繝悶ャ繧ｯ繧帝幕縺代∪縺帙ｓ縺ｧ縺励◆: %s", e)
        return ""
    try:
        ws = None
        for name in ("繝｡繧､繝ｳ", "繝｡繧､繝ｳ_", "Main"):
            if name in wb.sheetnames:
                ws = wb[name]
                break
        if ws is None:
            for sn in wb.sheetnames:
                if "繝｡繧､繝ｳ" in sn:
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
    繝｡繧､繝ｳ縲後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医阪�ｮ譁ｭ迚�縺後∝ｷ･蝣ｴ蜊倅ｽ阪�ｮ莨第･ｭ繝ｻ髱樒ｨｼ蜒阪ｒ諢丞袖縺吶ｋ縺具ｼ亥倶ｺｺ莨代∩縺�縺代ｒ隱､讀懷�ｺ縺励↑縺��ｼ峨�
    """
    c = unicodedata.normalize("NFKC", str(chunk or ""))
    if not c.strip():
        return False
    if re.search(r"閾ｨ譎�\s*莨第･ｭ", c):
        return True
    if "莨大�ｴ" in c:
        return True
    if re.search(r"蟾･蝣ｴ", c) and re.search(r"莨掃莨第･ｭ|莨代∩|蛛懈ｭ｢|縺贋ｼ代∩", c):
        return True
    if re.search(r"(?:蜈ｨ遉ｾ|蜈ｨ鬢ｨ|蜈ｨ蟾･蝣ｴ).{0,15}(?:莨掃莨第･ｭ|蛛懈ｭ｢)", c):
        return True
    if re.search(r"(?:遞ｼ蜒鋼逕溽肇|繝ｩ繧､繝ｳ).{0,12}(?:蛛懈ｭ｢|縺ｪ縺慾辟｡縺�)", c):
        return True
    if re.search(r"蜉�蟾･.{0,15}(?:縺励↑縺л辟｡縺慾縺ｪ縺慾縺贋ｼ代∩)", c):
        return True
    if "莨第･ｭ" in c and re.search(
        r"(?:蟾･蝣ｴ|蜈ｨ遉ｾ|譛ｬ遉ｾ|蠖捺律|蠑顔､ｾ|蠖鍋､ｾ|蜈ｨ蜩｡|遉ｾ蜈ｨ菴�)", c
    ):
        return True
    return False


def _md_slash_is_likely_fraction_not_date(t: str, start: int, end: int, mo: int, day: int) -> bool:
    """
    縲悟刈蟾･騾溷ｺｦ縺ｯ1/3縺ｨ縺励∪縺吶阪�ｮ 1/3 繧� 1譛�3譌･ 縺ｨ隱､隱阪＠縺ｪ縺�縲�
    縲�4/1縺ｯ蟾･蝣ｴ繧剃ｼ代∩縲阪�ｮ 4/1 縺ｯ譌･莉倥�ｮ縺ｾ縺ｾ�ｼ育峩蠕後′縲後�ｯ縲阪↑繧牙��謨ｰ謇ｱ縺�縺ｫ縺励↑縺��ｼ峨�
    """
    if mo <= 0 or day <= 0:
        return True
    before = t[max(0, start - 32) : start]
    after = t[end : min(len(t), end + 14)]
    after_st = after.lstrip()
    if after_st.startswith("縺ｯ"):
        return False
    if re.search(
        r"(?:蜉�蟾･騾溷ｺｦ|蜉�蟾･\s*繧ｹ繝斐�ｼ繝榎騾溷ｺｦ|蛟咲紫|繧ｹ繝斐�ｼ繝榎蜉ｹ邇�|蜑ｲ蜷�)(?:\s*縺ｯ)?\s*$",
        before,
    ):
        return True
    # 1/2繝ｻ1/3繝ｻ2/3 遲� + 縲後→縺励∪縺吶阪悟阪坂ｦ 縺ｯ蛻�謨ｰ繝ｻ豈皮紫蟇�繧奇ｼ医�3/1縺ｧ縺吶咲ｭ峨�ｮ譌･莉倥ｒ隱､繧ｹ繧ｭ繝�繝励＠縺ｪ縺�繧医≧ 縺ｧ縺�/縺ｧ縺ゅｋ 縺ｯ蜷ｫ繧√↑縺��ｼ�
    frac_pat = re.compile(
        r"^(?:縺ｨ縺励∪縺�?|縺ｨ縺吶ｋ|蛟鋼蜑ｲ蜷�|縺ｫ縺吶ｋ|縺ｫ險ｭ螳嘶縺上ｉ縺л遞句ｺｦ|縺ｫ蝗ｺ螳嘶縺ｫ螟画峩)"
    )
    if mo <= 12 and day <= 12 and frac_pat.match(after_st):
        if mo <= 2 or (mo == 3 and day <= 3):
            return True
    # 縲�1/2縺ｧ縺吶阪�1/10縺ｧ縺吶阪�ｮ繧医≧縺ｪ蛻�豈崎｡ｨ迴ｾ�ｼ亥�磯�ｭ縺� 1/ 縺ｮ縺ｿ�ｼ�
    if (
        mo == 1
        and 2 <= day <= 12
        and re.match(r"^縺ｧ縺處縺ｧ縺ゅｋ\b", after_st)
    ):
        return True
    return False


def _extract_calendar_dates_from_text(s: str, default_year: int) -> list[date]:
    """繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝亥��縺ｮ譌･莉倩｡ｨ險倥ｒ date 縺ｫ螟画鋤�ｼ亥渕貅門ｹｴ縺ｯ險育判縺ｮ蝓ｺ貅門ｹｴ�ｼ峨�"""
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
        r"(\d{4})\s*蟷ｴ\s*(\d{1,2})\s*譛�\s*(\d{1,2})\s*譌･?",
        t,
    ):
        add(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    for m in re.finditer(
        r"(\d{4})\s*[/\-\.�ｼ従\s*(\d{1,2})\s*[/\-\.�ｼ従\s*(\d{1,2})",
        t,
    ):
        add(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    for m in re.finditer(r"(\d{1,2})\s*譛�\s*(\d{1,2})\s*譌･", t):
        add(int(default_year), int(m.group(1)), int(m.group(2)))
    for m in re.finditer(
        r"(?<!\d)(\d{1,2})\s*[/�ｼ従\s*(\d{1,2})(?!\d)",
        t,
    ):
        mo_i, d_i = int(m.group(1)), int(m.group(2))
        if _md_slash_is_likely_fraction_not_date(t, m.start(), m.end(), mo_i, d_i):
            continue
        add(int(default_year), mo_i, d_i)
    return found


def _split_global_comment_into_chunks(blob: str) -> list[str]:
    """
    繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医ｒ縲檎峡遶九＠縺滓欠遉ｺ縲阪�ｮ蝪翫↓蛻�縺代ｋ縲�
    謾ｹ陦鯉ｼ�Excel 縺ｮ Alt+Enter繝ｻUnicode 謾ｹ陦悟性繧�ｼ峨〒蠢�縺壼��蜑ｲ縺励∝酔荳陦悟��縺ｯ 縲�;�ｼ� 縺ｧ邯壹￠縺ｦ蛻�蜑ｲ縲�
    """
    t = unicodedata.normalize("NFKC", str(blob or "").strip())
    if not t:
        return []
    lines = [ln.strip() for ln in re.split(r"[\n\r\v\f\u2028\u2029]+", t) if ln.strip()]
    if not lines:
        return []
    chunks: list[str] = []
    for line in lines:
        subs = [c.strip() for c in re.split(r"[縲�;�ｼ嫋+", line) if c.strip()]
        if subs:
            chunks.extend(subs)
        else:
            chunks.append(line)
    return chunks


def parse_factory_closure_dates_from_global_comment(
    text: str, default_year: int
) -> set[date]:
    """
    繝｡繧､繝ｳ繧ｷ繝ｼ繝医後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医阪↓縲∝ｷ･蝣ｴ閾ｨ譎ゆｼ第･ｭ縺ｪ縺ｩ縺ｨ譌･莉倥′譖ｸ縺九ｌ縺ｦ縺�繧句�ｴ蜷医↓
    縺昴�ｮ譌･繧貞ｷ･蝣ｴ莨代∩�ｼ亥�ｨ蜩｡髱樒ｨｼ蜒阪�ｻ驟榊床縺ｧ蜉�蟾･縺励↑縺��ｼ峨→縺励※謇ｱ縺�譌･莉倬寔蜷医ｒ霑斐☆縲�
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
    """蟾･蝣ｴ莨第･ｭ譌･: 蜍､諤�荳翫�ｯ蜈ｨ蜩｡ is_working=False 縺ｨ縺励√◎縺ｮ譌･縺ｯ險ｭ蛯吝牡莉倥ｒ陦後ｏ縺ｪ縺�縲�"""
    if not closure_dates or not attendance_data:
        return
    tag = "蟾･蝣ｴ莨第･ｭ�ｼ医Γ繧､繝ｳ繝ｻ繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝茨ｼ�"
    for d in sorted(closure_dates):
        if d not in attendance_data:
            logging.warning(
                "繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医�ｮ蟾･蝣ｴ莨第･ｭ譌･ %s 縺ｯ繝槭せ繧ｿ蜍､諤�縺ｫ陦後′縺ゅｊ縺ｾ縺帙ｓ縲�"
                " 縺昴�ｮ譌･縺ｯ險育判繝ｫ繝ｼ繝励↓蜷ｫ縺ｾ繧後↑縺�蝣ｴ蜷医�驟榊床荳翫�ｮ蜉ｹ譫懊′髯仙ｮ夂噪縺ｧ縺吶�",
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
    縲悟宛髯先彫蟒�縲阪後≠繧峨ｆ繧区擅莉ｶ縲咲ｭ�: 險ｭ蛯吝ｰよ怏繝ｻ譎ょ綾繧ｬ繝ｼ繝峨∪縺ｧ蜷ｫ繧�驟榊床蛻ｶ邏�繧堤ｷｩ繧√ｋ�ｼ�abolish_all_scheduling_limits�ｼ峨�
    """
    b = unicodedata.normalize("NFKC", str(blob or ""))
    strong = (
        "蛻ｶ髯先彫蟒�",
        "蛻ｶ髯舌ｒ謦､蟒�",
        "縺吶∋縺ｦ縺ｮ蛻ｶ髯�",
        "蜈ｨ縺ｦ縺ｮ蛻ｶ髯�",
        "縺ゅｉ繧�繧句宛髯�",
        "縺ゅｉ繧�繧区擅莉ｶ",
        "縺吶∋縺ｦ縺ｮ譚｡莉ｶ",
        "蜈ｨ縺ｦ縺ｮ譚｡莉ｶ",
        "謦､蟒�縺励※",
        "謦､蟒�縺�",
    )
    if any(k in b for k in strong):
        out = dict(coerced)
        out["abolish_all_scheduling_limits"] = True
        out["ignore_skill_requirements"] = True
        out["ignore_need_minimum"] = True
        logging.warning(
            "繝｡繧､繝ｳ蜀榊━蜈育音險�: 蛻ｶ髯先彫蟒�繧ｭ繝ｼ繝ｯ繝ｼ繝峨ｒ讀懷�ｺ縲りｨｭ蛯吝ｰよ怏繝ｻ譎ょ綾繧ｬ繝ｼ繝峨ｒ蜷ｫ繧�驟榊床荳翫�ｮ蛻ｶ邏�繧堤ｷｩ繧√∪縺吶�"
        )
        return out
    return coerced


def _maybe_fill_global_speed_rules_from_scheduler_notes(coerced: dict) -> dict:
    """
    AI 縺� global_speed_rules 繧堤ｩｺ縺ｫ縺励◆縺� scheduler_notes 縺ｫ蜈ｷ菴薙ヱ繧ｿ繝ｼ繝ｳ縺後≠繧句�ｴ蜷医�ｮ陬懷ｮ後�
    蠎�縺乗耳貂ｬ縺励↑縺��ｼ育�ｱ陞咲捩�ｼ区､懈渊�ｼ�1/3 邉ｻ縺ｮ縺ｿ�ｼ峨�
    """
    if not isinstance(coerced, dict):
        return coerced
    if coerced.get("global_speed_rules"):
        return coerced
    sn = str(coerced.get("scheduler_notes_ja") or "")
    t = unicodedata.normalize("NFKC", sn)
    if "辭ｱ陞咲捩" not in t or "讀懈渊" not in t:
        return coerced
    if not re.search(r"(?:1\s*/\s*3|�ｼ曾s*/\s*3|荳牙��縺ｮ荳|3\s*蛻�縺ｮ\s*1)", t):
        return coerced
    out = dict(coerced)
    out["global_speed_rules"] = [
        {
            "process_contains": "辭ｱ陞咲捩",
            "machine_contains": "讀懈渊",
            "speed_multiplier": 1.0 / 3.0,
        }
    ]
    logging.info(
        "繝｡繧､繝ｳ蜀榊━蜈育音險�: scheduler_notes_ja 縺九ｉ global_speed_rules 繧定｣懷ｮ鯉ｼ育�ｱ陞咲捩繝ｻ讀懈渊繝ｻ1/3�ｼ�"
    )
    return out


def _finalize_global_priority_override(blob: str, coerced: dict) -> dict:
    """繧ｽ繝ｭ陬懈ｭ｣縺ｮ蠕後∥bolish 縺� true 縺ｪ繧峨せ繧ｭ繝ｫ繝ｻ莠ｺ謨ｰ繧ょｼｷ蛻ｶ繧ｪ繝ｳ縲�"""
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
    縲御ｸ莠ｺ縺ｧ諡�蠖薙阪悟腰迢ｬ縲咲ｭ峨〒莠ｺ謨ｰ縺�縺醍ｷｩ繧√※繧ゅ∵欠蜷阪Γ繝ｳ繝舌�ｼ縺後せ繧ｭ繝ｫ髱櫁ｩｲ蠖薙□縺ｨ驟榊床縺輔ｌ縺ｪ縺�縲�
    縺昴�ｮ蝣ｴ蜷医�ｯ繧ｹ繧ｭ繝ｫ辟｡隕悶ｒ蜷梧凾縺ｫ遶九※繧九�
    """
    if not coerced.get("ignore_need_minimum") or coerced.get("ignore_skill_requirements"):
        return coerced
    b = unicodedata.normalize("NFKC", str(blob or ""))
    solo_kw = ("荳莠ｺ", "縺ｲ縺ｨ繧�", "蜊倡峡", "�ｼ台ｺｺ", "1莠ｺ", "迢ｬ閾ｪ", "蜊倩ｺｫ")
    if any(k in b for k in solo_kw):
        out = dict(coerced)
        out["ignore_skill_requirements"] = True
        logging.info(
            "繝｡繧､繝ｳ蜀榊━蜈育音險�: 蜊倡峡邉ｻ繧ｭ繝ｼ繝ｯ繝ｼ繝峨�ｮ縺溘ａ ignore_skill_requirements 繧定｣懷勧逧�縺ｫ true 縺ｫ縺励∪縺励◆縲�"
        )
        return out
    return coerced


def _coerce_task_preferred_operators_dict(raw_val) -> dict:
    """AI 縺ｮ task_preferred_operators 繧� {萓晞�ｼNO: 豌丞錐} 縺ｫ豁｣隕丞喧縲�"""
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
    AI 縺ｾ縺溘�ｯ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ縺ｮ譌･莉倥Μ繧ｹ繝医ｒ YYYY-MM-DD 譁�蟄怜�励�ｮ譏�鬆�繝ｦ繝九�ｼ繧ｯ縺ｫ豁｣隕丞喧縲�
    隕∫ｴ�縺ｯ ISO 譁�蟄怜�励�ｻExcel 譌･莉倥�ｻ縲�4/1縲咲ｨ句ｺｦ縺ｮ遏ｭ譁�縺ｧ繧ょ庄縲�
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
    Gemini 縺ｮ global_speed_rules 繧呈ｭ｣隕丞喧縲�
    蜷�隕∫ｴ�: process_contains / machine_contains�ｼ医＞縺壹ｌ縺句ｿ�鬆医�ｻ驛ｨ蛻�荳閾ｴ逕ｨ�ｼ�, speed_multiplier�ｼ域里蟄倬溷ｺｦ縺ｫ荵礼ｮ励�0雜�縲�10莉･荳具ｼ峨�
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
    """sub 縺檎ｩｺ縺ｧ縺ｪ縺代ｌ縺ｰ縲∝ｷ･遞句錐縺ｾ縺溘�ｯ讖滓｢ｰ蜷阪�ｮ縺�縺壹ｌ縺九↓驛ｨ蛻�荳閾ｴ縺吶ｌ縺ｰ True縲�"""
    if not sub_nfkc:
        return True
    return sub_nfkc in pnorm or sub_nfkc in mnorm


def _global_speed_multiplier_for_row(process_name: str, machine_name: str, rules: list) -> float:
    """
    蟾･遞句錐繝ｻ讖滓｢ｰ蜷阪↓荳閾ｴ縺吶ｋ繝ｫ繝ｼ繝ｫ縺ｮ speed_multiplier 繧呈寺縺大粋繧上○繧具ｼ井ｸ閾ｴ縺ｪ縺励�ｯ 1.0�ｼ峨�

    process_contains / machine_contains 縺ｯ縺昴ｌ縺槭ｌ **蟾･遞句錐縺ｾ縺溘�ｯ讖滓｢ｰ蜷阪�ｮ縺ｩ縺｡繧峨°** 縺ｫ蜷ｫ縺ｾ繧後ｌ縺ｰ繧医＞縲�
    荳｡譁ｹ謖�螳壽凾縺ｯ AND�ｼ井ｾ�: 縲檎�ｱ陞咲捩縲阪→縲梧､懈渊縲阪′縲∝�励�ｮ邨�縺ｿ蜷医ｏ縺帙〒荳｡譁ｹ迴ｾ繧後ｋ陦後↓繝槭ャ繝√�
    繝槭せ繧ｿ荳翫〒蟾･遞�=讀懈渊繝ｻ讖滓｢ｰ=辭ｱ陞咲捩讖� 縺ｮ繧医≧縺ｫ繧ｭ繝ｼ繝ｯ繝ｼ繝峨′騾�蛛ｴ縺ｮ蛻励↓縺ゅ▲縺ｦ繧ょ酔縺倥Ν繝ｼ繝ｫ縺ｧ蜉ｹ縺上�
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
    Gemini 縺� task_preferred_operators 縺ｫ隱､縺｣縺ｦ髟ｷ譁�繧貞�･繧後◆蝣ｴ蜷医↑縺ｩ縲�
    閾ｪ辟ｶ險隱樊妙迚�縺九ｉ global_day_process_operator_rules 逶ｸ蠖薙ｒ謗ｨ螳壹☆繧具ｼ井ｿ晏ｮ育噪�ｼ峨�
    萓�: 縲�2026/4/4 蟾･遞句錐:EC 譽ｮ荳九→螳ｮ蟲ｶ繧帝�榊床縲�
    """
    t = unicodedata.normalize("NFKC", str(text or "")).strip()
    if len(t) < 6:
        return []
    dates = _extract_calendar_dates_from_text(t, int(ref_y))
    if not dates:
        return []
    d0 = dates[0]
    proc_m = re.search(
        r"蟾･遞句錐?\s*[:�ｼ咯?\s*([A-Za-z0-9荳-鮴ｯ繝ｼ繝ｻ縲�縲�]+)",
        t,
    )
    pc = proc_m.group(1).strip() if proc_m else ""
    if not pc:
        m2 = re.search(r"([\dA-Za-z荳-鮴ｯ繝ｼ繝ｻ縲�縲�]{1,12})\s*蟾･遞�", t)
        pc = m2.group(1).strip() if m2 else ""
    if not pc:
        return []
    names: list[str] = []
    for m in re.finditer(
        r"([\u3040-\u9FFF縲�繝ｼ繝ｻA-Za-z繝ｻ縲�縲�]{1,16}?)\s*縺ｨ\s*([\u3040-\u9FFF縲�繝ｼ繝ｻA-Za-z繝ｻ縲�縲�]{1,16}?)\s*繧�?\s*(?:驟榊床|驟榊ｱ桍邨�縺ｾ縺斈蜷御ｸ繝√�ｼ繝�)",
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
    Gemini 縺� task_preferred_operators 縺ｫ **驟榊��**繧�隱､繧ｹ繧ｭ繝ｼ繝橸ｼ�workstation_id 遲会ｼ峨ｒ霑斐＠縺溘→縺阪�
    謐ｨ縺ｦ縺壹↓ global_day_process_operator_rules / scheduler_notes_ja 縺ｸ謨第ｸ医☆繧九�
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
            if isinstance(v, str) and len(v) > 35 and ("驟�" in v or "蟾･遞�" in v):
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
    """Gemini 謌ｻ繧翫ｒ驟榊床逕ｨ繝輔Λ繧ｰ繝ｻ蟾･蝣ｴ莨第･ｭ譌･繝ｪ繧ｹ繝医↓豁｣隕丞喧縲�"""
    y0 = int(reference_year) if reference_year is not None else date.today().year

    def as_bool(v):
        if v is True:
            return True
        if v is False:
            return False
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        s = unicodedata.normalize("NFKC", str(v).strip()).lower()
        return s in ("true", "1", "yes", "縺ｯ縺�", "on")

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
    """Gemini 蠢懃ｭ斐°繧� JSON 繧ｪ繝悶ず繧ｧ繧ｯ繝�1縺､繧貞叙繧雁�ｺ縺呻ｼ�```json 繝輔ぉ繝ｳ繧ｹ莉倥″縺ｧ繧ょ庄�ｼ峨�"""
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
    """Gemini 譛ｪ菴ｿ逕ｨ繝ｻ蠢懃ｭ碑ｧ｣驥亥､ｱ謨玲凾: 繝ｫ繝ｼ繝ｫ繝吶�ｼ繧ｹ縺ｧ蟾･蝣ｴ莨第･ｭ譌･縺�縺題｣懷ｮ鯉ｼ亥ｾ捺擂莠呈鋤�ｼ峨�"""
    out = dict(coerced)
    rx = parse_factory_closure_dates_from_global_comment(blob, ref_y)
    out["factory_closure_dates"] = sorted({d.isoformat() for d in rx})
    return out


def analyze_global_priority_override_comment(
    text: str, members: list, reference_year: int, ai_sheet_sink: dict | None = None
) -> dict:
    """
    繝｡繧､繝ｳ繧ｷ繝ｼ繝医後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医搾ｼ�UI 荳翫�ｮ閾ｪ逕ｱ險倩ｿｰ�ｼ峨ｒ **Gemini 縺ｧ荳諡ｬ隗｣驥�**縺励�驟榊床縺ｫ蜉ｹ縺� JSON 縺ｫ關ｽ縺ｨ縺吶�
    閾ｪ辟ｶ險隱槭�ｮ譁�閼亥��繧雁��縺代�ｻ謾ｹ陦後�ｮ蛻･謖�遉ｺ隗｣驥医�ｯ AI 縺ｫ莉ｻ縺帙∵綾繧雁､縺ｮ繧ｭ繝ｼ縺�縺代す繧ｹ繝�繝�縺梧ｩ滓｢ｰ驕ｩ逕ｨ縺吶ｋ縲�

    - factory_closure_dates: **蟾･蝣ｴ蜈ｨ菴�**縺ｧ遞ｼ蜒阪＠縺ｪ縺�譌･�ｼ亥�ｨ蜩｡髱樒ｨｼ蜒肴桶縺��ｼ峨�ｮ YYYY-MM-DD 譁�蟄怜�励�ｮ驟榊�励りｩｲ蠖薙↑縺励�ｯ []縲�
    - ignore_skill_requirements / ignore_need_minimum / abolish_all_scheduling_limits / task_preferred_operators: 蠕捺擂縺ｩ縺翫ｊ縲�
    - global_speed_rules: **蟾･遞句錐繝ｻ讖滓｢ｰ蜷�**縺ｸ縺ｮ驛ｨ蛻�荳閾ｴ�ｼ亥推繧ｭ繝ｼ繝ｯ繝ｼ繝峨�ｯ **縺ｩ縺｡繧峨�ｮ蛻励↓縺ゅ▲縺ｦ繧ょ庄**�ｼ峨〒縲∵里蟄倥�ｮ蜉�蟾･騾溷ｺｦ�ｼ医す繝ｼ繝茨ｼ丈ｸ頑嶌縺榊ｾ鯉ｼ峨↓ **荵礼ｮ�**縺吶ｋ繝ｫ繝ｼ繝ｫ縺ｮ驟榊�励りｩｲ蠖薙↑縺励�ｯ []縲�
    - global_day_process_operator_rules: **譌･莉假ｼ句ｷ･遞句錐縺ｮ驛ｨ蛻�荳閾ｴ�ｼ玖､�謨ｰ繝｡繝ｳ繝舌�ｼ**繧偵∝ｽ捺律縺昴�ｮ蟾･遞九�ｮ繧ｿ繧ｹ繧ｯ縺ｮ**繝√�ｼ繝�蜈ｨ蜩｡縺ｫ蠢�縺壼性繧√ｋ**繝ｫ繝ｼ繝ｫ縺ｮ驟榊�励りｩｲ蠖薙↑縺励�ｯ []縲�
    - scheduler_notes_ja: 荳願ｨ倥↓關ｽ縺ｨ縺励″繧後↑縺�陬懆ｶｳ繧�驕狗畑繝｡繝｢�ｼ磯溷ｺｦ縺ｯ蜿ｯ閭ｽ縺ｪ繧� global_speed_rules 繧ゆｽｵ險假ｼ峨�

    API 繧ｭ繝ｼ辟｡縺励�ｻJSON 隗｣驥亥､ｱ謨玲凾: 荳願ｨ倥ヶ繝ｼ繝ｫ繝ｻ謖�蜷阪�ｯ譌｢螳壼､縲∝ｷ･蝣ｴ莨第･ｭ譌･縺ｮ縺ｿ蠕捺擂縺ｮ繝ｫ繝ｼ繝ｫ繝吶�ｼ繧ｹ隗｣譫舌〒陬懷ｮ後�
    """
    ref_y = int(reference_year) if reference_year is not None else date.today().year
    empty = _coerce_global_priority_override_dict({}, ref_y)
    if not text or not str(text).strip():
        if ai_sheet_sink is not None:
            ai_sheet_sink["繝｡繧､繝ｳ蜀榊━蜈育音險論AI_API"] = "繧ｹ繧ｭ繝�繝暦ｼ医Γ繧､繝ｳ蜴滓枚縺ｪ縺暦ｼ�"
        return empty
    blob = str(text).strip()
    mem_sig = ",".join(sorted(str(m).strip() for m in (members or []) if m))
    cache_fingerprint = f"{GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX}{ref_y}\n{blob}\n{mem_sig}"
    cache_key = hashlib.sha256(cache_fingerprint.encode("utf-8")).hexdigest()
    ai_cache = load_ai_cache()
    cached = get_cached_ai_result(ai_cache, cache_key, content_key=cache_fingerprint)
    if cached is not None:
        logging.info("繝｡繧､繝ｳ蜀榊━蜈育音險�: 繧ｭ繝｣繝�繧ｷ繝･繝偵ャ繝茨ｼ�Gemini 縺ｯ蜻ｼ縺ｳ縺ｾ縺帙ｓ�ｼ峨�")
        if ai_sheet_sink is not None:
            ai_sheet_sink["繝｡繧､繝ｳ蜀榊━蜈育音險論AI_API"] = "縺ｪ縺暦ｼ医く繝｣繝�繧ｷ繝･菴ｿ逕ｨ�ｼ�"
        return _finalize_global_priority_override(
            blob, _coerce_global_priority_override_dict(cached, ref_y)
        )

    if not API_KEY:
        logging.info("GEMINI_API_KEY 譛ｪ險ｭ螳壹�ｮ縺溘ａ繝｡繧､繝ｳ蜀榊━蜈育音險倥�ｮ AI 隗｣譫舌ｒ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆縲�")
        if ai_sheet_sink is not None:
            ai_sheet_sink["繝｡繧､繝ｳ蜀榊━蜈育音險論AI_API"] = "縺ｪ縺暦ｼ�API繧ｭ繝ｼ譛ｪ險ｭ螳壹�ｻ蟾･蝣ｴ莨第･ｭ縺ｮ縺ｿ繝ｫ繝ｼ繝ｫ陬懷ｮ鯉ｼ�"
        coerced = _apply_regex_factory_closure_fallback(
            _coerce_global_priority_override_dict({}, ref_y), blob, ref_y
        )
        return _finalize_global_priority_override(blob, coerced)

    member_sample = ", ".join(str(m) for m in (members or [])[:80])
    if len(members or []) > 80:
        member_sample += " 窶ｦ"

    prompt = f"""縺ゅ↑縺溘�ｯ蟾･蝣ｴ縺ｮ驟榊床險育判繧ｷ繧ｹ繝�繝�逕ｨ繧｢繧ｷ繧ｹ繧ｿ繝ｳ繝医〒縺吶�
Excel 繝｡繧､繝ｳ繧ｷ繝ｼ繝医�ｮ **縲後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医�**�ｼ郁�ｪ逕ｱ險倩ｿｰ繝ｻ閾ｪ辟ｶ險隱橸ｼ峨�ｮ **蜈ｨ譁�** 繧定ｪｭ縺ｿ縲∵ｬ｡縺ｮ繧ｭ繝ｼ縺�縺代ｒ謖√▽ JSON 繧�1縺､霑斐＠縺ｦ縺上□縺輔＞縲�

縲仙ｽｹ蜑ｲ縲�
繝ｦ繝ｼ繧ｶ繝ｼ縺ｯ謾ｹ陦後ｄ蜿･轤ｹ縺ｧ隍�謨ｰ縺ｮ謖�遉ｺ繧呈嶌縺上％縺ｨ縺後≠繧翫∪縺吶�**譁�閼医ｒ隱ｭ縺ｿ蛻�縺�**縲�驟榊床繧ｷ繧ｹ繝�繝�縺� **讖滓｢ｰ逧�縺ｫ驕ｩ逕ｨ縺ｧ縺阪ｋ蛟､** 縺ｫ關ｽ縺ｨ縺苓ｾｼ繧薙〒縺上□縺輔＞縲�
謗ｨ貂ｬ縺ｧ繝悶�ｼ繝ｫ繧� true 縺ｫ縺励↑縺�縺薙→縲よ�ｹ諡�縺梧�守｢ｺ縺ｪ縺ｨ縺阪□縺� true縲�

縲先怙蜆ｪ蜈医�
縺薙�ｮ谺�縺ｮ蜀�螳ｹ縺ｯ繝槭せ繧ｿ繝ｻ繧ｹ繧ｭ繝ｫ繝ｻneed繝ｻ繧ｿ繧ｹ繧ｯ陦後�ｻ迚ｹ蛻･謖�螳喟蛯呵�縺ｮ AI 謖�蜷阪ｈ繧雁━蜈医＆繧後ｋ萓句､匁欠遉ｺ縺ｨ縺励※謇ｱ繧上ｌ縺ｾ縺吶�

縲先隼陦後�ｻ隍�謨ｰ陦後�
蜷�陦後�ｻ蜷�譁�縺ｯ **蜴溷援縺ｨ縺励※迢ｬ遶九＠縺滓欠遉ｺ** 縺ｧ縺吶り｡後ｒ縺ｾ縺溘＞縺ｧ1縺､縺ｫ縺ｾ縺ｨ繧√◆繧翫�**蜑ｲ蜷郁｡ｨ迴ｾ�ｼ井ｾ� 1/3�ｼ峨ｒ譌･莉倥→邨舌�ｳ莉倥￠縺溘ｊ縺励↑縺�**縺薙→縲�

縲舌く繝ｼ蛻･繝ｫ繝ｼ繝ｫ縲�

A) **factory_closure_dates** �ｼ磯�榊�励�ｻ蠢�鬆茨ｼ�
   - **蟾･蝣ｴ蜈ｨ菴�**縺檎ｨｼ蜒阪＠縺ｪ縺�譌･�ｼ郁�ｨ譎ゆｼ第･ｭ繝ｻ蜈ｨ蟾･蝣ｴ莨代∩繝ｻ縺昴�ｮ譌･縺ｯ蜉�蟾･縺励↑縺�遲会ｼ峨�ｮ譌･莉倥ｒ **YYYY-MM-DD** 縺ｮ譁�蟄怜�励〒蛻玲嫌縲�
   - **蛟倶ｺｺ縺ｮ莨代∩繝ｻ迚ｹ螳壹Λ繧､繝ｳ縺�縺�**縺ｮ蛛懈ｭ｢縺ｯ縺薙％縺ｫ **蜷ｫ繧√↑縺�**�ｼ�[]�ｼ峨�
   - 隧ｲ蠖薙′縺ｪ縺代ｌ縺ｰ **遨ｺ驟榊�� []**�ｼ医く繝ｼ逵∫払荳榊庄�ｼ峨�
   - 蟷ｴ縺檎怐逡･縺輔ｌ縺ｦ縺�繧後�ｰ隘ｿ證ｦ {ref_y} 蟷ｴ縺ｨ縺励※隗｣驥医�

B) **ignore_skill_requirements** / **ignore_need_minimum** / **abolish_all_scheduling_limits** / **task_preferred_operators**
   - 蠕捺擂縺ｩ縺翫ｊ�ｼ磯�榊床縺ｮ繧ｹ繧ｭ繝ｫ辟｡隕悶�ｻ莠ｺ謨ｰ1蝗ｺ螳壹�ｻ蛻ｶ髯先彫蟒�繝ｻ萓晞�ｼNO竊剃ｸｻ諡�蠖徹P謖�蜷搾ｼ峨りｩｲ蠖薙↑縺代ｌ縺ｰ false 縺ｾ縺溘�ｯ {{}}縲�

C) **global_speed_rules** �ｼ磯�榊�励�ｻ蠢�鬆茨ｼ�
   - 迚ｹ螳壹�ｮ **蟾･遞句錐**�ｼ�Excel縲悟ｷ･遞句錐縲榊�暦ｼ峨ｄ **讖滓｢ｰ蜷�**�ｼ医梧ｩ滓｢ｰ蜷阪榊�暦ｼ峨↓蟇ｾ縺励�**譌｢蟄倥�ｮ蜉�蟾･騾溷ｺｦ縺ｫ謗帙￠繧句咲紫** 繧呈欠螳壹☆繧九が繝悶ず繧ｧ繧ｯ繝医�ｮ繝ｪ繧ｹ繝医�
   - 蜷�繧ｪ繝悶ず繧ｧ繧ｯ繝医�ｮ繧ｭ繝ｼ:
     - "process_contains": 譁�蟄怜�暦ｼ育怐逡･蜿ｯ�ｼ峨�**蟾･遞句錐縺ｾ縺溘�ｯ讖滓｢ｰ蜷阪�ｮ縺�縺壹ｌ縺�**縺ｫ **驛ｨ蛻�荳閾ｴ**�ｼ�NFKC 諠ｳ螳夲ｼ峨�
     - "machine_contains": 譁�蟄怜�暦ｼ育怐逡･蜿ｯ�ｼ峨�**蟾･遞句錐縺ｾ縺溘�ｯ讖滓｢ｰ蜷阪�ｮ縺�縺壹ｌ縺�**縺ｫ **驛ｨ蛻�荳閾ｴ**縲�
     - "speed_multiplier": 豁｣縺ｮ謨ｰ縲�**1/3 縺ｮ騾溷ｺｦ**縺ｪ繧臥ｴ� **0.333333**�ｼ域里蟄倬溷ｺｦ ﾃ� 縺薙�ｮ蛟､�ｼ峨�**2蛟埼�**縺ｪ繧� 2.0縲�
   - **荳｡譁ｹ謖�螳壽凾縺ｯ AND**�ｼ�2縺､縺ｮ繧ｭ繝ｼ繝ｯ繝ｼ繝峨′縲�**荳｡譁ｹ縺ｨ繧�**縲悟ｷ･遞句錐繝ｻ讖滓｢ｰ蜷阪�ｮ縺ｩ縺｡繧峨°縲阪↓迴ｾ繧後ｋ陦鯉ｼ峨ゆｾ�: 蟾･遞�=讀懈渊繝ｻ讖滓｢ｰ=辭ｱ陞咲捩讖� 縺ｧ繧ゅ∝ｷ･遞�=辭ｱ陞咲捩繝ｻ讖滓｢ｰ=讀懈渊逕ｨ險ｭ蛯� 縺ｧ繧ゅ�槭ャ繝√＠縺�繧九�
   - 縺ｩ縺｡繧峨°荳譁ｹ縺�縺第欠螳壹☆繧後�ｰ縲√◎縺ｮ繧ｭ繝ｼ繝ｯ繝ｼ繝峨′蟾･遞句錐縺ｾ縺溘�ｯ讖滓｢ｰ蜷阪�ｮ縺ｩ縺｡繧峨°縺ｫ縺ゅｌ縺ｰ繝槭ャ繝√�
   - 隧ｲ蠖捺欠遉ｺ縺後↑縺代ｌ縺ｰ **遨ｺ驟榊�� []**縲�
   - 萓�: 縲檎�ｱ陞咲捩繧剃ｽｿ縺�讀懈渊縺ｮ蜉�蟾･騾溷ｺｦ縺ｯ1/3縲坂��
     [{{"process_contains":"辭ｱ陞咲捩","machine_contains":"讀懈渊","speed_multiplier":0.333333}}]
     �ｼ医檎�ｱ陞咲捩縲阪→縲梧､懈渊縲阪′蟾･遞句錐繝ｻ讖滓｢ｰ蜷阪�ｮ邨�縺ｿ蜷医ｏ縺帙〒謠�縺�繧ｿ繧ｹ繧ｯ縺ｮ騾溷ｺｦ縺檎ｴ�1/3縺ｫ縺ｪ繧具ｼ�

D) **scheduler_notes_ja** �ｼ域枚蟄怜�励�ｻ蠢�鬆茨ｼ�
   - 荳願ｨ倥く繝ｼ縺ｫ關ｽ縺ｨ縺励″繧後↑縺�陬懆ｶｳ縲る溷ｺｦ縺ｯ **global_speed_rules 縺ｧ讒矩�蛹悶〒縺阪ｋ縺ｨ縺阪�ｯ蠢�縺壹◎縺｡繧峨↓繧ょ�ｺ縺�**�ｼ医％縺薙�ｯ莠ｺ髢灘髄縺題ｦ∫ｴ�縺ｧ繧ゅｈ縺��ｼ峨ら┌縺代ｌ縺ｰ ""縲�

E) **interpretation_ja** �ｼ域枚蟄怜�励�ｻ蠢�鬆茨ｼ�
   - 蜴滓枚縺ｮ隕∫ｴ�繧�1譁��ｼ�200譁�蟄嶺ｻ･蜀��ｼ峨�

F) **global_day_process_operator_rules** �ｼ磯�榊�励�ｻ蠢�鬆茨ｼ�
   - **迚ｹ螳壹�ｮ遞ｼ蜒肴律**縺九▽ **蟾･遞句錐�ｼ�Excel縲悟ｷ･遞句錐縲榊�暦ｼ峨�ｮ驛ｨ蛻�荳閾ｴ** 縺ｫ蠖薙※縺ｯ縺ｾ繧九ち繧ｹ繧ｯ縺ｫ縺､縺�縺ｦ縲�
     蛻玲嫌縺励◆ **蜈ｨ繝｡繝ｳ繝舌�ｼ繧貞酔荳繝√�ｼ繝�縺ｫ蠢�縺壼性繧√ｋ** 繝ｫ繝ｼ繝ｫ�ｼ�**OP/AS 縺ｩ縺｡繧峨�ｮ繧ｹ繧ｭ繝ｫ縺ｧ繧ょ庄**縲よｰ丞錐隗｣豎ｺ縺ｯ **諡�蠖徹P謖�蜷阪→蜷後§**�ｼ峨�
   - **萓晞�ｼNO縺悟��縺九ｋ荳ｻ諡�蠖薙�ｮ1蜷肴欠蜷�**縺ｯ **task_preferred_operators** 繧剃ｽｿ縺�縺薙→縲ょ次譁�縺� **縲娯留譛遺留譌･縺ｮ笆ｳ蟾･遞九↓�ｼ｡縺ｨ�ｼ｢繧帝�榊床縲�** 縺ｮ繧医≧縺ｫ **譌･莉倥�ｻ蟾･遞九�ｻ隍�謨ｰ蜷�**縺ｮ縺ｨ縺阪�ｯ **譛ｬ驟榊��**縺ｸ關ｽ縺ｨ縺吶�
   - 蜷�繧ｪ繝悶ず繧ｧ繧ｯ繝医�ｮ繧ｭ繝ｼ:
     - "date": **YYYY-MM-DD**�ｼ医◎縺ｮ譌･縺ｫ蜑ｲ繧雁ｽ薙※繧九Ο繝ｼ繝ｫ縺ｫ驕ｩ逕ｨ�ｼ�
     - "process_contains": 蟾･遞句錐縺ｫ **驛ｨ蛻�荳閾ｴ**�ｼ�NFKC 諠ｳ螳夲ｼ峨ゆｾ�: "EC"
     - "operator_names": 豌丞錐縺ｮ驟榊�暦ｼ井ｾ�: ["譽ｮ荳�", "螳ｮ蟲ｶ縲闃ｱ蟄�"]�ｼ�
   - 隧ｲ蠖捺欠遉ｺ縺後↑縺代ｌ縺ｰ **遨ｺ驟榊�� []**縲�

縲占ｿ皮ｭ泌ｽ｢蠑上�
蜈磯�ｭ縺� {{ 縺ｧ邨ゅｏ繧翫′ }} 縺ｮ **JSON 繧ｪ繝悶ず繧ｧ繧ｯ繝�1縺､縺ｮ縺ｿ**�ｼ郁ｪｬ譏取枚繝ｻ繝槭�ｼ繧ｯ繝繧ｦ繝ｳ遖∵ｭ｢�ｼ峨�

蠢�鬆医く繝ｼ荳隕ｧ:
- "factory_closure_dates": string 縺ｮ驟榊�暦ｼ�YYYY-MM-DD�ｼ�
- "ignore_skill_requirements": true 縺ｾ縺溘�ｯ false
- "ignore_need_minimum": true 縺ｾ縺溘�ｯ false
- "abolish_all_scheduling_limits": true 縺ｾ縺溘�ｯ false
- "task_preferred_operators": **JSON 繧ｪ繝悶ず繧ｧ繧ｯ繝医�ｮ縺ｿ**�ｼ医く繝ｼ=萓晞�ｼNO繝ｻ蛟､=荳ｻ諡�蠖捺ｰ丞錐�ｼ峨�**驟榊�励↓縺励※縺ｯ縺ｪ繧峨↑縺�**縲りｩｲ蠖薙↑縺励�ｯ {{}}
- "global_speed_rules": 繧ｪ繝悶ず繧ｧ繧ｯ繝医�ｮ驟榊�暦ｼ郁ｩｲ蠖薙↑縺励�ｯ []�ｼ�
- "global_day_process_operator_rules": 繧ｪ繝悶ず繧ｧ繧ｯ繝医�ｮ驟榊�暦ｼ郁ｩｲ蠖薙↑縺励�ｯ []�ｼ�
- "scheduler_notes_ja": 譁�蟄怜��
- "interpretation_ja": 譁�蟄怜��

縲仙渕貅門ｹｴ縲� 譌･莉倩ｨ蜿翫′縺ゅｌ縺ｰ隘ｿ證ｦ {ref_y} 蟷ｴ縺ｨ縺励※隗｣驥医＠縺ｦ繧医＞縲�

縲千匳骭ｲ繝｡繝ｳ繝舌�ｼ蜷阪�ｮ蜿り�縲托ｼ育�ｧ蜷育畑縲�JSON 繧ｭ繝ｼ縺ｫ縺ｯ蜷ｫ繧√↑縺��ｼ�
{member_sample}

縲舌げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医�ｻ蜴滓枚縲�
{blob}
"""
    try:
        ppath = os.path.join(log_dir, "ai_global_priority_override_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("繝｡繧､繝ｳ蜀榊━蜈育音險�: 繝励Ο繝ｳ繝励ヨ蜈ｨ譁� 竊� %s", ppath)
    except OSError as ex:
        logging.warning("繝｡繧､繝ｳ蜀榊━蜈育音險�: 繝励Ο繝ｳ繝励ヨ菫晏ｭ伜､ｱ謨�: %s", ex)

    client = genai.Client(api_key=API_KEY)
    try:
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        parsed = _parse_global_priority_override_gemini_response(res)
        if parsed is None:
            logging.warning(
                "繝｡繧､繝ｳ蜀榊━蜈育音險�: AI 蠢懃ｭ斐°繧� JSON 繧定ｧ｣驥医〒縺阪∪縺帙ｓ縺ｧ縺励◆縲ゅく繝｣繝�繧ｷ繝･縺帙★縲∵ｬ｡蝗槫�崎ｩｦ陦後＆繧後∪縺吶�"
            )
            try:
                rpath = os.path.join(log_dir, "ai_global_priority_override_last_response.txt")
                with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                    rf.write(_gemini_result_text(res) or "")
            except OSError:
                pass
            if ai_sheet_sink is not None:
                ai_sheet_sink["繝｡繧､繝ｳ蜀榊━蜈育音險論AI_API"] = "縺ゅｊ�ｼ�JSON隗｣驥亥､ｱ謨励�ｻ蟾･蝣ｴ莨第･ｭ縺ｯ繝ｫ繝ｼ繝ｫ陬懷ｮ鯉ｼ�"
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
            "繝｡繧､繝ｳ蜀榊━蜈育音險�: AI 隗｣驥� factory莨第･ｭ=%s譌･ 騾溷ｺｦ繝ｫ繝ｼ繝ｫ=%s莉ｶ 譌･ﾃ怜ｷ･遞九メ繝ｼ繝�=%s莉ｶ skill=%s need1=%s abolish=%s task_pref=%s莉ｶ 窶� %s",
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
            ai_sheet_sink["繝｡繧､繝ｳ蜀榊━蜈育音險論AI_API"] = "縺ゅｊ"
        return coerced
    except Exception as e:
        logging.warning("繝｡繧､繝ｳ蜀榊━蜈育音險�: Gemini 蜻ｼ縺ｳ蜃ｺ縺怜､ｱ謨�: %s", e)
        if ai_sheet_sink is not None:
            ai_sheet_sink["繝｡繧､繝ｳ蜀榊━蜈育音險論AI_API"] = f"螟ｱ謨�: {e}"[:500]
        coerced = _apply_regex_factory_closure_fallback(
            _coerce_global_priority_override_dict({}, ref_y), blob, ref_y
        )
        return _finalize_global_priority_override(blob, coerced)


def default_result_task_sheet_column_order(max_history_len: int) -> list:
    """邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ譌｢螳壼�鈴���ｼ亥ｱ･豁ｴ蛻玲焚縺ｯ螳溯｡梧凾縺ｫ豎ｺ縺ｾ繧具ｼ峨�"""
    hist = [f"螻･豁ｴ{i+1}" for i in range(max_history_len)]
    return [
        "繧ｹ繝�繝ｼ繧ｿ繧ｹ",
        "繧ｿ繧ｹ繧ｯID",
        "蟾･遞句錐",
        "讖滓｢ｰ蜷�",
        "蜆ｪ蜈亥ｺｦ",
        RESULT_TASK_COL_DISPATCH_TRIAL_ORDER,
        *hist,
        "蠢�隕＾P(荳頑嶌)",
        "繧ｿ繧ｹ繧ｯ蜉ｹ邇�",
        "蜉�蟾･騾比ｸｭ",
        "迚ｹ蛻･謖�螳壹≠繧�",
        "諡�蠖徹P謖�蜷�",
        "蝗樒ｭ皮ｴ肴悄",
        "謖�螳夂ｴ肴悄",
        "險育判蝓ｺ貅也ｴ肴悄",
        TASK_COL_RAW_INPUT_DATE,
        "邏肴悄邱頑･",
        "蜉�蟾･髢句ｧ区律",
        "驟榊ｮ契蜉�蟾･髢句ｧ�",
        "驟榊ｮ契蜉�蟾･邨ゆｺ�",
        RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16,
        "邱丞刈蟾･驥�",
        "谿句刈蟾･驥�",
        "螳御ｺ�邇�(螳溯｡梧凾轤ｹ)",
        "迚ｹ蛻･謖�螳喟AI",
    ]


def _task_date_key_for_result_sheet_sort(val):
    """邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ荳ｦ縺ｹ譖ｿ縺育畑縲よｬ�謳阪�ｻ隗｣驥井ｸ崎�ｽ縺ｯ譛蠕鯉ｼ�date.max�ｼ峨�"""
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
    """蝗樒ｭ皮ｴ肴悄繝ｻ謖�螳夂ｴ肴悄縺ｪ縺ｩ繧� date 縺ｫ豁｣隕丞喧�ｼ域ｬ�謳阪�ｯ None�ｼ峨�"""
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
    邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ逕ｨ: 縲碁�榊ｮ契蜉�蟾･邨ゆｺ�縲咲嶌蠖薙�ｮ譛邨らｵゆｺ�縺後�
    蝗樒ｭ皮ｴ肴悄縺ｮ譌･莉� + PLAN_DUE_DAY_COMPLETION_TIME�ｼ域里螳� 16:00�ｼ我ｻ･荳九°繧貞愛螳壹�
    蝗樒ｭ皮ｴ肴悄縺檎┌縺�陦後�ｯ謖�螳夂ｴ肴悄縺ｮ譌･莉� + 16:00 縺ｧ蛻､螳壹�
    荳｡譁ｹ辟｡縺�蝣ｴ蜷医�ｯ縲檎ｴ肴悄縺ｪ縺励阪�
    """
    if not plan_window or len(plan_window) < 2:
        return "譛ｪ蜑ｲ蠖�"
    _pe = plan_window[1]
    if _pe is None:
        return "譛ｪ蜑ｲ蠖�"
    dd = _coerce_planning_date_for_deadline(answer_due)
    if dd is None:
        dd = _coerce_planning_date_for_deadline(specified_due)
    if dd is None:
        return "邏肴悄縺ｪ縺�"
    try:
        deadline_dt = datetime.combine(dd, PLAN_DUE_DAY_COMPLETION_TIME)
        if _pe <= deadline_dt:
            return "縺ｯ縺�"
        return "縺�縺�縺�"
    except Exception:
        return "蛻､螳壻ｸ崎�ｽ"


def _result_task_sheet_sort_key(t: dict):
    """
    邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ陦ｨ遉ｺ鬆�縲や蔵驟榊床隧ｦ陦碁��逡ｪ�ｼ�generate_plan 蜀帝�ｭ縺ｧ繧ｭ繝･繝ｼ鬆�縺ｫ莉倅ｸ弱＠縺� 1..n�ｼ画��鬆�縲�
    谺�謳阪�ｻ髱樊焚縺ｯ譛蠕後ょ酔荳隧ｦ陦碁��蜀�縺ｯ萓晞�ｼNO繝ｻ讖滓｢ｰ蜷阪∫ｶ壹￠縺ｦ蜉�蟾･髢句ｧ区律繝ｻ邏肴悄縺ｧ螳牙ｮ壼喧縲�
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
    """蛻苓ｨｭ螳壹す繝ｼ繝医〒縲悟ｱ･豁ｴ縲�1陦後ｒ鄂ｮ縺上→螻･豁ｴ1�ｽ柤 繧偵◎縺ｮ菴咲ｽｮ縺ｫ螻暮幕縺吶ｋ縲�"""
    if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
        return False
    s = unicodedata.normalize("NFKC", str(cell_val).strip())
    return s in ("螻･豁ｴ", "螻･豁ｴ*")


def _result_task_column_alias_map(df_columns) -> dict:
    """隕句�ｺ縺励�ｮ NFKC 豁｣隕丞喧繧ｭ繝ｼ 竊� DataFrame 荳翫�ｮ螳溷�怜錐縲�"""
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
    # 譌ｧ蛻怜錐�ｼ郁ｨ育判蝓ｺ貅也ｴ肴悄繝吶�ｼ繧ｹ�ｼ俄�� 驟榊ｮ契蝗樒ｭ疲欠螳�16譎ゅ∪縺ｧ
    if nk == _nfkc_column_aliases("驟榊ｮ契蝓ｺ貅�16譎ゅ∪縺ｧ"):
        return col_by_norm.get(
            _nfkc_column_aliases(RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16)
        )
    return None


def _parse_column_visible_cell(val) -> bool:
    """陦ｨ遉ｺ蛻�: 遨ｺ繝ｻ譛ｪ險伜�･縺ｯ True�ｼ郁｡ｨ遉ｺ�ｼ峨�FALSE/0/縺�縺�縺� 遲峨〒髱櫁｡ｨ遉ｺ縲�"""
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
    if s in ("", "true", "1", "縺ｯ縺�", "yes", "on", "陦ｨ遉ｺ", "笳�"):
        return True
    if s in ("false", "flase", "0", "縺�縺�縺�", "no", "off", "髱櫁｡ｨ遉ｺ", "髫�縺�", "ﾃ�"):
        return False
    return True


def parse_result_task_column_config_dataframe(
    df_cfg: pd.DataFrame | None, max_history_len: int
) -> list | None:
    """
    縲悟�苓ｨｭ螳喟邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縲咲嶌蠖薙�ｮ DataFrame 縺九ｉ (蛻励Λ繝吶Ν, 陦ｨ遉ｺ) 繧剃ｸ翫°繧芽ｪｭ繧縲�
    隕句�ｺ縺励悟�怜錐縲阪→縲瑚｡ｨ遉ｺ縲搾ｼ育┌縺�蝣ｴ蜷医�ｯ陦ｨ遉ｺ縺ｯ縺吶∋縺ｦ True�ｼ峨�
    縲悟ｱ･豁ｴ縲阪悟ｱ･豁ｴ*縲阪�ｮ1陦後�ｯ螻･豁ｴ1�ｽ槫ｱ･豁ｴn 縺ｫ螻暮幕縺励∝酔荳陦後�ｮ陦ｨ遉ｺ繝輔Λ繧ｰ繧貞�ｱ譛峨☆繧九�
    蜷御ｸ蛻怜錐�ｼ�NFKC繝ｻ蛻･蜷肴ｭ｣隕丞喧蠕鯉ｼ峨′隍�謨ｰ陦後≠繧句�ｴ蜷医�ｯ蜈磯�ｭ陦後�ｮ縺ｿ謗｡逕ｨ縺励∽ｻ･髯阪�ｯ繝ｭ繧ｰ縺ｫ蜃ｺ縺励※謐ｨ縺ｦ繧九�
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
                "蛻苓ｨｭ螳壹�%s縲�: 驥崎､�蛻怜錐縲�%s縲阪ｒ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆�ｼ井ｸ翫�ｮ陦後ｒ蜆ｪ蜈茨ｼ峨�",
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
                _try_add(f"螻･豁ｴ{j+1}", vis)
            continue
        if raw is None or (isinstance(raw, float) and pd.isna(raw)):
            continue
        s = unicodedata.normalize("NFKC", str(raw).strip())
        if not s or s.lower() in ("nan", "none"):
            continue
        _try_add(s, vis)
    return out or None


def _xlwings_write_column_config_sheet_ab(xw_sheet, rows: list[tuple[str, bool]]) -> None:
    """蛻苓ｨｭ螳壹す繝ｼ繝医�ｮ A:B 繧� 蛻怜錐繝ｻ陦ｨ遉ｺ 縺ｮ縺ｿ縺ｧ荳頑嶌縺搾ｼ�1陦檎岼隕句�ｺ縺暦ｼ九ョ繝ｼ繧ｿ�ｼ峨�"""
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
    TASK_INPUT_WORKBOOK 縺ｮ縲悟�苓ｨｭ螳喟邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縲阪す繝ｼ繝医°繧� (蛻励Λ繝吶Ν, 陦ｨ遉ｺ) 繧剃ｸ翫°繧芽ｪｭ繧縲�
    """
    wb = TASKS_INPUT_WORKBOOK
    if not wb or not os.path.exists(wb):
        return None
    if _workbook_should_skip_openpyxl_io(wb):
        logging.info(
            "蛻苓ｨｭ螳�: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ pandas(openpyxl) 縺ｧ縺ｮ縲�%s縲崎ｪｭ霎ｼ繧偵せ繧ｭ繝�繝暦ｼ域里螳壼�鈴��繧剃ｽｿ縺�縺ｾ縺呻ｼ峨�",
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
            "繧ｷ繝ｼ繝医�%s縲�: 隱ｭ縺ｿ霎ｼ縺ｿ縺ｫ螟ｱ謨励＠縺溘◆繧∵里螳壹�ｮ蛻鈴��繧剃ｽｿ縺�縺ｾ縺� (%s)",
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
    蛻苓ｨｭ螳壹す繝ｼ繝医′縺ゅｌ縺ｰ縺昴�ｮ鬆�繝ｻ陦ｨ遉ｺ繧貞━蜈医＠縲∫┌縺�蛻励�ｯ譌｢螳夐��縺ｧ蠕後ｍ縺ｫ霑ｽ險假ｼ郁｡ｨ遉ｺ縺ｯ True�ｼ峨�
    config_dataframe 繧呈ｸ｡縺励◆蝣ｴ蜷医�ｯ繝輔ぃ繧､繝ｫ繧定ｪｭ縺ｾ縺壹◎縺ｮ蜀�螳ｹ繧貞�苓ｨｭ螳壹→縺ｿ縺ｪ縺呻ｼ�xlwings 螳溯｡梧凾逕ｨ�ｼ峨�
    謌ｻ繧雁､: (荳ｦ縺ｹ譖ｿ縺亥ｾ� DataFrame, 螳滄圀縺ｮ蛻怜錐繝ｪ繧ｹ繝�, 險ｭ螳壹た繝ｼ繧ｹ隱ｬ譏取枚蟄怜��, 蛻怜錐竊定｡ｨ遉ｺbool)
    """
    default_order = default_result_task_sheet_column_order(max_history_len)
    if config_dataframe is not None:
        user_rows = parse_result_task_column_config_dataframe(config_dataframe, max_history_len)
    else:
        user_rows = load_result_task_column_rows_from_input_workbook(max_history_len)
    if user_rows:
        primary = user_rows
        source = (
            f"繝槭け繝ｭ繝悶ャ繧ｯ縲鶏COLUMN_CONFIG_SHEET_NAME}縲�"
            if config_dataframe is None
            else f"繧ｷ繝ｼ繝医鶏COLUMN_CONFIG_SHEET_NAME}縲搾ｼ亥ｮ溯｡御ｸｭ繝悶ャ繧ｯ�ｼ�"
        )
    else:
        primary = [(n, True) for n in default_order]
        source = "譌｢螳�"

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
            "蛻苓ｨｭ螳�: 邨先棡縺ｫ辟｡縺�蛻怜錐繧堤┌隕悶＠縺ｾ縺励◆�ｼ域怙螟ｧ20莉ｶ�ｼ�: %s",
            ", ".join(unknown[:20]) + (" 窶ｦ" if len(unknown) > 20 else ""),
        )
    logging.info("邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ蛻鈴��繧ｽ繝ｼ繧ｹ: %s�ｼ�%s 蛻暦ｼ�", source, len(ordered))
    if not user_rows and config_dataframe is None:
        logging.info(
            "蛻鈴��繝ｻ陦ｨ遉ｺ縺ｮ繧ｫ繧ｹ繧ｿ繝槭う繧ｺ: 繝槭け繝ｭ螳溯｡後ヶ繝�繧ｯ縺ｫ繧ｷ繝ｼ繝医�%s縲阪ｒ霑ｽ蜉�縲�"
            " 隕句�ｺ縺励�%s縲阪�%s縲坂ｦ 陦ｨ遉ｺ縺� FALSE 縺ｮ蛻励�ｯ邨先棡繧ｷ繝ｼ繝医〒髱櫁｡ｨ遉ｺ縲�"
            " 1陦後悟ｱ･豁ｴ縲阪〒螻･豁ｴ1�ｽ柤 繧呈諺蜈･縲７BA 縺ｮ縲悟�苓ｨｭ螳喟邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ_繝√ぉ繝�繧ｯ繝懊ャ繧ｯ繧ｹ繧帝�咲ｽｮ縲阪〒繝√ぉ繝�繧ｯ繝懊ャ繧ｯ繧ｹ繧定｡ｨ遉ｺ蛻励↓騾｣蜍募庄閭ｽ縲�",
            COLUMN_CONFIG_SHEET_NAME,
            COLUMN_CONFIG_HEADER_COL,
            COLUMN_CONFIG_VISIBLE_COL,
        )
    return df[ordered], ordered, source, vis_map


def _xlwings_sheet_to_matrix(sheet) -> list:
    """xlwings Sheet 縺ｮ UsedRange 繧堤洸蠖｢縺ｮ list[list] 縺ｫ縺吶ｋ�ｼ�1陦後�ｮ縺ｿ縺ｧ繧�2谺｡蜈��ｼ峨�"""
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
    """1陦檎岼繧貞�怜錐縺ｨ縺ｿ縺ｪ縺� DataFrame 繧定ｿ斐☆縲らｩｺ縺ｪ繧� None縲�"""
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
    """邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ縲悟ｱ･豁ｴn縲榊�励°繧� n 縺ｮ譛螟ｧ繧定ｿ斐☆�ｼ育┌縺代ｌ縺ｰ 1�ｼ峨�"""
    imax = 0
    for c in columns:
        m = re.match(r"^螻･豁ｴ(\d+)$", str(c).strip())
        if m:
            imax = max(imax, int(m.group(1)))
    return max(imax, 1)


def apply_result_task_column_layout_via_xlwings(workbook_path: str | None = None) -> bool:
    """
    Excel 縺ｧ髢九＞縺ｦ縺�繧九�槭け繝ｭ繝悶ャ繧ｯ縺ｫ縺､縺�縺ｦ縲�
    縲悟�苓ｨｭ螳喟邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縲阪�ｮ蜀�螳ｹ縺ｫ蜷医ｏ縺帙※縲檎ｵ先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縲阪�ｮ蛻鈴��縺ｨ蛻鈴撼陦ｨ遉ｺ繧呈峩譁ｰ縺吶ｋ縲�
    繝悶ャ繧ｯ縺ｯ莠句燕縺ｫ菫晏ｭ倥＠縲∵悽蜃ｦ逅�荳ｭ繧� Excel 荳翫〒髢九＞縺溘∪縺ｾ縺ｫ縺吶ｋ縺薙→�ｼ�xlwings 縺梧磁邯壹☆繧具ｼ峨�
    """
    path = (workbook_path or "").strip() or TASKS_INPUT_WORKBOOK.strip()
    if not path:
        logging.error("邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ 蛻鈴←逕ｨ: 繝悶ャ繧ｯ繝代せ縺檎ｩｺ縺ｧ縺呻ｼ�TASK_INPUT_WORKBOOK 繧定ｨｭ螳壹＠縺ｦ縺上□縺輔＞�ｼ峨�")
        return False
    try:
        import xlwings as xw
    except ImportError:
        logging.error("邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ 蛻鈴←逕ｨ: xlwings 縺� import 縺ｧ縺阪∪縺帙ｓ縲Ｑip install xlwings 繧堤｢ｺ隱阪＠縺ｦ縺上□縺輔＞縲�")
        return False

    try:
        wb = xw.Book(path)
    except Exception as e:
        logging.error("邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ 蛻鈴←逕ｨ: 繝悶ャ繧ｯ縺ｫ謗･邯壹〒縺阪∪縺帙ｓ: %s", e)
        return False

    try:
        ws_res = wb.sheets[RESULT_TASK_SHEET_NAME]
        ws_cfg = wb.sheets[COLUMN_CONFIG_SHEET_NAME]
    except Exception as e:
        logging.error("邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ 蛻鈴←逕ｨ: 蠢�隕√す繝ｼ繝医′隕九▽縺九ｊ縺ｾ縺帙ｓ: %s", e)
        return False

    mat_res = _xlwings_sheet_to_matrix(ws_res)
    mat_cfg = _xlwings_sheet_to_matrix(ws_cfg)
    df_res = _matrix_to_dataframe_header_first(mat_res)
    df_cfg = _matrix_to_dataframe_header_first(mat_cfg)
    if df_res is None or df_res.empty:
        logging.error("邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ 蛻鈴←逕ｨ: 縲�%s縲阪↓繝�繝ｼ繧ｿ縺後≠繧翫∪縺帙ｓ縲�", RESULT_TASK_SHEET_NAME)
        return False
    if df_cfg is None:
        logging.error("邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ 蛻鈴←逕ｨ: 縲�%s縲阪�ｮ隕句�ｺ縺励ｒ隱ｭ繧√∪縺帙ｓ縲�", COLUMN_CONFIG_SHEET_NAME)
        return False

    max_h = _max_history_len_from_result_task_df_columns(df_res.columns)
    rows_cfg = parse_result_task_column_config_dataframe(df_cfg, max_h)
    if not rows_cfg:
        logging.error(
            "邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ 蛻鈴←逕ｨ: 縲�%s縲阪↓譛牙柑縺ｪ蛻怜錐陦後′縺ゅｊ縺ｾ縺帙ｓ縲�",
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
                logging.warning("蛻鈴撼陦ｨ遉ｺ縺ｫ螟ｱ謨暦ｼ亥��%s %s�ｼ�: %s", ci, col_name, e)

    try:
        wb.save()
    except Exception as e:
        logging.warning("邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ 蛻鈴←逕ｨ: 菫晏ｭ倥〒隴ｦ蜻奇ｼ医ョ繝ｼ繧ｿ縺ｯ繧ｷ繝ｼ繝井ｸ翫�ｯ譖ｴ譁ｰ貂医∩縺ｮ蜿ｯ閭ｽ諤ｧ�ｼ�: %s", e)

    logging.info(
        "邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ 蛻鈴←逕ｨ螳御ｺ�: %s�ｼ�%s 蛻励�髱櫁｡ｨ遉ｺ=%s�ｼ�",
        source,
        len(ordered),
        sum(1 for c in ordered if not vis_map.get(c, True)),
    )
    return True


def apply_result_task_column_layout_only() -> bool:
    """迺ｰ蠅�螟画焚 TASK_INPUT_WORKBOOK 縺ｮ繝悶ャ繧ｯ縺ｫ蟇ｾ縺怜�苓ｨｭ螳壹ｒ驕ｩ逕ｨ縺吶ｋ�ｼ�VBA 繝懊ち繝ｳ逕ｨ�ｼ峨�"""
    p = os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK
    return apply_result_task_column_layout_via_xlwings(p)

_PLAN_INPUT_XLWINGS_ORIG_ROW = "__orig_sheet_row__"


def refresh_plan_input_dispatch_trial_order_via_xlwings(
    workbook_path: str | None = None,
) -> bool:
    """
    Excel 縺ｧ髢九＞縺溘�槭け繝ｭ繝悶ャ繧ｯ蜀�縺ｮ縲碁�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙阪↓縺､縺�縺ｦ縲�
    谿ｵ髫�2 縺ｨ蜷後§ ``fill_plan_dispatch_trial_order_column_stage1`` 縺ｧ縲碁�榊床隧ｦ陦碁��逡ｪ縲阪ｒ
    蜀堺ｻ倅ｸ弱＠縲∵ｮｵ髫�1 蜃ｺ蜉帷峩蜑阪→蜷後§謇矩��縺ｧ陦後ｒ荳ｦ縺ｹ譖ｿ縺医ｋ縲�
    �ｼ域悴菫晏ｭ倥�ｮ邱ｨ髮�蛻�繧� xlwings 縺ｧ蜿肴丐縺輔○繧九◆繧� read_excel 縺ｯ菴ｿ繧上↑縺��ｼ�

    逶ｴ蜑阪↓ ``_apply_planning_sheet_post_load_mutations`` 繧貞ｮ溯｡後☆繧九◆繧√�
    ``load_planning_tasks_df``�ｼ域ｮｵ髫�2 隱ｭ霎ｼ�ｼ峨→蜷後§邨瑚ｷｯ縺ｧ縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九阪�ｮ蜷梧悄蜿肴丐繝ｻ
    蛻�蜑ｲ陦後�ｮ閾ｪ蜍暮�榊床荳崎ｦ√�ｻ繝ｫ繝ｼ繝ｫ驕ｩ逕ｨ縺瑚｡後ｏ繧後ｋ縲ゅ碁�榊床荳崎ｦ√阪ｒ謇句虚縺ｧ繧ｯ繝ｪ繧｢縺励◆陦後�ｯ縲�
    繝ｫ繝ｼ繝ｫ縺ｧ蜀阪�ｳ yes 縺ｫ縺ｪ繧峨↑縺�髯舌ｊ驟榊床蟇ｾ雎｡縺ｫ謌ｻ繧翫∬ｩｦ陦碁��縺ｫ蜷ｫ縺ｾ繧後ｋ縲�
    """
    path = (workbook_path or "").strip() or os.environ.get(
        "TASK_INPUT_WORKBOOK", ""
    ).strip() or TASKS_INPUT_WORKBOOK.strip()
    if not path:
        logging.error("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: 繝悶ャ繧ｯ繝代せ縺檎ｩｺ縺ｧ縺吶�")
        return False
    try:
        import xlwings as xw
    except ImportError:
        logging.error("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: xlwings 縺後≠繧翫∪縺帙ｓ縲�")
        return False
    try:
        wb = xw.Book(path)
        ws = wb.sheets[PLAN_INPUT_SHEET_NAME]
    except Exception as e:
        logging.error("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: 繧ｷ繝ｼ繝域磁邯壹↓螟ｱ謨�: %s", e)
        return False

    mat = _xlwings_sheet_to_matrix(ws)
    df = _matrix_to_dataframe_header_first(mat)
    if df is None or df.empty:
        logging.warning("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: 繝�繝ｼ繧ｿ陦後′縺ゅｊ縺ｾ縺帙ｓ縲�")
        return False

    df = df.copy()
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(df, plan_input_sheet_column_order())
    for c in plan_input_sheet_column_order():
        if c not in df.columns:
            df[c] = ""

    df.insert(0, _PLAN_INPUT_XLWINGS_ORIG_ROW, range(len(df)))

    _apply_planning_sheet_post_load_mutations(df, path, "驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ")

    dto_col = RESULT_TASK_COL_DISPATCH_TRIAL_ORDER
    if dto_col not in df.columns:
        logging.error("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: 蛻励�%s縲阪′縺ゅｊ縺ｾ縺帙ｓ縲�", dto_col)
        return False

    dto_idx = df.columns.get_loc(dto_col)
    if isinstance(dto_idx, slice):
        logging.error("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: 蛻励�%s縲阪′隍�謨ｰ縺ゅｊ縺ｾ縺吶�", dto_col)
        return False
    if pd.api.types.is_numeric_dtype(df[dto_col]):
        df[dto_col] = float("nan")
    else:
        df[dto_col] = ""

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
        logging.exception("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: master 隱ｭ霎ｼ縺ｫ螟ｱ謨�: %s", e)
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
        logging.exception("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: 隧ｦ陦碁��險育ｮ励↓螟ｱ謨�: %s", e)
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
        logging.exception("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: 繧ｷ繝ｼ繝域嶌霎ｼ縺ｫ螟ｱ謨�: %s", e)
        return False

    try:
        wb.save()
    except Exception as e:
        logging.warning("驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: Save 隴ｦ蜻�: %s", e)

    logging.info(
        "驟榊床隧ｦ陦碁��逡ｪ譖ｴ譁ｰ: 縲�%s縲阪ｒ %s 陦後〒譖ｴ譁ｰ縺励∪縺励◆縲�",
        PLAN_INPUT_SHEET_NAME,
        len(df_sorted),
    )
    return True


def refresh_plan_input_dispatch_trial_order_only() -> bool:
    """TASK_INPUT_WORKBOOK 縺ｫ蟇ｾ縺吶ｋ驟榊床隧ｦ陦碁��逡ｪ蜀崎ｨ育ｮ暦ｼ�VBA / cmd 邨檎罰縺ｮ繧ｨ繝ｳ繝医Μ�ｼ峨�"""
    p = os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK
    return refresh_plan_input_dispatch_trial_order_via_xlwings(p)


def apply_plan_input_column_layout_only() -> bool:
    """
    驟榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙�ｮ蛻鈴��繝ｻ陦ｨ遉ｺ縺ｮ縺ｿ繧帝←逕ｨ縺吶ｋ莠亥ｮ夲ｼ�VBA 逕ｨ�ｼ峨�
    譛ｪ螳溯｣�縲ょ�励�ｮ荳ｦ縺ｳ縺ｯ谿ｵ髫�1蜃ｺ蜉帙∪縺溘�ｯ謇句虚謨ｴ逅�繧剃ｽｿ逕ｨ縺励※縺上□縺輔＞縲�
    """
    logging.warning("apply_plan_input_column_layout_only: not implemented")
    return False



def dedupe_result_task_column_config_sheet_via_xlwings(workbook_path: str | None = None) -> bool:
    """
    縲悟�苓ｨｭ螳喟邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縲阪�ｮ A:B 縺�縺代ｒ縲�驥崎､�蛻怜錐繧帝勁縺�縺滉ｸ隕ｧ縺ｧ譖ｸ縺咲峩縺呻ｼ亥�医�ｮ陦後ｒ蜆ｪ蜈茨ｼ峨�
    縲檎ｵ先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縲阪′縺ゅｌ縺ｰ螻･豁ｴ蛻玲焚縺ｮ隗｣驥医↓菴ｿ縺�縲らｵ先棡繧ｷ繝ｼ繝医�ｯ螟画峩縺励↑縺�縲�
    """
    path = (workbook_path or "").strip() or TASKS_INPUT_WORKBOOK.strip()
    if not path:
        logging.error("蛻苓ｨｭ螳� 驥崎､�謨ｴ逅�: 繝悶ャ繧ｯ繝代せ縺檎ｩｺ縺ｧ縺吶�")
        return False
    try:
        import xlwings as xw
    except ImportError:
        logging.error("蛻苓ｨｭ螳� 驥崎､�謨ｴ逅�: xlwings 縺� import 縺ｧ縺阪∪縺帙ｓ縲�")
        return False
    try:
        wb = xw.Book(path)
        ws_cfg = wb.sheets[COLUMN_CONFIG_SHEET_NAME]
    except Exception as e:
        logging.error("蛻苓ｨｭ螳� 驥崎､�謨ｴ逅�: 謗･邯壹∪縺溘�ｯ繧ｷ繝ｼ繝亥叙蠕励↓螟ｱ謨�: %s", e)
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
        logging.error("蛻苓ｨｭ螳� 驥崎､�謨ｴ逅�: 縲�%s縲阪�ｮ隕句�ｺ縺励ｒ隱ｭ繧√∪縺帙ｓ縲�", COLUMN_CONFIG_SHEET_NAME)
        return False
    rows = parse_result_task_column_config_dataframe(df_cfg, max_h)
    if not rows:
        logging.warning("蛻苓ｨｭ螳� 驥崎､�謨ｴ逅�: 譛牙柑縺ｪ繝�繝ｼ繧ｿ陦後′縺ゅｊ縺ｾ縺帙ｓ縲�")
        return False
    _xlwings_write_column_config_sheet_ab(ws_cfg, rows)
    try:
        wb.save()
    except Exception as e:
        logging.warning("蛻苓ｨｭ螳� 驥崎､�謨ｴ逅�: 菫晏ｭ倩ｭｦ蜻�: %s", e)
    logging.info(
        "蛻苓ｨｭ螳壹�%s縲阪ｒ驥崎､�髯､蜴ｻ貂医∩縺ｧ %s 陦後↓謨ｴ逅�縺励∪縺励◆�ｼ亥ｱ･豁ｴ螻暮幕蠕後�ｮ陦梧焚�ｼ峨�",
        COLUMN_CONFIG_SHEET_NAME,
        len(rows),
    )
    return True


def dedupe_result_task_column_config_sheet_only() -> bool:
    """迺ｰ蠅�螟画焚 TASK_INPUT_WORKBOOK 縺ｮ繝悶ャ繧ｯ縺ｮ蛻苓ｨｭ螳壹す繝ｼ繝医□縺鷹㍾隍�謨ｴ逅��ｼ�VBA 逕ｨ�ｼ峨�"""
    p = os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK
    return dedupe_result_task_column_config_sheet_via_xlwings(p)


def _apply_result_task_sheet_column_visibility(worksheet, column_names: list, vis_map: dict):
    """邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｧ縲」is_map 縺� False 縺ｮ蛻励ｒ髱櫁｡ｨ遉ｺ縺ｫ縺吶ｋ縲�"""
    for idx, col_name in enumerate(column_names, 1):
        if not vis_map.get(col_name, True):
            worksheet.column_dimensions[get_column_letter(idx)].hidden = True


def _norm_history_member_label(name: str) -> str:
    """螻･豁ｴ縺ｮ諡�蠖灘錐豈碑ｼ�逕ｨ�ｼ亥�ｨ隗堤ｩｺ逋ｽ繧貞濠隗�1蛟句喧繝ｻ蜑榊ｾ荊rim繝ｻ騾｣邯夂ｩｺ逋ｽ縺ｮ蝨ｧ邵ｮ�ｼ峨�"""
    t = str(name or "").replace("\u3000", " ").strip()
    return " ".join(t.split())


def _history_team_text_main_assignment_only(h: dict) -> str:
    """
    邨先棡繧ｷ繝ｼ繝医梧球蠖薙肴ｬ�逕ｨ: 繝｡繧､繝ｳ蜑ｲ莉倡｢ｺ螳壽凾轤ｹ縺ｮ蜷榊燕�ｼ井ｽ吝鴨霑ｽ險倥し繝悶�ｯ蜷ｫ繧√↑縺��ｼ峨�
    append_surplus 蠕後�ｮ h['team'] 縺九ｉ post_dispatch_surplus_names 繧帝勁螟悶☆繧九�
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
    """邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ螻･豁ｴ繧ｻ繝ｫ譁�蟄怜�暦ｼ育ｵ�蜷医○陦ｨ縺ｮ謗｡逕ｨ陦栗D繝ｻ繝｡繧､繝ｳ霑ｽ蜉�莠ｺ謨ｰ繝ｻ菴吝鴨霑ｽ險倥�ｮ譏守､ｺ繧貞性繧�ｼ峨�"""
    um = task.get("unit_m") or 0
    try:
        done_r = int(h["done_m"] / um) if um else 0
    except (TypeError, ValueError, ZeroDivisionError):
        done_r = 0
    dm = h.get("done_m", 0)
    parts_out: list[str] = [f"繝ｻ縲須h.get('date', '')}縲托ｼ嘴done_r}R ({dm}m)"]
    cid = h.get("combo_sheet_row_id")
    if cid is not None:
        try:
            parts_out.append(f"邨�蜷医○陦ｨ#{int(cid)}")
        except (TypeError, ValueError):
            parts_out.append(f"邨�蜷医○陦ｨ#{cid}")
    parts_out.append(f"諡�蠖甜{_history_team_text_main_assignment_only(h)}]")
    sm = h.get("surplus_member_names") or []
    if sm:
        parts_out.append(f"霑ｽ蜉�[{','.join(str(x) for x in sm)}]")
    ps = h.get("post_dispatch_surplus_names") or []
    if ps:
        parts_out.append(f"菴吝鴨霑ｽ險麓{','.join(str(x) for x in ps)}]")
    return " ".join(parts_out)


_RESULT_TASK_HISTORY_RICH_HEAD_RE = re.compile(r"^繝ｻ(縲深^縲曽*縲�)(.*)$", re.DOTALL)


def _apply_result_task_history_rich_text(worksheet, column_names: list):
    """
    螻･豁ｴ蛻�: 縲後�ｻ縲先律莉倥托ｼ壺ｦ縲阪�ｮ譌･莉俶峡蠑ｧ驛ｨ蛻�繧帝搨濶ｲ繝ｪ繝�繝√ユ繧ｭ繧ｹ繝医↓縺吶ｋ縲�
    openpyxl 3.1 譛ｪ貅縺ｧ縺ｯ繧ｹ繧ｭ繝�繝暦ｼ域枚蟄怜�励�ｮ縲舌代�ｮ縺ｿ�ｼ峨�
    """
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        from openpyxl.styles.colors import Color
    except ImportError:
        return

    hist_cols = [
        i + 1 for i, c in enumerate(column_names) if str(c).startswith("螻･豁ｴ")
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
            if not isinstance(v, str) or not v.startswith("繝ｻ縲�"):
                continue
            m = _RESULT_TASK_HISTORY_RICH_HEAD_RE.match(v)
            if not m:
                continue
            bracketed, rest = m.group(1), m.group(2)
            cell.value = CellRichText(
                TextBlock(plain_if, "繝ｻ"),
                TextBlock(blue_if, bracketed),
                TextBlock(plain_if, rest),
            )
            cell.alignment = top


def _apply_result_task_date_columns_blue_font(worksheet, column_names: list):
    """
    邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ: 蝗樒ｭ皮ｴ肴悄繝ｻ謖�螳夂ｴ肴悄繝ｻ險育判蝓ｺ貅也ｴ肴悄繝ｻ蜴溷渚謚募�･譌･繝ｻ蜉�蟾･髢句ｧ区律縺ｮ繧ｻ繝ｫ繧帝搨濶ｲ縺ｫ縺吶ｋ縲�
    �ｼ亥ｱ･豁ｴ蛻励�ｮ縲先律莉倥代�ｯ _apply_result_task_history_rich_text 蛛ｴ縲り牡縺ｯ 0070C0 縺ｧ邨ｱ荳�ｼ�
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
    need縲碁�榊床譎りｿｽ蜉�莠ｺ謨ｰ縲咲嶌蠖薙〒蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ繧定ｶ�縺医※謗｡逕ｨ縺励◆繝悶Ο繝�繧ｯ縲√∪縺溘�ｯ
    繝｡繧､繝ｳ螳御ｺ�蠕後�ｮ菴吝鴨霑ｽ險倥〒繧ｵ繝悶′蠅励∴縺溘ヶ繝ｭ繝�繧ｯ縺ｫ蟇ｾ蠢懊☆繧九悟ｱ･豁ｴn縲阪そ繝ｫ繧定埋鮟�縺ｫ蝪励ｋ縲�
    """
    hist_cols: list[tuple[int, int]] = []
    for col_idx, col_name in enumerate(column_names, 1):
        m = re.match(r"^螻･豁ｴ(\d+)$", str(col_name).strip())
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
    蜉�蟾･蜀�螳ｹ縺ｫ蟾･遞句錐縺悟性縺ｾ繧後↑縺�陦後�ｮ縲後ち繧ｹ繧ｯID縲阪そ繝ｫ繧定ｵ､閭梧勹繝ｻ逋ｽ譁�蟄励↓縺吶ｋ�ｼ亥��繝�繝ｼ繧ｿ荳肴紛蜷医�ｮ隕冶ｪ咲畑�ｼ峨�
    """
    task_id_col_idx = None
    for col_idx, col_name in enumerate(column_names, 1):
        if str(col_name) == "繧ｿ繧ｹ繧ｯID":
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
    蛻励碁�榊ｮ契蝗樒ｭ疲欠螳�16譎ゅ∪縺ｧ縲阪′縲後＞縺�縺医阪�ｮ繧ｻ繝ｫ繧定ｵ､閭梧勹繝ｻ逋ｽ譁�蟄励�ｻ螟ｪ蟄励↓縺吶ｋ縲�
    蛻苓ｨｭ螳壹〒譌ｧ蜷阪碁�榊ｮ契蝓ｺ貅�16譎ゅ∪縺ｧ縲阪�ｮ縺ｾ縺ｾ縺ｮ隕句�ｺ縺励↓繧ょｯｾ蠢懊�
    """
    target_names = frozenset(
        {
            RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16,
            "驟榊ｮ契蝓ｺ貅�16譎ゅ∪縺ｧ",
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
        if s != "縺�縺�縺�":
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
    邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ縲後ち繧ｹ繧ｯID縲阪そ繝ｫ縺ｫ縲∫ｵ先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ縺ｧ蠖楢ｩｲ繧ｿ繧ｹ繧ｯ縺梧怙蛻昴↓迴ｾ繧後ｋ繧ｻ繝ｫ縺ｸ縺ｮ蜀�驛ｨ繝上う繝代�ｼ繝ｪ繝ｳ繧ｯ繧剃ｻ倅ｸ弱☆繧九�
    譎る俣蜑ｲ縺ｫ迴ｾ繧後↑縺�繧ｿ繧ｹ繧ｯ�ｼ域悴蜑ｲ蠖薙�ｮ縺ｿ遲会ｼ峨�ｯ繝ｪ繝ｳ繧ｯ縺ｪ縺励�
    """
    if not task_id_to_schedule_cell or worksheet_tasks.max_row < 2:
        return
    task_id_col_idx = None
    for col_idx, col_name in enumerate(column_names, 1):
        if str(col_name) == "繧ｿ繧ｹ繧ｯID":
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
    """陦ｨ遉ｺ蛻励↓ TRUE/FALSE 繝ｪ繧ｹ繝茨ｼ医メ繧ｧ繝�繧ｯ縺ｮ莉｣繧上ｊ縺ｫ繝励Ν繝繧ｦ繝ｳ�ｼ峨ｒ莉倅ｸ弱�"""
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
    pandas/openpyxl 縺ｧ譁ｰ隕丈ｽ懈�舌＠縺溽ｵ先棡繝悶ャ繧ｯ縺ｫ縺ｯ蝗ｳ蠖｢縺悟性縺ｾ繧後↑縺�縲�
    譌｢螳壹〒譛牙柑�ｼ育腸蠅�螟画焚縺ｧ 0/false/no/off 縺ｮ縺ｨ縺咲┌蜉ｹ�ｼ峨ょ�･蜉帙ヶ繝�繧ｯ縺ｮ
    縲悟�苓ｨｭ螳喟邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縲堺ｸ翫�ｮ Shapes 繧堤ｵ先棡繝悶ャ繧ｯ縺ｮ蜷悟錐繧ｷ繝ｼ繝医∈繧ｳ繝斐�ｼ縺励�
    蜷�蝗ｳ蠖｢縺ｮ Left/Top/Width/Height�ｼ医♀繧医�ｳ蜿悶ｌ繧九→縺� Placement�ｼ峨ｒ蜈･蜉帛�ｴ縺ｨ蜷後§縺ｫ謌ｻ縺吶�
    openpyxl 縺ｫ繧医ｋ蠖楢ｩｲ繝悶ャ繧ｯ縺ｸ縺ｮ菫晏ｭ倥′縺吶∋縺ｦ邨ゅｏ縺｣縺溷ｾ後↓蜻ｼ縺ｶ縺薙→縲�
    """
    if not STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT:
        return
    rp = (result_path or "").strip()
    ip = (input_path or "").strip()
    if not rp or not os.path.isfile(rp):
        logging.warning(
            "蛻苓ｨｭ螳壹す繝ｼ繝亥峙蠖｢繧ｳ繝斐�ｼ: 邨先棡繝代せ縺檎┌蜉ｹ縺ｮ縺溘ａ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆縲�"
        )
        return
    if not ip or not os.path.isfile(ip):
        logging.warning(
            "蛻苓ｨｭ螳壹す繝ｼ繝亥峙蠖｢繧ｳ繝斐�ｼ: TASK_INPUT_WORKBOOK 縺檎┌蜉ｹ縺ｮ縺溘ａ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆縲�"
        )
        return
    try:
        import xlwings as xw
    except ImportError:
        logging.warning(
            "蛻苓ｨｭ螳壹す繝ｼ繝亥峙蠖｢繧ｳ繝斐�ｼ: xlwings 縺� import 縺ｧ縺阪∪縺帙ｓ縲�"
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
                "蛻苓ｨｭ螳壹す繝ｼ繝亥峙蠖｢繧ｳ繝斐�ｼ: 邨先棡繝悶ャ繧ｯ縺ｫ繧ｷ繝ｼ繝医�%s縲阪′縺ゅｊ縺ｾ縺帙ｓ縲�",
                COLUMN_CONFIG_SHEET_NAME,
            )
            return
        try:
            ws_in = wb_in.sheets[COLUMN_CONFIG_SHEET_NAME]
        except Exception:
            logging.warning(
                "蛻苓ｨｭ螳壹す繝ｼ繝亥峙蠖｢繧ｳ繝斐�ｼ: 蜈･蜉帙ヶ繝�繧ｯ縺ｫ繧ｷ繝ｼ繝医�%s縲阪′縺ゅｊ縺ｾ縺帙ｓ縲�",
                COLUMN_CONFIG_SHEET_NAME,
            )
            return
        n_shapes = int(ws_in.api.Shapes.Count)
        if n_shapes <= 0:
            logging.info(
                "蛻苓ｨｭ螳壹す繝ｼ繝亥峙蠖｢繧ｳ繝斐�ｼ: 蜈･蜉帛�ｴ縺ｫ蝗ｳ蠖｢縺後≠繧翫∪縺帙ｓ�ｼ医せ繧ｭ繝�繝暦ｼ峨�"
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
            "蛻苓ｨｭ螳壹す繝ｼ繝亥峙蠖｢繧ｳ繝斐�ｼ: 蜈･蜉帙°繧� %s 蛟九�ｮ蝗ｳ蠖｢繧堤ｵ先棡繝悶ャ繧ｯ縺ｸ隍�陬ｽ縺励∪縺励◆縲�",
            n_shapes,
        )
    except Exception as e:
        logging.warning(
            "蛻苓ｨｭ螳壹す繝ｼ繝亥峙蠖｢繧ｳ繝斐�ｼ: 螟ｱ謨励＠縺ｾ縺励◆�ｼ�%s�ｼ峨�Excel 蜊�譛峨�ｻCOM 繧ｨ繝ｩ繝ｼ遲峨�ｮ蜿ｯ閭ｽ諤ｧ縺後≠繧翫∪縺吶�",
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
    """蜉�蟾･螳溽ｸｾDATA 縺ｮ1陦後°繧� (髢句ｧ�, 邨ゆｺ�) 繧貞ｾ励ｋ縲りｧ｣縺代↑縺代ｌ縺ｰ (None, None)縲�"""
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
    繝槭け繝ｭ繝悶ャ繧ｯ縺ｮ縲悟刈蟾･螳溽ｸｾDATA縲阪ｒ隱ｭ繧�ｼ育┌縺代ｌ縺ｰ遨ｺ DataFrame�ｼ峨�
    Power Query 遲峨〒逕ｨ諢上＠縺溘す繝ｼ繝医ｒ諠ｳ螳壹�
    """
    if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
        return pd.DataFrame()
    try:
        df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=ACTUALS_SHEET_NAME)
    except ValueError:
        logging.info(
            f"繧ｷ繝ｼ繝医鶏ACTUALS_SHEET_NAME}縲阪′辟｡縺�縺溘ａ縲√ぎ繝ｳ繝医�ｮ螳溽ｸｾ陦後�ｯ蜃ｺ蜉帙＠縺ｾ縺帙ｓ縲�"
        )
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(df, ACTUAL_HEADER_CANONICAL)
    logging.info(
        f"蜉�蟾･螳溽ｸｾ: '{TASKS_INPUT_WORKBOOK}' 縺ｮ '{ACTUALS_SHEET_NAME}' 繧� {len(df)} 陦瑚ｪｭ縺ｿ霎ｼ縺ｿ縲�"
    )
    return df


def build_actual_timeline_events(df, equipment_list, sorted_dates):
    """
    螳溽ｸｾ繧ｷ繝ｼ繝医�ｮ蜷�陦後ｒ繧ｬ繝ｳ繝育畑繧､繝吶Φ繝医∈螟画鋤縲�
    險育判陦ｨ遉ｺ譌･�ｼ�sorted_dates�ｼ峨°縺､險ｭ蛯吶�槭せ繧ｿ縺ｫ荳閾ｴ縺吶ｋ縲悟ｷ･遞句錐縲阪□縺大ｯｾ雎｡縲�
    蟾･遞句錐縺ｯ NFKC繝ｻ遨ｺ逋ｽ豁｣隕丞喧蠕後↓繝槭せ繧ｿ蛻怜錐縺ｸ繝槭ャ繝斐Φ繧ｰ縺吶ｋ縲�
    譎ょ綾縺ｯ DEFAULT_START_TIME / DEFAULT_END_TIME 縺ｮ譫�蜀�縺ｫ繧ｯ繝ｪ繝�繝励�
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
            f"蜉�蟾･螳溽ｸｾDATA: 蟾･遞句錐縺後�槭せ繧ｿ險ｭ蛯吶→荳閾ｴ縺励↑縺�陦後ｒ {bad_eq} 莉ｶ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆�ｼ育ｩｺ逋ｽ遲峨�ｯ豁｣隕丞喧貂医∩�ｼ峨�"
        )
        if mismatch_norm_samples:
            logging.info(
                "  荳堺ｸ閾ｴ縺ｨ縺ｪ縺｣縺溷ｷ･遞句錐縺ｮ豁｣隕丞喧蠕後し繝ｳ繝励Ν: "
                + " | ".join(mismatch_norm_samples[:12])
            )
    if bad_time:
        logging.info(
            f"蜉�蟾･螳溽ｸｾDATA: 髢句ｧ�/邨ゆｺ�譌･譎ゅ′隗｣驥医〒縺阪↑縺�陦後ｒ {bad_time} 莉ｶ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆縲�"
        )
    if no_plan_overlap and sorted_dates:
        logging.info(
            f"蜉�蟾･螳溽ｸｾDATA: 險ｭ蛯吶�ｻ譌･譎ゅ�ｯ譛牙柑縺�縺後∬ｨ育判蟇ｾ雎｡譌･�ｼ亥ｽ捺律莉･髯阪�ｮ蜍､諤�譌･ﾃ養DEFAULT_START_TIME}�ｽ桀DEFAULT_END_TIME}�ｼ峨→驥阪↑繧峨↑縺�陦後′ {no_plan_overlap} 莉ｶ縺ゅｊ縺ｾ縺励◆縲�"
        )
    if not events and len(df) > 0:
        logging.info(
            "蜉�蟾･螳溽ｸｾDATA: 繧ｬ繝ｳ繝育畑繧ｻ繧ｰ繝｡繝ｳ繝医′0莉ｶ縺ｧ縺吶る℃蜴ｻ譌･縺ｮ螳溽ｸｾ縺ｮ縺ｿ縺ｮ蝣ｴ蜷医∬ｨ育判縺ｮ陦ｨ遉ｺ譌･�ｼ�sorted_dates�ｼ峨↓蜷ｫ縺ｾ繧後↑縺�縺溘ａ謠冗判縺輔ｌ縺ｾ縺帙ｓ縲�"
        )
    logging.info(f"蜉�蟾･螳溽ｸｾDATA 縺九ｉ繧ｬ繝ｳ繝育畑繧ｻ繧ｰ繝｡繝ｳ繝� {len(events)} 莉ｶ繧堤函謌舌＠縺ｾ縺励◆縲�")
    return events


TASK_SPECIAL_AI_LAST_RESPONSE_FILE = "ai_task_special_remark_last.txt"
# 蜍､諤�蛯呵�繧ｭ繝｣繝�繧ｷ繝･縺ｨ繧ｭ繝ｼ遨ｺ髢薙ｒ蛻�髮｢�ｼ亥酔荳SHA陦晉ｪ√ｒ驕ｿ縺代ｋ�ｼ峨よ欠邏九↓蝓ｺ貅門ｹｴ繧貞性繧∵律莉倩ｧ｣驥医�ｮ繧ｺ繝ｬ繧帝亟縺舌�
TASK_SPECIAL_CACHE_KEY_PREFIX = "TASK_SPECIAL_v3|"
# 繝｡繧､繝ｳ繧ｷ繝ｼ繝医後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医堺ｸ九�ｮ閾ｪ逕ｱ險倩ｿｰ 竊� Gemini 隗｣驥茨ｼ磯�榊床縺ｮ譛蜆ｪ蜈医が繝ｼ繝舌�ｼ繝ｩ繧､繝会ｼ�
GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX = "GLOBAL_PRIO_v8|"


def _normalize_special_task_id_for_ai(val):
    """
    萓晞�ｼNO繧偵く繝｣繝�繧ｷ繝･繧ｭ繝ｼ繝ｻ繝励Ο繝ｳ繝励ヨ陦後〒荳雋ｫ縺輔○繧九�
    Excel 縺ｮ謨ｰ蛟､繧ｻ繝ｫ縺ｯ float 縺ｫ縺ｪ繧翫′縺｡縺ｪ縺ｮ縺ｧ 12345.0 竊� \"12345\" 縺ｫ謠�縺医ｋ縲�
    譁�蟄怜�励�ｯ NFKC�ｼ亥�ｨ隗定恭謨ｰ蟄励↑縺ｩ�ｼ峨〒陦ｨ險倥ｆ繧後ｒ蜷ｸ蜿趣ｼ亥酔荳螳滉ｽ薙�ｮ蜀喉PI蜻ｼ縺ｳ蜃ｺ縺励ｒ貂帙ｉ縺呻ｼ峨�
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
    # 譁�蟄怜�励→縺励※縺ｮ "20010.0" 遲会ｼ�Excel繝ｻCSV�ｼ峨ｒ謨ｴ謨ｰ陦ｨ險倥�ｮ萓晞�ｼNO縺ｫ蟇�縺帙ｋ
    if re.fullmatch(r"-?\d+\.0+", s):
        try:
            return str(int(float(s)))
        except ValueError:
            pass
    return s or None


def planning_task_id_str_from_scalar(val) -> str:
    """驟榊床繝ｻ谿ｵ髫�1繝槭�ｼ繧ｸ繝ｻ繧ｭ繝･繝ｼ讒狗ｯ峨〒逕ｨ縺�繧倶ｾ晞�ｼNO縺ｮ螳牙ｮ壽枚蟄怜�暦ｼ育ｩｺ縺ｪ繧� \"\"�ｼ峨�"""
    return _normalize_special_task_id_for_ai(val) or ""


def planning_task_id_str_from_plan_row(row) -> str:
    """驥崎､�隕句�ｺ縺怜�励〒繧ょ�磯�ｭ繧ｹ繧ｫ繝ｩ繝ｼ繧呈鏡縺�縲∽ｾ晞�ｼNO繧� planning_task_id_str_from_scalar 縺ｫ貂｡縺吶�"""
    return planning_task_id_str_from_scalar(_planning_df_cell_scalar(row, TASK_COL_TASK_ID))


def _cell_text_task_special_remark(val):
    """
    迚ｹ蛻･謖�螳喟蛯呵�繧偵�励Ο繝ｳ繝励ヨ逕ｨ縺ｫ蜿悶ｊ蜃ｺ縺吶ゆｻ墓ｧ倥←縺翫ｊ **strip 縺ｮ縺ｿ**
    �ｼ亥�磯�ｭ譛ｫ蟆ｾ縺ｮ遨ｺ逋ｽ繝ｻExcel 縺ｮ蛛ｽ遨ｺ逋ｽ繧帝勁縺阪∵枚荳ｭ縺ｮ謾ｹ陦後�ｻ繧ｹ繝壹�ｼ繧ｹ縺ｯ菫晄戟縲よ焚蛟､繧ｻ繝ｫ縺ｯ陦ｨ險倥ｒ蝗ｺ螳夲ｼ峨�
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
        # 蛯呵�蛻励↓謨ｰ蛟､縺�縺大�･縺｣縺ｦ縺�繧句�ｴ蜷医�ｮ陦ｨ險倥ｆ繧後ｒ貂帙ｉ縺�
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
    """繝励Ο繝ｳ繝励ヨ縺ｫ霈峨○繧玖｡後Μ繧ｹ繝茨ｼ医た繝ｼ繝亥燕�ｼ峨よｭ｣隕丞喧縺ｯ荳願ｨ倥�倥Ν繝代�ｼ縺ｫ邨ｱ荳縲�"""
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
        proc_disp = proc if proc else "�ｼ育ｩｺ�ｼ�"
        macn_disp = macn if macn else "�ｼ育ｩｺ�ｼ�"
        lines.append(
            f"- 萓晞�ｼNO縲須tid}縲掃 蟾･遞句錐縲鶏proc_disp}縲� | 讖滓｢ｰ蜷阪鶏macn_disp}縲� | 蛯呵�譛ｬ譁�: {rem}"
        )
    return lines


def _repair_task_special_ai_wrong_top_level_keys(parsed: dict, tasks_df) -> dict:
    """
    蛯呵�縺悟刀逡ｪ繝ｻ蜴溷渚繧ｳ繝ｼ繝会ｼ井ｾ�: 20010 縺ｧ蟋九∪繧区焚蟄怜�暦ｼ峨〒蟋九∪繧九→縲√Δ繝�繝ｫ縺後◎縺ｮ蛻励ｒ JSON 繝医ャ繝励く繝ｼ縺ｫ
    隱､逕ｨ縺吶ｋ縺薙→縺後≠繧九ゆｾ晞�ｼNO縲絶ｦ縲代→荳閾ｴ縺励↑縺�謨ｰ蟄励�ｮ縺ｿ縺ｮ繧ｭ繝ｼ繧偵∝ｽ楢ｩｲ蛯呵�繧呈戟縺､陦後�ｮ萓晞�ｼNO縺ｸ莉倥￠譖ｿ縺医ｋ縲�
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
                or r.startswith(sk + "繝ｼ")
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
                "繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: JSON 繝医ャ繝励く繝ｼ隱､繧翫ｒ菫ｮ蠕ｩ�ｼ�%r 縺ｯ萓晞�ｼNO縺ｧ縺ｯ縺ｪ縺� 竊� %r�ｼ�",
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
    restrict 縺檎┌縺�繝ｻ遨ｺ縺ｪ繧牙宛髯舌↑縺暦ｼ�True�ｼ峨�
    髱樒ｩｺ縺ｪ繧� Excel 蛛ｴ縺ｮ蛟､縺ｨ縺ゅ＞縺ｾ縺�荳閾ｴ�ｼ磯Κ蛻�荳閾ｴ蜿ｯ�ｼ峨�
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
    """restrict_to_* 縺檎┌縺�縺ｨ縺阪�ｯ蜷御ｸ萓晞�ｼNO縺ｮ蜈ｨ陦後↓驕ｩ逕ｨ縲�"""
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
    JSON 縺ｮ process_name / machine_name 縺後∝ｽ楢ｩｲ Excel 陦後�ｮ蟾･遞句錐繝ｻ讖滓｢ｰ蜷阪→荳閾ｴ縺吶ｋ縺九�
    �ｼ医�励Ο繝ｳ繝励ヨ縺ｧ貂｡縺励◆縲悟ｙ閠�縺後≠縺｣縺溯｡後阪→蟇ｾ蠢懊▼縺代ｋ縲ら援譁ｹ縺�縺台ｸ閾ｴ縺ｧ繧ょ庄�ｼ�
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
    """restrict_to_* 縺檎┌縺�繝ｻ遨ｺ�ｼ晏酔荳萓晞�ｼNO縺ｮ蜈ｨ蟾･遞玖｡後↓蜉ｹ縺九○繧区欠螳壹�"""
    if not isinstance(entry, dict):
        return False
    a = _normalize_task_special_scope_str(entry.get("restrict_to_process_name"))
    b = _normalize_task_special_scope_str(entry.get("restrict_to_machine_name"))
    return not a and not b


def _select_ai_task_special_entry_for_tid_value(val, row):
    """1萓晞�ｼNO縺ｫ蟇ｾ縺吶ｋ蛟､縺� dict 縺ｾ縺溘�ｯ dict 縺ｮ驟榊�励�ｮ縺ｩ縺｡繧峨〒繧り｡後↓蜷医≧隕∫ｴ�繧定ｿ斐☆縲�"""
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
    analyze_task_special_remarks 縺ｮ謌ｻ繧翫°繧牙ｽ楢ｩｲ陦後�ｮ繧ｨ繝ｳ繝医Μ繧貞叙繧九�
    繝励Ο繝ｳ繝励ヨ繧ｭ繝ｼ縺ｯ豁｣隕丞喧貂医∩萓晞�ｼNO縺ｪ縺ｮ縺ｧ縲・xcel 縺� 12345.0 縺ｧ繧ゅヲ繝�繝医☆繧九�
    restrict_to_process_name / restrict_to_machine_name 縺檎┌縺�繝ｻ遨ｺ縺ｮ縺ｨ縺阪�ｯ
    蜷御ｸ萓晞�ｼNO縺ｮ蟾･遞九�ｻ讖滓｢ｰ縺檎焚縺ｪ繧句�ｨ陦後↓蜷後§謖�遉ｺ繧帝←逕ｨ縺吶ｋ縲�
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


# 1 蝗槭�ｮ Python 螳溯｡鯉ｼ域ｮｵ髫�1 縺ｾ縺溘�ｯ 谿ｵ髫�2�ｼ牙腰菴阪〒繝ｪ繧ｻ繝�繝医☆繧�
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
    """API_Payment 蜀�縺ｮ邏ｯ險� JSON 繧定ｪｭ繧縲ら┌縺�繝ｻ螢翫ｌ縺ｦ縺�繧後�ｰ蛻晄悄蠖｢繧定ｿ斐☆縲�"""
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
        logging.debug("Gemini 邏ｯ險� JSON 縺ｮ菫晏ｭ倥↓螟ｱ謨�: %s", ex)


def _gemini_buckets_ensure_structure(data: dict) -> None:
    """邏ｯ險� JSON 縺ｫ譛滄俣蛻･繝舌こ繝�繝育畑縺ｮ霎樊嶌繧堤畑諢上☆繧具ｼ域里蟄� v1 繝輔ぃ繧､繝ｫ繧ゅ�槭�ｼ繧ｸ�ｼ峨�"""
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
        "period_key 縺ｯ PC 繝ｭ繝ｼ繧ｫ繝ｫ譎ょ綾�ｼ�datetime.now�ｼ峨〒莉倅ｸ弱ゆｻ� PC 縺ｨ縺ｮ髮�險医�ｯ豺ｷ縺懊↑縺�縺ｧ縺上□縺輔＞縲�",
    )


def _gemini_time_bucket_keys(dt: datetime) -> tuple[str, str, str, str, str]:
    """蟷ｴ繝ｻ譛医�ｻISO騾ｱ繝ｻ譌･繝ｻ譎� 縺ｮ繧ｭ繝ｼ�ｼ域枚蟄怜�励た繝ｼ繝医〒譎らｳｻ蛻玲ｯ碑ｼ�縺励ｄ縺吶＞蠖｢�ｼ峨�"""
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
    """1 蝗槭�ｮ API 蜻ｼ蜃ｺ縺励ｒ蟷ｴ繝ｻ譛医�ｻ騾ｱ繝ｻ譌･繝ｻ譎ゅ�ｮ蜷�繝舌こ繝�繝医↓蜉�邂励☆繧九�"""
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
    """1 蝗槭�ｮ API 蠢懃ｭ泌��繧堤ｴｯ險� JSON 縺ｫ蜉�邂励☆繧具ｼ医Ο繧ｰ縺ｫ蜊倡匱譁咎≡縺ｯ蜃ｺ縺輔↑縺��ｼ峨�"""
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
    """generate_content 縺ｮ蠢懃ｭ斐°繧� usage_metadata 繧帝寔險医☆繧具ｼ医そ繝�繧ｷ繝ｧ繝ｳ�ｼ狗ｴｯ險� JSON�ｼ峨�"""
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
        logging.debug("Gemini 邏ｯ險医�ｮ譖ｴ譁ｰ縺ｧ萓句､厄ｼ育ｶ夊｡鯉ｼ�: %s", ex)


def _gemini_estimate_cost_usd(
    model_id: str, prompt_tok: int, cand_tok: int, thoughts_tok: int
) -> float | None:
    m = str(model_id).strip().lower()
    rin, rout = None, None
    if "flash" in m:
        rin, rout = _GEMINI_FLASH_IN_PER_M, _GEMINI_FLASH_OUT_PER_M
    elif "pro" in m:
        # 逶ｮ螳会ｼ域悴菴ｿ逕ｨ繝｢繝�繝ｫ蜷代￠繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ�ｼ�
        rin, rout = 1.25, 5.0
    if rin is None:
        return None
    out_equiv = cand_tok + thoughts_tok
    return (prompt_tok / 1_000_000.0) * rin + (out_equiv / 1_000_000.0) * rout


def _gemini_daily_trend_series(
    cum: dict, *, max_days: int | None = None
) -> tuple[list[str], list[float], str] | None:
    """邏ｯ險� JSON 縺ｮ by_day 縺九ｉ縲∵律莉倥く繝ｼ�ｼ亥商竊呈眠�ｼ峨�ｻ蛟､繝ｻ邉ｻ蛻怜錐縲ら┌縺代ｌ縺ｰ None縲�"""
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
    label = "蜻ｼ蜃ｺ縺怜屓謨ｰ" if use_calls else "謗ｨ螳啅SD"
    return (keys, series, label)


def _gemini_daily_total_tokens_for_days(cum: dict, day_keys: list[str]) -> list[int]:
    """by_day 縺ｮ蜷�繧ｭ繝ｼ縺ｫ縺､縺�縺ｦ縲》otal_tokens�ｼ育┌縺代ｌ縺ｰ prompt+candidates+thoughts�ｼ峨ｒ霑斐☆縲�"""
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
    """繝�繧ｭ繧ｹ繝亥�ｴ縺ｯ繧ｰ繝ｩ繝募盾辣ｧ縺ｨ CSV 譯亥��縺ｮ縺ｿ�ｼ�ASCII 繧ｹ繝代�ｼ繧ｯ繝ｩ繧､繝ｳ縺ｯ蜃ｺ縺輔↑縺��ｼ峨�"""
    ser = _gemini_daily_trend_series(cum)
    if ser is None:
        return []
    keys, _, label = ser
    b = cum.get("buckets")
    lines = [
        "縲先耳遘ｻ繧ｰ繝ｩ繝輔第侭驥代�ｻ蜻ｼ蜃ｺ縺�: Q縲彝 蛻暦ｼ上ヨ繝ｼ繧ｯ繝ｳ驥�: S縲弋 蛻暦ｼ亥推繧ｰ繝ｩ繝輔�ｻ閾ｪ蜍墓峩譁ｰ�ｼ峨ｒ蜿ら�ｧ",
        f"  邉ｻ蛻�1: 譌･谺｡ {label}�ｼ�{keys[0]} �ｽ� {keys[-1]}�ｼ�",
        "  邉ｻ蛻�2: 譌･谺｡ 蜷郁ｨ医ヨ繝ｼ繧ｯ繝ｳ�ｼ�API 蝣ｱ蜻� total 縺ｾ縺溘�ｯ蜀�險ｳ蜷郁ｨ茨ｼ�",
        f"  蟷ｴ繝ｻ譛医�ｻ騾ｱ繝ｻ譎ゅ↑縺ｩ縺ｮ蜀�險ｳ: log\\{GEMINI_USAGE_BUCKETS_CSV_FILE}�ｼ�Excel 縺ｧ繧ｰ繝ｩ繝募庄�ｼ�",
    ]
    if isinstance(b, dict):
        note = b.get("timezone_note")
        if note:
            lines.append(f"  �ｼ�{note}�ｼ�")
    return lines


def _gemini_resolve_main_sheet_xlwings(book) -> object | None:
    """xlwings Book 縺九ｉ繝｡繧､繝ｳ逶ｸ蠖薙す繝ｼ繝医ｒ霑斐☆縲ら┌縺代ｌ縺ｰ None縲�"""
    for name in ("繝｡繧､繝ｳ", "繝｡繧､繝ｳ_", "Main"):
        try:
            return book.sheets[name]
        except Exception:
            continue
    try:
        for sht in book.sheets:
            try:
                if "繝｡繧､繝ｳ" in str(sht.name):
                    return sht
            except Exception:
                continue
    except Exception:
        pass
    return None


def _strip_gemini_usage_charts_xlwings(ws) -> None:
    """蠖捺ｩ溯�ｽ縺檎ｮ｡逅�縺吶ｋ謚倥ｌ邱夲ｼ亥錐蜑阪∪縺溘�ｯ繧ｰ繝ｩ繝輔ち繧､繝医Ν�ｼ峨ｒ蜑企勁縺吶ｋ縲�"""
    managed_names = (
        GEMINI_USAGE_XLW_CHART_NAME,
        GEMINI_USAGE_XLW_CHART_TOKENS_NAME,
    )
    title_markers = (
        "Gemini API 譌･谺｡謗ｨ遘ｻ",
        "Gemini API 譌･谺｡繝医�ｼ繧ｯ繝ｳ",
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
    """髢九＞縺溘ヶ繝�繧ｯ荳翫〒 Q縲彝繝ｻS縲弋 繧貞沂繧√∵釜繧檎ｷ壹げ繝ｩ繝輔ｒ 2 譛ｬ縺ｾ縺ｧ鄂ｮ縺擾ｼ�xlwings�ｼ峨�"""
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

    ws.range((hr, cdt)).value = "譌･莉�"
    ws.range((hr, cvl)).value = val_label
    for i, (dk, val) in enumerate(zip(day_keys, values)):
        r = hr + 1 + i
        ws.range((r, cdt)).value = dk
        ws.range((r, cvl)).value = val
    try:
        vrng = ws.range((hr + 1, cvl), (hr + n, cvl))
        vrng.number_format = "0.000000" if val_label == "謗ｨ螳啅SD" else "0"
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
        ca.ChartTitle.Text = "Gemini API 譌･谺｡謗ｨ遘ｻ"
        ca.HasLegend = False
    except Exception:
        pass

    tok_vals = _gemini_daily_total_tokens_for_days(cum, day_keys)
    if not tok_vals or max(tok_vals) <= 0:
        return

    tok_label = "蜷郁ｨ医ヨ繝ｼ繧ｯ繝ｳ"
    ws.range((hr, cts)).value = "譌･莉�"
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
        ca2.ChartTitle.Text = "Gemini API 譌･谺｡繝医�ｼ繧ｯ繝ｳ"
        ca2.HasLegend = False
    except Exception:
        pass


def _write_main_sheet_gemini_usage_via_xlwings(
    macro_wb_path: str, text: str, log_prefix: str
) -> bool:
    """Excel 縺ｧ繝悶ャ繧ｯ縺碁幕縺�縺ｦ縺�繧九→縺阪√Γ繧､繝ｳ P 蛻励�ｻQ縲弋繝ｻ謗ｨ遘ｻ繧ｰ繝ｩ繝包ｼ域怙螟ｧ2譛ｬ�ｼ峨ｒ xlwings 縺ｧ譖ｴ譁ｰ縺励※ Save縲�"""
    attached = _xlwings_attach_open_macro_workbook(macro_wb_path, log_prefix)
    if attached is None:
        logging.info(
            "%s: xlwings 縺ｧ繝槭け繝ｭ繝悶ャ繧ｯ縺ｫ謗･邯壹〒縺阪★縲√Γ繧､繝ｳ AI 繧ｵ繝槭Μ繧偵せ繧ｭ繝�繝励＠縺ｾ縺励◆縲�",
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
                "%s: 繝｡繧､繝ｳ繧ｷ繝ｼ繝医′辟｡縺�縺溘ａ xlwings 縺ｧ縺ｮ AI 繧ｵ繝槭Μ繧偵せ繧ｭ繝�繝励＠縺ｾ縺励◆縲�",
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
                "%s: 繝｡繧､繝ｳ繧ｷ繝ｼ繝� P%d 莉･髯阪�ｻGemini 謗ｨ遘ｻ繧ｰ繝ｩ繝包ｼ域侭驥�/蜻ｼ蜃ｺ縺励�ｻ繝医�ｼ繧ｯ繝ｳ�ｼ峨ｒ xlwings 縺ｧ菫晏ｭ倥＠縺ｾ縺励◆縲�",
                log_prefix,
                start_r,
            )
        finally:
            _xlwings_app_save_perf_state_pop(xw_book.app, _perf_snap)
    except Exception as ex:
        logging.warning(
            "%s: 繝｡繧､繝ｳ AI 繧ｵ繝槭Μ縺ｮ xlwings 菫晏ｭ倥↓螟ｱ謨�: %s", log_prefix, ex
        )
        ok = False
    finally:
        _xlwings_release_book_after_mutation(xw_book, info, ok)
    return ok


def _gemini_kv_table_lines(title: str, rows: list[tuple[str, str]]) -> list[str]:
    """邏ｯ險医�ｻ蠖灘ｮ溯｡悟髄縺代�ｮ 2 蛻励ユ繧ｭ繧ｹ繝郁｡ｨ�ｼ亥ｱ･豁ｴ陦後�ｯ蜷ｫ繧√↑縺��ｼ峨�"""
    out = [title]
    if not rows:
        return out
    lw = min(22, max(len(a) for a, _ in rows))
    sep = "  " + ("笏" * (lw + 2 + 28))
    out.append(sep)
    for a, b in rows:
        out.append(f"  {a:<{lw}}  {b}")
    return out


def _export_gemini_buckets_csv_for_charts(cum: dict) -> None:
    """Excel 謚倥ｌ邱壹�ｻ譽偵げ繝ｩ繝募髄縺代↓髟ｷ蠖｢蠑� CSV 繧� log 縺ｫ譖ｸ縺榊�ｺ縺吶�"""
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
        logging.debug("Gemini 繝舌こ繝�繝� CSV 縺ｮ菫晏ｭ倥↓螟ｱ謨�: %s", ex)


def build_gemini_usage_summary_text() -> str:
    """繝｡繧､繝ｳ陦ｨ遉ｺ繝ｻ邨先棡繝ｭ繧ｰ逕ｨ縺ｮ隍�謨ｰ陦後ユ繧ｭ繧ｹ繝茨ｼ医％縺ｮ螳溯｡悟���ｼ狗ｴｯ險� JSON�ｼ峨�"""
    cum = _load_gemini_cumulative_payload()
    ct_tot = int(cum.get("calls_total") or 0)
    if not _gemini_usage_session and ct_tot <= 0:
        return ""

    lines: list[str] = []
    ts = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    if _gemini_usage_session:
        lines.append(f"髮�險域凾蛻ｻ: {ts}�ｼ医％縺ｮ螳溯｡後〒縺ｮ Gemini API�ｼ�")
        tot_calls = sum(b["calls"] for b in _gemini_usage_session.values())
        tot_p = sum(b["prompt"] for b in _gemini_usage_session.values())
        tot_c = sum(b["candidates"] for b in _gemini_usage_session.values())
        tot_th = sum(b["thoughts"] for b in _gemini_usage_session.values())
        tot_t = sum(b["total"] for b in _gemini_usage_session.values())
        sess_rows: list[tuple[str, str]] = [
            ("蜻ｼ蜃ｺ縺�", f"{tot_calls:,} 蝗�"),
            ("蜈･蜉帙ヨ繝ｼ繧ｯ繝ｳ", f"{tot_p:,}"),
            ("蜃ｺ蜉帙ヨ繝ｼ繧ｯ繝ｳ", f"{tot_c:,}"),
        ]
        if tot_th:
            sess_rows.append(("諤晁�繝医�ｼ繧ｯ繝ｳ", f"{tot_th:,}"))
        sess_rows.append(("total 蝣ｱ蜻�", f"{tot_t:,}"))
        lines.extend(_gemini_kv_table_lines("縲舌％縺ｮ螳溯｡後�", sess_rows))
        grand_usd = 0.0
        any_price = False
        for mid in sorted(_gemini_usage_session.keys()):
            b = _gemini_usage_session[mid]
            mrows: list[tuple[str, str]] = [
                ("繝｢繝�繝ｫ", mid),
                ("蜻ｼ蜃ｺ縺�", f"{b['calls']:,} 蝗�"),
                ("蜈･蜉帙ヨ繝ｼ繧ｯ繝ｳ", f"{b['prompt']:,}"),
                ("蜃ｺ蜉帙ヨ繝ｼ繧ｯ繝ｳ", f"{b['candidates']:,}"),
            ]
            if b.get("thoughts", 0):
                mrows.append(("諤晁�繝医�ｼ繧ｯ繝ｳ", f"{b['thoughts']:,}"))
            mrows.append(("total_token_count", f"{b['total']:,}"))
            est = _gemini_estimate_cost_usd(
                mid, b["prompt"], b["candidates"], b.get("thoughts", 0)
            )
            if est is not None:
                any_price = True
                grand_usd += est
                mrows.append(("謗ｨ螳啅SD", f"${est:.6f}"))
                mrows.append(
                    (
                        "謗ｨ螳哽PY",
                        f"ﾂ･{est * GEMINI_JPY_PER_USD:.2f}�ｼ�{GEMINI_JPY_PER_USD:.0f}蜀�/USD�ｼ�",
                    )
                )
            else:
                mrows.append(("謗ｨ螳壽侭驥�", "�ｼ亥腰萓｡譛ｪ逋ｻ骭ｲ繝｢繝�繝ｫ�ｼ�"))
            lines.append("")
            lines.extend(_gemini_kv_table_lines(f"縲舌％縺ｮ螳溯｡後�ｻ繝｢繝�繝ｫ蛻･縲�", mrows))
        if any_price:
            lines.append("")
            lines.extend(
                _gemini_kv_table_lines(
                    "縲舌％縺ｮ螳溯｡後�ｻ謗ｨ螳壽侭驥大粋險医�",
                    [
                        ("USD", f"${grand_usd:.6f}"),
                        (
                            "JPY",
                            f"ﾂ･{grand_usd * GEMINI_JPY_PER_USD:.2f}�ｼ�{GEMINI_JPY_PER_USD:.0f}蜀�/USD�ｼ�",
                        ),
                    ],
                )
            )
    else:
        lines.append(f"髮�險域凾蛻ｻ: {ts}")
        lines.append("�ｼ医％縺ｮ螳溯｡後〒縺ｮ Gemini API 蜻ｼ蜃ｺ縺励�ｯ縺ゅｊ縺ｾ縺帙ｓ�ｼ�")
    lines.append("窶ｻ 繝医�ｼ繧ｯ繝ｳ縺ｯ API 縺ｮ usage_metadata 縺ｫ蝓ｺ縺･縺阪∪縺吶�")
    lines.append(
        "窶ｻ USD 蜊倅ｾ｡縺ｯ繧ｳ繝ｼ繝会ｼ冗腸蠅�螟画焚縺ｮ逶ｮ螳峨〒縺吶ょｮ溯ｪｲ驥代�ｯ Google 縺ｮ隲区ｱゅｒ蜿ら�ｧ縺励※縺上□縺輔＞縲�"
    )
    lines.append(
        "窶ｻ 蜷� API 蜻ｼ蜃ｺ縺励＃縺ｨ縺ｮ謗ｨ螳壽侭驥代�ｯ繧ｳ繝ｳ繧ｽ繝ｼ繝ｫ縺ｫ蜃ｺ縺輔★縲∽ｸ玖ｨ倡ｴｯ險� JSON 縺ｫ縺ｮ縺ｿ遨阪∩荳翫￡縺ｾ縺吶�"
    )

    if ct_tot > 0:
        lines.append("")
        cum_hdr = (
            f"縲千ｴｯ險医捜GEMINI_USAGE_CUMULATIVE_JSON_FILE} "
            "�ｼ�API_Payment 繝輔か繝ｫ繝繝ｻ蜈ｨ螳溯｡後�ｮ謗ｨ螳壼､�ｼ�"
        )
        pt0 = int(cum.get("prompt_total") or 0)
        cc0 = int(cum.get("candidates_total") or 0)
        th0 = int(cum.get("thoughts_total") or 0)
        tt0 = int(cum.get("total_tokens_reported") or 0)
        cum_rows: list[tuple[str, str]] = [
            ("譛邨よ峩譁ｰ", str(cum.get("updated_at") or "窶�")),
            ("蜻ｼ蜃ｺ縺�", f"{ct_tot:,} 蝗�"),
            ("蜈･蜉帙ヨ繝ｼ繧ｯ繝ｳ", f"{pt0:,}"),
            ("蜃ｺ蜉帙ヨ繝ｼ繧ｯ繝ｳ", f"{cc0:,}"),
        ]
        if th0:
            cum_rows.append(("諤晁�繝医�ｼ繧ｯ繝ｳ", f"{th0:,}"))
        cum_rows.append(("total 蝣ｱ蜻�", f"{tt0:,}"))
        usd_all = float(cum.get("estimated_cost_usd_total") or 0.0)
        if usd_all > 0:
            cum_rows.append(("謗ｨ螳啅SD 邏ｯ險�", f"${usd_all:.6f}"))
            cum_rows.append(
                (
                    "謗ｨ螳哽PY 邏ｯ險�",
                    f"ﾂ･{usd_all * GEMINI_JPY_PER_USD:.2f}�ｼ�{GEMINI_JPY_PER_USD:.0f}蜀�/USD�ｼ�",
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
                    ("繝｢繝�繝ｫ", mid),
                    ("蜻ｼ蜃ｺ縺�", f"{int(m.get('calls') or 0):,} 蝗�"),
                    (
                        "蜈･蜉� / 蜃ｺ蜉�",
                        f"{int(m.get('prompt') or 0):,} / {int(m.get('candidates') or 0):,}",
                    ),
                ]
                if int(m.get("thoughts") or 0):
                    mrows2.append(("諤晁�繝医�ｼ繧ｯ繝ｳ", f"{int(m.get('thoughts') or 0):,}"))
                mud = float(m.get("estimated_cost_usd") or 0.0)
                if mud > 0:
                    mrows2.append(("謗ｨ螳啅SD 邏ｯ險�", f"${mud:.6f}"))
                    mrows2.append(
                        ("謗ｨ螳哽PY 邏ｯ險�", f"ﾂ･{mud * GEMINI_JPY_PER_USD:.2f}")
                    )
                lines.append("")
                lines.extend(_gemini_kv_table_lines("縲千ｴｯ險医�ｻ繝｢繝�繝ｫ蛻･縲�", mrows2))
        trend = _gemini_usage_trend_caption_lines(cum)
        if trend:
            lines.append("")
            lines.extend(trend)
    return "\n".join(lines)


def write_main_sheet_gemini_usage_summary(wb_path: str, log_prefix: str) -> None:
    """Gemini 蛻ｩ逕ｨ繧ｵ繝槭Μ繧� log 縺ｫ譖ｸ縺阪』lwings 縺ｧ繝｡繧､繝ｳ P 蛻励�ｻ謗ｨ遘ｻ繧ｰ繝ｩ繝輔∈菫晏ｭ假ｼ磯幕縺�縺ｦ縺�繧九ヶ繝�繧ｯ蜷代￠�ｼ峨�"""
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
                "%s: AI 繧ｵ繝槭Μ縺ｮ xlwings 譖ｸ縺崎ｾｼ縺ｿ縺ｧ萓句､�: %s", log_prefix, ex
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
        logging.debug("Gemini 繝舌こ繝�繝� CSV 蜃ｺ蜉帙〒萓句､厄ｼ育ｶ夊｡鯉ｼ�: %s", ex)
    if xw_ok:
        return
    if text.strip():
        logging.info(
            "%s: 繝｡繧､繝ｳ P 蛻励�ｻ繧ｰ繝ｩ繝輔ｒ xlwings 縺ｧ菫晏ｭ倥〒縺阪∪縺帙ｓ縺ｧ縺励◆縲�"
            " %s 縺ｫ蜃ｺ蜉帶ｸ医∩ 竊� 繝槭け繝ｭ縲後Γ繧､繝ｳ繧ｷ繝ｼ繝�_Gemini蛻ｩ逕ｨ繧ｵ繝槭Μ繧単蛻励↓蜿肴丐縲阪〒 P 蛻励�ｮ縺ｿ蜿肴丐縺ｧ縺阪∪縺吶�",
            log_prefix,
            path,
        )
    else:
        logging.info(
            "%s: Gemini 譛ｪ菴ｿ逕ｨ: 繧ｵ繝槭Μ繧堤ｩｺ縺ｧ %s 縺ｫ蜃ｺ蜉帙�",
            log_prefix,
            path,
        )


def _try_write_main_sheet_gemini_usage_summary(phase: str) -> None:
    try:
        write_main_sheet_gemini_usage_summary(TASKS_INPUT_WORKBOOK, phase)
    except Exception as ex:
        logging.warning(
            "%s: 繝｡繧､繝ｳ繧ｷ繝ｼ繝医∈縺ｮ AI 蛻ｩ逕ｨ繧ｵ繝槭Μ譖ｸ縺崎ｾｼ縺ｿ縺ｧ萓句､厄ｼ育ｶ夊｡鯉ｼ�: %s", phase, ex
        )


def _plan_sheet_write_global_parse_block_to_ws(
    ws,
    global_priority_override: dict,
    when_str: str,
) -> None:
    """譌｢縺ｫ髢九＞縺ｦ縺�繧九碁�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙咲嶌蠖薙す繝ｼ繝医∈ AX:AY 縺ｮ繧ｰ繝ｭ繝ｼ繝舌Ν隗｣譫舌ヶ繝ｭ繝�繧ｯ繧呈嶌縺上�"""
    gpo = global_priority_override or {}
    lc = PLAN_SHEET_GLOBAL_PARSE_LABEL_COL
    vc = PLAN_SHEET_GLOBAL_PARSE_VALUE_COL
    max_r = PLAN_SHEET_GLOBAL_PARSE_MAX_ROWS
    for i in range(max_r):
        ws.cell(row=1 + i, column=lc, value=None)
        ws.cell(row=1 + i, column=vc, value=None)
    align_top = Alignment(wrap_text=True, vertical="top")
    pairs: list[tuple[str, str]] = [
        ("縲舌げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫舌�", "蜿ら�ｧ逕ｨ繝ｻ谿ｵ髫�2縺ｧ閾ｪ蜍戊ｨ倬鹸"),
        (
            "窶ｻ莠碁㍾驕ｩ逕ｨ縺ｫ縺､縺�縺ｦ",
            "驟榊床縺ｸ縺ｮ蜿肴丐縺ｯ繝｡繧､繝ｳ繧ｷ繝ｼ繝医後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医阪°繧峨�ｮ縺ｿ陦後ｏ繧後∪縺吶�"
            "縺薙�ｮAX縲廣Y蛻励�ｯ隱ｭ縺ｿ蜿悶ｉ繧後∪縺帙ｓ縲らｷｨ髮�縺励※繧よｬ｡蝗槫ｮ溯｡後∪縺ｧ驟榊床縺ｫ蜉ｹ縺阪∪縺帙ｓ縲�"
            "蜴滓枚縺ｯ繝｡繧､繝ｳ谺�繧貞盾辣ｧ縺励※縺上□縺輔＞縲�",
        ),
        ("險育判蝓ｺ貅匁律譎�", (when_str or "").strip() or "窶�"),
        (
            "蟾･蝣ｴ莨第･ｭ譌･",
            ", ".join(str(x) for x in (gpo.get("factory_closure_dates") or []))
            if gpo.get("factory_closure_dates")
            else "�ｼ医↑縺暦ｼ�",
        ),
        (
            "繧ｹ繧ｭ繝ｫ隕∽ｻｶ繧堤┌隕�",
            "縺ｯ縺�" if gpo.get("ignore_skill_requirements") else "縺�縺�縺�",
        ),
        (
            "need莠ｺ謨ｰ1蝗ｺ螳�",
            "縺ｯ縺�" if gpo.get("ignore_need_minimum") else "縺�縺�縺�",
        ),
        (
            "驟榊床蛻ｶ髯舌�ｮ謦､蟒�",
            "縺ｯ縺�" if gpo.get("abolish_all_scheduling_limits") else "縺�縺�縺�",
        ),
        (
            "繧ｰ繝ｭ繝ｼ繝舌ΝOP謖�蜷�",
            json.dumps(gpo.get("task_preferred_operators") or {}, ensure_ascii=False)
            if gpo.get("task_preferred_operators")
            else "�ｼ医↑縺暦ｼ�",
        ),
        (
            "譌･莉佚怜ｷ･遞九メ繝ｼ繝�謖�蜷�",
            json.dumps(
                gpo.get("global_day_process_operator_rules") or [],
                ensure_ascii=False,
            )
            if gpo.get("global_day_process_operator_rules")
            else "�ｼ医↑縺暦ｼ�",
        ),
        (
            "繧ｰ繝ｭ繝ｼ繝舌Ν騾溷ｺｦ繝ｫ繝ｼ繝ｫ",
            json.dumps(gpo.get("global_speed_rules") or [], ensure_ascii=False)
            if gpo.get("global_speed_rules")
            else "�ｼ医↑縺暦ｼ�",
        ),
        (
            "譛ｪ驕ｩ逕ｨ繝｡繝｢(AI)",
            str(gpo.get("scheduler_notes_ja") or "").strip() or "�ｼ医↑縺暦ｼ�",
        ),
        (
            "AI隕∫ｴ�",
            str(gpo.get("interpretation_ja") or "").strip() or "�ｼ医↑縺暦ｼ�",
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
    log_prefix: str = "谿ｵ髫�2",
) -> bool:
    """
    縲碁�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙阪す繝ｼ繝医�ｮ蜿ｳ遶ｯ莉倩ｿ托ｼ�AX:AY�ｼ峨↓縲√げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医�ｮ隗｣譫千ｵ先棡繧呈嶌縺崎ｾｼ繧縲�
    繝｡繧､繝ｳ蜴滓枚縺ｯ縺薙％縺ｫ霆｢險倥＠縺ｪ縺��ｼ医Γ繧､繝ｳ谺�縺ｨ縺ｮ驥崎､�繝ｻ隱､隗｣繧帝∩縺代ｋ�ｼ峨よ悽蛻励�ｯ蜀崎ｪｭ霎ｼ縺輔ｌ縺壼盾辣ｧ蟆ら畑縲�
    Excel 縺ｧ繝悶ャ繧ｯ繧帝幕縺�縺溘∪縺ｾ縺�縺ｨ菫晏ｭ倥↓螟ｱ謨励☆繧九％縺ｨ縺後≠繧具ｼ井ｻ悶�ｮ openpyxl 譖ｸ霎ｼ縺ｨ蜷梧ｧ假ｼ峨�
    """
    if not wb_path or not os.path.isfile(wb_path):
        return False
    gpo = global_priority_override or {}
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = None
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "%s: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ openpyxl 縺ｧ繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫舌ｒ驟榊床繧ｷ繝ｼ繝医∈譖ｸ縺崎ｾｼ縺ｿ縺ｾ縺帙ｓ縲�",
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
            "%s: 繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫舌�ｮ驟榊床繧ｷ繝ｼ繝域嶌霎ｼ縺ｮ縺溘ａ繝悶ャ繧ｯ繧帝幕縺代∪縺帙ｓ: %s",
            log_prefix,
            ex,
        )
        return False
    try:
        if sheet_name not in wb.sheetnames:
            logging.info(
                "%s: 繧ｷ繝ｼ繝� '%s' 縺檎┌縺�縺溘ａ繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫舌�ｮ蜿肴丐繧偵せ繧ｭ繝�繝励�",
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
            "%s: 縲�%s縲�%s:%s 蛻励↓繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫舌ｒ菫晏ｭ倥＠縺ｾ縺励◆縲�",
            log_prefix,
            sheet_name,
            get_column_letter(lc),
            get_column_letter(vc),
        )
        return True
    except OSError as ex:
        logging.warning(
            "%s: 繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫舌ｒ驟榊床繧ｷ繝ｼ繝医∈菫晏ｭ倥〒縺阪∪縺帙ｓ縺ｧ縺励◆�ｼ�Excel 縺ｧ髢九＞縺溘∪縺ｾ遲会ｼ�: %s",
            log_prefix,
            ex,
        )
        return False
    except Exception as ex:
        logging.warning(
            "%s: 繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫舌�ｮ驟榊床繧ｷ繝ｼ繝域嶌霎ｼ縺ｧ萓句､�: %s", log_prefix, ex
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
            log_prefix="谿ｵ髫�2",
        )
    except Exception as ex:
        logging.warning(
            "谿ｵ髫�2: 驟榊床繧ｷ繝ｼ繝医∈縺ｮ繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫先嶌縺崎ｾｼ縺ｿ縺ｧ萓句､厄ｼ育ｶ夊｡鯉ｼ�: %s",
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
            log_prefix="谿ｵ髫�2",
        )
    except Exception as ex:
        logging.warning(
            "谿ｵ髫�2: 驟榊床繧ｷ繝ｼ繝医∈縺ｮ繧ｰ繝ｭ繝ｼ繝舌Ν隗｣譫撰ｼ狗泝逶ｾ逹濶ｲ�ｼ�1蝗樔ｿ晏ｭ假ｼ峨〒萓句､厄ｼ育ｶ夊｡鯉ｼ�: %s",
            ex,
        )


def _log_task_special_ai_response(raw_text, parsed, extracted_json_str, prompt_text=None):
    """迚ｹ蛻･謖�螳喟蛯呵�蜷代￠ Gemini 縺ｮ繝励Ο繝ｳ繝励ヨ繝ｻ逕溘ユ繧ｭ繧ｹ繝医�ｻ謚ｽ蜃ｺJSON繝ｻ繝代�ｼ繧ｹ邨先棡繧�1繝輔ぃ繧､繝ｫ縺ｫ谿九☆縲�"""
    path = os.path.join(log_dir, TASK_SPECIAL_AI_LAST_RESPONSE_FILE)
    try:
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            if prompt_text is not None and str(prompt_text).strip():
                f.write("=== Gemini 縺ｸ騾∽ｿ｡縺励◆繝励Ο繝ｳ繝励ヨ�ｼ亥�ｨ譁��ｼ� ===\n")
                f.write(str(prompt_text).strip())
                f.write("\n\n")
            f.write("=== Gemini 霑泌唆繝�繧ｭ繧ｹ繝茨ｼ医Δ繝�繝ｫ蜃ｺ蜉帙◎縺ｮ縺ｾ縺ｾ�ｼ� ===\n")
            f.write(raw_text or "")
            f.write(
                "\n\n=== AI 縺瑚ｿ斐＠縺溘ユ繧ｭ繧ｹ繝医°繧峨け繝ｩ繧､繧｢繝ｳ繝医′蛻�繧雁�ｺ縺励◆ JSON 譁�蟄怜�� ===\n"
                "�ｼ遺ｻ繝ｦ繝ｼ繧ｶ繝ｼ迚ｹ蛻･謖�螳壹�ｮ隗｣譫舌↓豁｣隕剰｡ｨ迴ｾ縺ｯ菴ｿ縺｣縺ｦ縺�縺ｾ縺帙ｓ縲ゅΔ繝�繝ｫ蠢懃ｭ斐�ｮ繝代�ｼ繧ｹ逕ｨ縺ｧ縺呻ｼ噂n"
            )
            f.write(extracted_json_str if extracted_json_str else "(謚ｽ蜃ｺ縺ｪ縺�)")
            f.write("\n\n=== json.loads 蠕鯉ｼ井ｾ晞�ｼNO繧ｭ繝ｼ�ｼ� ===\n")
            if isinstance(parsed, dict):
                f.write(json.dumps(parsed, ensure_ascii=False, indent=2))
            else:
                f.write("(繝代�ｼ繧ｹ縺ｧ縺阪★)")
        logging.info(
            "繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: 繝励Ο繝ｳ繝励ヨ�ｼ帰I蠢懃ｭ斐�ｮ隧ｳ邏ｰ 竊� %s",
            path,
        )
    except OSError as ex:
        logging.warning("繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: AI蠢懃ｭ斐ヵ繧｡繧､繝ｫ菫晏ｭ倥↓螟ｱ謨�: %s", ex)
    if isinstance(parsed, dict) and parsed:
        logging.info(
            "繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: 隗｣譫舌＆繧後◆萓晞�ｼNO: %s",
            ", ".join(sorted(parsed.keys(), key=lambda x: str(x))),
        )
        for tid_k in sorted(parsed.keys(), key=lambda x: str(x)):
            logging.info(
                "  萓晞�ｼNO [%s] AI隗｣譫舌ヵ繧｣繝ｼ繝ｫ繝�: %s",
                tid_k,
                json.dumps(parsed[tid_k], ensure_ascii=False),
            )


def _parse_and_log_task_special_gemini_response(res, prompt_text=None):
    """
    API 繝ｬ繧ｹ繝昴Φ繧ｹ繧� JSON 蛹悶＠繝ｭ繧ｰ�ｼ上ヵ繧｡繧､繝ｫ縺ｸ險倬鹸縲ょ､ｱ謨玲凾縺ｯ None縲�
    繝ｦ繝ｼ繧ｶ繝ｼ縺ｮ迚ｹ蛻･謖�螳壽枚險縺ｫ縺ｯ隗ｦ繧後★縲√Δ繝�繝ｫ蜃ｺ蜉帙°繧� JSON 繝悶Ο繝�繧ｯ繧貞叙繧雁�ｺ縺吝�ｦ逅�縺ｮ縺ｿ縲�
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
            "繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: AI蠢懃ｭ斐°繧� JSON 繧呈歓蜃ｺ縺ｧ縺阪∪縺帙ｓ縺ｧ縺励◆縲ら函繝�繧ｭ繧ｹ繝亥�磯�ｭ 3000 譁�蟄�:\n%s",
            (raw[:3000] if raw else "(遨ｺ)"),
        )
        return None
    extracted = match.group(0)
    try:
        parsed = json.loads(extracted)
    except json.JSONDecodeError as je:
        _log_task_special_ai_response(raw, None, extracted, prompt_text)
        logging.warning("繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: JSON 繝代�ｼ繧ｹ螟ｱ謨�: %s", je)
        return None
    if not isinstance(parsed, dict):
        _log_task_special_ai_response(raw, None, extracted, prompt_text)
        logging.warning("繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: 繝医ャ繝励Ξ繝吶Ν縺� JSON 繧ｪ繝悶ず繧ｧ繧ｯ繝医〒縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲�")
        return None
    _log_task_special_ai_response(raw, parsed, extracted, prompt_text)
    return parsed


def analyze_task_special_remarks(tasks_df, reference_year=None, ai_sheet_sink: dict | None = None):
    """
    縲碁�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙阪�ｮ縲檎音蛻･謖�螳喟蛯呵�縲阪ｒ AI 縺ｧ讒矩�蛹厄ｼ医そ繝ｫ縺ｫ蛟､縺後≠繧矩��逶ｮ縺ｯ蠕梧ｮｵ縺ｧ繧ｻ繝ｫ繧貞━蜈茨ｼ峨�
    縲碁�榊床荳崎ｦ√阪′繧ｪ繝ｳ縺ｪ陦後�ｯ繝励Ο繝ｳ繝励ヨ縺ｫ霈峨○縺ｪ縺��ｼ�API 遽邏�繝ｻ蠖楢ｩｲ陦後�ｯ驟榊床縺励↑縺�縺溘ａ�ｼ峨�
    諡�蠖徹P謖�蜷阪�ｯ繝励Ο繝ｳ繝励ヨ縺ｮ霑泌唆螂醍ｴ�縺ｧ繝｢繝�繝ｫ縺ｫ preferred_operator 繧貞�ｺ蜉帙＆縺帙ｋ�ｼ亥ｙ閠�繧呈ｭ｣隕剰｡ｨ迴ｾ縺ｧ蛻�繧雁�ｺ縺吝�ｦ逅�縺ｯ陦後ｏ縺ｪ縺��ｼ峨�
    json/ai_remarks_cache.json 縺ｫ TTL AI_CACHE_TTL_SECONDS 縺ｧ繧ｭ繝｣繝�繧ｷ繝･�ｼ亥酔荳蜈･蜉帙�ｻ蜷御ｸ蝓ｺ貅門ｹｴ縺ｪ繧� API 繧貞他縺ｰ縺ｪ縺��ｼ峨�
    萓晞�ｼNO縺ｯ謨ｰ蛟､陦ｨ險倥�ｻ蜈ｨ隗偵↑縺ｩ繧呈ｭ｣隕丞喧縺励※繧ｭ繝ｼ繧貞ｮ牙ｮ壼喧縺励∝渕貅門ｹｴ縺ｯ謖�邏九↓蜷ｫ繧√※譌･莉倩ｧ｣驥医�ｮ螟牙喧縺ｨ繧ｭ繝｣繝�繧ｷ繝･縺ｮ鬟溘＞驕輔＞繧帝亟縺舌�

    謌ｻ繧雁､縺ｮ萓�: 萓晞�ｼNO -> 繧ｪ繝悶ず繧ｧ繧ｯ繝医√∪縺溘�ｯ蜷御ｸ萓晞�ｼNO縺ｫ蛯呵�陦後′隍�謨ｰ縺ゅｋ蝣ｴ蜷医�ｯ繧ｪ繝悶ず繧ｧ繧ｯ繝医�ｮ驟榊�励�
      process_name, machine_name 窶ｦ 蠖楢ｩｲ蛯呵�繧ｻ繝ｫ縺後≠繧玖｡後�ｮ蟾･遞句錐繝ｻ讖滓｢ｰ蜷搾ｼ医�励Ο繝ｳ繝励ヨ縺ｮ陦後→荳閾ｴ�ｼ�
      restrict_to_process_name, restrict_to_machine_name 窶ｦ 逵∫払縺ｾ縺溘�ｯ遨ｺ縺ｪ繧牙酔荳萓晞�ｼNO縺ｮ蜈ｨ蟾･遞九�ｻ蜈ｨ讖滓｢ｰ陦後↓驕ｩ逕ｨ縲�
      縺昴�ｮ莉� required_op, speed_override, task_efficiency, priority, start_date, start_time,
      target_completion_date, ship_by_date, preferred_operator 縺ｪ縺ｩ縲�
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
            "繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: AI 隗｣譫仙ｯｾ雎｡縺後≠繧翫∪縺帙ｓ�ｼ医�%s縲榊�励�ｯ%s�ｼ峨�"
            "諤ｻ陦梧焚=%s縲∽ｾ晞�ｼNO縺ｮ縺ゅｋ陦�=%s縲∝ｙ閠�縺悟�･縺｣縺ｦ縺�繧玖｡�=%s縲�"
            "谿ｵ髫�2螳溯｡悟燕縺ｫ繝悶ャ繧ｯ繧剃ｿ晏ｭ倥＠縲∵悽蠖薙↓縲�%s縲榊�励∈蜈･蜉帙＠縺ｦ縺�繧九°遒ｺ隱阪＠縺ｦ縺上□縺輔＞縲�",
            PLAN_COL_SPECIAL_REMARK,
            "隕九▽縺九ｊ縺ｾ縺帙ｓ" if miss_col else "遨ｺ縺ｮ蜿ｯ閭ｽ諤ｧ縺後≠繧翫∪縺�",
            n_rows,
            n_tid_raw,
            n_rem_only,
            PLAN_COL_SPECIAL_REMARK,
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["迚ｹ蛻･謖�螳壼ｙ閠ダAI_API"] = "繧ｹ繧ｭ繝�繝暦ｼ亥ｯｾ雎｡陦後↑縺暦ｼ�"
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
            "繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: 繧ｭ繝｣繝�繧ｷ繝･繝偵ャ繝茨ｼ�%s 莉ｶ繝ｻ蝓ｺ貅門ｹｴ=%s�ｼ峨�Gemini 縺ｯ蜻ｼ縺ｳ縺ｾ縺帙ｓ縲�",
            len(lines),
            ref_y,
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["迚ｹ蛻･謖�螳壼ｙ閠ダAI_API"] = "縺ｪ縺暦ｼ医く繝｣繝�繧ｷ繝･菴ｿ逕ｨ�ｼ�"
        out = copy.deepcopy(cached_parsed)
        if isinstance(out, dict):
            _repair_task_special_ai_wrong_top_level_keys(out, tasks_df)
        return out

    logging.info(
        "繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: 繧ｭ繝｣繝�繧ｷ繝･縺ｪ縺励�Gemini 縺ｧ %s 莉ｶ縺ｮ蛯呵�繧定ｧ｣譫舌＠縺ｾ縺呻ｼ亥渕貅門ｹｴ=%s�ｼ峨�",
        len(lines),
        ref_y,
    )

    if not API_KEY:
        logging.info("GEMINI_API_KEY 譛ｪ險ｭ螳壹�ｮ縺溘ａ繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳壹�ｮAI隗｣譫舌ｒ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆縲�")
        if ai_sheet_sink is not None:
            ai_sheet_sink["迚ｹ蛻･謖�螳壼ｙ閠ダAI_API"] = "縺ｪ縺暦ｼ�API繧ｭ繝ｼ譛ｪ險ｭ螳夲ｼ�"
        return {}

    prompt = f"""
縺ゅ↑縺溘�ｯ蟾･蝣ｴ縺ｮ驟榊床險育判蜷代￠縺ｫ縲・xcel縲檎音蛻･謖�螳喟蛯呵�縲肴ｬ�縺ｸ縺ｮ閾ｪ逕ｱ險倩ｿｰ繧定ｪｭ縺ｿ縲�驟榊床繝ｭ繧ｸ繝�繧ｯ縺御ｽｿ縺医ｋ繝輔ぅ繝ｼ繝ｫ繝峨□縺代↓關ｽ縺ｨ縺苓ｾｼ繧繧｢繧ｷ繧ｹ繧ｿ繝ｳ繝医〒縺吶�

縲先怙驥崎ｦ√�
1) 縲千音蛻･謖�螳壼次譁�縲代�ｮ蜷�陦後�ｯ縲√Θ繝ｼ繧ｶ繝ｼ縺後そ繝ｫ縺ｫ蜈･蜉帙＠縺滓枚蟄怜�励ｒ **謾ｹ螟峨�ｻ隕∫ｴ�繝ｻ譁ｭ縺｡蛻�繧翫�ｯ縺励※縺翫ｉ縺�**�ｼ亥�磯�ｭ譛ｫ蟆ｾ縺ｮ遨ｺ逋ｽ縺ｮ縺ｿ髯､蜴ｻ�ｼ峨√◎縺ｮ縺ｾ縺ｾ貂｡縺励※縺�縺ｾ縺吶�**蜴滓枚縺ｮ莠句ｮ溘ｄ諢丞峙繧貞挨縺ｮ譁�險縺ｫ鄂ｮ縺肴鋤縺医↑縺�縺ｧ縺上□縺輔＞縲�**
2) 縺ゅ↑縺溘�ｮ蠢懃ｭ斐�ｯ **1蛟九�ｮ JSON 繧ｪ繝悶ず繧ｧ繧ｯ繝医�ｮ縺ｿ**�ｼ亥�磯�ｭ縺� {{ 縲∵忰蟆ｾ縺� }} �ｼ峨りｪｬ譏取枚繝ｻ繝槭�ｼ繧ｯ繝繧ｦ繝ｳ繝ｻ繧ｳ繝ｼ繝峨ヵ繧ｧ繝ｳ繧ｹ縺ｯ遖∵ｭ｢縲�
3) JSON 縺ｮ繝医ャ繝励Ξ繝吶Ν繧ｭ繝ｼ縺ｯ縲∝推陦後�ｮ **萓晞�ｼNO縲舌→縲代�ｮ髢薙�ｮ譁�蟄怜�励�ｮ縺ｿ** 縺ｨ **螳悟�ｨ荳閾ｴ** 縺輔○繧九％縺ｨ縲�**蛯呵�譛ｬ譁�**縺ｫ譖ｸ縺九ｌ縺溷刀逡ｪ繝ｻ蜴溷渚蜷阪�ｻ陬ｽ蜩√さ繝ｼ繝会ｼ井ｾ�: 20010 縺ｧ蟋九∪繧狗分蜿ｷ蛻暦ｼ峨ｒ繧ｭ繝ｼ縺ｫ縺励※縺ｯ縺ｪ繧峨↑縺�縲ょｙ閠�縺後◎縺ｮ繧医≧縺ｪ逡ｪ蜿ｷ縺ｧ蟋九∪縺｣縺ｦ縺�縺ｦ繧ゅ√く繝ｼ縺ｯ蠢�縺壹舌大��縺ｮ萓晞�ｼNO縺�縺代→縺吶ｋ縲�

縲占ｿ泌唆JSON縺ｮ螂醍ｴ��ｼ医％縺ｮ遽縺ｩ縺翫ｊ縺ｫ蜃ｺ蜉帙☆繧九％縺ｨ�ｼ峨�
笆� 繝医ャ繝励Ξ繝吶Ν
- 繧ｭ繝ｼ: 荳願ｨ倥千音蛻･謖�螳壼次譁�縲代�ｮ **萓晞�ｼNO縲絶ｦ縲代�ｮ諡ｬ蠑ｧ蜀�** 縺ｮ譁�蟄怜�励→ **螳悟�ｨ荳閾ｴ**�ｼ郁｡ｨ險倥�ｻ繝上う繝輔Φ繝ｻ闍ｱ螟ｧ譁�蟄怜ｰ乗枚蟄励ｒ蜴滓枚縺ｩ縺翫ｊ�ｼ峨ょｙ閠�譛ｬ譁�荳ｭ縺ｮ謨ｰ蟄怜�励ｒ繧ｭ繝ｼ縺ｫ縺励↑縺�縲�
- 蛟､: 谺｡縺ｮ縺�縺壹ｌ縺九�
  (A) **JSON繧ｪ繝悶ず繧ｧ繧ｯ繝�1縺､** 窶ｦ 蠖楢ｩｲ萓晞�ｼNO縺ｮ蛯呵�縺後�励Ο繝ｳ繝励ヨ荳� **1陦後□縺�** 縺ｮ縺ｨ縺阪�
  (B) **JSON驟榊��**�ｼ郁ｦ∫ｴ�縺ｯ繧ｪ繝悶ず繧ｧ繧ｯ繝茨ｼ俄ｦ 蜷御ｸ萓晞�ｼNO縺ｧ蟾･遞句錐繝ｻ讖滓｢ｰ蜷阪′逡ｰ縺ｪ繧句ｙ閠�陦後′ **隍�謨ｰ** 縺ゅｋ縺ｨ縺阪りｦ∫ｴ�縺ｮ鬆�縺ｯ繝励Ο繝ｳ繝励ヨ縺ｮ陦碁��縺ｨ蟇ｾ蠢懊＆縺帙ｋ縲�

笆� process_name�ｼ域枚蟄怜�暦ｼ峨�ｻmachine_name�ｼ域枚蟄怜�暦ｼ俄� **蠢�鬆�**
- 蠖楢ｩｲ蛯呵�縺ｫ蟇ｾ蠢懊☆繧九�励Ο繝ｳ繝励ヨ陦後�ｮ **蟾･遞句錐縲娯ｦ縲�**繝ｻ**讖滓｢ｰ蜷阪娯ｦ縲�** 縺ｮ蛟､縺ｨ **荳閾ｴ** 縺輔○繧具ｼ医鯉ｼ育ｩｺ�ｼ峨阪�ｮ縺ｨ縺阪�ｯ遨ｺ譁�蟄怜�� ""�ｼ峨�
- 繝ｭ繧ｰ繝ｻ繝医Ξ繝ｼ繧ｹ逕ｨ縲ら怐逡･荳榊庄縲�

笆� restrict_to_process_name�ｼ域枚蟄怜�暦ｼ峨�ｻrestrict_to_machine_name�ｼ域枚蟄怜�暦ｼ俄� **莉ｻ諢�**
- **蜴滓枚縺後檎音螳壹�ｮ蟾･遞九□縺代阪後％縺ｮ讖滓｢ｰ縺�縺代阪↑縺ｩ縲�驕ｩ逕ｨ遽�蝗ｲ繧堤ｵ槭▲縺ｦ縺�繧九→縺阪□縺�** 蜃ｺ蜉帙☆繧九�
- **蜴滓枚縺ｫ蟾･遞句錐繝ｻ讖滓｢ｰ蜷阪�ｮ髯仙ｮ壹′辟｡縺�**�ｼ井ｾ晞�ｼ蜈ｨ菴薙�ｻ蜈ｨ陦檎ｨ九∈縺ｮ謖�遉ｺ�ｼ峨→縺阪�ｯ **荳｡譁ｹ縺ｨ繧ら怐逡･** 縺吶ｋ縺� **遨ｺ譁�蟄怜�� ""** 縺ｨ縺吶ｋ縲�
- 縺昴�ｮ蝣ｴ蜷医�驟榊床繝ｭ繧ｸ繝�繧ｯ縺ｯ **蜷御ｸ萓晞�ｼNO縺ｮ蛻･陦鯉ｼ井ｾ�: 繧ｨ繝ｳ繝懊せ陦後→蛻�蜑ｲ陦鯉ｼ峨↓繧ょ酔縺俶欠遉ｺ繧帝←逕ｨ** 縺吶ｋ縲�
- 邨槭ｋ蝣ｴ蜷医�ｯ縲∝次譁�縺ｧ遉ｺ縺輔ｌ縺溯ｭ伜挨蜷阪ｒ蜈･繧後ｋ�ｼ�Excel 縺ｮ蟾･遞句錐繝ｻ讖滓｢ｰ蜷阪→辣ｧ蜷医＠繧�縺吶＞陦ｨ險假ｼ峨�

笆� preferred_operator�ｼ域枚蟄怜�暦ｼ俄� 譚｡莉ｶ莉倥″**蠢�鬆�**
- **蠢�隕∵擅莉ｶ**: 蠖楢ｩｲ萓晞�ｼ縺ｮ蜴滓枚繧定ｪｭ縺ｿ縲√�**隱ｰ縺後％縺ｮ蜉�蟾･繝ｻ菴懈･ｭ縺ｮ荳ｻ諡�蠖難ｼ�OP�ｼ峨→縺励※蜑ｲ繧雁ｽ薙※縺溘＞縺�**縲阪′ **諢丞袖縺ｨ縺励※** 隱ｭ縺ｿ蜿悶ｌ繧九→縺阪�
  萓�: 迚ｹ螳壹�ｮ莠ｺ縺ｫ繧�縺｣縺ｦ繧ゅｉ縺��ｼ上◎縺ｮ莠ｺ縺ｫ莉ｻ縺帙ｋ�ｼ乗球蠖薙�ｯ縺ゅ�ｮ莠ｺ�ｼ衆P縺ｯ縲懶ｼ上懊＆繧難ｼ域ｰ丞錐�ｼ峨↓萓晞�ｼ縲√↑縺ｩ縲�**陦ｨ迴ｾ縺ｮ蝙九↓萓晏ｭ倥○縺�**縲∵枚縺ｮ諢丞袖縺ｧ蛻､譁ｭ縺吶ｋ縲�
- **貅縺溘＠縺溘→縺阪�ｮ蜃ｺ蜉帷ｾｩ蜍�**: 荳願ｨ倥�ｮ諢丞袖縺梧�千ｫ九☆繧九→蛻､譁ｭ縺励◆繧ｪ繝悶ず繧ｧ繧ｯ繝医〒縺ｯ縲�**蠢�縺�** 繧ｭ繝ｼ `preferred_operator` 繧貞性繧√∝､縺ｯ **遨ｺ縺ｧ縺ｪ縺�譁�蟄怜��** 縺ｨ縺吶ｋ縲ゆｽｵ縺帙※ **process_name / machine_name 縺ｯ蠢�鬆�**�ｼ井ｾ�: `{{"process_name":"窶ｦ","machine_name":"窶ｦ","preferred_operator":"窶ｦ"}}`�ｼ峨�
- **蛟､縺ｮ蠖｢蠑�**: 蜴滓枚縺ｧ遉ｺ縺輔ｌ縺� **諡�蠖楢�縺ｮ隴伜挨蜷阪ｒ1蜷榊��**�ｼ亥ｧ薙�ｻ蜷阪�ｻ繝九ャ繧ｯ繝阪�ｼ繝�遲峨∝次譁�縺ｫ迴ｾ繧後◆陦ｨ險倥ｒ邯ｭ謖��ｼ峨よ忰蟆ｾ縺ｮ謨ｬ遘ｰ�ｼ医＆繧薙�ｻ蜷帙�ｻ豌擾ｼ峨�ｮ縺ｿ髯､蜴ｻ縲ゆｾ�:縲梧｣ｮ蟯｡縺輔ｓ縺ｫ繧�縺｣縺ｦ繧ゅｉ縺�縺ｾ縺吶坂�� `"譽ｮ蟯｡"`縲�
- **蜃ｺ蜉帙＠縺ｦ縺ｯ縺�縺代↑縺�縺ｨ縺�**: 蜴滓枚縺ｫ諡�蠖楢�縺ｮ謖�諢上′ **荳蛻�縺ｪ縺�** 縺ｨ蛻､譁ｭ縺励◆萓晞�ｼNO縺ｧ縺ｯ `preferred_operator` 繧ｭ繝ｼ閾ｪ菴薙ｒ **逵∫払** 縺吶ｋ�ｼ育ｩｺ譁�蟄怜�励ｂ莉倥￠縺ｪ縺��ｼ峨�

笆� 縺昴�ｮ莉悶ヵ繧｣繝ｼ繝ｫ繝会ｼ�required_op, speed_override, task_efficiency, priority, start_date, start_time, target_completion_date, ship_by_date�ｼ�
- 蜴滓枚縺九ｉ **譏守｢ｺ縺ｫ** 隱ｭ縺ｿ蜿悶ｌ繧句�ｴ蜷医�ｮ縺ｿ蜃ｺ蜉帙りｪｭ縺ｿ蜿悶ｌ縺ｪ縺�謨ｰ蛟､繝ｻ譌･莉倥�ｯ **逵∫払**�ｼ域耳貂ｬ縺ｧ蝓九ａ縺ｪ縺��ｼ峨�

縲仙酔荳萓晞�ｼNO繝ｻ隍�謨ｰ蟾･遞九�ｮ萓九�
萓晞�ｼNO Y4-2 縺ｫ縲後お繝ｳ繝懊せ縲阪→縲悟��蜑ｲ縲阪�ｮ陦後′縺ゅｊ縲∝ｙ閠�縺後�4/5縺ｾ縺ｧ縺ｫ邨ゅｏ繧峨○繧九阪�ｮ縺ｿ縺ｧ蟾･遞九�ｮ髯仙ｮ壹′辟｡縺�蝣ｴ蜷�:
- process_name / machine_name 縺ｯ **蛯呵�縺梧嶌縺九ｌ縺溯｡�** 縺ｮ蛟､繧貞�･繧後ｋ縲�
- restrict_to_* 縺ｯ **蜃ｺ縺輔↑縺�縺狗ｩｺ** 縺ｫ縺励�**繧ｨ繝ｳ繝懊せ陦後�ｻ蛻�蜑ｲ陦後�ｮ荳｡譁ｹ** 縺ｫ蜷後§蜆ｪ蜈亥ｺｦ繝ｻ譌･莉倡ｭ峨′蜉ｹ縺上ｈ縺�縺ｫ縺吶ｋ縲�

縲仙渕貅門ｹｴ�ｼ亥ｹｴ縺ｪ縺玲律莉倡畑�ｼ峨�
縲�4/5縲阪�4/5縺ｫ蜃ｺ闕ｷ縲阪�ｮ繧医≧縺ｫ **蟷ｴ縺檎┌縺�** 譌･莉倥�ｯ蜴溷援 **隘ｿ證ｦ {ref_y} 蟷ｴ** 縺ｨ縺励〆YYY-MM-DD 縺ｧ蜃ｺ蜉帙�

縲舌ヵ繧｣繝ｼ繝ｫ繝我ｸ隕ｧ�ｼ亥梛縺ｮ蜿り��ｼ峨�
- process_name, machine_name: 譁�蟄怜�暦ｼ亥ｿ�鬆医ゅ�励Ο繝ｳ繝励ヨ陦後→荳閾ｴ�ｼ�
- restrict_to_process_name, restrict_to_machine_name: 譁�蟄怜�暦ｼ井ｻｻ諢上る剞螳壹↑繧会ｼ�
- preferred_operator: 譁�蟄怜�暦ｼ井ｸ願ｨ伜･醍ｴ�縺ｫ蠕薙≧�ｼ�
- required_op: 豁｣縺ｮ謨ｴ謨ｰ
- speed_override: 豁｣縺ｮ謨ｰ�ｼ�m/蛻��ｼ�
- task_efficiency: 0縲�1
- priority: 謨ｴ謨ｰ�ｼ亥ｰ上＆縺�縺ｻ縺ｩ蜈医↓蜑ｲ莉假ｼ�
- start_date: YYYY-MM-DD / start_time: HH:MM
- target_completion_date, ship_by_date: YYYY-MM-DD

縲占ｧ｣驥医�ｮ謖�驥昴�
- 縲碁俣縺ｫ蜷医≧繧医≧縺ｫ縲阪檎ｹｰ繧贋ｸ翫￡繧九坂�� priority 繧剃ｸ翫￡繧具ｼ域焚蛟､繧剃ｸ九￡繧具ｼ峨よ律莉倥′譁�荳ｭ縺ｫ縺ゅｌ縺ｰ target_completion_date 縺ｾ縺溘�ｯ ship_by_date 縺ｫ蜈･繧後ｋ縲�
- 諡�蠖楢�謖�蜷阪�ｯ **諢丞袖逅�隗｣** 縺ｧ preferred_operator 繧呈ｱｺ繧√ｋ�ｼ育音螳壹�ｮ繧ｭ繝ｼ繝ｯ繝ｼ繝牙�玲嫌縺ｫ鬆ｼ繧峨↑縺��ｼ峨�
- 謨ｰ蛟､繝ｻ譌･莉倥�ｯ謗ｨ貂ｬ縺ｧ陬懊ｏ縺ｪ縺�縲�
- **蛯呵�縺檎音螳壹�ｮ蟾･遞九�ｻ讖滓｢ｰ縺ｫ縺�縺題ｨ蜿翫＠縺ｦ縺�縺ｪ縺�髯舌ｊ**縲〉estrict_to_* 縺ｯ遨ｺ縺ｫ縺励∝酔荳萓晞�ｼNO縺ｮ莉冶｡後↓繧る←逕ｨ縺輔ｌ繧句ｽ｢縺ｫ縺吶ｋ縲�

縲仙�ｺ蜉帷峩蜑阪�ｮ閾ｪ蟾ｱ讀懆ｨｼ�ｼ亥ｿ�縺壼ｮ溯｡後＠縺ｦ縺九ｉ JSON 繧帝哩縺倥ｋ�ｼ峨�
- 縲千音蛻･謖�螳壼次譁�縲代�ｮ **蜷�陦�** 縺ｫ縺､縺�縺ｦ縲∝ｯｾ蠢懊☆繧九が繝悶ず繧ｧ繧ｯ繝医↓ **process_name** 縺ｨ **machine_name** 縺後≠繧九°縲�
- 蜷御ｸ萓晞�ｼNO縺瑚､�謨ｰ陦後≠繧九→縺阪�ｯ **驟榊��** 縺ｧ蜷�陦後↓1繧ｪ繝悶ず繧ｧ繧ｯ繝医√∪縺溘�ｯ驕ｩ蛻�縺ｫ繝槭�ｼ繧ｸ縺励◆蜊倅ｸ繧ｪ繝悶ず繧ｧ繧ｯ繝茨ｼ脚estrict 縺ｮ驕狗畑繧剃ｸ雋ｫ縺輔○繧九�
- 縲御ｸｻ諡�蠖徹P縺ｮ謖�諢上阪′縺ゅｋ陦後〒縺ｯ **髱樒ｩｺ縺ｮ preferred_operator** 繧剃ｻ倥￠繧九�

縲仙�ｺ蜉帛ｽ｢蠑上�ｮ萓九托ｼ井ｾ晞�ｼNO繝ｻ蛟､縺ｯ螳溘ョ繝ｼ繧ｿ縺ｫ蜷医ｏ縺帶崛縺医ｋ縺薙→�ｼ�
{{
  "W3-14": {{
    "process_name": "讀懈渊",
    "machine_name": "繝ｩ繧､繝ｳA",
    "preferred_operator": "譽ｮ蟯｡"
  }},
  "Y3-26": {{
    "process_name": "繧ｳ繝ｼ繝�繧｣繝ｳ繧ｰ",
    "machine_name": "",
    "priority": 1,
    "ship_by_date": "{ref_y}-04-05",
    "target_completion_date": "{ref_y}-04-05"
  }},
  "Y4-2": {{
    "process_name": "繧ｨ繝ｳ繝懊せ",
    "machine_name": "E1",
    "priority": 2,
    "restrict_to_process_name": "",
    "restrict_to_machine_name": ""
  }}
}}

縲千音蛻･謖�螳壼次譁�縲托ｼ�Excel 縺九ｉ縺昴�ｮ縺ｾ縺ｾ縲�1陦鯉ｼ昜ｾ晞�ｼNO縺ｨ蛯呵�縺ｮ繝壹い�ｼ�
{blob}
"""
    try:
        ppath = os.path.join(log_dir, "ai_task_special_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: 莉雁屓 Gemini 縺ｫ貂｡縺励◆繝励Ο繝ｳ繝励ヨ蜈ｨ譁� 竊� %s", ppath)
    except OSError as ex:
        logging.warning("繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: 繝励Ο繝ｳ繝励ヨ菫晏ｭ伜､ｱ謨�: %s", ex)

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
            logging.info("繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: AI隗｣譫舌′螳御ｺ�縺励∪縺励◆縲�")
            if ai_sheet_sink is not None:
                ai_sheet_sink["迚ｹ蛻･謖�螳壼ｙ閠ダAI_API"] = "縺ゅｊ"
            return parsed
        if ai_sheet_sink is not None:
            ai_sheet_sink["迚ｹ蛻･謖�螳壼ｙ閠ダAI_API"] = "縺ゅｊ�ｼ�JSON隗｣驥亥､ｱ謨暦ｼ�"
        return {}
    except Exception as e:
        err_text = str(e)
        is_quota = ("429" in err_text) or ("RESOURCE_EXHAUSTED" in err_text)
        is_unavailable = ("503" in err_text) or ("UNAVAILABLE" in err_text)
        retry_sec = extract_retry_seconds(err_text) if is_quota else None
        if is_quota and retry_sec is not None:
            wait_sec = min(max(retry_sec, 1.0), 90.0)
            logging.warning(f"繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳� AI 429縲�{wait_sec:.1f}遘貞ｾ�讖溘＠縺ｦ蜀崎ｩｦ陦後＠縺ｾ縺吶�")
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
                        ai_sheet_sink["迚ｹ蛻･謖�螳壼ｙ閠ダAI_API"] = "縺ゅｊ�ｼ�429蜀崎ｩｦ陦悟ｾ鯉ｼ�"
                    return parsed
            except Exception as e2:
                logging.warning(f"繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳� AI 蜀崎ｩｦ陦悟､ｱ謨�: {e2}")
        elif is_unavailable:
            wait_sec = 8.0
            logging.warning(
                f"繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳� AI 503/UNAVAILABLE縲�{wait_sec:.1f}遘貞ｾ�讖溘＠縺ｦ蜀崎ｩｦ陦後＠縺ｾ縺吶�"
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
                    logging.info("繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: AI蜀崎ｩｦ陦後〒隗｣譫舌′螳御ｺ�縺励∪縺励◆縲�")
                    if ai_sheet_sink is not None:
                        ai_sheet_sink["迚ｹ蛻･謖�螳壼ｙ閠ダAI_API"] = "縺ゅｊ�ｼ�503蜀崎ｩｦ陦悟ｾ鯉ｼ�"
                    return parsed
                logging.warning("繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳� AI 503蜀崎ｩｦ陦�: JSON 謚ｽ蜃ｺ縺ｫ螟ｱ謨励＠縺ｾ縺励◆縲�")
            except Exception as e2:
                logging.warning(f"繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳� AI 503蜀崎ｩｦ陦悟､ｱ謨�: {e2}")
        else:
            logging.warning(f"繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳� AI 繧ｨ繝ｩ繝ｼ: {e}")
        logging.warning(
            "繧ｿ繧ｹ繧ｯ迚ｹ蛻･謖�螳�: AI隗｣譫千ｵ先棡繧貞叙蠕励〒縺阪↑縺九▲縺溘◆繧√∫音蛻･謖�螳喟蛯呵�縺ｮ髢句ｧ区律/蜆ｪ蜈域欠遉ｺ縺ｯ蜿肴丐縺輔ｌ縺ｾ縺帙ｓ縲�"
            "�ｼ亥�励悟刈蟾･髢句ｧ区律_謖�螳壹阪梧欠螳夂ｴ肴悄_荳頑嶌縺阪阪�ｯ蟒�豁｢貂医∩縲ょｙ閠�縺ｮ蜀崎ｨ倩ｼ峨∪縺溘�ｯ蠕後°繧� AI 蜀榊ｮ溯｡後ｒ讀懆ｨ弱＠縺ｦ縺上□縺輔＞縲ゑｼ�"
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["迚ｹ蛻･謖�螳壼ｙ閠ダAI_API"] = f"螟ｱ謨�: {e}"[:500]
        return {}


def _merge_preferred_operator_cell_and_ai(row, ai_for_tid):
    """Excel縲梧球蠖徹P_謖�螳壹阪ｒ蜆ｪ蜈医＠縲∫ｩｺ縺ｪ繧� AI 縺ｮ preferred_operator縲�"""
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
    繝｡繧､繝ｳ縲悟�榊━蜈育音蛻･險倩ｼ峨阪�ｮ task_preferred_operators縲�
    繧ｭ繝ｼ縺ｯ萓晞�ｼNO�ｼ亥､ｧ譁�蟄励�ｻ蟆乗枚蟄励�ｮ蟾ｮ縺ｯ辟｡隕厄ｼ峨�
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
    荳頑嶌縺榊�励�ｯ蜉�蟾･騾溷ｺｦ_荳頑嶌縺阪�ｻ蜴溷渚謚募�･譌･_荳頑嶌縺咲ｭ峨�ｮ縺ｿ�ｼ郁ｨ育判繧ｷ繝ｼ繝茨ｼ峨ゅ◎縺ｮ莉悶�ｯ迚ｹ蛻･謖�螳壼ｙ閠� AI 縺九ｉ縲�
    allow_ai_dispatch_priority_from_remark 縺� False 縺ｮ縺ｨ縺阪、I 縺ｮ required_op / task_efficiency / priority /
    start_date / start_time 縺ｯ謗｡逕ｨ縺励↑縺��ｼ亥ｙ閠�縺ｫ邏肴悄邉ｻ譁�險縺檎┌縺�陦悟髄縺托ｼ峨�
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
    迚ｹ蛻･謖�螳喟蛯呵�縺ｫ萓昴ｋ AI 隗｣譫千ｵ先棡縺ｨ縲∵�守､ｺ繧ｻ繝ｫ縺ｮ荳｡譁ｹ縺ｫ蛟､縺後≠繧企｣溘＞驕輔≧蛻励ｒ霑斐☆縲�
    蛯呵�繝ｻAI縺�縺壹ｌ縺区ｬ�縺代ｋ蝣ｴ蜷医�ｯ遨ｺ髮�蜷医�
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
    """Excel 陦檎分蜿ｷ(1蟋九∪繧翫�ｻ繝倥ャ繝繝ｼ=1陦檎岼) -> 遏帷崟縺後≠縺｣縺溷�怜錐縺ｮ髮�蜷�"""
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
    """譌｢縺ｫ髢九＞縺ｦ縺�繧矩�榊床險育判繧ｷ繝ｼ繝医∈縲∫泝逶ｾ蛻励�ｮ逹濶ｲ�ｼ郁埋鮟�繝ｪ繧ｻ繝�繝遺�定ｵ､�ｼ峨ｒ驕ｩ逕ｨ縺吶ｋ縲ゆｿ晏ｭ倥�ｯ蜻ｼ縺ｳ蜃ｺ縺怜�ｴ縲�"""
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
            # 繝輔か繝ｳ繝医�ｯ荳頑嶌縺阪＠縺ｪ縺��ｼ医ヶ繝�繧ｯ譌｢螳壹�ｻ繝ｦ繝ｼ繧ｶ繝ｼ險ｭ螳壹ｒ邯ｭ謖��ｼ�

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
    log_prefix: str = "谿ｵ髫�2",
) -> bool:
    """
    谿ｵ髫�2蜷代￠: 繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝郁ｧ｣譫舌ヶ繝ｭ繝�繧ｯ�ｼ�AX:AY�ｼ峨→遏帷崟繝上う繝ｩ繧､繝医ｒ **1蝗槭�ｮ load/save** 縺ｧ蜿肴丐縺吶ｋ縲�
    蠕捺擂縺ｯ蛻･髢｢謨ｰ縺ｧ繝悶ャ繧ｯ繧�2蝗樣幕縺�縺ｦ縺�縺溘◆繧√�.xlsm 縺悟､ｧ縺阪＞迺ｰ蠅�縺ｧ蜊∵焚遘貞腰菴阪�ｮ遏ｭ邵ｮ縺ｫ縺ｪ繧九�
    """
    if not wb_path or not os.path.isfile(wb_path):
        return False
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "%s: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ openpyxl 縺ｧ繧ｰ繝ｭ繝ｼ繝舌Ν隗｣譫舌�ｻ遏帷崟逹濶ｲ繧偵せ繧ｭ繝�繝励＠縺ｾ縺励◆縲�",
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
            "%s: 驟榊床繧ｷ繝ｼ繝井ｸ諡ｬ譖ｸ霎ｼ縺ｮ縺溘ａ繝悶ャ繧ｯ繧帝幕縺代∪縺帙ｓ: %s",
            log_prefix,
            ex,
        )
        return False
    try:
        if sheet_name not in wb.sheetnames:
            logging.info(
                "%s: 繧ｷ繝ｼ繝� '%s' 縺檎┌縺�縺溘ａ繧ｰ繝ｭ繝ｼ繝舌Ν隗｣譫舌�ｻ遏帷崟逹濶ｲ繧偵せ繧ｭ繝�繝励�",
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
                "%s: 驟榊床繧ｷ繝ｼ繝医∈縺ｮ荳諡ｬ菫晏ｭ倥↓螟ｱ謨暦ｼ�Excel 縺ｧ髢九＞縺溘∪縺ｾ遲会ｼ峨�"
                " 遏帷崟繝上う繝ｩ繧､繝医�ｯ '%s' 縺ｫ譖ｸ縺榊�ｺ縺励∪縺励◆縲ゅげ繝ｭ繝ｼ繝舌Ν隗｣譫舌�ｯ譛ｪ菫晏ｭ倥�ｮ蜿ｯ閭ｽ諤ｧ縺後≠繧翫∪縺吶� (%s)",
                log_prefix,
                _planning_conflict_sidecar_path(),
                e,
            )
            return False
        _remove_planning_conflict_sidecar_safe()
        _n_conf = len(conflicts_by_row) if conflicts_by_row else 0
        if _n_conf:
            logging.info(
                "%s: 縲�%s縲�%s:%s 蛻励↓繧ｰ繝ｭ繝ｼ繝舌Ν隗｣譫舌ｒ菫晏ｭ倥＠縲�"
                "迚ｹ蛻･謖�螳喟蛯呵�縺ｨ蛻励�ｮ遏帷崟 %s 陦後ｒ蜷後§菫晏ｭ倥〒繝上う繝ｩ繧､繝医＠縺ｾ縺励◆縲�",
                log_prefix,
                sheet_name,
                get_column_letter(lc),
                get_column_letter(vc),
                _n_conf,
            )
        else:
            logging.info(
                "%s: 縲�%s縲�%s:%s 蛻励↓繧ｰ繝ｭ繝ｼ繝舌Ν隗｣譫舌ｒ菫晏ｭ倥＠縺ｾ縺励◆�ｼ育泝逶ｾ陦後↑縺暦ｼ峨�",
                log_prefix,
                sheet_name,
                get_column_letter(lc),
                get_column_letter(vc),
            )
        return True
    except OSError as ex:
        logging.warning(
            "%s: 驟榊床繧ｷ繝ｼ繝井ｸ諡ｬ菫晏ｭ倥〒 OSError: %s",
            log_prefix,
            ex,
        )
        return False
    except Exception as ex:
        logging.warning(
            "%s: 驟榊床繧ｷ繝ｼ繝医∈縺ｮ繧ｰ繝ｭ繝ｼ繝舌Ν隗｣譫撰ｼ狗泝逶ｾ逹濶ｲ�ｼ井ｸ諡ｬ�ｼ峨〒萓句､�: %s",
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
    驟榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙す繝ｼ繝医�ｮ繝�繝ｼ繧ｿ陦後ｒ縲∫泝逶ｾ蛻励�ｮ縺ｿ襍､蝨ｰ繝ｻ逋ｽ螟ｪ蟄励↓縺吶ｋ縲�
    莠句燕繝代せ縺ｧ縺ｯ荳頑嶌縺榊�･蜉帛�励ｒ谿ｵ髫�1縺ｨ蜷後§阮�鮟�濶ｲ縺ｫ謌ｻ縺励√ヵ繧ｩ繝ｳ繝医�ｯ螟画峩縺励↑縺��ｼ井ｽ楢｣∫ｶｭ謖��ｼ峨�
    AI隗｣譫仙�励�ｯ逹濶ｲ縺励↑縺��ｼ域ｮｵ髫�1縺ｮ莉墓ｧ倥↓蜷医ｏ縺帙ｋ�ｼ峨�
    .xlsm 縺ｯ keep_vba=True 縺ｧ菫晏ｭ倥☆繧九�
    """
    if not wb_path or not os.path.exists(wb_path):
        return
    if _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "遏帷崟譖ｸ蠑�: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ openpyxl 縺ｧ縺ｮ繝上う繝ｩ繧､繝医ｒ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆縲�",
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )
        return
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = load_workbook(wb_path, keep_vba=keep_vba)
    try:
        if sheet_name not in wb.sheetnames:
            logging.warning(f"遏帷崟譖ｸ蠑�: 繧ｷ繝ｼ繝� '{sheet_name}' 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲�")
            return
        ws = wb[sheet_name]
        _plan_sheet_apply_conflict_styles_to_ws(ws, num_data_rows, conflicts_by_row)

        try:
            wb.save(wb_path)
        except OSError as e:
            write_planning_conflict_highlight_sidecar(sheet_name, num_data_rows, conflicts_by_row)
            logging.warning(
                "驟榊床繧ｷ繝ｼ繝医∈縺ｮ遏帷崟繝上う繝ｩ繧､繝医ｒ繝輔ぃ繧､繝ｫ菫晏ｭ倥〒縺阪∪縺帙ｓ縺ｧ縺励◆�ｼ�Excel 縺ｧ繝悶ャ繧ｯ繧帝幕縺�縺溘∪縺ｾ遲会ｼ峨�"
                " '%s' 縺ｫ謖�遉ｺ繧呈嶌縺榊�ｺ縺励∪縺励◆縲ゅ�槭け繝ｭ縺後す繝ｼ繝井ｸ翫↓逶ｴ謗･驕ｩ逕ｨ縺励∪縺吶� (%s)",
                _planning_conflict_sidecar_path(),
                e,
            )
        else:
            _remove_planning_conflict_sidecar_safe()
            if conflicts_by_row:
                logging.info(
                    f"迚ｹ蛻･謖�螳喟蛯呵�縺ｨ蛻励�ｮ遏帷崟: {len(conflicts_by_row)} 陦後ｒ '{sheet_name}' 縺ｧ繝上う繝ｩ繧､繝医＠縺ｾ縺励◆縲�"
                )
    finally:
        wb.close()


def _ai_planning_target_due_date(ai_dict):
    """AI JSON 縺ｮ螳御ｺ�繝ｻ蜃ｺ闕ｷ逶ｮ讓呎律縺九ｉ縲�驟榊床縺ｮ逶ｮ讓呎律1縺､繧呈ｱｺ繧√ｋ�ｼ郁､�謨ｰ縺ゅｌ縺ｰ譛繧よ掠縺�譌･�ｼ晏宍縺励＞譁ｹ�ｼ峨�"""
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
    迚ｹ蛻･謖�螳喟蛯呵�縺ｫ縲∫ｴ肴悄繝ｻ譛滄剞繝ｻ譛蜆ｪ蜈医↑縺ｩ縲碁�榊床隧ｦ陦後ｒ蜑阪↓蜃ｺ縺吶肴э蝗ｳ縺ｮ譁�險縺後≠繧九→縺� True縲�
    蛯呵�縺瑚ｨ伜�･縺輔ｌ縺ｦ縺�繧九□縺代〒縺ｯ True 縺ｫ縺励↑縺��ｼ�AI 逕ｱ譚･縺ｮ逶ｮ讓呎律繝ｻ髢句ｧ区律繝ｻ蜆ｪ蜈亥ｺｦ縺ｯ菴ｿ繧上↑縺��ｼ峨�
    """
    if not remark_raw:
        return False
    s = str(remark_raw).strip()
    if not s or s.lower() in ("nan", "none"):
        return False
    n = unicodedata.normalize("NFKC", s)
    n_lower = n.casefold()
    needles = (
        "邏肴悄",
        "謖�螳夂ｴ肴悄",
        "蝗樒ｭ皮ｴ肴悄",
        "險育判蝓ｺ貅�",
        "譛滓律",
        "邱�蛻�",
        "邱�繧∝��繧�",
        "譛滄剞",
        "譛蜆ｪ蜈�",
        "閾ｳ諤･",
        "諤･縺�",
        "逶ｴ縺｡縺ｫ",
        "譌ｩ諤･",
        "蜃ｺ闕ｷ",
        "邏榊�･",
        "蠢�逹",
        "deadline",
        "繝�繝�繝峨Λ繧､繝ｳ",
        "蜑榊偵＠",
        "譌ｩ繧√↓",
        "蜴ｳ螳�",
        "縺ｾ縺ｧ縺ｫ",
        "髢薙↓蜷医ｏ",
        "髢薙↓蜷医＞",
        "驕�繧後↑縺�",
        "驕�蟒ｶ荳榊庄",
        "蜆ｪ蜈磯�榊床",
        "蜈医↓驟榊床",
        "螳御ｺ�莠亥ｮ�",
        "譛ｬ邏肴悄",
        "蝗樒ｭ疲悄髯�",
    )
    return any(w.casefold() in n_lower for w in needles)


def _task_id_same_machine_due_tiebreak_key(task_id) -> tuple:
    """
    邏肴悄蝓ｺ貅厄ｼ亥屓遲披�呈欠螳夲ｼ峨�ｻ讖滓｢ｰ蜷阪′蜷後§蟶ｯ縺ｧ縺ｮ隧ｦ陦碁��縲�
    Y3-24 縺ｯ譛ｫ蟆ｾ縺ｮ謨ｰ蛟､縲�Y4-1-1 縺ｮ繧医≧縺ｫ繝上う繝輔Φ縺�2縺､莉･荳翫≠繧九→縺阪�ｯ縲梧怙蛻昴�ｮ - 縺ｮ逶ｴ蠕後阪�ｮ謨ｰ蛟､驛ｨ繧呈治逕ｨ縲�
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
# 驟榊床逕ｨ繧ｿ繧ｹ繧ｯ繧ｭ繝･繝ｼ
#   驟榊床險育判 DataFrame 1陦� 竊� 蜑ｲ莉倥い繝ｫ繧ｴ繝ｪ繧ｺ繝�逕ｨ dict 縺ｸ縺ｮ螟画鋤�ｼ亥━蜈亥ｺｦ繝ｻ邏肴悄繝ｻAI 荳頑嶌縺阪ｒ髮�邏��ｼ�
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
    ``generate_plan`` 蜀�縺ｧ蜻ｼ縺ｰ繧後ｋ縲ょｮ御ｺ�貂医∩繝ｻ驟榊床荳崎ｦ∬｡後ｒ髯､縺阪∵ｮ九ｊ繧� task_queue 縺ｫ遨阪�縲�
    ai_by_tid 縺� None 縺ｮ縺ｨ縺阪□縺大��驛ｨ縺ｧ analyze_task_special_remarks 繧貞ｮ溯｡後☆繧九�
    """
    if ai_by_tid is None:
        ai_by_tid = analyze_task_special_remarks(tasks_df, reference_year=run_date.year)
    task_queue = []
    n_exclude_plan = 0
    seq_by_tid = _collect_process_content_order_by_task_id(tasks_df)
    same_tid_line_seq = defaultdict(int)
    # 萓晞�ｼNO逶ｴ蛻鈴�榊床縺ｮ鬆�蠎冗畑: iterrows 縺ｮ隱ｭ縺ｿ霎ｼ縺ｿ鬆��ｼ�0 蟋九∪繧奇ｼ峨Ｕask_queue.sort 蠕後ｂ荳榊､峨�
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
        # 邏肴悄蝓ｺ貅�: 竭�蝗樒ｭ皮ｴ肴悄�ｼ育ｩｺ縺ｧ縺ｪ縺代ｌ縺ｰ�ｼ俄贈蛻励梧欠螳夂ｴ肴悄縲搾ｼ亥�励梧欠螳夂ｴ肴悄_荳頑嶌縺阪阪�ｯ蟒�豁｢貂医∩�ｼ�
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
                "蜴溷渚謚募�･譌･_荳頑嶌縺阪ｒ謗｡逕ｨ: 萓晞�ｼNO=%s 繧ｷ繝ｼ繝亥次蜿肴兜蜈･譌･=%s 荳頑嶌縺�=%s",
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
                "繝｡繧､繝ｳ蜀榊━蜈育音險�: 萓晞�ｼNO=%s 縺ｮ諡�蠖徹P繧偵げ繝ｭ繝ｼ繝舌Ν謖�蜷阪〒荳頑嶌縺� %r�ｼ医そ繝ｫ繝ｻ迚ｹ蛻･謖�螳壼ｙ閠アI繧医ｊ蜆ｪ蜈茨ｼ�",
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
                "繝｡繧､繝ｳ繧ｰ繝ｭ繝ｼ繝舌Ν: 萓晞�ｼNO=%s 蟾･遞�=%r 讖滓｢ｰ蜷�=%r 縺ｫ speed_multiplier 邏ｯ遨�=%s 繧帝←逕ｨ�ｼ磯溷ｺｦ %s 竊� %s�ｼ�",
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

        # 邏肴悄縺ｯ蜆ｪ蜈磯��菴阪�ｻ邱頑･蠎ｦ縺ｫ縺ｯ菴ｿ縺�縺後�髢句ｧ区律縺ｮ荳矩剞縺ｫ縺ｯ菴ｿ繧上↑縺��ｼ井ｽ吝鴨縺後≠繧後�ｰ蜑榊偵＠髢句ｧ九☆繧九◆繧��ｼ峨�
        if due_basis is None:
            due_urgent = False
        else:
            due_urgent = due_basis <= run_date

        # 髢句ｧ区律繝ｫ繝ｼ繝ｫ:
        # 1) 蜴溷渚謚募�･譌･縺後≠繧九→縺阪�ｯ縲悟次蜿肴兜蜈･譌･ 13:00 莉･髯阪阪ｒ髢句ｧ句庄閭ｽ譌･譎ゅ�ｮ荳矩剞縺ｨ縺吶ｋ縲�
        #    �ｼ域律莉倅ｸ矩剞: max(run_date, raw_input_date)縲∝酔譌･譎る俣荳矩剞: 13:00�ｼ�
        # 2) 迚ｹ蛻･謖�螳夲ｼ医そ繝ｫ/AI�ｼ峨�ｮ髢句ｧ区律縺後≠繧句�ｴ蜷医ｂ縲∝次蜿肴兜蜈･譌･繧医ｊ蜑榊偵＠縺ｫ縺ｯ縺励↑縺��ｼ�date 荳矩剞繧堤ｶｭ謖��ｼ�
        # 3) 蜴溷渚縺檎┌縺�縺ｨ縺阪�ｯ run_date
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
                    "髢句ｧ区律荳頑嶌縺阪�ｯ蜴溷渚謚募�･譌･繧医ｊ蜑榊偵＠荳榊庄: 萓晞�ｼNO=%s 謖�螳夐幕蟋区律=%s 蜴溷渚謚募�･譌･=%s 謗｡逕ｨ髢句ｧ区律=%s",
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
                # 邏肴悄蠕後ｍ蛟偵＠蜀崎ｩｦ陦後〒 due_basis_date 繧貞��驛ｨ +1 縺励※繧ゅ∫ｵ先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ蠖灘�暦ｼ亥�怜錐縺ｯ莠呈鋤縺ｧ縲瑚ｨ育判蝓ｺ貅也ｴ肴悄縲搾ｼ峨�ｯ縺薙�ｮ蛟､縺ｮ縺ｾ縺ｾ
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
        "task_queue 讒狗ｯ牙ｮ御ｺ�: total=%s�ｼ磯�榊床荳崎ｦ√↓繧医ｊ繧ｹ繧ｭ繝�繝� %s 陦鯉ｼ�",
        len(task_queue),
        n_exclude_plan,
    )
    return task_queue


def _task_id_priority_key(task_id):
    """
    萓晞�ｼNO縺ｮ蜷梧擅莉ｶ繧ｿ繧､繝悶Ξ繝ｼ繧ｯ逕ｨ繧ｭ繝ｼ縲�
    萓�: Y3-24, Y3-34 縺ｮ繧医≧縺ｪ蝣ｴ蜷医�ｯ繝上う繝輔Φ蠕悟濠縺ｮ謨ｰ蛟､縺悟ｰ上＆縺�譁ｹ繧貞━蜈医�
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
    萓晞�ｼNO逶ｴ蛻鈴�榊床縺ｮ蜃ｦ逅�鬆�縲ょ推萓晞�ｼNO縺ｫ縺､縺�縺ｦ **驟榊床隧ｦ陦碁��逡ｪ縺ｮ譛蟆丞､** 縺悟ｰ上＆縺�萓晞�ｼ繧貞�医↓螳瑚ｵｰ縺輔○繧�
    �ｼ亥酔荳萓晞�ｼ蜀�縺ｮ隍�謨ｰ陦後�ｯ譛蟆丞ｹ�縺ｮ隧ｦ陦碁��縺ｧ莉｣陦ｨ�ｼ峨ゅち繧､繝悶Ξ繝ｼ繧ｯ縺ｯ險育判繧ｷ繝ｼ繝井ｸ翫�ｮ蜈郁｡瑚｡�
    �ｼ�planning_sheet_row_seq�ｼ峨→萓晞�ｼNO繧ｭ繝ｼ縲�
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
    譌｢蟄倥す繝ｼ繝茨ｼ�read_excel�ｼ臥罰譚･縺ｮ繧ｹ繧ｫ繝ｩ繝ｼ繧偵�驟榊床險育判 DataFrame 縺ｮ譁�蟄怜�怜�暦ｼ�StringDtype�ｼ峨∈
    莉｣蜈･縺ｧ縺阪ｋ str 縺ｫ豁｣隕丞喧縺吶ｋ縲�Excel 縺梧焚蛟､縺ｨ縺励※菫晄戟縺励◆蜆ｪ蜈亥ｺｦ 1 竊� \"1\" 縺ｪ縺ｩ縲�
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
    繝悶ャ繧ｯ蜀�縺ｮ縲碁�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙阪↓繝ｦ繝ｼ繧ｶ繝ｼ縺悟�･蜉帙＠縺滉ｸ頑嶌縺榊�励ｒ縲�
    谿ｵ髫�1縺ｮ謚ｽ蜃ｺ邨先棡縺ｸ (萓晞�ｼNO, 蟾･遞句錐) 蜊倅ｽ阪〒蠑輔″邯吶＄縲�
    遨ｺ縺ｮ繧ｻ繝ｫ縺ｯ繝槭�ｼ繧ｸ縺励↑縺��ｼ域眠隕乗歓蜃ｺ蛛ｴ縺ｮ遨ｺ縺ｮ縺ｾ縺ｾ�ｼ峨�
    """
    if out_df is None or out_df.empty:
        return out_df
    if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
        return out_df
    try:
        df_old = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=PLAN_INPUT_SHEET_NAME)
    except Exception as e:
        logging.info("谿ｵ髫�1: 譌｢蟄倥�ｮ驟榊床繧ｷ繝ｼ繝医ｒ隱ｭ繧√↑縺�縺溘ａ荳頑嶌縺咲ｶ呎価繧偵せ繧ｭ繝�繝� (%s)", e)
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
            "谿ｵ髫�1: 譌｢蟄倥す繝ｼ繝医°繧我ｸ頑嶌縺榊�励ｒ %s 陦後∈蠑輔″邯吶℃縺ｾ縺励◆�ｼ医く繝ｼ: 萓晞�ｼNO+蟾･遞句錐�ｼ峨�",
            merged_rows,
        )
    return out_df


# ---------------------------------------------------------------------------
# 驟榊床荳崎ｦ��ｼ�2邉ｻ邨ｱ�ｼ�
#   (A) DataFrame 荳翫�ｮ繝ｫ繝ｼ繝ｫ 窶ｦ 蜷御ｸ萓晞�ｼNOﾃ怜酔荳讖滓｢ｰ縺ｧ縲悟��蜑ｲ縲崎｡後↓ yes�ｼ域焔蜈･蜉帙�ｯ荳頑嶌縺阪＠縺ｪ縺��ｼ�
#   (B) 繝槭け繝ｭ繝悶ャ繧ｯ縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九坂ｦ 蟾･遞�+讖滓｢ｰ縺斐→縺ｮ C/D/E 蛻励；emini 縺ｧ D竊脱縲�
#       菫晏ｭ倥Ο繝�繧ｯ譎ゅ�ｯ xlwings 縺ｧ A:E 蜷梧悄竊担ave 縺ｮ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ縺ゅｊ
#   縺�縺壹ｌ繧� apply_exclude_rules_config_to_plan_df 縺ｧ險育判 DataFrame 縺ｫ蜿肴丐縺輔ｌ繧九�
# ---------------------------------------------------------------------------

def _auto_exclude_cell_empty_for_autofill(v) -> bool:
    """驟榊床荳崎ｦ√そ繝ｫ縺梧悴蜈･蜉帙�ｮ縺ｨ縺阪□縺題�ｪ蜍輔〒 yes 繧呈嶌縺崎ｾｼ繧縲�"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return True
    if isinstance(v, str):
        s = str(v).strip()
        return not s or s.lower() in ("nan", "none")
    return False


def _normalize_task_id_for_dup_grouping(raw) -> str:
    """蜷御ｸ萓晞�ｼNO縺ｮ繧ｰ繝ｫ繝ｼ繝斐Φ繧ｰ逕ｨ�ｼ郁｡ｨ險倥ｆ繧後�ｻ闍ｱ蟄励�ｮ螟ｧ蟆上ｒ蟇�縺帙ｋ�ｼ峨�"""
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
    """蟾･遞句錐縺後悟��蜑ｲ縲搾ｼ育ｩｺ逋ｽ髯､蜴ｻ繝ｻNFKC 蠕鯉ｼ峨�"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\s縲]+", "", t)
    return t == "蛻�蜑ｲ"


def _apply_auto_exclude_bunkatsu_duplicate_machine(
    df: pd.DataFrame, log_prefix: str = "谿ｵ髫�1"
) -> pd.DataFrame:
    """
    蜷御ｸ萓晞�ｼNO縺�2陦御ｻ･荳翫≠繧翫√°縺､遨ｺ縺ｧ縺ｪ縺�蜷御ｸ讖滓｢ｰ蜷阪′2陦御ｻ･荳翫≠繧九げ繝ｫ繝ｼ繝励〒縺ｯ縲�
    蟾･遞句錐縺後悟��蜑ｲ縲阪�ｮ陦後�ｮ縲碁�榊床荳崎ｦ√阪↓ yes 繧貞�･繧後ｋ�ｼ医そ繝ｫ縺檎ｩｺ縺ｮ縺ｨ縺阪�ｮ縺ｿ�ｼ峨�
    讖滓｢ｰ蜷阪�ｯ _normalize_equipment_match_key 縺ｧ驥崎､�蛻､螳壹�
    """
    if df is None or df.empty:
        return df
    need_cols = (TASK_COL_TASK_ID, TASK_COL_MACHINE, TASK_COL_MACHINE_NAME)
    for c in need_cols:
        if c not in df.columns:
            return df
    if PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns:
        df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = ""
    # read_excel 遲峨〒 StringDtype 縺ｫ縺ｪ繧九→謨ｰ蛟､繝ｻ逵溷⊃縺ｮ .at 莉｣蜈･縺ｧ TypeError 縺ｫ縺ｪ繧九◆繧� object 縺ｫ蟇�縺帙ｋ
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
            # 蛻励′ StringDtype 縺ｮ縺ｨ縺� int 莉｣蜈･縺ｧ TypeError 縺ｫ縺ｪ繧九◆繧∵枚蟄怜�励↓縺吶ｋ�ｼ�_plan_row_exclude_from_assignment 縺ｯ yes 繧堤悄縺ｨ縺ｿ縺ｪ縺呻ｼ�
            df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
            n_set += 1

    if n_set:
        logging.info(
            "%s: 蜷御ｸ萓晞�ｼNO縺九▽蜷御ｸ讖滓｢ｰ蜷阪′隍�謨ｰ陦後≠繧九げ繝ｫ繝ｼ繝励〒縲∝ｷ･遞句錐縲悟��蜑ｲ縲阪�ｮ陦� %s 莉ｶ縺ｫ縲碁�榊床荳崎ｦ√�=yes 繧定�ｪ蜍戊ｨｭ螳壹＠縺ｾ縺励◆縲�",
            log_prefix,
            n_set,
        )
    return df


def _normalize_process_name_for_rule_match(raw) -> str:
    """蟾･遞句錐縺ｮ繝ｫ繝ｼ繝ｫ辣ｧ蜷茨ｼ�NFKC繝ｻ遨ｺ逋ｽ髯､蜴ｻ�ｼ峨�"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\s縲]+", "", t)
    return t


def _exclude_rules_sheet_header_map(ws) -> dict:
    """1陦檎岼隕句�ｺ縺� 竊� 蛻礼分蜿ｷ(1蟋九∪繧�)縲�
    openpyxl 縺ｯ譁ｰ隕上す繝ｼ繝育峩蠕後↓ max_column 縺� 0 縺ｮ縺ｾ縺ｾ縺ｮ縺薙→縺後≠繧翫∬ｦ句�ｺ縺励′隱ｭ繧√★菫晏ｭ伜燕縺ｫ return 縺励※縺励∪縺�縲�
    縺昴�ｮ縺溘ａ譛菴� A�ｽ昿 蛻励�ｯ蠢�縺夊ｵｰ譟ｻ縺吶ｋ縲�
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
    1陦檎岼縺ｫ讓呎ｺ冶ｦ句�ｺ縺暦ｼ亥ｷ･遞句錐繝ｻ讖滓｢ｰ蜷阪�ｻ驟榊床荳崎ｦ√�ｻ驟榊床荳崎�ｽ繝ｭ繧ｸ繝�繧ｯ繝ｻ繝ｭ繧ｸ繝�繧ｯ蠑擾ｼ峨′縺ゅｋ縺薙→繧剃ｿ晁ｨｼ縺吶ｋ縲�
    謇句虚縺ｧ遨ｺ繧ｷ繝ｼ繝医□縺題ｿｽ蜉�縺励◆蝣ｴ蜷医�ｯ A1:E1 縺檎ｩｺ縺ｮ縺溘ａ縲√％縺薙〒譖ｸ縺崎ｾｼ繧薙〒蛻礼分蜿ｷ繧定ｿ斐☆縲�
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
        "%s: 縲�%s縲阪�ｮ隕句�ｺ縺励′辟｡縺��ｼ丞�怜錐縺御ｸ閾ｴ縺励↑縺�縺溘ａ縲∵ｨ呎ｺ悶�ｮ1陦檎岼�ｼ�A1:E1�ｼ峨ｒ險ｭ螳壹＠縺ｾ縺励◆縲�",
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
    2 陦檎岼莉･髯阪°繧峨檎ｩｺ陦後阪ｒ髯､縺�縺ｦ荳翫↓隧ｰ繧√ｋ�ｼ亥��縺ｮ荳ｦ縺ｳ縺ｯ邯ｭ謖√√た繝ｼ繝医＠縺ｪ縺��ｼ峨�
    遨ｺ陦�: 蟾･遞句錐縺檎ｩｺ縲√∪縺溘�ｯ A�ｽ昿 逶ｸ蠖薙�ｮ5繧ｻ繝ｫ縺後☆縺ｹ縺ｦ遨ｺ逋ｽ逶ｸ蠖薙�
    Returns (谿九＠縺溘ョ繝ｼ繧ｿ陦梧焚, 蜑企勁縺励◆陦梧焚).
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
                "%s: 縲�%s縲阪�ｯ譛牙柑縺ｪ繝�繝ｼ繧ｿ陦後′辟｡縺九▲縺溘◆繧√√ョ繝ｼ繧ｿ陦� %s 陦後ｒ蜑企勁縺励∪縺励◆縲�",
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
            "%s: 縲�%s縲阪°繧臥ｩｺ陦後ｒ %s 莉ｶ蜑企勁縺励�%s 陦後↓隧ｰ繧√∪縺励◆�ｼ井ｸｦ縺ｳ鬆�縺ｯ邯ｭ謖��ｼ峨�",
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
    """C蛻励碁�榊床荳崎ｦ√阪′繧ｪ繝ｳ�ｼ医％縺ｮ蟾･遞�+讖滓｢ｰ繝代ち繝ｼ繝ｳ縺ｯ蟶ｸ縺ｫ驟榊床荳崎ｦ��ｼ峨�"""
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
    return s in ("yes", "true", "1", "y", "縺ｯ縺�", "笳�", "縲�", "笳�")


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
    """蜉�蟾･險育判DATA 縺九ｉ縲∵ｮｵ髫�1縺ｨ蜷後§謚ｽ蜃ｺ譚｡莉ｶ縺ｧ (蟾･遞句錐, 讖滓｢ｰ蜷�) 縺ｮ荳隕ｧ�ｼ磯㍾隍�髯､縺上�ｻ鬆�蠎冗ｶｭ謖��ｼ峨�"""
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
    """Gemini�ｼ拾蛻励°繧牙ｾ励◆ dict 縺碁�榊床荳崎ｦ√Ν繝ｼ繝ｫ縺ｨ縺励※譛牙柑縺九�"""
    if not isinstance(o, dict):
        return None
    if int(o.get("version") or 0) != 1:
        return None
    mode = str(o.get("mode") or "").strip().lower()
    if mode not in ("always_exclude", "conditions"):
        return None
    return o


def _exclude_rule_de_cache_key(stripped_blob: str) -> str:
    """縲碁�榊床荳崎�ｽ繝ｭ繧ｸ繝�繧ｯ縲肴枚險�ｼ域ｭ｣隕丞喧貂医∩�ｼ峨↓蟇ｾ縺吶ｋ ai_remarks_cache 逕ｨ繧ｭ繝ｼ縲�"""
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
        "縲舌せ繧ｭ繝ｼ繝� version 縺ｯ蠢�縺� 1縲曾n"
        "1) 蟶ｸ縺ｫ驟榊床荳崎ｦ��ｼ郁ｪｬ譏弱′譚｡莉ｶ縺ｪ縺励〒螟悶☆諢丞袖�ｼ峨�ｮ縺ｨ縺�:\n"
        '{"version":1,"mode":"always_exclude"}\n\n'
        "2) 蛻励�ｮ譚｡莉ｶ縺ｧ驟榊床荳崎ｦ√→縺吶ｋ縺ｨ縺�:\n"
        '{"version":1,"mode":"conditions","require_all": true 縺ｾ縺溘�ｯ false,"conditions":[ ... ]}\n\n'
        "conditions 縺ｮ蜷�隕∫ｴ�:\n"
        "- {\"column\":\"蛻怜錐\",\"op\":\"empty\"} 窶ｦ 繧ｻ繝ｫ縺檎ｩｺ\n"
        "- {\"column\":\"蛻怜錐\",\"op\":\"not_empty\"}\n"
        "- {\"column\":\"蛻怜錐\",\"op\":\"eq\",\"value\":\"譁�蟄怜�予"} / ne / contains / not_contains / regex�ｼ域ｭ｣隕剰｡ｨ迴ｾ�ｼ噂n"
        "- {\"column\":\"蛻怜錐\",\"op\":\"gt\"|\"gte\"|\"lt\"|\"lte\",\"value\":謨ｰ蛟､} 窶ｦ 謨ｰ蛟､豈碑ｼ��ｼ亥�励�ｯ謨ｰ縺ｨ縺励※隗｣驥茨ｼ噂n\n"
        f"縲蝉ｽｿ逕ｨ蜿ｯ閭ｽ縺ｪ蛻怜錐縺ｮ縺ｿ縲托ｼ医％繧御ｻ･螟悶�ｯ菴ｿ繧上↑縺��ｼ�:\n{allowed}\n"
    )


def _parse_exclude_rule_json_array_response(text: str) -> list | None:
    """繝｢繝�繝ｫ蠢懃ｭ斐°繧� JSON 驟榊�励ｒ蜿悶ｊ蜃ｺ縺呻ｼ�```json 繝輔ぉ繝ｳ繧ｹ莉倥″蜿ｯ�ｼ峨�"""
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
        logging.warning("驟榊床荳崎ｦ√Ν繝ｼ繝ｫ: 譛ｪ蟇ｾ蠢懊�ｮ蛻怜錐繧偵せ繧ｭ繝�繝励＠縺ｾ縺励◆: %s", col)
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
    E蛻励�ｮ JSON�ｼ�version=1�ｼ峨ｒ隧穂ｾ｡縺励∝ｽ楢ｩｲ繧ｿ繧ｹ繧ｯ陦後ｒ驟榊床荳崎ｦ√→縺吶∋縺阪↑繧� True縲�
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
    D蛻励�ｮ閾ｪ辟ｶ險隱槭ｒ Gemini 縺ｧ JSON 繝ｫ繝ｼ繝ｫ縺ｫ螟画鋤縲ょ､ｱ謨玲凾 None縲�
    json/ai_remarks_cache.json 縺ｫ TTL 莉倥″縺ｧ繧ｭ繝｣繝�繧ｷ繝･�ｼ亥酔荳譁�險縺ｪ繧� API 繧貞他縺ｰ縺ｪ縺��ｼ峨�
    """
    blob = str(natural_language or "").strip()
    if not blob:
        return None
    ai_cache = load_ai_cache()
    hit = _cache_get_exclude_rule_de_parsed(ai_cache, blob)
    if hit is not None:
        logging.info("驟榊床荳崎ｦ√Ν繝ｼ繝ｫ: AI繧ｭ繝｣繝�繧ｷ繝･繝偵ャ繝茨ｼ磯�榊床荳崎�ｽ繝ｭ繧ｸ繝�繧ｯ竊谷SON�ｼ�")
        return hit
    if not API_KEY:
        return None
    schema = _exclude_rule_logic_gemini_schema_instructions()
    prompt = (
        "縺ゅ↑縺溘�ｯ蟾･蝣ｴ縺ｮ驟榊床繧ｷ繧ｹ繝�繝�逕ｨ縺ｧ縺吶よｬ｡縺ｮ縲碁�榊床荳崎�ｽ縺ｮ隱ｬ譏弱阪ｒ縲√ち繧ｹ繧ｯ1陦後ｒ蛻､螳壹☆繧区ｩ滓｢ｰ蜿ｯ隱ｭ繝ｫ繝ｼ繝ｫ縺ｫ螟画鋤縺励※縺上□縺輔＞縲�\n\n"
        "縲仙�ｺ蜉帙大�磯�ｭ縺� { 縺ｧ邨ゅｏ繧翫′ } 縺ｮ JSON 繧ｪ繝悶ず繧ｧ繧ｯ繝�1縺､縺ｮ縺ｿ�ｼ郁ｪｬ譏弱�ｻ繝槭�ｼ繧ｯ繝繧ｦ繝ｳ遖∵ｭ｢�ｼ峨�\n\n"
        f"{schema}\n"
        f"縲占ｪｬ譏取枚縲曾n{blob}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("驟榊床荳崎ｦ√Ν繝ｼ繝ｫ: 繝励Ο繝ｳ繝励ヨ 竊� %s", ppath)
    except OSError as ex:
        logging.warning("驟榊床荳崎ｦ√Ν繝ｼ繝ｫ: 繝励Ο繝ｳ繝励ヨ菫晏ｭ伜､ｱ謨�: %s", ex)
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
        logging.warning("驟榊床荳崎ｦ√Ν繝ｼ繝ｫ: Gemini 螟画鋤螟ｱ謨�: %s", e)
        return None


def _ai_compile_exclude_rule_logics_batch(blobs: list[str]) -> list[dict | None]:
    """
    隍�謨ｰ縺ｮ D 蛻玲枚險繧� 1 蝗槭�ｮ Gemini 蜻ｼ縺ｳ蜃ｺ縺励〒 JSON 蛹悶ょ､ｱ謨励�ｻ隕∫ｴ�謨ｰ荳堺ｸ閾ｴ譎ゅ�ｯ 1 莉ｶ縺壹▽縺ｫ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ縲�
    json/ai_remarks_cache.json 縺ｫ繝偵ャ繝医＠縺滓枚險縺ｯ API 繧貞他縺ｰ縺ｪ縺�縲�
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
            "驟榊床荳崎ｦ√Ν繝ｼ繝ｫ: AI繧ｭ繝｣繝�繧ｷ繝･縺ｮ縺ｿ縺ｧ D竊脱 繝舌ャ繝� %s 莉ｶ繧貞ｮ檎ｵ撰ｼ�API 蜻ｼ縺ｳ蜃ｺ縺励↑縺暦ｼ峨�",
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
        "縺ゅ↑縺溘�ｯ蟾･蝣ｴ縺ｮ驟榊床繧ｷ繧ｹ繝�繝�逕ｨ縺ｧ縺吶ゆｻ･荳九�ｮ N 蛟九�ｮ縲碁�榊床荳崎�ｽ縺ｮ隱ｬ譏弱阪ｒ縲∽ｸ弱∴縺滄��蠎上〒縺昴ｌ縺槭ｌ JSON 繝ｫ繝ｼ繝ｫ縺ｫ螟画鋤縺励※縺上□縺輔＞縲�\n\n"
        f"縲仙�ｺ蜉帙遷SON 驟榊�励�ｮ縺ｿ縲ょ�磯�ｭ縺� [ 縺ｧ邨ゅｏ繧翫′ ] 縲りｦ∫ｴ�謨ｰ縺ｯ蠢�縺� {m}�ｼ�Markdown繝ｻ隱ｬ譏守ｦ∵ｭ｢�ｼ峨�\n"
        f"驟榊�励�ｮ蜈磯�ｭ隕∫ｴ�縺� [1]縲�2 逡ｪ逶ｮ縺� [2] 窶ｦ 縺ｫ蟇ｾ蠢懊＠縺ｾ縺吶�\n\n"
        f"{schema}\n"
        f"縲占ｪｬ譏取枚縲曾n{numbered}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_batch_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("驟榊床荳崎ｦ√Ν繝ｼ繝ｫ(繝舌ャ繝�): 繝励Ο繝ｳ繝励ヨ 竊� %s", ppath)
    except OSError as ex:
        logging.warning("驟榊床荳崎ｦ√Ν繝ｼ繝ｫ(繝舌ャ繝�): 繝励Ο繝ｳ繝励ヨ菫晏ｭ伜､ｱ謨�: %s", ex)
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
                "驟榊床荳崎ｦ√Ν繝ｼ繝ｫ: 繝舌ャ繝∝ｿ懃ｭ斐′荳肴ｭ｣�ｼ郁ｦ∫ｴ�謨ｰ %s縲∵悄蠕� %s�ｼ峨�1 莉ｶ縺壹▽蜀崎ｩｦ陦後＠縺ｾ縺吶�",
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
        logging.warning("驟榊床荳崎ｦ√Ν繝ｼ繝ｫ: 繝舌ャ繝� Gemini 螟ｱ謨励∝腰逋ｺ縺ｫ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ: %s", e)
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
    縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九阪�ｮ菫晏ｮ亥�ｦ逅�縺ｮ繧､繝吶Φ繝医Ο繧ｰ縲�

    險ｭ螳壹す繝ｼ繝亥�ｦ逅�縺ｮ謌仙凄繧� log/exclude_rules_sheet_debug.txt 縺ｫ霑ｽ險倥＠縲‘xecution_log 縺ｫ繧ゅち繧ｰ莉倥″縺ｧ蜃ｺ蜉帙☆繧九�
    event 萓�: START, OPEN_OK, OPEN_RETRY, OPEN_FAIL, HEADER_FIX, SYNC_ROWS, OPENPYXL_SAVE_OK, OPENPYXL_SAVE_FAIL,
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
        logging.warning("exclude_rules_sheet_debug.txt 縺ｸ譖ｸ縺代∪縺帙ｓ: %s", wex)

    tag = "[險ｭ螳喟驟榊床荳崎ｦ∝ｷ･遞犠"
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
    """繝�繧｣繧ｹ繧ｯ繝代せ縺ｨ xlwings Book.full_name 縺悟酔荳繝輔ぃ繧､繝ｫ繧呈欠縺吶°�ｼ郁｡ｨ險倥ｆ繧後ｒ螟壼ｰ大精蜿趣ｼ峨�"""
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
    """襍ｷ蜍穂ｸｭ縺ｮ Excel 縺九ｉ繝代せ荳閾ｴ縺吶ｋ xlwings Book 繧定ｿ斐☆縲ら┌縺代ｌ縺ｰ None縲�"""
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
    """譌｢蟄倥�ｮ Excel.App 縺ｧ Workbooks.Open 繧定ｩｦ縺吶よ�仙粥譎� Book縲∝､ｱ謨玲凾 None縲�"""
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
    """蟆ら畑襍ｷ蜍輔＠縺� Excel 縺ｯ邨ゆｺ�縺吶ｋ縲ょｮ溯｡御ｸｭ Excel 縺ｧ縺�縺� Open 縺励◆繝悶ャ繧ｯ縺ｯ螟ｱ謨玲凾縺ｮ縺ｿ髢峨§繧九�"""
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
    繝槭け繝ｭ繝悶ャ繧ｯ繧� xlwings 縺ｧ蜿門ｾ励☆繧具ｼ域悽逡ｪ繝ｻ繝�繧ｹ繝亥�ｱ騾夲ｼ峨�
    謌ｻ繧雁､: (Book, release_info) / 螟ｱ謨玲凾 None縲�
    release_info: mode 縺� keep 縺ｾ縺溘�ｯ quit_excel縲｛pened_wb_here 縺� bool縲�
    """
    try:
        import xlwings as xw  # noqa: F401
    except ImportError:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_UNAVAILABLE",
            log_prefix,
            "xlwings 縺� import 縺ｧ縺阪∪縺帙ｓ�ｼ�pip install xlwings 繧堤｢ｺ隱搾ｼ峨�",
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
            "xlwings 縺ｧ繝悶ャ繧ｯ繧帝幕縺代∪縺帙ｓ縺ｧ縺励◆縲�",
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
    讀懆ｨｼ繧ｹ繧ｯ繝ｪ繝励ヨ逕ｨ: 襍ｷ蜍穂ｸｭ繝悶ャ繧ｯ繧貞━蜈医＠縲∝ｿ�隕√↑繧芽｡ｨ遉ｺ莉倥″ Excel 縺ｧ髢九￥縲�
    謌ｻ繧雁､: (Book, info, 隱ｬ譏取枚蟄怜��) 縺ｾ縺溘�ｯ None縲�
    """
    abs_path = os.path.abspath(book_path)
    book = _xlwings_find_book_on_running_instances(abs_path)
    if book is not None:
        return book, {"mode": "keep", "opened_wb_here": False}, f"{label}:譌｢蟄倥う繝ｳ繧ｹ繧ｿ繝ｳ繧ｹ"
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
    """VBA 蛛ｴ縺ｮ繧ｹ繝励Λ繝�繧ｷ繝･繝昴�ｼ繝ｪ繝ｳ繧ｰ縺ｨ遶ｶ蜷医＠縺ｫ縺上￥縺吶ｋ縺溘ａ縲∝酔譛溘�ｻ菫晏ｭ倥�ｮ遏ｭ譎る俣縺�縺� Excel 繧帝撕縺九↓縺吶ｋ縲�"""
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
    openpyxl 縺ｧ菫晏ｭ倥〒縺阪↑縺�縺ｨ縺阪』lwings 縺ｧ縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九喉:E 繧偵Γ繝｢繝ｪ荳翫�ｮ蛟､縺ｧ荳頑嶌縺阪＠ Save縲�

    陦ｨ遉ｺ荳ｭ繧ｷ繝ｼ繝医↓蟇ｾ縺吶ｋ荳諡ｬ .value 縺�縺代□縺ｨ縲√せ繝励Λ繝�繧ｷ繝･�ｼ九�昴�ｼ繝ｪ繝ｳ繧ｰ�ｼ�D3=true�ｼ我ｸ九〒
    Range 莉｣蜈･縺梧焚蛻�縺九°繧玖ｨ域ｸｬ縺後≠繧雁ｾ励ｋ縲ょ酔譛滉ｸｭ縺ｮ縺ｿ繧ｷ繝ｼ繝医ｒ荳譎る撼陦ｨ遉ｺ縺ｫ縺� api.Value2 縺ｧ譖ｸ縺上�
    """
    global _exclude_rules_effective_read_path

    attached = _xlwings_attach_open_macro_workbook(wb_path, log_prefix)
    if attached is None:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_SKIP",
            log_prefix,
            "xlwings 縺ｧ繝悶ャ繧ｯ縺ｫ謗･邯壹〒縺阪★ A:E 蜷梧悄繧偵せ繧ｭ繝�繝励�",
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
        # 蜈ｨ繧ｷ繝ｼ繝亥錐繧貞�玲嫌縺吶ｋ縺ｨ繧ｷ繝ｼ繝域焚蛻�縺ｮ COM 蠕蠕ｩ縺ｫ縺ｪ繧翫．3=true 譎ゅ�ｯ VBA 繝昴�ｼ繝ｪ繝ｳ繧ｰ縺ｨ遶ｶ蜷医＠縺ｦ
        # 1 繧ｷ繝ｼ繝域焚遘偵懷香謨ｰ遘偵°縺九ｋ縺薙→縺後≠繧具ｼ郁ｨ域ｸｬ縺ｧ 40 繧ｷ繝ｼ繝遺沿213s�ｼ峨ょ錐蜑阪〒逶ｴ謗･隗｣豎ｺ縺吶ｋ縲�
        try:
            sht = xw_book.sheets[EXCLUDE_RULES_SHEET_NAME]
        except Exception:
            _log_exclude_rules_sheet_debug(
                "XLWINGS_SYNC_SKIP",
                log_prefix,
                f"xlwings 蛛ｴ縺ｫ繧ｷ繝ｼ繝医鶏EXCLUDE_RULES_SHEET_NAME}縲阪′縺ゅｊ縺ｾ縺帙ｓ縲�",
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
                    sht.api.Visible = 0  # xlSheetHidden�ｼ亥酔譛滉ｸｭ縺�縺代ょ�肴緒逕ｻ繝ｻ繧ｦ繧｣繝ｳ繝峨え譖ｴ譁ｰ雋�闕ｷ繧呈椛縺医ｋ�ｼ�
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
            "xlwings 邨檎罰縺ｧ險ｭ螳壹す繝ｼ繝� A縲廢 繧貞酔譛溘＠繝悶ャ繧ｯ繧剃ｿ晏ｭ倥＠縺ｾ縺励◆縲�",
            details=f"path={wb_path} rows={max_r}",
        )
        logging.info(
            "%s: 險ｭ螳壹す繝ｼ繝医ｒ xlwings 縺ｧ繝槭け繝ｭ繝悶ャ繧ｯ縺ｫ菫晏ｭ倥＠縺ｾ縺励◆�ｼ�A縲廢�ｼ峨�",
            log_prefix,
        )
        return True
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "XLWINGS_SYNC_FAIL",
            log_prefix,
            "xlwings 縺ｧ縺ｮ A:E 蜷梧悄縺ｾ縺溘�ｯ Save 縺ｫ螟ｱ謨励＠縺ｾ縺励◆縲�",
            details=f"path={wb_path}",
            exc=ex,
        )
        return False
    finally:
        _xlwings_release_book_after_mutation(xw_book, info, ok)


# 險ｭ螳壹す繝ｼ繝医�ｮ蛻礼ｯ�蝗ｲ�ｼ�A縲廢�ｼ峨Ｙlwings 蜷梧悄繝ｻVBA 陦悟�� TSV 蜃ｺ蜉帙〒繧ゆｽｿ逕ｨ縲�
EXCLUDE_RULES_SHEET_COM_SYNC_MAX_COL = 5
EXCLUDE_RULES_MATRIX_CLIP_MAX_COL = 5


def _persist_exclude_rules_workbook(_wb, wb_path: str, ws, log_prefix: str) -> bool:
    """
    險ｭ螳壹す繝ｼ繝医�ｮ繝�繧｣繧ｹ繧ｯ蜿肴丐縲よ里螳壹�ｯ xlwings 縺ｧ A:E 蜷梧悄竊担ave�ｼ�EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1 縺ｮ縺ｨ縺阪�ｮ縺ｿ openpyxl save 繧定ｩｦ陦鯉ｼ峨�
    菫晏ｭ倥〒縺阪↑縺�縺ｨ縺阪�ｯ log 縺ｫ陦悟�� TSV 繧貞�ｺ縺励〃BA縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞祇A縺九ｉE_TSV縺九ｉ蜿肴丐縲阪〒蜿肴丐縺吶ｋ縲�

    _wb 窶ｦ 邱ｨ髮�貂医∩ openpyxl 繝悶ャ繧ｯ�ｼ�openpyxl 邨瑚ｷｯ譎ゅ�ｮ縺ｿ save 縺ｫ菴ｿ逕ｨ�ｼ峨�
    """
    global _exclude_rules_effective_read_path

    def _openpyxl_persist_ok(which: str) -> bool:
        try:
            _wb.save(wb_path)
        except Exception as ex:
            _log_exclude_rules_sheet_debug(
                "OPENPYXL_SAVE_FAIL",
                log_prefix,
                f"openpyxl 縺ｧ縺ｮ .xlsm 菫晏ｭ倥↓螟ｱ謨励＠縺ｾ縺励◆ {which}�ｼ�Excel 縺ｧ髢九″縺｣縺ｱ縺ｪ縺励�ｻ繝ｭ繝�繧ｯ縺ｮ蜿ｯ閭ｽ諤ｧ�ｼ峨�",
                details=f"path={wb_path}",
                exc=ex,
            )
            return False
        _exclude_rules_effective_read_path = wb_path
        _clear_exclude_rules_e_apply_files()
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_OK",
            log_prefix,
            "openpyxl 縺ｧ險ｭ螳壹す繝ｼ繝医ｒ蜷ｫ繧繝悶ャ繧ｯ繧剃ｿ晏ｭ倥＠縺ｾ縺励◆�ｼ�A縲廢�ｼ峨�",
            details=f"path={wb_path} {which}",
        )
        logging.info(
            "%s: 險ｭ螳壹す繝ｼ繝医ｒ openpyxl 縺ｧ繝槭け繝ｭ繝悶ャ繧ｯ縺ｫ菫晏ｭ倥＠縺ｾ縺励◆縲�%s",
            log_prefix,
            which,
        )
        return True

    saved_openpyxl = False
    if EXCLUDE_RULES_SKIP_OPENPYXL_SAVE:
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_SKIPPED_EXCLUDE_RULES_POLICY",
            log_prefix,
            "險ｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九�ｮ菫晏ｭ倥〒縺ｯ openpyxl save 繧定ｩｦ陦後＠縺ｾ縺帙ｓ�ｼ�xlwings 蜷梧悄繧貞�郁｡後ょ�崎ｩｦ陦後☆繧句�ｴ蜷医�ｯ EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1�ｼ峨�",
            details=f"path={wb_path}",
        )
        logging.info(
            "%s: 險ｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九�ｯ openpyxl 繧定ｩｦ縺輔★ xlwings 蜷梧悄竊担ave 繧定ｩｦ縺ｿ縺ｾ縺呻ｼ井ｸ榊庄縺ｪ繧� VBA 逕ｨ陦悟�� TSV�ｼ峨�",
            log_prefix,
        )
    elif not _workbook_should_skip_openpyxl_io(wb_path):
        logging.info(
            "%s: 險ｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九�ｯ openpyxl 縺ｧ菫晏ｭ倥＠縺ｾ縺呻ｼ井ｸ榊庄縺ｮ縺ｨ縺阪�ｯ xlwings 蜷梧悄竊担ave縲√◎繧後ｂ荳榊庄縺ｪ繧� VBA 逕ｨ陦悟�� TSV�ｼ峨�",
            log_prefix,
        )
        labels = ("(1/4)", "(2/4)", "(3/4)", "(4/4)")
        for i, label in enumerate(labels):
            if i:
                _log_exclude_rules_sheet_debug(
                    "OPENPYXL_RETRY_WAIT",
                    log_prefix,
                    f"openpyxl 蜀堺ｿ晏ｭ倥∪縺ｧ 2 遘貞ｾ�縺｡縺ｾ縺� {label}縲�",
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
            f"繝悶ャ繧ｯ縺ｫ縲鶏OPENPYXL_INCOMPATIBLE_SHEET_MARKER}縲阪′縺ゅｋ縺溘ａ openpyxl 縺ｧ縺ｮ菫晏ｭ倥ｒ隧ｦ縺ｿ縺ｾ縺帙ｓ縲�",
            details=f"path={wb_path}",
        )
        logging.info(
            "%s: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ openpyxl save 繧偵せ繧ｭ繝�繝励＠縲』lwings 縺ｾ縺溘�ｯ陦悟�� TSV 縺ｫ蛻�繧頑崛縺医∪縺吶�",
            log_prefix,
            OPENPYXL_INCOMPATIBLE_SHEET_MARKER,
        )

    if saved_openpyxl:
        return True

    if _xlwings_sync_exclude_rules_sheet_from_openpyxl(wb_path, ws, log_prefix):
        return True

    if _write_exclude_rules_matrix_vba_tsv(wb_path, ws, log_prefix):
        logging.warning(
            "%s: 險ｭ螳壹す繝ｼ繝医ｒ log\\%s 縺ｫ蜃ｺ蜉帙＠縺ｾ縺励◆縲�"
            " Excel 縺ｧ繝槭け繝ｭ縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞祇A縺九ｉE_TSV縺九ｉ蜿肴丐縲阪ｒ螳溯｡後＠縺ｦ縺上□縺輔＞縲�",
            log_prefix,
            EXCLUDE_RULES_MATRIX_VBA_FILENAME,
        )

    _log_exclude_rules_sheet_debug(
        "OPENPYXL_VBA_FALLBACK",
        log_prefix,
        "openpyxl 菫晏ｭ倥↓螟ｱ謨励＠縺溘◆繧� VBA 逕ｨ陦悟�� TSV 繧貞�ｺ蜉帙＠縺ｾ縺励◆�ｼ医ヶ繝�繧ｯ縺ｯ Excel 荳翫〒謇句虚蜿肴丐縺悟ｿ�隕√↑蝣ｴ蜷医′縺ゅｊ縺ｾ縺呻ｼ峨�",
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
    """VBA 逕ｨ: 險ｭ螳壹す繝ｼ繝� 1 陦檎岼縲� max_row 縺ｮ A縲廢 繧� Base64(UTF-8) 莉倥″ TSV 縺ｧ蜃ｺ蜉帙☆繧九�"""
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
            "險ｭ螳壹す繝ｼ繝� A縲廢 繧� VBA 蜿肴丐逕ｨ TSV 縺ｫ譖ｸ縺榊�ｺ縺励∪縺励◆�ｼ�openpyxl 菫晏ｭ倅ｸ榊庄譎ゑｼ峨�",
            details=f"path={path} rows={max_r}",
        )
        return True
    except OSError as ex:
        logging.warning("%s: 陦悟�� VBA 逕ｨ TSV 繧呈嶌縺代∪縺帙ｓ: %s", log_prefix, ex)
        return False


def _build_exclude_rules_list_from_openpyxl_ws(
    ws, c_proc: int, c_mach: int, c_flag: int, c_e: int
) -> list[dict]:
    """openpyxl 荳翫�ｮ險ｭ螳壹す繝ｼ繝医°繧� _load_exclude_rules_from_workbook 縺ｨ蜷悟ｽ｢縺ｮ繝ｪ繧ｹ繝医ｒ讒狗ｯ峨�"""
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
    """VBA 逕ｨ: 陦檎分蜿ｷ縺ｨ Base64(UTF-8) 繧ｻ繝ｫ譁�蟄怜�励�ｮ TSV縲�"""
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
            "E 蛻励ｒ VBA 蜿肴丐逕ｨ TSV 縺ｫ譖ｸ縺榊�ｺ縺励∪縺励◆�ｼ井ｿ晏ｭ伜､ｱ謨玲凾縺ｮ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ逕ｨ�ｼ峨�",
            details=f"path={path_tsv} cells={len(cells)}",
        )
    except OSError as ex:
        logging.warning("%s: E 蛻� VBA 逕ｨ TSV 繧呈嶌縺代∪縺帙ｓ: %s", log_prefix, ex)


def _write_exclude_rules_e_apply_artifacts(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> None:
    """
    E 蛻暦ｼ磯撼遨ｺ�ｼ峨ｒ JSON 繧ｵ繧､繝峨き繝ｼ繝峨→ VBA 逕ｨ TSV 縺ｫ譖ｸ縺上らｩｺ縺ｪ繧我ｸ｡繝輔ぃ繧､繝ｫ繧貞炎髯､縲�
    Python 谺｡蝗櫁ｵｷ蜍墓凾縺ｮ E 蠕ｩ蜈�逕ｨ JSON 縺ｨ縲√�槭け繝ｭ縺九ｉ縺ｮ E 譖ｸ霎ｼ縺ｿ逕ｨ TSV縲�
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
        logging.warning("%s: E 蛻� JSON 繧呈嶌縺代∪縺帙ｓ: %s", log_prefix, ex)
    _write_exclude_rules_e_vba_tsv_from_cells(wb_path, c_e, cells, log_prefix)
    _log_exclude_rules_sheet_debug(
        "E_APPLY_FILES_WRITTEN",
        log_prefix,
        "E 蛻励ｒ JSON 縺ｨ VBA 逕ｨ TSV 縺ｫ譖ｸ縺榊�ｺ縺励∪縺励◆�ｼ医�槭け繝ｭ縺ｧ E 蛻励ｒ蜿肴丐蠕後√ヵ繧｡繧､繝ｫ蜑企勁�ｼ峨�",
        details=f"cells={len(cells)}",
    )


def _try_apply_pending_exclude_rules_e_column(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> int:
    """
    蜑榊屓菫晏ｭ倥↓螟ｱ謨励＠縺溘→縺肴嶌縺榊�ｺ縺励◆ JSON 縺九ｉ E 蛻励ｒ蠕ｩ蜈�縺吶ｋ縲�
    繝悶ャ繧ｯ繝代せ縺御ｸ閾ｴ縺励↑縺代ｌ縺ｰ菴輔ｂ縺励↑縺�縲る←逕ｨ蠕後�ｯ繧ｵ繧､繝峨き繝ｼ繝峨ｒ蜑企勁縺吶ｋ縲�
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
            f"譛ｪ菫晏ｭ倥□縺｣縺� E 蛻励ｒ繧ｵ繧､繝峨き繝ｼ繝峨°繧� {n} 繧ｻ繝ｫ蠕ｩ蜈�縺励∪縺励◆縲�",
            details=path_sc,
        )
        logging.info(
            "%s: %s 縺ｮ蜀�螳ｹ繧偵す繝ｼ繝医�ｮ繝ｭ繧ｸ繝�繧ｯ蠑丞�励∈驕ｩ逕ｨ縺励∪縺励◆�ｼ育ｶ壹￠縺ｦ菫晏ｭ倥ｒ隧ｦ縺ｿ縺ｾ縺呻ｼ峨�",
            log_prefix,
            path_sc,
        )
    return n


def _read_exclude_rules_d_cells_data_only_for_rows(
    wb_path: str, rows: list[int], c_d: int
) -> dict[int, object]:
    """
    D 蛻励′謨ｰ蠑上�ｮ縺ｨ縺阪｛penpyxl 縺ｮ騾壼ｸｸ隱ｭ霎ｼ縺ｧ縺ｯ '=...' 縺励°蜿悶ｌ縺ｪ縺�縲�
    data_only=True 縺ｧ繧ｭ繝｣繝�繧ｷ繝･蛟､繧定ｪｭ繧�ｼ�Excel 縺御ｸ蠎ｦ縺ｧ繧ゆｿ晏ｭ倥�ｻ險育ｮ玲ｸ医∩縺ｮ繝悶ャ繧ｯ縺ｧ譛牙柑�ｼ峨�
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
    縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九阪�ｮ陦悟酔譛溘�ｻD竊脱 縺ｮ AI 陬懷ｮ後�ｻ繝�繧｣繧ｹ繧ｯ蜿肴丐�ｼ域里螳壹�ｯ xlwings 縺ｧ A縲廢 蜷梧悄竊担ave縲Ａ`EXCLUDE_RULES_TRY_OPENPYXL_SAVE=1`` 縺ｮ縺ｨ縺� openpyxl save 繧定ｩｦ陦鯉ｼ峨�

    xlwings 縺ｧ繧ゆｿ晏ｭ倥〒縺阪↑縺�縺ｨ縺阪�ｯ ``log/exclude_rules_matrix_vba.tsv`` 繧呈ｮ九＠縲√�槭け繝ｭ
    ``險ｭ螳喟驟榊床荳崎ｦ∝ｷ･遞祇A縺九ｉE_TSV縺九ｉ蜿肴丐`` 縺ｧ A縲廢 繧貞渚譏�縺吶ｋ縲�
    菴ｵ縺帙※蠕捺擂縺ｩ縺翫ｊ E 蛻励�ｮ縺ｿ縺ｮ ``exclude_rules_e_column_vba.tsv`` 繧ょ�ｺ蜉帙＆繧悟ｾ励ｋ�ｼ郁｡悟�� TSV 蜆ｪ蜈医〒蜿肴丐蠕後�ｯ蜑企勁�ｼ峨�
    菫晏ｭ俶�仙粥譎ゅ�ｯ TSV/JSON 縺ｯ蜑企勁縺輔ｌ繧九�

    ``json/exclude_rules_e_column_pending.json`` 縺ｯ Python 谺｡蝗櫁ｵｷ蜍墓凾縺ｮ E 蛻怜ｾｩ蜈�逕ｨ縲�
    繧ｷ繝ｼ繝医�ｮ譁ｰ隕丈ｽ懈�舌→ 1 陦檎岼隕句�ｺ縺励�ｯ VBA縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞祇繧ｷ繝ｼ繝医ｒ遒ｺ菫昴阪�
    """
    if not wb_path:
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_PATH",
            log_prefix,
            "TASK_INPUT_WORKBOOK 縺檎ｩｺ縺ｮ縺溘ａ險ｭ螳壹す繝ｼ繝亥�ｦ逅�繧偵＠縺ｾ縺帙ｓ縲�",
        )
        return
    if not os.path.exists(wb_path):
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_FILE",
            log_prefix,
            "繝悶ャ繧ｯ縺悟ｭ伜惠縺励∪縺帙ｓ縲�",
            details=f"path={wb_path}",
        )
        return

    _log_exclude_rules_sheet_debug(
        "START",
        log_prefix,
        "險ｭ螳壹す繝ｼ繝井ｿ晏ｮ磯幕蟋�",
        details=f"path={wb_path} pairs={len(pairs)}",
    )
    global _exclude_rules_effective_read_path
    _exclude_rules_effective_read_path = None

    if _workbook_should_skip_openpyxl_io(wb_path):
        _log_exclude_rules_sheet_debug(
            "SKIP_OPENPYXL_INCOMPATIBLE_BOOK",
            log_prefix,
            f"繝悶ャ繧ｯ縺ｫ縲鶏OPENPYXL_INCOMPATIBLE_SHEET_MARKER}縲阪′蜷ｫ縺ｾ繧後ｋ縺溘ａ縲｛penpyxl 縺ｫ繧医ｋ險ｭ螳壹す繝ｼ繝井ｿ晏ｮ医�ｯ陦後＞縺ｾ縺帙ｓ縲�",
            details=f"path={wb_path}",
        )
        logging.warning(
            "%s: 縲�%s縲榊性譛峨�ｮ縺溘ａ縲�%s縲阪�ｮ openpyxl 菫晏ｮ医ｒ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆�ｼ�Excel�ｼ駿lwings 縺ｧ邱ｨ髮�縺励※縺上□縺輔＞�ｼ峨�",
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
                "keep_vba=True 縺ｧ繝悶ャ繧ｯ繧帝幕縺代★ keep_vba=False 縺ｧ蜀崎ｩｦ陦後＠縺ｾ縺呻ｼ医�槭け繝ｭ縺悟､ｱ繧上ｌ繧句庄閭ｽ諤ｧ�ｼ峨�",
                exc=e1,
            )
            try:
                wb = load_workbook(wb_path, keep_vba=False, read_only=False, data_only=False)
            except Exception as e2:
                _log_exclude_rules_sheet_debug(
                    "OPEN_FAIL",
                    log_prefix,
                    "繝悶ャ繧ｯ繧帝幕縺代∪縺帙ｓ縲ゅす繝ｼ繝医�ｯ菴懈�舌�ｻ菫晏ｭ倥＆繧後∪縺帙ｓ縲�",
                    details=f"path={wb_path}",
                    exc=e2,
                )
                return
        else:
            _log_exclude_rules_sheet_debug(
                "OPEN_FAIL",
                log_prefix,
                "繝悶ャ繧ｯ繧帝幕縺代∪縺帙ｓ縲ゅす繝ｼ繝医�ｯ菴懈�舌�ｻ菫晏ｭ倥＆繧後∪縺帙ｓ縲�",
                details=f"path={wb_path}",
                exc=e1,
            )
            return

    _log_exclude_rules_sheet_debug(
        "OPEN_OK",
        log_prefix,
        "繝悶ャ繧ｯ繧帝幕縺阪∪縺励◆縲�",
        details=f"keep_vba={keep_vba} sheets={len(wb.sheetnames)}",
    )

    try:
        if EXCLUDE_RULES_SHEET_NAME not in wb.sheetnames:
            _log_exclude_rules_sheet_debug(
                "SKIP_NO_SHEET",
                log_prefix,
                "繧ｷ繝ｼ繝医′縺ゅｊ縺ｾ縺帙ｓ縲７BA 縺ｮ縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞祇繧ｷ繝ｼ繝医ｒ遒ｺ菫昴阪ｒ螳溯｡後☆繧九°縲∵ｮｵ髫�1/2 繧偵�槭け繝ｭ縺九ｉ襍ｷ蜍輔＠縺ｦ縺上□縺輔＞縲�",
                details=f"path={wb_path}",
            )
            logging.error(
                "%s: 縲�%s縲阪′縺ゅｊ縺ｾ縺帙ｓ縲１ython 縺ｧ縺ｯ繧ｷ繝ｼ繝医ｒ菴懈�舌＠縺ｾ縺帙ｓ縲�",
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
                "1陦檎岼縺ｫ讓呎ｺ冶ｦ句�ｺ縺励ｒ譖ｸ縺崎ｾｼ縺ｿ縺ｾ縺励◆�ｼ育ｩｺ繧ｷ繝ｼ繝医�ｻ蛻怜錐荳堺ｸ閾ｴ縺ｮ陬懈ｭ｣�ｼ峨�",
                details=f"cols=({c_proc},{c_mach},{c_flag},{c_d},{c_e})",
            )

        # 蜑榊屓繝悶ャ繧ｯ菫晏ｭ倥↓螟ｱ謨励＠縺溘→縺埼驕ｿ縺励◆ E 蛻励ｒ縲∝�医↓繝ｯ繝ｼ繧ｯ繧ｷ繝ｼ繝医∈謌ｻ縺呻ｼ育ｶ壹￥菫晏ｭ倥〒繝�繧｣繧ｹ繧ｯ縺ｸ霈峨ｋ�ｼ�
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
                f"蟾･遞�+讖滓｢ｰ縺ｮ陦後ｒ {added} 莉ｶ霑ｽ蜉�縺励∪縺励◆縲�",
            )
            logging.info(
                "%s: 縲�%s縲阪↓蟾･遞�+讖滓｢ｰ縺ｮ邨�縺ｿ蜷医ｏ縺帙ｒ %s 陦瑚ｿｽ蜉�縺励∪縺励◆縲�",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                added,
            )

        # 蜉�蟾･險育判縺九ｉ繝壹い縺�1莉ｶ繧ょ叙繧後★縲√す繝ｼ繝医↓繧ゅョ繝ｼ繧ｿ陦後′辟｡縺�縺ｨ縺阪�ｯ萓玖｡後�ｮ縺ｿ�ｼ亥ｾ捺擂縺ｮ譁ｰ隕上す繝ｼ繝育嶌蠖難ｼ�
        if added == 0 and not existing_keys:
            ws.append(["譴ｱ蛹�", "", "yes", "", ""])
            existing_keys.add(
                (_normalize_process_name_for_rule_match("譴ｱ蛹�"), _normalize_equipment_match_key(""))
            )
            _log_exclude_rules_sheet_debug(
                "EXAMPLE_ROW",
                log_prefix,
                "繝�繝ｼ繧ｿ陦後′辟｡縺九▲縺溘◆繧∽ｾ具ｼ域｢ｱ蛹�=yes�ｼ峨ｒ1陦瑚ｿｽ蜉�縲�",
            )
            logging.info(
                "%s: 縲�%s縲阪↓繝�繝ｼ繧ｿ陦後′辟｡縺九▲縺溘◆繧√∽ｾ具ｼ域｢ｱ蛹�=yes�ｼ峨ｒ1陦瑚ｿｽ蜉�縺励∪縺励◆縲�",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
            )

        # 遨ｺ陦瑚ｩｰ繧√�ｯ AI 繧医ｊ蜈医↓陦後≧�ｼ亥ｾ後°繧芽ｩｰ繧√ｋ縺ｨ縲∵嶌縺崎ｾｼ繧薙□陦檎分蜿ｷ縺ｨ逕ｻ髱｢荳翫�ｮ陦後′縺壹ｌ繧具ｼ�
        n_kept, n_removed_empty = _compact_exclude_rules_data_rows(
            ws, c_proc, c_mach, c_flag, c_d, c_e, log_prefix
        )
        if n_removed_empty:
            _log_exclude_rules_sheet_debug(
                "DATA_COMPACT",
                log_prefix,
                "遨ｺ陦後ｒ蜑企勁縺励※繝�繝ｼ繧ｿ陦後ｒ隧ｰ繧√∪縺励◆�ｼ井ｸｦ縺ｳ鬆�縺ｯ邯ｭ謖��ｼ峨�AI 陬懷ｮ後ｈ繧雁燕縲�",
                details=f"rows={n_kept} removed_empty={n_removed_empty}",
            )

        max_r = int(ws.max_row or 1)
        pending_rows: list[int] = []
        for r in range(2, max_r + 1):
            dv = ws.cell(row=r, column=c_d).value
            ev = ws.cell(row=r, column=c_e).value
            # C 蛻励�ｮ譛臥┌縺ｫ髢｢菫ゅ↑縺上． 縺ｫ隱ｬ譏弱′縺ゅｊ E 縺檎ｩｺ縺ｪ繧� D竊脱 繧定ｩｦ縺�
            if _cell_is_blank_for_rule(dv):
                continue
            if not _cell_is_blank_for_rule(ev):
                continue
            pending_rows.append(r)

        # D 縺梧焚蠑上�ｮ縺ｨ縺阪�ｯ騾壼ｸｸ隱ｭ霎ｼ縺ｧ縺ｯ '=...' 縺�縺大叙繧後ｋ縲Ｅata_only 縺ｧ繧ｭ繝｣繝�繧ｷ繝･陦ｨ遉ｺ蛟､繧定｣懊≧縲�
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
                        "%s: 縲�%s縲�%s 陦檎岼縺ｮ D 蛻励′謨ｰ蠑上〒縲√く繝｣繝�繧ｷ繝･蛟､繧定ｪｭ繧√∪縺帙ｓ縺ｧ縺励◆�ｼ�Excel 縺ｧ荳蠎ｦ菫晏ｭ倥☆繧九° D 繧貞､縺ｫ縺励※縺上□縺輔＞�ｼ峨�",
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
                        "%s: 縲�%s縲�%s 陦檎岼縺ｮ D 蛻励ｒ JSON 縺ｫ縺ｧ縺阪∪縺帙ｓ縺ｧ縺励◆�ｼ�API繧ｭ繝ｼ繝ｻ蠢懃ｭ斐ｒ遒ｺ隱搾ｼ峨�",
                        log_prefix,
                        EXCLUDE_RULES_SHEET_NAME,
                        r,
                    )
                    continue
                jstr = json.dumps(parsed, ensure_ascii=False)
                ws.cell(row=r, column=c_e, value=jstr)
                cell_addr = f"{get_column_letter(c_e)}{r}"
                ai_e_cell_addrs.append(cell_addr)
                preview = jstr if len(jstr) <= 160 else (jstr[:160] + "窶ｦ")
                logging.info(
                    "%s: 縲�%s縲阪Ο繧ｸ繝�繧ｯ蠑丞�励�%s縲阪そ繝ｫ %s 縺ｫ JSON 繧呈嶌縺崎ｾｼ縺ｿ: %s",
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
                f"D竊脱 縺ｮ AI 陬懷ｮ後ｒ {ai_filled} 陦悟ｮ滓命縲�",
                details="cells=" + ",".join(ai_e_cell_addrs),
            )
            logging.info(
                "%s: 縲�%s縲阪〒 D竊脱 縺ｮ AI 陬懷ｮ後ｒ %s 陦鯉ｼ医そ繝ｫ: %s�ｼ峨�",
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
                f'E蛻� {_e_addr} 縺ｫ繝�繧ｹ繝医〒 "1234" 繧呈嶌縺崎ｾｼ縺ｿ',
                details=f"row={_er_row}",
            )
            logging.warning(
                '%s: 縲舌ユ繧ｹ繝医�%s 縺ｫ "1234" 繧呈嶌縺崎ｾｼ縺ｿ�ｼ�EXCLUDE_RULES_TEST_E1234�ｼ峨�',
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
                "%s: 險ｭ螳壹す繝ｼ繝医�ｮ openpyxl 菫晏ｭ倥↓螟ｱ謨励＠縺ｾ縺励◆縲�"
                " log 縺ｮ陦悟�� TSV 繧偵�槭け繝ｭ縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞祇A縺九ｉE_TSV縺九ｉ蜿肴丐縲阪�"
                "縺ｾ縺溘�ｯ E 蛻励�ｮ縺ｿ縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞祇E蛻誉TSV縺九ｉ蜿肴丐縲阪〒蜿肴丐縺励※縺上□縺輔＞縲�",
                log_prefix,
            )
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "FATAL",
            log_prefix,
            "險ｭ螳壹す繝ｼ繝亥�ｦ逅�荳ｭ縺ｫ譛ｪ謐墓拷萓句､悶′逋ｺ逕溘＠縺ｾ縺励◆縲�",
            exc=ex,
        )
        logging.exception("%s: 險ｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九�ｮ蜃ｦ逅�縺ｧ萓句､�", log_prefix)
    finally:
        if wb is not None:
            wb.close()
            _log_exclude_rules_sheet_debug("CLOSED", log_prefix, "繝悶ャ繧ｯ繧偵け繝ｭ繝ｼ繧ｺ縺励∪縺励◆縲�")


def _resolve_exclude_rules_workbook_path_for_read(wb_path: str) -> str:
    """逶ｴ蜑阪�ｮ菫晏ｮ医〒螳溷柑繝代せ縺悟､峨ｏ縺｣縺溘→縺搾ｼ磯壼ｸｸ縺ｯ菫晏ｭ俶�仙粥蠕後�ｮ蜈�繝悶ャ繧ｯ�ｼ峨↓縺昴ｌ繧剃ｽｿ縺�縲�"""
    p = _exclude_rules_effective_read_path
    if p and os.path.exists(p):
        return p
    return wb_path


def _load_exclude_rules_from_workbook(wb_path: str) -> list[dict]:
    """繧ｷ繝ｼ繝医°繧峨Ν繝ｼ繝ｫ陦後ｒ隱ｭ縺ｿ縲∬ｩ穂ｾ｡逕ｨ繝ｪ繧ｹ繝医ｒ霑斐☆縲�"""
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
            "驟榊床荳崎ｦ√Ν繝ｼ繝ｫ: 繝悶ャ繧ｯ縺ｫ縲�%s縲阪′縺ゅｋ縺溘ａ pandas(openpyxl) 縺ｧ縺ｮ縲�%s縲崎ｪｭ霎ｼ繧偵せ繧ｭ繝�繝励＠縺ｾ縺励◆�ｼ医Ν繝ｼ繝ｫ縺ｯ譛ｪ驕ｩ逕ｨ�ｼ峨�",
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
    """險ｭ螳壹す繝ｼ繝医↓蝓ｺ縺･縺阪碁�榊床荳崎ｦ√阪ｒ險ｭ螳夲ｼ�C=yes 縺ｾ縺溘�ｯ E 縺ｮ JSON 縺檎悄�ｼ峨�"""
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
        logging.info("%s: 險ｭ螳壹�%s縲阪↓繧医ｊ驟榊床荳崎ｦ�=yes 繧� %s 陦後↓險ｭ螳壹＠縺ｾ縺励◆縲�", log_prefix, EXCLUDE_RULES_SHEET_NAME, n)
    return df


def _sort_stage1_plan_df_by_dispatch_trial_order_asc(plan_df: "pd.DataFrame") -> "pd.DataFrame":
    """
    谿ｵ髫�1蜃ｺ蜉帷峩蜑�: 驟榊床隧ｦ陦碁��逡ｪ縺ｮ譏�鬆�縺ｫ陦後ｒ荳ｦ縺ｹ譖ｿ縺医◆ DataFrame 繧定ｿ斐☆縲�
    豁｣縺ｮ謨ｴ謨ｰ縺ｧ縺ｪ縺�繧ｻ繝ｫ縺ｯ譛蠕鯉ｼ亥酔蟶ｯ蜀�縺ｯ蜈�縺ｮ陦碁���ｼ峨�
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
# 谿ｵ髫�1繧ｨ繝ｳ繝医Μ�ｼ�task_extract_stage1.py 竊� run_stage1_extract�ｼ�
#   蜉�蟾･險育判DATA 隱ｭ蜿� 竊� 驟榊床荳崎ｦ∬�ｪ蜍募�ｦ逅� 竊� 險ｭ螳壹す繝ｼ繝井ｿ晏ｮ� 竊� plan_input_tasks.xlsx 蜃ｺ蜉�
# =============================================================================
def run_stage1_extract():
    """
    谿ｵ髫�1: 蜉�蟾･險育判DATA 縺九ｉ驟榊床逕ｨ繧ｿ繧ｹ繧ｯ荳隕ｧ繧呈歓蜃ｺ縺� output/plan_input_tasks.xlsx 縺ｸ蜃ｺ蜉帙�
    蜷御ｸ萓晞�ｼNO縺ｧ蜷御ｸ讖滓｢ｰ蜷阪′隍�謨ｰ陦後≠繧九→縺阪∝ｷ･遞句錐縲悟��蜑ｲ縲崎｡後�ｮ遨ｺ縺ｮ縲碁�榊床荳崎ｦ√阪↓ yes 繧定�ｪ蜍戊ｨｭ螳壹☆繧九�
    繝槭け繝ｭ繝悶ャ繧ｯ縺ｮ縲瑚ｨｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九阪〒蟾･遞�+讖滓｢ｰ縺斐→縺ｮ驟榊床荳崎ｦ√�ｻ譚｡莉ｶ蠑擾ｼ�AI�ｼ峨ｒ邂｡逅�縺吶ｋ�ｼ医す繝ｼ繝井ｽ懈�舌�ｯ VBA�ｼ峨�
    """
    if not TASKS_INPUT_WORKBOOK:
        logging.error("TASK_INPUT_WORKBOOK 縺梧悴險ｭ螳壹〒縺吶�")
        return False
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        logging.error(f"TASK_INPUT_WORKBOOK 縺悟ｭ伜惠縺励∪縺帙ｓ: {TASKS_INPUT_WORKBOOK}")
        return False
    reset_gemini_usage_tracker()
    df_src = load_tasks_df()
    try:
        _pm_pairs = _collect_process_machine_pairs_for_exclude_rules(df_src)
        run_exclude_rules_sheet_maintenance(TASKS_INPUT_WORKBOOK, _pm_pairs, "谿ｵ髫�1")
    except Exception:
        logging.exception("谿ｵ髫�1: 險ｭ螳喟驟榊床荳崎ｦ∝ｷ･遞九�ｮ菫晏ｮ医〒萓句､厄ｼ育ｶ夊｡鯉ｼ�")
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
        # 蟾･遞句錐 + 讖滓｢ｰ蜷� 繧停懷屏蟄絶昴→縺励※陦ｨ遉ｺ逕ｨ縺ｫ霑ｽ蜉��ｼ亥ｾ梧ｮｵ縺ｯ險育ｮ励く繝ｼ縺ｫ繧ゆｽｿ逕ｨ�ｼ�
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
        logging.warning("谿ｵ髫�1: 謚ｽ蜃ｺ蟇ｾ雎｡繧ｿ繧ｹ繧ｯ縺後≠繧翫∪縺帙ｓ縲�")
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
        logging.error("谿ｵ髫�1繧剃ｸｭ譁ｭ: 繝槭せ繧ｿ skills 縺ｮ讀懆ｨｼ繧ｨ繝ｩ繝ｼ�ｼ亥━蜈亥ｺｦ縺ｮ謨ｰ蛟､驥崎､�縺ｪ縺ｩ�ｼ峨�")
        raise
    except Exception as e:
        logging.info("谿ｵ髫�1: 繝槭せ繧ｿ need 繧定ｪｭ繧√★蜈�蛻励�ｯ need 縺ｪ縺励〒蝓九ａ縺ｾ縺� (%s)", e)
        req_map, need_rules = {}, []
        equipment_list_stage1 = []
        need_combo_col_index_stage1 = {}
    out_df = _merge_plan_sheet_user_overrides(out_df)
    _refresh_plan_reference_columns(out_df, req_map, need_rules)
    try:
        _apply_auto_exclude_bunkatsu_duplicate_machine(out_df, log_prefix="谿ｵ髫�1")
    except Exception as ex:
        logging.exception("谿ｵ髫�1: 蛻�蜑ｲ陦後�ｮ驟榊床荳崎ｦ∬�ｪ蜍戊ｨｭ螳壹〒萓句､厄ｼ亥�ｺ蜉帙�ｯ邯夊｡鯉ｼ�: %s", ex)
    try:
        out_df = apply_exclude_rules_config_to_plan_df(out_df, TASKS_INPUT_WORKBOOK, "谿ｵ髫�1")
    except Exception as ex:
        logging.warning("谿ｵ髫�1: 險ｭ螳壹す繝ｼ繝医↓繧医ｋ驟榊床荳崎ｦ�驕ｩ逕ｨ縺ｧ萓句､厄ｼ育ｶ夊｡鯉ｼ�: %s", ex)
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
        logging.warning("谿ｵ髫�1: 驟榊床隧ｦ陦碁��逡ｪ蛻励�ｮ險育ｮ励ｒ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆�ｼ育ｶ夊｡鯉ｼ�: %s", ex)
    out_df = _sort_stage1_plan_df_by_dispatch_trial_order_asc(out_df)
    out_path = os.path.join(output_dir, STAGE1_OUTPUT_FILENAME)
    out_df.to_excel(out_path, sheet_name="繧ｿ繧ｹ繧ｯ荳隕ｧ", index=False)
    _apply_excel_date_columns_date_only_display(out_path, "繧ｿ繧ｹ繧ｯ荳隕ｧ")
    _apply_plan_input_visual_format(out_path, "繧ｿ繧ｹ繧ｯ荳隕ｧ")
    logging.info(f"谿ｵ髫�1螳御ｺ�: '{out_path}' 繧貞�ｺ蜉帙＠縺ｾ縺励◆縲ゅ�槭け繝ｭ縺ｧ '{PLAN_INPUT_SHEET_NAME}' 縺ｫ蜿悶ｊ霎ｼ繧薙〒縺上□縺輔＞縲�")
    _try_write_main_sheet_gemini_usage_summary("谿ｵ髫�1")
    return True


# 遞ｼ蜒阪Ν繝ｼ繝ｫ�ｼ医ョ繝輔か繝ｫ繝亥､繝ｻ2026蟷ｴ3譛亥渕貅厄ｼ�
TARGET_YEAR = 2026
TARGET_MONTH = 3
DEFAULT_START_TIME = time(8, 45)
DEFAULT_END_TIME = time(17, 0)
DEFAULT_BREAKS = [
    (time(12, 0), time(12, 50)),
    (time(14, 45), time(15, 0))
]
# 邨よ･ｭ逶ｴ蜑阪ョ繝輔ぃ繝ｼ: ASSIGN_END_OF_DAY_DEFER_MINUTES 縺梧ｭ｣縺ｮ縺ｨ縺阪》eam_end_limit 縺ｾ縺ｧ縺ｮ谿九ｊ縺後◎縺ｮ蛻�謨ｰ莉･荳九〒縲�
# 縺九▽ remaining_units�ｼ亥��繧贋ｸ翫￡�ｼ峨′ ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS 莉･荳九�ｮ縺ｨ縺阪√◎縺ｮ譌･縺ｮ髢句ｧ倶ｸ榊庄�ｼ�None�ｼ峨�
# 蜷後§繧ｦ繧｣繝ｳ繝峨え縺ｧ縲窟SSIGN_EOD_DEFER_MAX_REMAINING_ROLLS 繝ｭ繝ｼ繝ｫ蛻�莉･荳翫�ｯ蝗槭○縺ｪ縺�縲搾ｼ亥庶螳ｹ縺碁明蛟､譛ｪ貅�ｼ峨→縺阪�ｯ
# 譁ｰ隕上↓蜉�蟾･繧貞ｧ九ａ縺ｪ縺��ｼ�_eod_reject_capacity_units_below_threshold�ｼ峨�
# ASSIGN_END_OF_DAY_DEFER_MINUTES 譌｢螳� 45�ｼ亥���ｼ峨�0 繧呈�守､ｺ縺吶ｋ縺ｨ辟｡蜉ｹ�ｼ亥ｾ捺擂縺ｩ縺翫ｊ�ｼ峨�
# ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS 譌｢螳� 5縲ょ香蛻�螟ｧ縺阪↑蛟､�ｼ井ｾ�: 999999�ｼ峨↓縺吶ｋ縺ｨ螳溯ｳｪ縲梧ｮ九Ο繝ｼ繝ｫ縺ｫ萓昴ｉ縺夂ｵよ･ｭ逶ｴ蜑阪�ｯ荳榊庄縲阪�
# 莨第�ｩ: 蟶ｯ蜀�縺ｫ關ｽ縺｡縺滄幕蟋九�ｯ _defer_team_start_past_prebreak_and_end_of_day 縺ｧ莨第�ｩ邨ゆｺ�縺ｸ郢ｰ繧贋ｸ九￡縲�
# 莨第�ｩ繧偵∪縺溘＄騾｣邯夐�榊床縺ｯ _contiguous_work_minutes_until_next_break_or_limit 縺ｧ蜊ｴ荳九�
# �ｼ域立 ASSIGN_DEFER_MIN_REMAINING_ROLLS / ASSIGN_PRE_BREAK_DEFER_GAP_MINUTES 縺ｯ蟒�豁｢繝ｻ辟｡隕厄ｼ�
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
    """ASSIGN_END_OF_DAY_DEFER_MINUTES 縺梧ｭ｣縺ｮ縺ｨ縺阪�髢句ｧ九′邨よ･ｭ荳企剞縺ｮ縺昴�ｮ蛻�謨ｰ莉･蜀�縺九�"""
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
    邨よ･ｭ縺ｾ縺ｧ縺ゅ→ ASSIGN_END_OF_DAY_DEFER_MINUTES 蛻�莉･蜀�縺ｮ繧ｦ繧｣繝ｳ繝峨え蜀�縺ｧ縲�
    ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS 繝ｭ繝ｼ繝ｫ蛻�莉･荳翫�ｯ蝗槭○縺ｪ縺��ｼ亥庶螳ｹ繝ｭ繝ｼ繝ｫ謨ｰ縺碁明蛟､譛ｪ貅�ｼ峨→縺� True�ｼ域眠隕丞刈蟾･繧貞ｧ九ａ縺ｪ縺��ｼ晏呵｣懷唆荳具ｼ峨�
    """
    th = ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS
    if th <= 0:
        return False
    if not _eod_minutes_window_covers_start(team_start, team_end_limit):
        return False
    return int(units_fit_until_close) < int(th)


# =========================================================
# 1. 繧ｳ繧｢險育ｮ励Ο繧ｸ繝�繧ｯ (譌･譎ゅ�吶�ｼ繧ｹ)
#    莨第�ｩ蟶ｯ繧呈検繧薙□縲悟ｮ溷ロ蛻�縲肴鋤邂励�ｻ邨ゆｺ�譎ょ綾縺ｮ郢ｰ繧贋ｸ翫￡縲ょ牡莉倥Ν繝ｼ繝励�ｮ荳句屓繧翫�
# =========================================================
def merge_time_intervals(intervals):
    """譎ょ綾蛹ｺ髢薙�ｮ繝ｪ繧ｹ繝医ｒ繧ｽ繝ｼ繝医＠縲�驥阪↑繧句玄髢薙ｒ邨仙粋縺励※霑斐☆縲�"""
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
    start_dt 縺九ｉ谺｡縺ｮ莨第�ｩ髢句ｧ具ｼ医∪縺溘�ｯ邨よ･ｭ荳企剞�ｼ峨∪縺ｧ縺ｮ縲�騾｣邯壹＠縺ｦ螳溷ロ縺ｫ菴ｿ縺医ｋ蛻�謨ｰ縲�
    髢句ｧ九′莨第�ｩ蟶ｯ蜀�縺ｪ繧� 0�ｼ亥他縺ｳ蜃ｺ縺怜��縺ｧ蜊ｴ荳具ｼ峨Ｃreaks_dt 縺ｯ merge 貂医∩諠ｳ螳壹�
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
    莨第�ｩ蟶ｯ螟悶〒繧ゅ∵ｬ｡縺ｮ莨第�ｩ髢句ｧ九∪縺ｧ縺ｮ騾｣邯壼ｮ溷ロ縺� min_contiguous_mins 譛ｪ貅縺ｪ繧峨�
    縺昴�ｮ莨第�ｩ蛹ｺ髢薙�ｮ邨ゆｺ�譎ょ綾繧定ｿ斐☆�ｼ亥壕蠕御ｼ第�ｩ逶ｴ蜑阪↓ 1 繝ｭ繝ｼ繝ｫ蛻�縺悟庶縺ｾ繧峨↑縺�髢句ｧ九□縺鷹ｲ繧√ｋ�ｼ峨�
    邨よ･ｭ縺ｾ縺ｧ縺励°螳溷ロ縺檎ｶ壹°縺ｪ縺�蝣ｴ蜷医�ｯ None縲�
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
    - ASSIGN_END_OF_DAY_DEFER_MINUTES > 0 縺九▽ (team_end_limit - 隧ｦ陦碁幕蟋�) 縺後◎縺ｮ蛻�謨ｰ莉･荳九〒縲�
      remaining_units 蛻�繧贋ｸ翫￡縺� ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS 莉･荳九�ｮ縺ｨ縺阪∝ｽ捺律髢句ｧ倶ｸ榊庄�ｼ�None�ｼ峨�
    - 隧ｦ陦碁幕蟋九′莨第�ｩ蟶ｯ蜀�縺ｮ縺ｨ縺阪�ｯ **莨第�ｩ邨ゆｺ�譎ょ綾縺ｸ郢ｰ繧贋ｸ九￡**縺励～refloor_fn` 縺ｧ險ｭ蛯吩ｸ矩剞繝ｻavail 繧貞�埼←逕ｨ縺吶ｋ縲�
      郢ｰ繧贋ｸ九￡縺ｮ縺ゅ→邨よ･ｭ雜�驕弱�ｻEOD 繝�繝輔ぃ繝ｼ縺ｫ隧ｲ蠖薙☆繧後�ｰ None縲�
    - min_contiguous_work_mins 縺梧ｭ｣縺ｮ縺ｨ縺阪∝ｸｯ螟悶〒繧� **谺｡縺ｮ莨第�ｩ縺ｾ縺ｧ縺ｮ騾｣邯壼ｮ溷ロ**縺後◎繧梧悴貅縺ｪ繧�
      蠖楢ｩｲ莨第�ｩ縺ｮ邨ゆｺ�縺ｸ郢ｰ繧贋ｸ九￡�ｼ井ｸ翫→蜷梧ｧ倥↓ refloor 縺励Ν繝ｼ繝暦ｼ峨�
    """
    _tid = str(task.get("task_id", "") or "").strip()
    _team_txt = ", ".join(str(x) for x in team) if team else "窶�"

    def _trace_block(msg: str, *a) -> None:
        if not _trace_schedule_task_enabled(_tid):
            return
        _log_dispatch_trace_schedule(
            _tid,
            "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繝悶Ο繝�繧ｯ蛻､螳�: " + msg,
            _tid,
            *a,
        )

    ts = refloor_fn(team_start)
    for _ in range(64):
        if ts >= team_end_limit:
            _trace_block(
                "髢句ｧ倶ｸ榊庄(邨よ･ｭ雜�驕�) machine=%s team=%s rem=%.4f trial_start=%s end_limit=%s",
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
                "莨第�ｩ蟶ｯ蜀�縺ｮ縺溘ａ邨ゆｺ�縺ｸ郢ｰ繧贋ｸ九￡ machine=%s team=%s rem=%.4f break_end=%s trial_was=%s",
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
                    "莨第�ｩ逶ｴ蜑阪〒騾｣邯壼ｮ溷ロ荳崎ｶｳ縺ｮ縺溘ａ莨第�ｩ邨ゆｺ�縺ｸ郢ｰ繧贋ｸ九￡ machine=%s team=%s rem=%.4f need_contig_min=%s trial_was=%s break_end=%s",
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
                "髢句ｧ倶ｸ榊庄(邨よ･ｭ逶ｴ蜑阪�ｻ蟆乗ｮ九Ο繝ｼ繝ｫ) machine=%s team=%s rem_ceil=%s max_rem=%s trial_start=%s end_limit=%s gap_end_min=%s",
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
        "髢句ｧ倶ｸ榊庄(莨第�ｩ郢ｰ繧贋ｸ九￡謇灘��繧�) machine=%s team=%s rem=%.4f trial_start=%s",
        task.get("machine"),
        _team_txt,
        float(task.get("remaining_units") or 0),
        ts,
    )
    return None


def _expand_timeline_events_for_equipment_grid(timeline_events: list) -> list:
    """
    險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ繝ｻ繝｡繝ｳ繝舌�ｼ譌･遞九�ｻ遞ｼ蜒咲紫逕ｨ繧､繝ｳ繝�繝�繧ｯ繧ｹ蜷代￠縲�
    1 譛ｬ縺ｮ繧､繝吶Φ繝医′譌･繧偵∪縺溘＄蝣ｴ蜷医‘["date"] 縺�縺大ｽ捺律縺ｫ霈峨○繧九→鄙梧悃繧ｻ繧ｰ繝｡繝ｳ繝医′谺�縺代ｋ縺溘ａ縲�
    start_dt縲彳nd_dt 繧貞推蟆ｱ讌ｭ譌･ DEFAULT_START_TIME縲廛EFAULT_END_TIME 縺ｫ繧ｯ繝ｪ繝�繝励＠縺溯､�陬ｽ縺ｸ螻暮幕縺吶ｋ縲�
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
    start_dt 縺九ｉ end_dt 縺ｾ縺ｧ縺ｮ縲御ｼ第�ｩ繧帝勁縺�縺溷ｮ溷ロ蛻�謨ｰ縲阪�
    breaks_dt 窶ｦ (蛹ｺ髢馴幕蟋�, 蛹ｺ髢鍋ｵゆｺ�) 縺ｮ蛻暦ｼ�datetime 縺ｾ縺溘�ｯ time縲ょ他縺ｳ蜃ｺ縺怜��縺ｮ蜍､諤�繧､繝吶Φ繝医→謨ｴ蜷茨ｼ峨�
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
    start_dt 縺九ｉ螳溷ロ duration_minutes 蛻�騾ｲ繧√◆邨ゆｺ� datetime 繧呈ｱゅａ繧具ｼ井ｼ第�ｩ縺ｯ繧ｹ繧ｭ繝�繝暦ｼ峨�
    end_limit_dt 繧定ｶ�縺医↑縺�繧医≧謇薙■蛻�繧翫よ綾繧雁､: (邨ゆｺ�譎ょ綾, 螳滄圀縺ｫ騾ｲ繧√◆螳溷ロ蛻�, 谿九ｊ譛ｪ豸亥喧蛻�)
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
    need 繧ｷ繝ｼ繝医御ｾ晞�ｼNO譚｡莉ｶ縲肴ｬ�縺ｮ隗｣驥医�
    遨ｺ繝ｻ*繝ｻ蜈ｨ莉ｶ 竊� 蟶ｸ縺ｫ繝槭ャ繝√�
    prefix:ABC / 謗･鬆ｭ霎�:ABC 竊� 萓晞�ｼNO 縺後◎縺ｮ譁�蟄怜�励〒蟋九∪繧�
    regex:... / 豁｣隕剰｡ｨ迴ｾ:... 竊� 豁｣隕剰｡ｨ迴ｾ�ｼ磯Κ蛻�荳閾ｴ�ｼ�
    縺昴ｌ莉･螟悶�ｮ遏ｭ譁�縺ｯ謗･鬆ｭ霎槭→縺励※謇ｱ縺�縲ょｾ捺擂縺ｮ譌･譛ｬ隱樔ｾ九御ｾ晞�ｼNO縺繰R縺ｧ窶ｦ縲阪�ｯ JR 繧呈､懷�ｺ縺励◆繧画磁鬆ｭ霎曷R謇ｱ縺�縲�
    """
    cond = (condition_raw or "").strip()
    tid = str(task_id).strip()
    if not cond or cond in ("*", "蜈ｨ莉ｶ", "蜈ｨ縺ｦ", "any", "ANY"):
        return True
    low = cond.lower()
    cn = cond.replace("�ｼ�", ":")
    if low.startswith("prefix:") or low.startswith("謗･鬆ｭ霎�:"):
        pref = cn.split(":", 1)[1].strip() if ":" in cn else ""
        return bool(pref) and tid.startswith(pref)
    if low.startswith("regex:") or low.startswith("豁｣隕剰｡ｨ迴ｾ:"):
        pat = cn.split(":", 1)[1].strip() if ":" in cn else ""
        if not pat:
            return False
        try:
            return re.search(pat, tid) is not None
        except re.error:
            logging.warning(f"need 萓晞�ｼNO譚｡莉ｶ縺ｮ豁｣隕剰｡ｨ迴ｾ縺檎┌蜉ｹ縺ｧ縺�: {pat}")
            return False
    if "萓晞�ｼ" in cond and "JR" in cond.upper():
        return tid.upper().startswith("JR")
    return tid.startswith(cond)


def parse_need_sheet_special_rules(needs_df, label_col, equipment_list, cond_col):
    """迚ｹ蛻･謖�螳�1�ｽ�99 陦後°繧峨∬ｨｭ蛯吝挨縺ｮ蠢�隕∽ｺｺ謨ｰ荳頑嶌縺搾ｼ�1�ｽ�99�ｼ峨ｒ謚ｽ蜃ｺ�ｼ亥�医↓螳夂ｾｩ縺輔ｌ縺溽分蜿ｷ縺悟━蜈茨ｼ峨�"""
    rules = []
    for _, row in needs_df.iterrows():
        lab = str(row.get(label_col, "") or "").strip()
        m = re.match(r"迚ｹ蛻･謖�螳喀s*(\d+)", lab)
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
    need 繧ｷ繝ｼ繝医�ｮ縲悟ｷ･遞句錐 + 讖滓｢ｰ蜷阪阪〒蠢�隕＾P莠ｺ謨ｰ繧定ｧ｣豎ｺ�ｼ育音蛻･謖�螳�1縲�99縺ｯ order 縺悟ｰ上＆縺�縺ｻ縺ｩ蜆ｪ蜈茨ｼ峨�

    req_map 縺ｯ
      - f\"{process}+{machine_name}\"�ｼ亥宍蟇�繧ｭ繝ｼ�ｼ�
      - machine_name�ｼ域ｩ滓｢ｰ縺�縺代�ｮ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ�ｼ�
      - process�ｼ亥ｷ･遞九□縺代�ｮ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ�ｼ�
    縺ｮ縺�縺壹ｌ縺九〒 base 繧貞ｼ輔￠繧句燕謠舌�
    need_rules 縺ｮ overrides 繧ょ酔讒倥↓繧ｭ繝ｼ繧呈戟縺､縲�
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
    resolve_need_required_op 縺ｨ蜷悟､繧定ｿ斐＠縺､縺､縲√Ο繧ｰ逕ｨ縺ｫ蜿ら�ｧ蜈�縺ｮ隱ｬ譏取枚蟄怜�励ｒ莉倥￠繧九�
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
        base_src = f"req_map[讖滓｢ｰ蜷阪�ｮ縺ｿ {m!r}]={base}�ｼ郁､�蜷医く繝ｼ荳榊惠�ｼ�"
    elif p and p in req_map:
        base = req_map[p]
        base_src = f"req_map[蟾･遞句錐縺ｮ縺ｿ {p!r}]={base}�ｼ郁､�蜷医�ｻ讖滓｢ｰ繧ｭ繝ｼ荳榊惠�ｼ�"
    else:
        base = 1
        base_src = "req_map隧ｲ蠖薙↑縺冷�呈里螳�1"
    for rule in need_rules:
        if not match_need_sheet_condition(rule["condition"], task_id):
            continue
        order = rule.get("order", "?")
        if combo_key and combo_key in rule["overrides"]:
            v = int(rule["overrides"][combo_key])
            return v, f"need迚ｹ蛻･謖�螳嘴order} [{combo_key!r}]={v}"
        if m and m in rule["overrides"]:
            v = int(rule["overrides"][m])
            return v, f"need迚ｹ蛻･謖�螳嘴order} [讖滓｢ｰ蜷砿m!r}]={v}"
        if p and p in rule["overrides"]:
            v = int(rule["overrides"][p])
            return v, f"need迚ｹ蛻･謖�螳嘴order} [蟾･遞句錐{p!r}]={v}"
    return int(base), base_src


def _need_row_label_hints_surplus_add(label_a0: str) -> bool:
    """need 繧ｷ繝ｼ繝� A蛻�: 蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ縺ｮ逶ｴ荳九↓縺ゅｋ縲碁�榊床邨先棡縺ｧ菴吝臆縺悟�ｺ縺溘→縺阪�ｮ霑ｽ蜉�蠅怜藤荳企剞縲崎｡後°縲�"""
    s = unicodedata.normalize("NFKC", str(label_a0 or "").strip())
    if not s or s.startswith("迚ｹ蛻･謖�螳�"):
        return False
    if "萓晞�ｼ" in s and "譚｡莉ｶ" in s:
        return False
    if "霑ｽ蜉�" in s and ("莠ｺ謨ｰ" in s or "莠ｺ蜩｡" in s or "蠅怜藤" in s):
        return True
    if "蠅怜藤" in s or "菴吝臆" in s:
        return True
    if "驟榊床" in s and ("霑ｽ蜉�" in s or "蠅�" in s or "菴吝臆" in s):
        return True
    return False


def _find_need_surplus_add_row_index(
    needs_raw, base_row: int, col0: int, pm_cols: list
) -> int | None:
    """蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ陦後�ｮ谺｡陦後ｒ蜆ｪ蜈医ゅΛ繝吶Ν縺ｾ縺溘�ｯ謨ｰ蛟､縺ｧ霑ｽ蜉�莠ｺ謨ｰ陦後→蛻､螳壹�"""
    r = base_row + 1
    if r >= needs_raw.shape[0]:
        return None
    v0 = needs_raw.iat[r, col0]
    s0 = "" if pd.isna(v0) else str(v0).strip()
    if s0.startswith("迚ｹ蛻･謖�螳�"):
        return None
    if _need_row_label_hints_surplus_add(s0):
        return r
    nz = 0
    for col_idx, _, _ in pm_cols:
        if parse_optional_int(needs_raw.iat[r, col_idx]) is not None:
            nz += 1
    if nz > 0 and not unicodedata.normalize("NFKC", s0).startswith("迚ｹ蛻･"):
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
    need 繧ｷ繝ｼ繝医碁�榊床譎りｿｽ蜉�莠ｺ謨ｰ縲崎｡鯉ｼ亥ｷ･遞凝玲ｩ滓｢ｰ蛻暦ｼ峨�ｮ蛟､�ｼ晏ｿ�隕∽ｺｺ謨ｰ繧呈ｺ縺溘＠縺溘≧縺医〒
    縺輔ｉ縺ｫ蜑ｲ繧雁ｽ薙※蜿ｯ閭ｽ縺ｪ莠ｺ謨ｰ縺ｮ荳企剞�ｼ�0 縺ｪ繧牙ｾ捺擂縺ｩ縺翫ｊ蠢�隕∽ｺｺ謨ｰ縺｡繧�縺�縺ｩ縺ｮ縺ｿ�ｼ峨�
    need_rules 縺ｯ迴ｾ迥ｶ縺薙�ｮ陦後ｒ荳頑嶌縺阪＠縺ｪ縺��ｼ亥ｰ�譚･諡｡蠑ｵ逕ｨ縺ｫ task_id 繧貞女縺大叙繧具ｼ峨�
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
    """resolve_need_surplus_extra_max 縺ｨ蜷悟､�ｼ句盾辣ｧ蜈�隱ｬ譏趣ｼ医Ο繧ｰ逕ｨ�ｼ峨�"""
    val = resolve_need_surplus_extra_max(
        process, machine_name, task_id, surplus_map, need_rules
    )
    _ = need_rules
    if not surplus_map:
        return val, "surplus_map遨ｺ�ｼ磯�榊床譎りｿｽ蜉�莠ｺ謨ｰ陦後↑縺暦ｼ�"
    p = str(process).strip()
    m = str(machine_name).strip()
    combo_key = f"{p}+{m}" if p and m else None
    if combo_key and combo_key in surplus_map:
        raw = surplus_map[combo_key]
        return val, f"surplus_map[{combo_key!r}]={raw}"
    if m and m in surplus_map:
        raw = surplus_map[m]
        return val, f"surplus_map[讖滓｢ｰ蜷阪�ｮ縺ｿ {m!r}]={raw}�ｼ郁､�蜷医く繝ｼ荳榊惠�ｼ�"
    if p and p in surplus_map:
        raw = surplus_map[p]
        return val, f"surplus_map[蟾･遞句錐縺ｮ縺ｿ {p!r}]={raw}�ｼ郁､�蜷医く繝ｼ荳榊惠�ｼ�"
    return val, "surplus蠖薙く繝ｼ縺ｪ縺冷��0"


def _surplus_team_time_factor(
    rq_base: int, team_len: int, extra_max_allowed: int
) -> float:
    """
    蠢�隕∽ｺｺ謨ｰ繧定ｶ�縺医※蜈･繧後◆繝｡繝ｳ繝舌�ｼ縺ｫ繧医ｋ蜊倅ｽ肴凾髢薙∈縺ｮ菫よ焚�ｼ�1.0�ｼ晉洒邵ｮ縺ｪ縺暦ｼ峨�
    霑ｽ蜉�譫��ｼ�extra_max_allowed�ｼ峨ｒ菴ｿ縺�蛻�縺｣縺溘→縺阪〒繧ゅ∫洒邵ｮ縺ｯ SURPLUS_TEAM_MAX_SPEEDUP_RATIO 繧剃ｸ企剞縺ｨ縺吶ｋ邱壼ｽ｢繝｢繝�繝ｫ縲�
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
        return "(-莠ｺ謨ｰ, 髢句ｧ�, -蜊倅ｽ肴焚, 蜆ｪ蜈亥ｺｦ蜷郁ｨ�)"
    if TEAM_ASSIGN_START_SLACK_WAIT_MINUTES <= 0:
        return "(髢句ｧ�, -蜊倅ｽ肴焚, 蜆ｪ蜈亥ｺｦ蜷郁ｨ�)"
    return (
        f"譛譌ｩ髢句ｧ九°繧閲TEAM_ASSIGN_START_SLACK_WAIT_MINUTES}蛻�莉･蜀�縺ｯ"
        "(0,-莠ｺ謨ｰ,髢句ｧ�,-蜊倅ｽ肴焚,蜆ｪ蜈亥ｺｦ)縲∬ｶ�驕弱�ｯ(1,髢句ｧ�,-莠ｺ謨ｰ,-蜊倅ｽ肴焚,蜆ｪ蜈亥ｺｦ)"
    )


def _team_assignment_sort_tuple(
    team: tuple,
    team_start: datetime,
    units_today: int,
    team_prio_sum: int,
    t_min: datetime | None = None,
) -> tuple:
    """
    繝√�ｼ繝�蛟呵｣懊�ｮ蜆ｪ蜉｣逕ｨ繧ｿ繝励Ν�ｼ郁ｾ樊嶌蠑上〒蟆上＆縺�譁ｹ縺梧治逕ｨ�ｼ峨�
    - TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF: (-莠ｺ謨ｰ, 髢句ｧ�, -蜊倅ｽ肴焚, 蜆ｪ蜈亥ｺｦ蜷郁ｨ�)
    - 縺昴ｌ莉･螟悶°縺､ TEAM_ASSIGN_START_SLACK_WAIT_MINUTES>0 縺九▽ t_min 縺ゅｊ:
        譛譌ｩ髢句ｧ九°繧峨せ繝ｩ繝�繧ｯ莉･蜀� 竊� (0, -莠ｺ謨ｰ, 髢句ｧ�, -蜊倅ｽ肴焚, 蜆ｪ蜈亥ｺｦ) 窶ｦ 驕�繧後※繧ゆｺｺ謨ｰ繧貞字縺�
        繧ｹ繝ｩ繝�繧ｯ雜� 竊� (1, 髢句ｧ�, -莠ｺ謨ｰ, -蜊倅ｽ肴焚, 蜆ｪ蜈亥ｺｦ) 窶ｦ 髢句ｧ九ｒ蜆ｪ蜈�
    - 荳願ｨ倅ｻ･螟�: (髢句ｧ�, -蜊倅ｽ肴焚, 蜆ｪ蜈亥ｺｦ蜷郁ｨ�)
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


# skills 繧ｻ繝ｫ: OP / AS + 莉ｻ諢上�ｮ蜆ｪ蜈亥ｺｦ謨ｴ謨ｰ�ｼ井ｾ� OP1, AS 3�ｼ峨よ焚蛟､縺悟ｰ上＆縺�縺ｻ縺ｩ蜑ｲ蠖薙〒蜈医↓驕ｸ縺ｰ繧後ｋ縲�
_SKILL_OP_AS_CELL_RE = re.compile(r"^(OP|AS)(\d*)$", re.IGNORECASE)


def parse_op_as_skill_cell(cell_val):
    """
    master.xlsm縲茎kills縲阪�ｮ繧ｻ繝ｫ1縺､繧定ｧ｣驥医☆繧九�
    - 縲薫P縲阪∪縺溘�ｯ縲窟S縲阪�ｮ逶ｴ蠕後↓蜆ｪ蜈亥ｺｦ逕ｨ縺ｮ謨ｴ謨ｰ�ｼ育ｩｺ逋ｽ縺ｯ髯､蜴ｻ縺励※隗｣驥茨ｼ峨ゆｾ�: OP, OP1, AS3, AS 12
    - 蜆ｪ蜈亥ｺｦ縺ｯ蟆上＆縺�縺ｻ縺ｩ鬮伜━蜈茨ｼ亥酔荳譚｡莉ｶ縺ｮ繝√�ｼ繝�蛟呵｣懊°繧牙�医↓驕ｸ縺ｰ繧後ｋ�ｼ峨よ焚蟄礼怐逡･譎ゅ�ｯ 1縲�
    - OP/AS 縺ｧ蟋九∪繧峨↑縺�繝ｻ遨ｺ縺ｯ繧ｹ繧ｭ繝ｫ縺ｪ縺励�
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
    master縲茎kills縲阪�ｮ蜷�蛻暦ｼ亥ｷ･遞�+讖滓｢ｰ繧ｭ繝ｼ遲会ｼ峨↓縺､縺�縺ｦ縲＾P/AS 縺ｮ蜑ｲ蠖灘━蜈亥ｺｦ縺ｮ**謨ｰ蛟､**縺�
    繝｡繝ｳ繝舌�ｼ髢薙〒驥崎､�縺励※縺�縺ｪ縺�縺区､懆ｨｼ縺吶ｋ縲る㍾隍�譎ゅ�ｯ PlanningValidationError縲�
    �ｼ�OP1 縺ｨ AS1 縺ｮ繧医≧縺ｫ繝ｭ繝ｼ繝ｫ縺檎焚縺ｪ縺｣縺ｦ繧ょ酔荳謨ｰ蛟､縺ｪ繧蛾㍾隍�縺ｨ縺ｿ縺ｪ縺呻ｼ�
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
                errors.append(f'蛻励鶏ck}縲�: 蜆ｪ蜈亥ｺｦ {pr} 縺碁㍾隍� 竊� ' + "縲�".join(entries))
    if errors:
        cap = 50
        tail = errors[:cap]
        msg = (
            "繝槭せ繧ｿ縲茎kills縲阪〒縲∝酔荳蛻励�ｮ OP/AS 蜆ｪ蜈亥ｺｦ縺ｮ謨ｰ蛟､縺碁㍾隍�縺励※縺�縺ｾ縺吶�"
            " 蛻励＃縺ｨ縺ｫ謨ｰ蛟､縺ｯ1莠ｺ縺ｫ縺､縺�1遞ｮ鬘槭↓縺励※縺上□縺輔＞縲�\n"
            + "\n".join(tail)
        )
        if len(errors) > cap:
            msg += f"\n窶ｦ莉� {len(errors) - cap} 莉ｶ"
        raise PlanningValidationError(msg)


def build_member_assignment_priority_reference(
    skills_dict: dict,
    members: list | None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    邨先棡繝悶ャ繧ｯ逕ｨ: 繝槭せ繧ｿ skills 縺ｮ縲悟ｷ･遞句錐+讖滓｢ｰ蜷阪榊�励＃縺ｨ縺ｫ縲∝牡蠖薙い繝ｫ繧ｴ繝ｪ繧ｺ繝�縺ｨ蜷後§
    (蜆ｪ蜈亥ｺｦ蛟､譏�鬆�, 繝｡繝ｳ繝舌�ｼ蜷肴��鬆�) 縺ｧ荳ｦ縺ｹ縺溷盾閠�陦ｨ縺ｨ縲√Ν繝ｼ繝ｫ隱ｬ譏弱�ｮ陦ｨ繧定ｿ斐☆縲�
    蠖捺律縺ｮ蜃ｺ蜍､繝ｻ險ｭ蛯咏ｩｺ縺阪�ｻ蜷御ｸ萓晞�ｼ縺ｮ蟾･遞矩��繝ｻ繝√�ｼ繝�莠ｺ謨ｰ縺ｯ蜿肴丐縺励↑縺��ｼ医≠縺上∪縺ｧ繝槭せ繧ｿ荳翫�ｮ鬆�蠎擾ｼ峨�
    """
    mem_list = list(members) if members else list((skills_dict or {}).keys())
    mem_list = [str(m).strip() for m in mem_list if m and str(m).strip()]

    surplus_on = bool(TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF)
    slack_m = TEAM_ASSIGN_START_SLACK_WAIT_MINUTES
    if surplus_on:
        team_rule = (
            "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF=譛牙柑: "
            "(-莠ｺ謨ｰ, 髢句ｧ�, -蜊倅ｽ肴焚, 蜆ｪ蜈亥ｺｦ蜷郁ｨ�) 縺ｮ霎樊嶌蠑擾ｼ井ｺｺ謨ｰ譛蜆ｪ蜈医�ｻ蠕捺擂�ｼ峨�"
        )
    elif slack_m > 0:
        team_rule = (
            f"譌｢螳�: 縺昴�ｮ譌･縺ｮ謌千ｫ句呵｣懷�ｨ菴薙�ｮ縲梧怙譌ｩ髢句ｧ九阪ｒ蝓ｺ貅悶↓縲�"
            f"髢句ｧ九′縺昴�ｮ{slack_m}蛻�莉･蜀�縺ｮ驕�繧後↑繧我ｺｺ謨ｰ繧貞字縺丞━蜈茨ｼ�0,-莠ｺ謨ｰ,髢句ｧ�,-蜊倅ｽ肴焚,蜆ｪ蜈亥ｺｦ�ｼ峨�"
            f"縺昴ｌ繧医ｊ驕�縺�蛟呵｣懊�ｯ髢句ｧ九ｒ蜆ｪ蜈茨ｼ�1,髢句ｧ�,-莠ｺ謨ｰ,-蜊倅ｽ肴焚,蜆ｪ蜈亥ｺｦ�ｼ峨�"
            f"迺ｰ蠅�螟画焚 TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=0 縺ｧ辟｡蜉ｹ蛹悶�"
        )
    else:
        team_rule = (
            "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=0: "
            "(髢句ｧ�, -蜊倅ｽ肴焚, 蜆ｪ蜈亥ｺｦ蜷郁ｨ�) 縺ｮ縺ｿ�ｼ磯幕蟋区怙蜆ｪ蜈茨ｼ峨�"
        )

    legend_rows = [
        {
            "蛹ｺ蛻�": "繧ｹ繧ｭ繝ｫ蛻励�ｮ荳ｦ縺ｳ",
            "蜀�螳ｹ": "蜷�縲悟ｷ･遞句錐+讖滓｢ｰ蜷阪榊�励↓縺､縺�縺ｦ縲√そ繝ｫ縺� OP/AS�ｼ�+蜆ｪ蜈亥ｺｦ謨ｴ謨ｰ�ｼ峨�ｮ繝｡繝ｳ繝舌�ｼ縺ｮ縺ｿ蟇ｾ雎｡縲�"
            " 謨ｰ蛟､縺悟ｰ上＆縺�縺ｻ縺ｩ鬮伜━蜈医ら怐逡･譎ゅ�ｯ蜆ｪ蜈亥ｺｦ 1�ｼ�parse_op_as_skill_cell 縺ｨ蜷御ｸ�ｼ峨�"
            " 蜷御ｸ蛻励〒縺ｯ蜆ｪ蜈亥ｺｦ縺ｮ謨ｰ蛟､縺ｯ繝｡繝ｳ繝舌�ｼ髢薙〒驥崎､�荳榊庄�ｼ医�槭せ繧ｿ隱ｭ霎ｼ譎ゅ↓讀懆ｨｼ�ｼ峨�",
        },
        {
            "蛹ｺ蛻�": "蠖捺律縺ｨ縺ｮ蟾ｮ",
            "蜀�螳ｹ": "螳滄圀縺ｮ驟榊床縺ｯ縲√％縺ｮ鬆�縺ｮ縺�縺｡縺昴�ｮ譌･蜃ｺ蜍､縺九▽ AS/OP 隕∽ｻｶ繧呈ｺ縺溘☆閠�縺�縺代′蛟呵｣懊�"
            " 險ｭ蛯吶�ｮ遨ｺ縺阪�ｻ蜷御ｸ萓晞�ｼNO縺ｮ蟾･遞矩��繝ｻ蠢�隕∽ｺｺ謨ｰ繝ｻ蠅怜藤譫�繝ｻ謖�蜷弘P縺ｧ螟峨ｏ繧翫∪縺吶�",
        },
        {
            "蛹ｺ蛻�": "繝√�ｼ繝�蛟呵｣懊�ｮ豈碑ｼ�",
            "蜀�螳ｹ": team_rule,
        },
        {
            "蛹ｺ蛻�": "謖�蜷阪�ｻ繧ｰ繝ｭ繝ｼ繝舌Ν荳頑嶌縺�",
            "蜀�螳ｹ": "諡�蠖徹P_謖�螳壹�ｻ繝｡繧､繝ｳ縲悟�榊━蜈育音蛻･險倩ｼ峨阪�ｮ OP 謖�蜷阪�ｯ譛ｬ陦ｨ繧医ｊ蜆ｪ蜈医＆繧後∪縺吶�",
        },
        {
            "蛹ｺ蛻�": "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF",
            "蜀�螳ｹ": "1/譛牙柑�ｼ井ｺｺ謨ｰ譛蜆ｪ蜈医�ｻ蠕捺擂�ｼ�" if surplus_on else "0/辟｡蜉ｹ�ｼ域里螳夲ｼ�",
        },
        {
            "蛹ｺ蛻�": "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES",
            "蜀�螳ｹ": str(slack_m),
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
                    "蟾･遞句錐": proc,
                    "讖滓｢ｰ蜷�": mach,
                    "繧ｹ繧ｭ繝ｫ蛻励く繝ｼ": combo,
                    "蜆ｪ蜈磯��菴�": "",
                    "繝｡繝ｳ繝舌�ｼ": "�ｼ医↑縺暦ｼ�",
                    "繝ｭ繝ｼ繝ｫ": "",
                    "蜆ｪ蜈亥ｺｦ蛟､_蟆上＆縺�縺ｻ縺ｩ蜈�": "",
                    "skills繧ｻ繝ｫ蛟､": "",
                    "蛯呵�": "縺薙�ｮ蛻励↓ OP/AS 縺ｮ雉�譬ｼ繧ｻ繝ｫ縺後≠繧九Γ繝ｳ繝舌�ｼ縺後＞縺ｾ縺帙ｓ",
                }
            )
            continue
        for i, (pr, m, role, cell_s) in enumerate(ranked, start=1):
            out.append(
                {
                    "蟾･遞句錐": proc,
                    "讖滓｢ｰ蜷�": mach,
                    "繧ｹ繧ｭ繝ｫ蛻励く繝ｼ": combo,
                    "蜆ｪ蜈磯��菴�": i,
                    "繝｡繝ｳ繝舌�ｼ": m,
                    "繝ｭ繝ｼ繝ｫ": role,
                    "蜆ｪ蜈亥ｺｦ蛟､_蟆上＆縺�縺ｻ縺ｩ蜈�": pr,
                    "skills繧ｻ繝ｫ蛟､": cell_s,
                    "蛯呵�": "",
                }
            )

    df_tbl = pd.DataFrame(out)
    return df_legend, df_tbl


def _normalize_person_name_for_match(s):
    """諡�蠖楢�謖�蜷阪�ｮ縺ゅ＞縺ｾ縺�荳閾ｴ逕ｨ�ｼ�NFKC繝ｻ蟇檎伐/蜀ｨ逕ｰ縺ｮ陦ｨ險伜ｯ�縺帙�ｻ遨ｺ逋ｽ髯､蜴ｻ繝ｻ譛ｫ蟆ｾ謨ｬ遘ｰ縺ｮ縺ｿ髯､蜴ｻ�ｼ峨�"""
    if s is None:
        return ""
    t = unicodedata.normalize("NFKC", str(s).strip())
    if "蟇檎伐" in t:
        t = t.replace("蟇檎伐", "蜀ｨ逕ｰ")
    t = re.sub(r"[\s縲]+", "", t)
    t = re.sub(r"(縺輔ｓ|讒�|豌�)$", "", t)
    return t


def _split_person_sei_mei(s) -> tuple[str, str]:
    """
    豌丞錐繧貞ｧ薙�ｻ蜷阪↓蛻�縺代ｋ縲よ怙蛻昴�ｮ蜊願ｧ抵ｼ丞�ｨ隗堤ｩｺ逋ｽ縺ｮ謇句燕繧貞ｧ薙∽ｻ･髯阪ｒ蜷阪→縺吶ｋ縲�
    遨ｺ逋ｽ縺檎┌縺�蝣ｴ蜷医�ｯ (蜈ｨ菴�, '')�ｼ亥錐縺ｪ縺玲桶縺��ｼ峨�
    譛ｫ蟆ｾ縺ｮ 縺輔ｓ�ｼ乗ｧ假ｼ乗ｰ� 縺ｯ蛻�蜑ｲ蜑阪↓髯､蜴ｻ縺吶ｋ縲�
    """
    if s is None:
        return "", ""
    t = unicodedata.normalize("NFKC", str(s).strip())
    if not t or t.lower() in ("nan", "none", "null"):
        return "", ""
    t = re.sub(r"(縺輔ｓ|讒�|豌�)$", "", t)
    for i, ch in enumerate(t):
        if ch in " \u3000":
            sei = t[:i].strip()
            rest = t[i + 1 :]
            mei = re.sub(r"[\s縲]+", "", rest.strip())
            return sei, mei
    return t.strip(), ""


def _normalize_sei_for_match(sei: str) -> str:
    """蟋薙�ｮ縺ｿ豁｣隕丞喧縲り｡ｨ險倥ｆ繧後�ｯ險ｱ螳ｹ縺励↑縺�蜑肴署縺ｧ縲¨FKC繝ｻ蟇檎伐/蜀ｨ逕ｰ蟇�縺帙�ｻ遨ｺ逋ｽ髯､蜴ｻ縲�"""
    if not sei:
        return ""
    t = unicodedata.normalize("NFKC", str(sei).strip())
    if "蟇檎伐" in t:
        t = t.replace("蟇檎伐", "蜀ｨ逕ｰ")
    t = re.sub(r"[\s縲]+", "", t)
    return t


def _normalize_mei_for_match(mei: str) -> str:
    """蜷阪�ｮ豁｣隕丞喧�ｼ医ｆ繧瑚ｨｱ螳ｹ縺ｮ蜑榊�ｦ逅��ｼ峨�NFKC繝ｻ遨ｺ逋ｽ髯､蜴ｻ縲ょｧ鍋畑縺ｮ蟇檎伐鄂ｮ謠帙�ｯ陦後ｏ縺ｪ縺�縲�"""
    if not mei:
        return ""
    t = unicodedata.normalize("NFKC", str(mei).strip())
    t = re.sub(r"[\s縲]+", "", t)
    return t


def _has_duplicate_surname_among_members(member_names) -> bool:
    """skills 繝｡繝ｳ繝舌�ｼ荳隕ｧ縺ｫ縲∵ｭ｣隕丞喧蠕悟酔荳縺ｮ蟋薙′2莠ｺ莉･荳翫＞繧九°縲�"""
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
    """蜷御ｸ蟋薙′繝ｭ繧ｹ繧ｿ繝ｼ縺ｧ驥崎､�縺励↑縺�縺ｨ縺阪�ｮ縺ｿ菴ｿ縺�蜷阪�ｮ繧�繧瑚ｨｱ螳ｹ縲�"""
    if not r_mei_n and not m_mei_n:
        return True
    if not r_mei_n or not m_mei_n:
        return False
    if r_mei_n == m_mei_n:
        return True
    return r_mei_n in m_mei_n or m_mei_n in r_mei_n


def _resolve_preferred_name_to_capable_member(raw, capable_candidates, roster_member_names=None):
    """
    閾ｪ逕ｱ險倩ｿｰ縺ｮ謖�蜷阪ｒ縲∝ｽ捺律繧ｹ繧ｭ繝ｫ荳� OP/AS 縺ｮ繝｡繝ｳ繝舌�ｼ蜷搾ｼ�skills 繧ｷ繝ｼ繝医�ｮ陦後く繝ｼ�ｼ峨↓隗｣豎ｺ縺吶ｋ縲�
    capable_candidates: 縺昴�ｮ險ｭ蛯吶〒 OP 縺ｾ縺溘�ｯ AS 縺ｨ縺励※蜑ｲ蠖灘庄閭ｽ縺ｪ繝｡繝ｳ繝舌�ｼ蜷阪Μ繧ｹ繝医�
    roster_member_names: skills 縺ｮ蜈ｨ繝｡繝ｳ繝舌�ｼ蜷搾ｼ育怐逡･譎ゅ�ｯ capable_candidates�ｼ峨ょ酔荳蟋薙�ｮ驥崎､�蛻､螳壹↓菴ｿ逕ｨ縲�

    蜷榊燕縺ｮ陦ｨ險倥ｆ繧�:
    - 蟋薙�ｯ豁｣隕丞喧蠕後↓螳悟�ｨ荳閾ｴ縺ｮ縺ｿ�ｼ医ｆ繧瑚ｨｱ螳ｹ縺励↑縺�縲ょｯ檎伐/蜀ｨ逕ｰ縺ｮ縺ｿ蠕捺擂縺ｩ縺翫ｊ蟇�縺幢ｼ峨�
    - roster 縺ｫ蜷御ｸ蟋薙′2莠ｺ莉･荳翫＞縺ｪ縺�縺ｨ縺阪□縺代∝錐縺ｯ驛ｨ蛻�荳閾ｴ�ｼ医←縺｡繧峨°縺御ｻ匁婿繧貞性繧�ｼ峨∪縺溘�ｯ螳悟�ｨ荳閾ｴ繧定ｨｱ螳ｹ縲�
    - 蜷御ｸ蟋薙′繝ｭ繧ｹ繧ｿ繝ｼ縺ｫ縺�繧矩俣縺ｯ蜷阪ｂ螳悟�ｨ荳閾ｴ蠢�鬆医�
    - 蟋薙�ｮ縺ｿ縺ｮ蜈･蜉帙〒蜷阪ｆ繧後Δ繝ｼ繝峨�ｮ縺ｨ縺阪∝ｧ薙′荳閾ｴ縺吶ｋ蛟呵｣懊′隍�謨ｰ縺�繧後�ｰ隗｣豎ｺ荳榊庄�ｼ�None�ｼ峨�
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
    """蠖捺律繧ｹ繧ｭ繝ｫ荳� OP 縺ｮ縺ｿ縺ｸ隗｣豎ｺ�ｼ亥ｾ捺擂 API�ｼ峨ょｮ滉ｽ薙�ｯ `_resolve_preferred_name_to_capable_member`縲�"""
    return _resolve_preferred_name_to_capable_member(
        raw, op_candidates, roster_member_names
    )


def _task_process_matches_global_contains(machine_val: str, contains: str) -> bool:
    """蟾･遞句錐�ｼ医ち繧ｹ繧ｯ縺ｮ machine�ｼ峨↓驛ｨ蛻�荳閾ｴ�ｼ�NFKC繝ｻ螟ｧ蟆冗┌隕厄ｼ峨�"""
    m = unicodedata.normalize("NFKC", str(machine_val or "").strip()).casefold()
    c = unicodedata.normalize("NFKC", str(contains or "").strip()).casefold()
    if not c:
        return False
    return c in m


def _coerce_global_day_process_operator_rules(raw_val) -> list:
    """Gemini 縺ｮ global_day_process_operator_rules 繧呈ｭ｣隕丞喧�ｼ育ｩｺ繝ｻ荳肴ｭ｣縺ｯ髯､螟厄ｼ峨�"""
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
    繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝育罰譚･縺ｮ縲梧律莉佚怜ｷ･遞凝苓､�謨ｰ謖�蜷阪阪〒縲√◎縺ｮ譌･繝ｻ縺昴�ｮ蟾･遞九ち繧ｹ繧ｯ縺ｫ
    **繝√�ｼ繝�縺ｸ蠢�縺壼性繧√ｋ**繝｡繝ｳ繝舌�ｼ�ｼ�skills 陦後く繝ｼ�ｼ峨→隴ｦ蜻翫Γ繝�繧ｻ繝ｼ繧ｸ繧定ｿ斐☆縲�
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
                    "繝｡繧､繝ｳ繧ｰ繝ｭ繝ｼ繝舌Ν(譌･莉佚怜ｷ･遞�)謖�蜷�: "
                    f"萓晞�ｼNO={tid} 譌･莉�={current_date} 蟾･遞�={machine!r} 縺ｮ "
                    f"謖�蜷阪鶏raw_name}縲阪ｒ蠖捺律繧ｹ繧ｭ繝ｫ隧ｲ蠖薙Γ繝ｳ繝舌�ｼ縺ｫ隗｣豎ｺ縺ｧ縺阪∪縺帙ｓ"
                )
    return acc, warns


def _merge_global_day_process_and_pref_anchor(
    must_include: list, pref_mem, capable_members: list
) -> list[str]:
    """蠢�鬆医Γ繝ｳ繝舌�ｼ縺ｨ諡�蠖徹P謖�蜷阪ｒ1譛ｬ蛹厄ｼ�capable 縺ｫ縺�繧九ｂ縺ｮ縺�縺托ｼ峨�"""
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
# 2. 繝槭せ繧ｿ繝�繝ｼ繧ｿ繝ｻ蜃ｺ蜍､邁ｿ(繧ｫ繝ｬ繝ｳ繝繝ｼ) 縺ｨ AI隗｣譫�
#    master.xlsm 縺ｮ skills / need / 蜷�繝｡繝ｳ繝舌�ｼ蜍､諤�繧ｷ繝ｼ繝医ｒ隱ｭ縺ｿ縲�
#    蛯呵�繝ｻ莨第嚊蛹ｺ蛻�縺ｯ蠢�隕√↓蠢懊§縺ｦ Gemini 縺ｧ讒矩�蛹悶☆繧九�
# =========================================================
def load_skills_and_needs():
    """
    邨ｱ蜷医ヵ繧｡繧､繝ｫ(MASTER_FILE)縺九ｉ繧ｹ繧ｭ繝ｫ縺ｨ need 繧貞虚逧�縺ｫ隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺吶�

    謌ｻ繧雁､縺ｯ7隕∫ｴ�縲よ怙蠕後�ｯ need 繧ｷ繝ｼ繝井ｸ翫�ｮ縲悟ｷ･遞句錐+讖滓｢ｰ蜷阪榊�嶺ｽ咲ｽｮ�ｼ亥ｷｦ縺ｻ縺ｩ蟆上＆縺�謨ｴ謨ｰ�ｼ峨�ｮ霎樊嶌
    ``need_combo_col_index``�ｼ磯�榊床繧ｭ繝･繝ｼ繧ｽ繝ｼ繝育畑�ｼ峨�

    莉雁屓縺ｮ need 縺ｯ�ｼ�Excel荳翫〒�ｼ�
      蟾･遞句錐陦後�ｻ讖滓｢ｰ蜷崎｡後�ｮ縺ゅ→縲悟渕譛ｬ蠢�隕∽ｺｺ謨ｰ縲崎｡鯉ｼ�A蛻励↓縲悟ｿ�隕∽ｺｺ謨ｰ縲阪ｒ蜷ｫ繧�ｼ�
      縺昴�ｮ逶ｴ荳�: 驟榊床縺ｧ菴吝臆莠ｺ蜩｡縺後≠繧九→縺阪↓霑ｽ蜉�縺ｧ蜈･繧後ｉ繧後ｋ莠ｺ謨ｰ�ｼ亥ｷ･遞凝玲ｩ滓｢ｰ縺斐→縲よ悴險ｭ螳壹�ｯ 0�ｼ�
      莉･髯�: 迚ｹ蛻･謖�螳�1縲�99
    縺ｨ縺�縺�讒矩�縺ｮ縺溘ａ縲∝ｿ�隕＾P縺ｯ縲悟ｷ･遞句錐+讖滓｢ｰ蜷阪阪〒隗｣豎ｺ縺吶ｋ縲�

    skills 莠､蟾ｮ繧ｻ繝ｫ縺ｯ OP/AS 縺ｮ蠕後↓蜆ｪ蜈亥ｺｦ謨ｴ謨ｰ�ｼ井ｾ� OP1, AS3�ｼ峨よ焚蛟､縺悟ｰ上＆縺�縺ｻ縺ｩ蠖楢ｩｲ蟾･遞九∈縺ｮ蜑ｲ蠖薙〒蜆ｪ蜈医�
    謨ｰ蟄礼怐逡･縺ｮ OP/AS 縺ｯ蜆ｪ蜈亥ｺｦ 1縲�
    蜷御ｸ蛻暦ｼ亥酔荳蟾･遞凝玲ｩ滓｢ｰ�ｼ峨〒縺ｯ蜆ｪ蜈亥ｺｦ縺ｮ謨ｰ蛟､縺ｯ繝｡繝ｳ繝舌�ｼ髢薙〒驥崎､�荳榊庄�ｼ磯㍾隍�譎ゅ�ｯ PlanningValidationError�ｼ峨�
    """
    try:
        # 蜷御ｸ繝悶ャ繧ｯ繧� pd.read_excel 縺ｧ驛ｽ蠎ｦ髢九￥縺ｨ I/O 縺碁㍾縺�縺溘ａ縲・xcelFile 繧�1蝗槭□縺鷹幕縺�縺ｦ繧ｷ繝ｼ繝医ｒ parse 縺吶ｋ縲�
        with pd.ExcelFile(MASTER_FILE) as _master_xls:
            # skills 縺ｯ譁ｰ莉墓ｧ�:
            #   1陦檎岼: 蟾･遞句錐
            #   2陦檎岼: 讖滓｢ｰ蜷�
            #   A3莉･髯�: 繝｡繝ｳ繝舌�ｼ蜷�
            #   莠､蟾ｮ繧ｻ繝ｫ: OP 縺ｾ縺溘�ｯ AS 縺ｮ蠕後↓蜑ｲ蠖灘━蜈亥ｺｦ縺ｮ謨ｴ謨ｰ�ｼ井ｾ� OP1, AS3�ｼ峨よ焚蛟､縺悟ｰ上＆縺�縺ｻ縺ｩ蠖楢ｩｲ蟾･遞九∈蜆ｪ蜈亥牡蠖薙�
            #             謨ｰ蟄礼怐逡･縺ｮ OP/AS 縺ｯ蜆ｪ蜈亥ｺｦ 1�ｼ亥ｾ捺擂縺ｩ縺翫ｊ譛蜆ｪ蜈域桶縺��ｼ峨�
            # 繧貞渕譛ｬ縺ｨ縺励▽縺､縲∵立莉墓ｧ假ｼ�1陦後�倥ャ繝�ｼ峨↓繧ゅヵ繧ｩ繝ｼ繝ｫ繝舌ャ繧ｯ蟇ｾ蠢懊☆繧九�
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
                    "skills繧ｷ繝ｼ繝�: 2谿ｵ繝倥ャ繝蠖｢蠑上〒隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆�ｼ亥ｷ･遞�+讖滓｢ｰ=%s蛻�, 繝｡繝ｳ繝舌�ｼ=%s莠ｺ�ｼ峨�",
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
                    if c in ("繝｡繝ｳ繝舌�ｼ", "諡�蠖楢�", "豌丞錐", "菴懈･ｭ閠�"):
                        member_col = c
                        break
                if member_col is None and skill_cols:
                    member_col = skill_cols[0]
                    logging.warning(
                        "skills繧ｷ繝ｼ繝�: 繝｡繝ｳ繝舌�ｼ蛻怜錐縺梧ｨ呎ｺ悶→荳閾ｴ縺励↑縺�縺溘ａ縲∝�磯�ｭ蛻� '%s' 繧偵Γ繝ｳ繝舌�ｼ蛻励→縺励※謇ｱ縺�縺ｾ縺吶�",
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
                    "skills繧ｷ繝ｼ繝�: 1陦後�倥ャ繝蠖｢蠑擾ｼ域立莠呈鋤�ｼ峨〒隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆�ｼ医Γ繝ｳ繝舌�ｼ=%s莠ｺ�ｼ峨�",
                    len(members),
                )

            if not members:
                logging.error("skills繧ｷ繝ｼ繝医°繧峨Γ繝ｳ繝舌�ｼ繧定ｪｭ縺ｿ霎ｼ繧√∪縺帙ｓ縺ｧ縺励◆縲�")
            else:
                _validate_skills_op_as_priority_numbers_unique(
                    skills_dict, equipment_list
                )

            # need 縺ｯ header=None 縺ｧ隱ｭ縺ｿ縲∝�磯�ｭ縺ｮ隍�謨ｰ陦後ｒ窶懆ｦ句�ｺ縺苓｡娯昴→縺励※隗｣驥�
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
            if process_header_row is None and s0 == "蟾･遞句錐":
                process_header_row = r
            elif machine_header_row is None and s0 == "讖滓｢ｰ蜷�":
                machine_header_row = r
            if base_row is None and "蠢�隕∽ｺｺ謨ｰ" in s0 and not s0.startswith("迚ｹ蛻･謖�螳�"):
                base_row = r
            if process_header_row is not None and machine_header_row is not None and base_row is not None:
                break

        if process_header_row is None or machine_header_row is None or base_row is None:
            raise ValueError("need 繧ｷ繝ｼ繝医�ｮ繝倥ャ繝繝ｼ陦鯉ｼ亥ｷ･遞句錐/讖滓｢ｰ蜷�/蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ�ｼ峨′隕九▽縺九ｊ縺ｾ縺帙ｓ縲�")

        # 縲御ｾ晞�ｼNO譚｡莉ｶ縲榊�嶺ｽ咲ｽｮ�ｼ医ョ繝輔か繝ｫ繝医�ｯ 1蛻礼岼�ｼ�
        cond_col_idx = 1
        for r in range(needs_raw.shape[0]):
            c1 = needs_raw.iat[r, 1] if needs_raw.shape[1] > 1 else None
            c2 = needs_raw.iat[r, 2] if needs_raw.shape[1] > 2 else None
            if pd.isna(c1) or pd.isna(c2):
                continue
            if str(c1).strip() == NEED_COL_CONDITION and str(c2).strip() == NEED_COL_NOTE:
                cond_col_idx = 1
                break

        # 蟾･遞句錐ﾃ玲ｩ滓｢ｰ蜷� 縺ｮ蛻嶺ｸ隕ｧ�ｼ亥�礼分蜿ｷ縺ｯ Excel荳翫�ｮ螳溷�励ｒ菫晄戟�ｼ�
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
        # 蟾･遞句錐+讖滓｢ｰ蜷阪さ繝ｳ繝� 竊� need 繧ｷ繝ｼ繝井ｸ翫�ｮ蛻励う繝ｳ繝�繝�繧ｯ繧ｹ�ｼ亥ｷｦ縺ｻ縺ｩ蟆上＆縺��ｼ晞�榊床繧ｭ繝･繝ｼ縺ｧ蜈茨ｼ�
        need_combo_col_index: dict[str, int] = {}
        # need_rules: [{'order': int, 'condition': str, 'overrides': {combo_key/machine/process: int}}]
        need_rules = []

        # 蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ
        for col_idx, p_s, m_s in pm_cols:
            n = parse_optional_int(needs_raw.iat[base_row, col_idx])
            if n is None or n < 1:
                n = 1
            combo_key = f"{p_s}+{m_s}"
            need_combo_col_index[combo_key] = col_idx
            req_map[combo_key] = n
            # 繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ逕ｨ�ｼ域ｩ滓｢ｰ蜷� or 蟾･遞句錐縺�縺代〒蠑輔￠繧九ｈ縺�縺ｫ縺吶ｋ�ｼ�
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
                "need 繧ｷ繝ｼ繝�: 驟榊床譎りｿｽ蜉�莠ｺ謨ｰ陦後ｒ讀懷�ｺ�ｼ�Excel陦娯沿%s�ｼ峨ょ�励＃縺ｨ縺ｮ荳企剞繧定ｪｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�",
                surplus_row + 1,
            )
        else:
            logging.info(
                "need 繧ｷ繝ｼ繝�: 蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ縺ｮ逶ｴ荳九↓驟榊床譎りｿｽ蜉�莠ｺ謨ｰ陦後ｒ讀懷�ｺ縺ｧ縺阪∪縺帙ｓ縺ｧ縺励◆�ｼ育怐逡･蜿ｯ�ｼ峨�"
            )

        if TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW:
            logging.info(
                "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW 縺梧怏蜉ｹ: 驟榊床譎りｿｽ蜉�莠ｺ謨ｰ縺ｯ隱ｭ縺ｿ霎ｼ繧薙〒繧ょｸｸ縺ｫ 0 謇ｱ縺��ｼ医メ繝ｼ繝�縺ｯ蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ縺ｮ縺ｿ隧ｦ陦鯉ｼ峨�"
            )

        logging.info(
            "need莠ｺ謨ｰ繝槭せ繧ｿ: %s 縺ｮ need 繧ｷ繝ｼ繝医ｒ隱ｭ縺ｿ霎ｼ縺ｿ�ｼ�skills 縺ｨ蜷御ｸ ExcelFile 縺ｧ髢九＞縺溽峩蠕後Ｏeed 蟆ら畑繝�繧｣繧ｹ繧ｯ繧ｭ繝｣繝�繧ｷ繝･縺ｯ辟｡縺励�ｻAI json 縺ｨ縺ｯ辟｡髢｢菫ゑｼ峨�",
            os.path.abspath(MASTER_FILE),
        )
        for _ci, _ps, _ms in pm_cols:
            _ck = f"{_ps}+{_ms}"
            _bn = req_map.get(_ck)
            _sx = surplus_map.get(_ck, 0) if surplus_map else 0
            logging.info(
                "need蛻励し繝槭Μ combo=%r 蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ=%s 驟榊床譎りｿｽ蜉�莠ｺ謨ｰ荳企剞=%s",
                _ck,
                _bn,
                _sx,
            )

        # 迚ｹ蛻･謖�螳�
        for r in range(needs_raw.shape[0]):
            v0 = needs_raw.iat[r, col0]
            if pd.isna(v0):
                continue
            lab = str(v0).strip()
            m = re.match(r"迚ｹ蛻･謖�螳喀s*(\d+)", lab)
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
                    # 繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ逕ｨ
                    overrides[p_s] = n
                    overrides[m_s] = n

            if overrides:
                need_rules.append({"order": order, "condition": cond, "overrides": overrides})

        need_rules.sort(key=lambda rr: rr["order"])
        logging.info(f"need 迚ｹ蛻･謖�螳壹Ν繝ｼ繝ｫ: {len(need_rules)} 莉ｶ�ｼ亥ｷ･遞句錐+讖滓｢ｰ蜷阪く繝ｼ�ｼ峨�")

        logging.info(f"縲施MASTER_FILE}縲上°繧峨せ繧ｭ繝ｫ縺ｨ險ｭ蛯呵ｦ∽ｻｶ(need)繧定ｪｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�")
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
        logging.error(f"繝槭せ繧ｿ繝輔ぃ繧､繝ｫ({MASTER_FILE})縺ｮ繧ｹ繧ｭ繝ｫ/need隱ｭ縺ｿ霎ｼ縺ｿ繧ｨ繝ｩ繝ｼ: {e}")
        return {}, [], [], {}, [], [], {}


def load_team_combination_presets_from_master() -> dict[
    str, list[tuple[int, int | None, tuple[str, ...], int | None]]
]:
    """
    master.xlsm縲檎ｵ�縺ｿ蜷医ｏ縺幄｡ｨ縲阪ｒ隱ｭ縺ｿ縲∝ｷ･遞�+讖滓｢ｰ繧ｭ繝ｼ縺斐→縺ｫ
    [(邨�蜷医○蜆ｪ蜈亥ｺｦ, 蠢�隕∽ｺｺ謨ｰ縺ｾ縺溘�ｯNone, 繝｡繝ｳ繝舌�ｼ繧ｿ繝励Ν, 邨�蜷医○陦栗D縺ｾ縺溘�ｯNone), ...] 繧定ｿ斐☆縲�
    蜷御ｸ繧ｭ繝ｼ蜀�縺ｯ蜆ｪ蜈亥ｺｦ譏�鬆�縲∝酔鬆�菴阪�ｯ繧ｷ繝ｼ繝井ｸ翫�ｮ陦碁��縲�
    縲悟ｿ�隕∽ｺｺ謨ｰ縲榊�励�ｯ驟榊床譎ゅ↓ need 蝓ｺ譛ｬ莠ｺ謨ｰ繧医ｊ蜆ｪ蜈医☆繧具ｼ医Γ繝ｳ繝舌�ｼ蛻嶺ｺｺ謨ｰ縺ｨ荳閾ｴ縺吶ｋ縺薙→�ｼ峨�
    驟榊床縺ｧ縺ｯ謌千ｫ九＠縺溘�励Μ繧ｻ繝�繝医ｒ縺吶∋縺ｦ蛟呵｣懊↓霈峨○縲∫ｵ�蜷医○謗｢邏｢縺ｨ縺ｾ縺ｨ繧√※ team_start 遲峨〒譛濶ｯ繧帝∈縺ｶ
    �ｼ医す繝ｼ繝亥━蜈亥ｺｦ縺ｯ隧ｦ陦碁��縺ｮ縺ｿ縲ょ�磯�ｭ繝励Μ繧ｻ繝�繝医�ｮ蜊ｳ豎ｺ縺ｯ縺励↑縺��ｼ峨�
    A 蛻励檎ｵ�蜷医○陦栗D縲阪′辟｡縺��ｼ冗ｩｺ縺ｮ譌ｧ繧ｷ繝ｼ繝医〒縺ｯ ID 縺ｯ None縲�
    """
    if not TEAM_ASSIGN_USE_MASTER_COMBO_SHEET:
        return {}
    path = MASTER_FILE
    if not os.path.isfile(path):
        return {}
    try:
        df = pd.read_excel(path, sheet_name=MASTER_SHEET_TEAM_COMBINATIONS, header=0)
    except Exception as e:
        logging.info("邨�縺ｿ蜷医ｏ縺幄｡ｨ繧ｷ繝ｼ繝医�ｮ隱ｭ霎ｼ繧偵せ繧ｭ繝�繝励＠縺ｾ縺�: %s", e)
        return {}
    if df is None or df.empty:
        return {}

    def norm_cell(x) -> str:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip()

    colmap = {norm_cell(c): c for c in df.columns if norm_cell(c)}
    id_c = colmap.get("邨�蜷医○陦栗D") or colmap.get("繧､繝ｳ繝�繝�繧ｯ繧ｹ")
    proc_c = colmap.get("蟾･遞句錐")
    mach_c = colmap.get("讖滓｢ｰ蜷�")
    combo_c = colmap.get("蟾･遞�+讖滓｢ｰ")
    prio_c = colmap.get("邨�蜷医○蜆ｪ蜈亥ｺｦ")
    req_c = colmap.get("蠢�隕∽ｺｺ謨ｰ")

    def mem_col_order(c) -> int:
        m = re.search(r"繝｡繝ｳ繝舌�ｼ\s*(\d+)", norm_cell(c))
        return int(m.group(1)) if m else 9999

    mem_keys = sorted(
        [c for c in df.columns if norm_cell(str(c)).startswith("繝｡繝ｳ繝舌�ｼ")],
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
    謗｡逕ｨ繝√�ｼ繝�縺ｮ繝｡繝ｳ繝舌�ｼ髮�蜷茨ｼ�NFKC繝ｻtrim�ｼ峨′邨�縺ｿ蜷医ｏ縺幄｡ｨ繝励Μ繧ｻ繝�繝医�ｮ縺�縺壹ｌ縺九→荳閾ｴ縺吶ｋ縺ｨ縺阪�
    縺昴�ｮ陦後�ｮ邨�蜷医○陦栗D�ｼ�A蛻暦ｼ峨ｒ霑斐☆縲らｵ�蜷医○謗｢邏｢縺ｮ縺ｿ縺ｧ豎ｺ縺ｾ繧� combo_sheet_row_id 縺御ｻ倥＞縺ｦ縺�縺ｪ縺�
    螻･豁ｴ陦後�ｮ陬懷ｮ後↓菴ｿ縺�縲り､�謨ｰ荳閾ｴ譎ゅ�ｯ邨�蜷医○蜆ｪ蜈亥ｺｦ�ｼ域焚蛟､縺悟ｰ上＆縺�譁ｹ�ｼ峨ｒ謗｡逕ｨ縲�
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
    """master 繝｡繧､繝ｳ縺ｮ譎ょ綾繧ｻ繝ｫ�ｼ�datetime / time / 譁�蟄怜�暦ｼ峨ｒ time 縺ｫ縲りｧ｣驥井ｸ崎�ｽ縺ｯ None縲�"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, time):
        return v
    if isinstance(v, datetime):
        return v.time()
    return parse_time_str(v, None)


def _pick_master_main_sheet_name(sheetnames: list[str]) -> str | None:
    """
    master.xlsm 縺ｮ縲後Γ繧､繝ｳ縲崎ｨｭ螳壹す繝ｼ繝亥錐繧定ｧ｣豎ｺ縺吶ｋ�ｼ�VBA MasterGetMainWorksheet 縺ｨ蜷瑚ｶ｣譌ｨ�ｼ峨�
    縲後�譛医Γ繧､繝ｳ繧ｫ繝ｬ繝ｳ繝繝ｼ縲咲ｭ峨ｒ隱､謗｡逕ｨ縺励↑縺�繧医≧縲後き繝ｬ繝ｳ繝繝ｼ縲阪ｒ蜷ｫ繧蜷榊燕縺ｯ髯､螟悶＠縲�
    隍�謨ｰ蛟呵｣懊�ｯ繧ｷ繝ｼ繝亥錐縺梧怙遏ｭ縺ｮ繧ゅ�ｮ繧貞━蜈医☆繧九�
    """
    for prefer in ("繝｡繧､繝ｳ", "繝｡繧､繝ｳ_", "Main"):
        if prefer in sheetnames:
            return prefer
    cand = [sn for sn in sheetnames if "繝｡繧､繝ｳ" in sn and "繧ｫ繝ｬ繝ｳ繝繝ｼ" not in sn]
    if not cand:
        return None
    return min(cand, key=len)


def _read_master_main_factory_operating_times(master_path: str) -> tuple[time | None, time | None]:
    """
    master.xlsm 縺ｮ繝｡繧､繝ｳ繧ｷ繝ｼ繝� A12�ｼ育ｨｼ蜒埼幕蟋具ｼ峨�ｻB12�ｼ育ｨｼ蜒咲ｵゆｺ��ｼ峨ｒ隱ｭ繧縲�
    縺�縺壹ｌ縺区ｬ�謳阪�ｻ荳肴ｭ｣繝ｻ髢句ｧ�>=邨ゆｺ�縺ｮ縺ｨ縺阪�ｯ (None, None)縲�
    """
    p = (master_path or "").strip()
    if not p or not os.path.isfile(p):
        return None, None
    if _workbook_should_skip_openpyxl_io(p):
        return None, None
    try:
        wb = load_workbook(p, data_only=True, read_only=False)
    except Exception as e:
        logging.warning("蟾･蝣ｴ遞ｼ蜒肴凾蛻ｻ: master 繧� openpyxl 縺ｧ髢九￠縺ｾ縺帙ｓ縺ｧ縺励◆�ｼ域里螳壹�ｮ譌･蜀�譫�繧剃ｽｿ縺�縺ｾ縺呻ｼ�: %s", e)
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
                "蟾･蝣ｴ遞ｼ蜒肴凾蛻ｻ: master 繝｡繧､繝ｳ A12/B12 縺碁幕蟋�>=邨ゆｺ� (%s >= %s) 縺ｮ縺溘ａ譌｢螳壼､繧剃ｽｿ縺�縺ｾ縺吶�",
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
    master.xlsm 縺ｮ繝｡繧､繝ｳ繧ｷ繝ｼ繝� A15�ｼ亥ｮ壼ｸｸ髢句ｧ具ｼ峨�ｻB15�ｼ亥ｮ壼ｸｸ邨ゆｺ��ｼ峨ｒ隱ｭ繧縲�
    縺�縺壹ｌ縺区ｬ�謳阪�ｻ荳肴ｭ｣繝ｻ髢句ｧ�>=邨ゆｺ�縺ｮ縺ｨ縺阪�ｯ (None, None)縲�
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
            "螳壼ｸｸ譎ょ綾: master 繧� openpyxl 縺ｧ髢九￠縺ｾ縺帙ｓ縺ｧ縺励◆�ｼ育ｵ先棡繧ｷ繝ｼ繝医�ｮ螳壼ｸｸ螟也捩濶ｲ繧偵せ繧ｭ繝�繝暦ｼ�: %s",
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
                "螳壼ｸｸ譎ょ綾: master 繝｡繧､繝ｳ A15/B15 縺碁幕蟋�>=邨ゆｺ� (%s >= %s) 縺ｮ縺溘ａ逹濶ｲ繝ｻ豈碑ｼ�縺ｫ菴ｿ縺�縺ｾ縺帙ｓ縲�",
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
    """谿ｵ髫�2縺ｮ髢薙□縺� DEFAULT_START_TIME / DEFAULT_END_TIME 繧� master 繝｡繧､繝ｳ A12/B12 縺ｧ荳頑嶌縺阪�"""
    global DEFAULT_START_TIME, DEFAULT_END_TIME
    orig_s, orig_e = DEFAULT_START_TIME, DEFAULT_END_TIME
    ns, ne = _read_master_main_factory_operating_times(master_path)
    try:
        if ns is not None and ne is not None:
            DEFAULT_START_TIME = ns
            DEFAULT_END_TIME = ne
            logging.info(
                "蟾･蝣ｴ遞ｼ蜒肴棧: master.xlsm 繝｡繧､繝ｳ A12/B12 繧呈治逕ｨ 竊� %s �ｽ� %s�ｼ育ｵ先棡_* 縺ｮ譌･蜀�繧ｰ繝ｪ繝�繝峨�ｻ驟榊床譫��ｼ�",
                DEFAULT_START_TIME.strftime("%H:%M"),
                DEFAULT_END_TIME.strftime("%H:%M"),
            )
        yield
    finally:
        DEFAULT_START_TIME, DEFAULT_END_TIME = orig_s, orig_e


def infer_mid_break_from_reason(reason_text, start_t, end_t, break1_start=None, break1_end=None):
    """
    蛯呵�縺九ｉ荳ｭ謚懊￠譎る俣繧呈耳螳壹☆繧九Ο繝ｼ繧ｫ繝ｫ陬懈ｭ｣縲�
    AI縺御ｸｭ謚懊￠繧定ｿ斐＆縺ｪ縺�蝣ｴ蜷医�ｮ繝輔ぉ繧､繝ｫ繧ｻ繝ｼ繝輔→縺励※菴ｿ縺�縲�
    master.xlsm 繧ｫ繝ｬ繝ｳ繝繝ｼ逕ｱ譚･縺ｮ莨第嚊蛹ｺ蛻�: 蜑堺ｼ�=蜊亥燕蟷ｴ莨代�ｻ蜊亥ｾ後�ｮ縺ｿ蜍､蜍吶∝ｾ御ｼ�=蜊亥ｾ悟ｹｴ莨代�ｻ蜊亥燕縺ｮ縺ｿ蜍､蜍呻ｼ亥�ｺ蜍､邁ｿ.txt 縺ｨ蜷檎ｾｩ�ｼ峨�
    蜑堺ｼ代�ｻ蠕御ｼ代�ｮ蠅�逡後�ｯ繝｡繝ｳ繝舌�ｼ蜍､諤�縺ｮ莨第�ｩ譎る俣1_髢句ｧ�/邨ゆｺ��ｼ域悴謖�螳壽凾縺ｯ DEFAULT_BREAKS[0]�ｼ峨↓蜷医ｏ縺帙ｋ縲�
    """
    if reason_text is None:
        return None, None
    txt = str(reason_text).strip()
    if not txt or txt.lower() in ("nan", "none", "null", "騾壼ｸｸ"):
        return None, None

    b1_s = break1_start if break1_start is not None else DEFAULT_BREAKS[0][0]
    b1_e = break1_end if break1_end is not None else DEFAULT_BREAKS[0][1]

    noon_end = time(12, 0)
    afternoon_start = time(13, 0)
    # 繧ｫ繝ｬ繝ｳ繝繝ｼ險伜捷縺ｨ荳閾ｴ縺輔○繧具ｼ医す繝輔ヨ譎ょ綾縺瑚ｪ､縺｣縺ｦ縺�繧句�ｴ蜷医�ｮ陬懷ｮ檎畑縲よｭ｣縺励＞陦後〒縺ｯ蛹ｺ髢薙′遨ｺ縺ｫ縺ｪ繧願ｿｽ蜉�縺輔ｌ縺ｪ縺��ｼ�
    if txt == "蜑堺ｼ�":
        # 豁｣縺励＞陦後�ｯ蜃ｺ蜍､縺御ｼ第�ｩ1邨ゆｺ�莉･髯阪〒陬懷ｮ御ｸ崎ｦ√ょ�ｨ譌･繧ｷ繝輔ヨ縺ｮ隱､蜈･蜉帶凾縺ｯ縺昴％縺ｾ縺ｧ繧剃ｸｭ謚懊￠�ｼ亥壕蜑榊ｹｴ莨醍嶌蠖難ｼ�
        if start_t and start_t < b1_e:
            return start_t, b1_e
        return None, None
    if txt == "蠕御ｼ�":
        if end_t and b1_s < end_t:
            return b1_s, end_t
        return None, None

    # 1) 譏守､ｺ逧�縺ｪ譎ょ綾遽�蝗ｲ�ｼ井ｾ�: 11:00-14:00 / 11:00�ｽ�14:00�ｼ�
    m = re.search(r"(\d{1,2}[:�ｼ咯\d{2})\s*[~縲彌-�ｼ阪�ｼ]\s*(\d{1,2}[:�ｼ咯\d{2})", txt)
    if m:
        s = parse_time_str(m.group(1).replace("�ｼ�", ":"), None)
        e = parse_time_str(m.group(2).replace("�ｼ�", ":"), None)
        if s and e and s < e:
            return s, e

    # 2) 縺ゅ＞縺ｾ縺�隱橸ｼ亥壕蜑�/蜊亥ｾ�/邨よ律�ｼ� + 迴ｾ蝣ｴ髮｢閼ｱ繝ｻ莨第嚊邉ｻ繧ｭ繝ｼ繝ｯ繝ｼ繝�
    # 縲悟壕蠕御ｼ代∩縺ｧ縺吶咲ｭ峨�ｯ縲悟壕蠕後阪ｒ蜷ｫ繧縺後∵立繝ｭ繧ｸ繝�繧ｯ縺ｯ縲梧栢縺代咲ｭ峨�ｮ縺ｿ隕九※縺翫ｊ荳ｭ謚懊￠謗ｨ螳壹↓蛻ｰ驕斐＠縺ｪ縺九▲縺�
    leave_keywords = (
        "莠句漁謇", "莨夊ｭｰ", "謨呵ご", "遐比ｿｮ", "螟門�ｺ", "髮｢繧�", "謚懊￠", "荳ｭ謚懊￠", "謇灘粋縺�",
        "莨代∩", "莨第嚊", "谺�蜍､",
    )
    has_leave_hint = any(k in txt for k in leave_keywords)
    if not has_leave_hint:
        return None, None

    if ("邨よ律" in txt) or ("1譌･" in txt and "騾壼ｸｸ" not in txt):
        return start_t, end_t
    if ("蜊亥燕荳ｭ" in txt) or ("蜊亥燕" in txt):
        return start_t, noon_end
    if ("蜊亥ｾ�" in txt):
        return afternoon_start, end_t

    return None, None


# 邨先棡_繧ｫ繝ｬ繝ｳ繝繝ｼ(蜃ｺ蜍､邁ｿ) 縺ｮ騾蜍､陦ｨ遉ｺ縲７BA 蜃ｺ蜍､邁ｿ縲悟ｾ御ｼ代搾ｼ亥壕蠕悟ｹｴ莨托ｼ峨→蜷梧ｧ倥↓螳溯ｳｪ 莨第�ｩ譎る俣1_髢句ｧ九〒邨ゆｺ�縺ｨ縺ｿ縺ｪ縺吶�
_AFTERNOON_OFF_DISPLAY_END = DEFAULT_BREAKS[0][0]


def _reason_is_afternoon_off(reason: str) -> bool:
    """蠕御ｼ托ｼ亥壕蠕悟ｹｴ莨代�ｻ蜊亥燕縺ｮ縺ｿ蜍､蜍呻ｼ峨∪縺溘�ｯ蛯呵�縺ｮ蜊亥ｾ御ｼ醍ｳｻ縲�"""
    r = str(reason or "")
    return ("蜊亥ｾ�" in r and ("莨�" in r or "莨代∩" in r)) or ("蠕御ｼ�" in r)


def _reason_is_morning_off(reason: str) -> bool:
    """蜑堺ｼ托ｼ亥壕蜑榊ｹｴ莨代�ｻ蜊亥ｾ後�ｮ縺ｿ蜍､蜍呻ｼ峨ゅき繝ｬ繝ｳ繝繝ｼ逕ｱ譚･縺ｮ逡･蜿ｷ縺ｮ縺ｿ譏守､ｺ謇ｱ縺��ｼ井ｺ句漁謇蜍､蜍吶↑縺ｩ縺ｨ豺ｷ蜷後＠縺ｪ縺��ｼ峨�"""
    return "蜑堺ｼ�" in str(reason or "")


def _calendar_display_clock_out_for_calendar_sheet(entry: dict, day_date: date):
    """
    驟榊床縺ｯ breaks_dt 縺ｮ蜊亥ｾ御ｸｭ謚懊￠縺ｧ豁｣縺励￥縺ｪ繧倶ｸ譁ｹ縲‘nd_dt 縺� 17:00 縺ｮ縺ｾ縺ｾ縺�縺ｨ邨先棡繧ｫ繝ｬ繝ｳ繝繝ｼ縺ｮ騾蜍､蛻励□縺題ｪ､繧九�
    蠕御ｼ托ｼ亥壕蠕悟ｹｴ莨托ｼ峨∪縺溘�ｯ蛯呵�縺悟壕蠕御ｼ代∩邉ｻ縺ｧ縲∝ｮ壽凾縺ｾ縺ｧ邯壹￥蜊亥ｾ後�ｮ荳ｭ謚懊￠縺後≠繧九→縺阪□縺鷹蜍､陦ｨ遉ｺ繧剃ｼ第�ｩ譎る俣1_髢句ｧ九↓謠�縺医ｋ�ｼ�end_dt 譛ｬ菴薙�ｯ螟画峩縺励↑縺��ｼ峨�
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
    蛟倶ｺｺ_* 繧ｹ繧ｱ繧ｸ繝･繝ｼ繝ｫ縺ｮ10蛻�譫�縺御ｼ第�ｩ蟶ｯ縺ｫ蜈･繧九→縺阪�ｮ譁�險縲�
    譏ｼ鬟溘↑縺ｩ騾壼ｸｸ莨第�ｩ縺ｯ縲御ｼ第�ｩ縲阪ょｾ御ｼ托ｼ亥壕蠕悟ｹｴ莨托ｼ峨〒螳壽凾縺ｾ縺ｧ蟾･蝣ｴ縺ｫ縺�縺ｪ縺�蜊亥ｾ悟ｸｯ縺ｯ縲御ｼ第嚊縲阪�
    蜑堺ｼ托ｼ亥壕蜑榊ｹｴ莨托ｼ峨〒蜊亥燕縺ｮ谺�蜍､蛹ｺ髢薙′莨第�ｩ蟶ｯ縺ｨ縺励※蜈･縺｣縺ｦ縺�繧句�ｴ蜷医�ｯ縲御ｼ第嚊縲阪�
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
                return "莨第嚊"
            if morning_off and bs < DEFAULT_BREAKS[0][0]:
                be_t = b_e.time() if isinstance(b_e, datetime) else b_e
                if be_t <= time(13, 0):
                    return "莨第嚊"
        return "莨第�ｩ"
    return None


def _member_schedule_off_shift_label(
    day_date: date,
    grid_mid_dt: datetime,
    d_start_dt: datetime,
    d_end_dt: datetime,
    reason: str,
) -> str:
    """
    蛟倶ｺｺ_* 繧ｷ繝ｼ繝医〒謇螳壼�ｺ騾蜍､縺ｮ螟門�ｴ縺ｮ10蛻�譫�縲�
    蜑堺ｼ代�ｮ蜊亥燕�ｼ亥ｷ･蝣ｴ譌･縺ｮ謇螳夐幕蟋具ｽ槫壕蠕悟�ｺ蜍､縺ｾ縺ｧ�ｼ峨�ｯ蟷ｴ莨代∝ｾ御ｼ代�ｮ蜊亥ｾ後�ｯ蟷ｴ莨代ゅ◎繧御ｻ･螟悶�ｮ繧ｷ繝輔ヨ螟悶�ｯ蜍､蜍吝､悶�
    """
    r = str(reason or "")
    day_start = datetime.combine(day_date, DEFAULT_START_TIME)
    day_end = datetime.combine(day_date, DEFAULT_END_TIME)
    if grid_mid_dt < d_start_dt:
        if _reason_is_morning_off(r) and grid_mid_dt >= day_start:
            return "蟷ｴ莨�"
        return "蜍､蜍吝､�"
    if grid_mid_dt >= d_end_dt:
        if _reason_is_afternoon_off(r) and grid_mid_dt < day_end:
            return "蟷ｴ莨�"
        return "蜍､蜍吝､�"
    return "蜍､蜍吝､�"


def _member_schedule_full_day_off_label(entry) -> str:
    """
    蜈ｨ譌･髱槫共蜍呻ｼ�is_working=False�ｼ峨�ｮ蛟倶ｺｺ繧ｷ繝ｼ繝亥�励�ｮ陦ｨ遉ｺ縲�
    莨第嚊蛹ｺ蛻�縺悟ｹｴ莨托ｼ医き繝ｬ繝ｳ繝繝ｼ *�ｼ峨�ｮ縺ｨ縺阪�ｯ縲主ｹｴ莨代上ょｷ･蝣ｴ莨第律縺ｪ縺ｩ縺ｯ縲惹ｼ代上�
    """
    if not entry:
        return "莨�"
    r = str(entry.get("reason") or "").strip()
    if r == "蟷ｴ莨�" or r.startswith("蟷ｴ莨� "):
        return "蟷ｴ莨�"
    return "莨�"


def _attendance_remark_text(row) -> str:
    """
    蜍､諤�1陦後°繧峨悟ｙ閠�縲榊�励�ｮ繝�繧ｭ繧ｹ繝医�ｮ縺ｿ蜿門ｾ励☆繧九�
    蜍､諤�AI縺ｮ隗｣譫舌Μ繧ｹ繝医∈縺ｮ謚募�･縺ｯ縺薙�ｮ蛻励�ｮ縺ｿ縲Ｓeason 譁�蟄怜�励�ｯ load_attendance 縺ｧ蛯呵�縺ｨ莨第嚊蛹ｺ蛻�繧貞粋謌舌☆繧九�
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
    """蜍､諤�1陦後°繧峨御ｼ第嚊蛹ｺ蛻�縲榊�暦ｼ医き繝ｬ繝ｳ繝繝ｼ逕ｱ譚･縺ｮ 蜑堺ｼ�/蠕御ｼ� 遲会ｼ峨�"""
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
    """蜍､諤�蛯呵� AI 縺ｮ逵溷⊃蛟､�ｼ�bool / 謨ｰ蛟､ / 譁�蟄怜�励�ｮ謠ｺ繧後ｒ蜷ｸ蜿趣ｼ峨�"""
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
    if s in ("true", "1", "yes", "y", "縺ｯ縺�", "逵�", "on"):
        return True
    if s in ("false", "0", "no", "n", "縺�縺�縺�", "蛛ｽ", "off", ""):
        return False
    return default


def _parse_attendance_overtime_end_optional(v) -> time | None:
    """蜍､諤�縲梧ｮ区･ｭ邨よ･ｭ縲榊�励よ怏蜉ｹ縺ｪ譎ょ綾縺ｮ縺ｿ縲らｩｺ繝ｻ荳肴ｭ｣縺ｯ None�ｼ�_excel_scalar_to_time_optional 縺ｨ蜷瑚ｶ｣譌ｨ�ｼ峨�"""
    return _excel_scalar_to_time_optional(v)


def load_attendance_and_analyze(members):
    attendance_data = {}
    # 窶ｻ縲悟共諤�蛯呵�縲阪�ｯ master 蜷�繝｡繝ｳ繝舌�ｼ繧ｷ繝ｼ繝医�ｮ縲悟ｙ閠�縲榊�励�ｮ縺ｿ縲ゅΓ繧､繝ｳ蜀榊━蜈医�ｻ迚ｹ蛻･謖�螳喟蛯呵�縺ｯ蛻･API�ｼ�generate_plan 蛛ｴ縺ｧ霑ｽ險假ｼ峨�
    ai_log = {
        "�ｼ域ｳｨ�ｼ峨％縺ｮ繧ｷ繝ｼ繝医�ｮ隕区婿": "蜈磯�ｭ2陦後�ｯ蜍､諤�縲悟ｙ閠�縲阪�ｮ蜃ｺ騾蜍､AI縺ｮ縺ｿ縲ゅΓ繧､繝ｳ蜀榊━蜈医�ｻ迚ｹ蛻･謖�螳壹�ｯ荳区ｮｵ縺ｮJSON縺ｨ縲契*_AI_API縲崎｡後�",
        "蜍､諤�蛯呵ダAI_API": "縺ｪ縺�",
        "蜍､諤�蛯呵ダAI_隧ｳ邏ｰ": "隗｣譫仙ｯｾ雎｡縺ｮ蛯呵�陦後↑縺�",
    }
    
    # 1. 繝｡繝ｳ繝舌�ｼ蛻･繧ｷ繝ｼ繝医°繧峨�ｮ隱ｭ縺ｿ霎ｼ縺ｿ
    all_records = []
    try:
        xls = pd.ExcelFile(MASTER_FILE)
        for sheet_name in xls.sheet_names:
            if "繧ｫ繝ｬ繝ｳ繝繝ｼ" in sheet_name or sheet_name.lower() in ['skills', 'need', 'tasks']:
                continue 
                
            m_name = sheet_name.strip()
            if m_name not in members:
                continue 
                
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
            df_sheet.columns = df_sheet.columns.str.strip()
            df_sheet['繝｡繝ｳ繝舌�ｼ'] = m_name 
            all_records.append(df_sheet)
            
        if all_records:
            df = pd.concat(all_records, ignore_index=True)
            df['譌･莉�'] = pd.to_datetime(df['譌･莉�'], errors='coerce').dt.date
            df = df.dropna(subset=['譌･莉�'])
            logging.info(f"縲施MASTER_FILE}縲上�ｮ蜷�繝｡繝ｳ繝舌�ｼ縺ｮ蜍､諤�繧ｷ繝ｼ繝医ｒ隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�")
            _cols = {str(c).strip() for c in df.columns}
            if ATT_COL_REMARK in _cols and ATT_COL_LEAVE_TYPE in _cols:
                logging.info(
                    "蜍､諤�蛻�: AI 蜈･蜉帙�ｯ縲�%s縲阪�ｮ縺ｿ縲ょｙ閠�縺檎ｩｺ縺ｮ譌･縺ｯ縲�%s縲搾ｼ亥燕莨代�ｻ蠕御ｼ代�ｻ莉匁侠轤ｹ蜍､蜍吶↑縺ｩ�ｼ峨ｒ reason 縺ｫ蜿肴丐縺励∪縺吶�",
                    ATT_COL_REMARK,
                    ATT_COL_LEAVE_TYPE,
                )
            elif ATT_COL_REMARK not in _cols:
                logging.warning(
                    "蜍､諤�繝�繝ｼ繧ｿ縺ｫ縲�%s縲榊�励′縺ゅｊ縺ｾ縺帙ｓ縲ょｙ閠�繝吶�ｼ繧ｹ縺ｮ AI 隗｣譫舌�ｯ遨ｺ謇ｱ縺�縺ｫ縺ｪ繧翫∪縺吶�",
                    ATT_COL_REMARK,
                )
            if ATT_COL_OT_END in _cols:
                logging.info(
                    "蜍､諤�蛻�: 莉ｻ諢上�%s縲阪〒騾蜍､荳企剞譎ょ綾繧呈欠螳壹〒縺阪∪縺呻ｼ亥�ｨ譌･莨代∩陦後〒縺ｯ辟｡隕厄ｼ峨�",
                    ATT_COL_OT_END,
                )
        else:
            raise FileNotFoundError("譛牙柑縺ｪ繝｡繝ｳ繝舌�ｼ蛻･蜍､諤�繧ｷ繝ｼ繝医′隕九▽縺九ｊ縺ｾ縺帙ｓ縲�")
            
    except Exception as e:
        logging.warning(f"蜍､諤�繧ｷ繝ｼ繝郁ｪｭ縺ｿ霎ｼ縺ｿ繧ｨ繝ｩ繝ｼ: {e} 繝�繝輔か繝ｫ繝医き繝ｬ繝ｳ繝繝ｼ繧堤函謌舌＠縺ｾ縺吶�")
        default_dates = generate_default_calendar_dates(TARGET_YEAR, TARGET_MONTH)
        records = []
        for d in default_dates:
            for m in members: records.append({'譌･莉�': d, '繝｡繝ｳ繝舌�ｼ': m, '蛯呵�': '騾壼ｸｸ'})
        df = pd.DataFrame(records)

    # 2. AI 縺ｫ繧医ｋ蜍､諤�譁�閼医�ｮ隗｣譫撰ｼ亥ｙ閠�縺檎ｩｺ縺ｧ繧ゆｼ第嚊蛹ｺ蛻�縺ｮ縺ｿ縺ｮ陦後�ｯ AI 縺ｫ貂｡縺励∬｡ｨ險俶昭繧後�ｯ繝｢繝�繝ｫ縺ｫ隗｣驥医＆縺帙ｋ�ｼ�
    remarks_to_analyze = []
    for _, row in df.iterrows():
        m = str(row.get('繝｡繝ｳ繝舌�ｼ', '')).strip()
        if m not in members:
            continue
        rem = _attendance_remark_text(row)
        lt = _attendance_leave_type_text(row)
        d_str = row['譌･莉�'].strftime("%Y-%m-%d") if pd.notna(row['譌･莉�']) else ""
        if rem:
            remarks_to_analyze.append(f"{d_str}_{m} 縺ｮ蛯呵�: {rem}")
        elif lt and lt not in ("騾壼ｸｸ", ""):
            remarks_to_analyze.append(f"{d_str}_{m} 縺ｮ莨第嚊蛹ｺ蛻��ｼ亥ｙ閠�縺ｯ遨ｺ�ｼ�: {lt}")

    if remarks_to_analyze:
        remarks_blob = "\n".join(remarks_to_analyze)
        cache_key = hashlib.sha256(
            (remarks_blob + "\n" + ATTENDANCE_REMARK_AI_SCHEMA_ID).encode("utf-8")
        ).hexdigest()
        ai_cache = load_ai_cache()

        # 蜷御ｸ蛯呵�繧ｻ繝�繝医�ｯ繧ｭ繝｣繝�繧ｷ繝･繧貞━蜈亥茜逕ｨ縺励、PI繧ｳ繝ｼ繝ｫ繧堤ｯ邏�
        cached_data = get_cached_ai_result(ai_cache, cache_key)
        if cached_data is not None:
            ai_parsed = cached_data
            ai_log["蜍､諤�蛯呵ダAI_API"] = "縺ｪ縺�(繧ｭ繝｣繝�繧ｷ繝･菴ｿ逕ｨ)"
            ai_log["蜍､諤�蛯呵ダAI_隧ｳ邏ｰ"] = "繧ｭ繝｣繝�繧ｷ繝･繝偵ャ繝�"
        elif not API_KEY:
            ai_parsed = {}
            ai_log["蜍､諤�蛯呵ダAI_API"] = "縺ｪ縺�"
            ai_log["蜍､諤�蛯呵ダAI_隧ｳ邏ｰ"] = "GEMINI_API_KEY譛ｪ險ｭ螳壹�ｮ縺溘ａ蜍､諤�蛯呵アI繧偵せ繧ｭ繝�繝�"
            logging.info("GEMINI_API_KEY 譛ｪ險ｭ螳壹�ｮ縺溘ａ蛯呵アI隗｣譫舌ｒ繧ｹ繧ｭ繝�繝励＠縺ｾ縺励◆縲�")
        else:
            logging.info("笆� AI縺瑚､�謨ｰ譌･縺ｮ迚ｹ險倅ｺ矩��繧定ｧ｣譫蝉ｸｭ...")
            ai_log["蜍､諤�蛯呵ダAI_API"] = "縺ゅｊ"
            
            prompt = f"""
            莉･荳九�ｮ蜷�譌･繝ｻ繝｡繝ｳ繝舌�ｼ縺ｮ蛯呵�繧定ｪｭ縺ｿ蜿悶ｊ縲∝�ｺ騾蜍､譎ょ綾縺ｮ螟画峩繧�荳ｭ謚懊￠縲∽ｼ第律縺ｮ蛻､螳壹ｒ陦後＞縲゛SON蠖｢蠑上〒蜃ｺ蜉帙＠縺ｦ縺上□縺輔＞縲�
            繝槭�ｼ繧ｯ繝繧ｦ繝ｳ險伜捷(``` 遲�)縺ｯ荳蛻�蜷ｫ繧√★縲∫ｴ皮ｲ九↑JSON譁�蟄怜�励�ｮ縺ｿ繧定ｿ斐＠縺ｦ縺上□縺輔＞縲�

            縲辱SON縺ｮ蜃ｺ蜉帛ｽ｢蠑擾ｼ医く繝ｼ蜷阪ｒ蜴ｳ蟇�縺ｫ螳医ｋ縺薙→�ｼ峨�
            {{
              "YYYY-MM-DD_繝｡繝ｳ繝舌�ｼ蜷�": {{
                "蜃ｺ蜍､譎ょ綾": "HH:MM", 
                "騾蜍､譎ょ綾": "HH:MM", 
                "荳ｭ謚懊￠髢句ｧ�": "HH:MM",
                "荳ｭ謚懊￠邨ゆｺ�": "HH:MM",
                "菴懈･ｭ蜉ｹ邇�": 1.0,     
                "is_holiday": false,
                "驟榊床荳榊盾蜉�": false
              }}
            }}
            繝ｻ繧ｭ繝ｼ蜷阪�ｯ荳願ｨ倥�ｮ譌･譛ｬ隱槭く繝ｼ繧偵◎縺ｮ縺ｾ縺ｾ菴ｿ縺��ｼ郁恭隱槭く繝ｼ縺ｫ鄂ｮ縺肴鋤縺医↑縺��ｼ�
            繝ｻ蜃ｺ蜍､譎ょ綾/騾蜍､譎ょ綾: 蠖楢ｩｲ陦後�ｮ縲悟ｙ閠�縲阪∪縺溘�ｯ縲御ｼ第嚊蛹ｺ蛻��ｼ亥ｙ閠�縺ｯ遨ｺ�ｼ峨阪�ｮ譁�閼医°繧画耳貂ｬ縲ゆｸ肴�弱ｄ螟画峩縺ｪ縺励↑繧� null
            繝ｻ荳ｭ謚懊￠髢句ｧ�/邨ゆｺ�: 荳譎ら噪縺ｪ髮｢閼ｱ�ｼ井ｸｭ謚懊￠繝ｻ莠句漁謇繝ｻ莨夊ｭｰ縺ｪ縺ｩ�ｼ峨′縺ゅｋ蝣ｴ蜷医√◎縺ｮ髢句ｧ九�ｻ邨ゆｺ�縲ゅ↑縺�蝣ｴ蜷医�ｯ null
            繝ｻ譖匁乂隱槭�ｮ隗｣驥井ｾ�:
              - 縲悟壕蜑堺ｸｭ縺ｯ莠句漁謇縺ｧ菴懈･ｭ縲�=> 荳ｭ謚懊￠髢句ｧ� "08:45", 荳ｭ謚懊￠邨ゆｺ� "12:00"
              - 縲悟壕蠕後�ｯ莨夊ｭｰ縲�=> 荳ｭ謚懊￠髢句ｧ� "13:00", 荳ｭ謚懊￠邨ゆｺ� "17:00"
            繝ｻis_holiday: 縺昴�ｮ譌･縺御ｼ夂､ｾ縺ｫ譚･縺ｪ縺�繝ｻ邨よ律莨第嚊繝ｻ谺�蜍､縺ｪ縺ｩ **蜍､蜍呵�ｪ菴薙′縺ｪ縺�** 縺ｨ蛻､譁ｭ縺ｧ縺阪ｋ蝣ｴ蜷医�ｮ縺ｿ true縲ょ壕蜑堺ｼ代�ｻ蜊亥ｾ御ｼ代↑縺ｩ驛ｨ蛻�逧�縺ｪ莨代∩縺ｯ false�ｼ井ｸｭ謚懊￠繧�譎ょ綾縺ｧ陦ｨ迴ｾ�ｼ�
            繝ｻ驟榊床荳榊盾蜉�: 蜍､蜍吶�ｯ縺ゅｋ縺� **蜉�蟾･繝ｩ繧､繝ｳ縺ｸ縺ｮ驟榊床�ｼ�OP/AS 縺ｮ蜑ｲ蠖難ｼ峨↓霈峨○縺ｦ縺ｯ縺�縺代↑縺�** 縺ｨ隱ｭ縺ｿ蜿悶ｌ繧句�ｴ蜷医�ｯ true縲り｡ｨ險倥�ｯ蝠上ｏ縺壽э蜻ｳ縺ｧ蛻､譁ｭ縺吶ｋ縺薙→縲�
              萓�: 縲碁�榊床荳榊庄縲阪碁�榊床�ｼｮ�ｼｧ縲阪後Λ繧､繝ｳ縺ｫ荵励ｉ縺ｪ縺�縲阪梧怦谺｡轤ｹ讀懊�ｮ縺ｿ縲阪檎せ讀懊〒荳譌･縲阪御ｺ句漁縺ｮ縺ｿ縲阪梧蕗閧ｲ縺ｧ迴ｾ蝣ｴ荳榊庄縲阪梧焔驟阪↑縺励阪後い繧ｵ繧､繝ｳ荳崎ｦ√阪↑縺ｩ縺ｮ謠ｺ繧後ｄ蟀画峇陦ｨ迴ｾ繧ょ性繧縲�
              騾壼ｸｸ蜍､蜍吶〒迚ｹ縺ｫ蛻ｶ髯舌′隱ｭ縺ｿ蜿悶ｌ縺ｪ縺�蝣ｴ蜷医�ｯ false
            繝ｻ菴懈･ｭ蜉ｹ邇�: 0.0縲�1.0縺ｮ謨ｰ蛟､
            
            縲千音險倅ｺ矩��繝ｪ繧ｹ繝医�
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
                    ai_log["蜍､諤�蛯呵ダAI_隧ｳ邏ｰ"] = "隗｣譫先�仙粥"
                else:
                    ai_parsed = {}
                    ai_log["蜍､諤�蛯呵ダAI_隧ｳ邏ｰ"] = "JSON繝代�ｼ繧ｹ螟ｱ謨�"
            except Exception as e:
                err_text = str(e)
                is_quota_or_rate = ("429" in err_text) or ("RESOURCE_EXHAUSTED" in err_text)
                retry_sec = extract_retry_seconds(err_text) if is_quota_or_rate else None

                if is_quota_or_rate and retry_sec is not None:
                    wait_sec = min(max(retry_sec, 1.0), 90.0)
                    logging.warning(f"AI騾壻ｿ｡ 429/RESOURCE_EXHAUSTED縲�{wait_sec:.1f}遘貞ｾ�讖溘＠縺ｦ1蝗槭□縺大�崎ｩｦ陦後＠縺ｾ縺吶�")
                    time_module.sleep(wait_sec)
                    try:
                        res = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                        match = re.search(r'\{.*\}', res.text, re.DOTALL)
                        if match:
                            ai_parsed = json.loads(match.group(0))
                            put_cached_ai_result(ai_cache, cache_key, ai_parsed)
                            save_ai_cache(ai_cache)
                            ai_log["蜍､諤�蛯呵ダAI_隧ｳ邏ｰ"] = "蜀崎ｩｦ陦後〒隗｣譫先�仙粥"
                        else:
                            ai_parsed = {}
                            ai_log["蜍､諤�蛯呵ダAI_隧ｳ邏ｰ"] = "蜀崎ｩｦ陦悟ｾ繰SON繝代�ｼ繧ｹ螟ｱ謨�"
                    except Exception as e2:
                        ai_parsed = {}
                        logging.warning(f"AI蜀崎ｩｦ陦後お繝ｩ繝ｼ: {e2}")
                        ai_log["蜍､諤�蛯呵ダAI_隧ｳ邏ｰ"] = f"429蠕悟�崎ｩｦ陦悟､ｱ謨�: {e2}"
                else:
                    ai_parsed = {}
                    logging.warning(f"AI騾壻ｿ｡繧ｨ繝ｩ繝ｼ: {e}")
                    ai_log["蜍､諤�蛯呵ダAI_隧ｳ邏ｰ"] = str(e)
    else:
        ai_parsed = {}

    # 3. 譌･莉倥＃縺ｨ縺ｮ蛻ｶ邏�霎樊嶌繧呈ｧ狗ｯ�
    for _, row in df.iterrows():
        if pd.isna(row['譌･莉�']): continue
        curr_date = row['譌･莉�']
        m = str(row.get('繝｡繝ｳ繝舌�ｼ', '')).strip()
        if m not in members: continue

        if curr_date not in attendance_data:
            attendance_data[curr_date] = {}

        original_reason = _attendance_remark_text(row)
        leave_type = _attendance_leave_type_text(row)

        key = f"{curr_date.strftime('%Y-%m-%d')}_{m}"
        ai_info = ai_parsed.get(key, {})

        is_empty_shift = pd.isna(row.get('蜃ｺ蜍､譎る俣')) and pd.isna(row.get('騾蜍､譎る俣')) and not ai_info
        is_holiday = _ai_json_bool(ai_info.get("is_holiday"), False) or is_empty_shift
        exclude_from_line = _ai_json_bool(ai_info.get("驟榊床荳榊盾蜉�"), False)

        ai_eff = ai_info.get("菴懈･ｭ蜉ｹ邇�")
        excel_eff = row.get('菴懈･ｭ蜉ｹ邇�')
        
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
                and leave_type not in ("騾壼ｸｸ", "")
                and leave_type not in original_reason
            ):
                reason = f"{leave_type} {original_reason}"
            else:
                reason = original_reason
        elif leave_type and leave_type not in ("騾壼ｸｸ", ""):
            reason = leave_type
        else:
            reason = '騾壼ｸｸ' if not is_empty_shift else '莨第律繧ｷ繝輔ヨ'

        # 繝槭せ繧ｿ縺ｫ蜃ｺ蜍､繝ｻ騾蜍､縺ｮ荳｡譁ｹ縺悟�･縺｣縺ｦ縺�繧区律縺ｯ縲∝共諤�AI縺ｮ蜃ｺ蜍､/騾蜍､譎ょ綾縺ｧ荳頑嶌縺阪＠縺ｪ縺��ｼ井ｼ第嚊蛹ｺ蛻�縺ｮ縺ｿ縺ｮ陦後〒隱､謗ｨ螳壹＆繧後≧繧具ｼ�
        excel_s = row.get("蜃ｺ蜍､譎る俣")
        excel_e = row.get("騾蜍､譎る俣")
        if not pd.isna(excel_s) and not pd.isna(excel_e):
            start_t = parse_time_str(excel_s, DEFAULT_START_TIME)
            end_t = parse_time_str(excel_e, DEFAULT_END_TIME)
        else:
            start_t = parse_time_str(ai_info.get("蜃ｺ蜍､譎ょ綾") or excel_s, DEFAULT_START_TIME)
            end_t = parse_time_str(ai_info.get("騾蜍､譎ょ綾") or excel_e, DEFAULT_END_TIME)
        base_end_t = end_t

        b1_s = parse_time_str(row.get('莨第�ｩ譎る俣1_髢句ｧ�'), DEFAULT_BREAKS[0][0])
        b1_e = parse_time_str(row.get('莨第�ｩ譎る俣1_邨ゆｺ�'), DEFAULT_BREAKS[0][1])
        b2_s = parse_time_str(row.get('莨第�ｩ譎る俣2_髢句ｧ�'), DEFAULT_BREAKS[1][0])
        b2_e = parse_time_str(row.get('莨第�ｩ譎る俣2_邨ゆｺ�'), DEFAULT_BREAKS[1][1])

        # 笘�霑ｽ蜉�: AI縺九ｉ荳ｭ謚懊￠譎る俣繧貞叙蠕�
        mid_break_s = parse_time_str(ai_info.get("荳ｭ謚懊￠髢句ｧ�"), None)
        mid_break_e = parse_time_str(ai_info.get("荳ｭ謚懊￠邨ゆｺ�"), None)
        # AI縺御ｸｭ謚懊￠繧定ｿ斐＆縺ｪ縺九▲縺溷�ｴ蜷医�ｯ縲∝ｙ閠�譁�險縺九ｉ繝ｭ繝ｼ繧ｫ繝ｫ謗ｨ螳壹〒陬懷ｮ�
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
                "蜍､諤� %s %s: 谿区･ｭ邨よ･ｭ驕ｩ逕ｨ蠕後↓騾蜍､縺悟�ｺ蜍､莉･蜑阪→縺ｪ縺｣縺溘◆繧√∵ｮ区･ｭ邨よ･ｭ繧堤┌隕悶＠縺ｦ螳壽凾騾蜍､縺ｫ謌ｻ縺励∪縺吶�",
                curr_date,
                m,
            )
            end_t = base_end_t
            end_dt = combine_dt(end_t)
        breaks_dt = []
        
        # 騾壼ｸｸ縺ｮ莨第�ｩ繧定ｿｽ蜉�
        if b1_s and b1_e: breaks_dt.append((combine_dt(b1_s), combine_dt(b1_e)))
        if b2_s and b2_e: breaks_dt.append((combine_dt(b2_s), combine_dt(b2_e)))
        
        # 笘�霑ｽ蜉�: 荳ｭ謚懊￠譎る俣縺後≠繧句�ｴ蜷医�ｯ縲∫音蛻･縺ｪ縲御ｼ第�ｩ縲阪→縺励※繧ｹ繧ｱ繧ｸ繝･繝ｼ繝ｫ險育ｮ励↓霑ｽ蜉�
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
# 蜈ｨ萓晞�ｼ蜈ｱ騾�: 蜉�蟾･蜀�螳ｹ蛻励�ｮ蟾･遞矩��蠎� / 蛟句挨: EC竊呈､懈渊繝ｭ繝ｼ繝ｫ繝代う繝励Λ繧､繝ｳ
# ---------------------------------------------------------------------------
ROLL_PIPELINE_EC_PROCESS = "EC"
ROLL_PIPELINE_EC_MACHINE = "EC讖溘貉門漉"
ROLL_PIPELINE_INSP_PROCESS = "讀懈渊"
ROLL_PIPELINE_INSP_MACHINE = "辭ｱ陞咲捩讖溘貉門漉"
# ﾂｧB-3: 蠕檎ｶ壹�ｯ B-2 縺ｮ縲梧､懈渊縲阪↓逶ｸ蠖薙☆繧句ｷ･遞九→縺励※蟾ｻ霑斐＠�ｼ亥酔荳萓晞�ｼ縺ｧ EC 蜈郁｡後�ｻ繝ｭ繝ｼ繝ｫ譫�繝ｻ繝ｪ繝ｯ繧､繝ｳ繝臥ｭ峨�ｯ B-2 縺ｨ蜷瑚ｶ｣譌ｨ�ｼ�
ROLL_PIPELINE_REWIND_PROCESS = "蟾ｻ霑斐＠"
ROLL_PIPELINE_REWIND_MACHINE = "EC讖溘貉門漉"
ROLL_PIPELINE_INITIAL_BUFFER_ROLLS = 2
# 讀懈渊縺ｮ蜑ｲ蠖謎ｸ企剞 min 縺ｫ菴ｿ縺�縲ょ酔荳萓晞�ｼ縺ｫ EC 陦後′辟｡縺�縺ｨ縺阪�ｯ need繝ｻ繧ｹ繧ｭ繝ｫ縺ｫ蠕薙＞騾壼ｸｸ驟榊床縺吶ｋ�ｼ�ec_done=0 蝗ｺ螳壹〒豌ｸ荵�繧ｹ繧ｭ繝�繝励＠縺ｪ縺��ｼ峨�
ROLL_PIPELINE_INSP_UNCAPPED_ROOM = 1.0e18


# 蜍､諤�縺ｫ霈峨▲縺ｦ縺�繧区怙邨よ律縺ｾ縺ｧ縺ｧ蜑ｲ莉倥′邨ゅｏ繧峨↑縺�縺ｨ縺阪∵怙邨よ律縺ｨ蜷後§繧ｷ繝輔ヨ蝙九〒譌･莉倥ｒ蟒ｶ髟ｷ縺吶ｋ�ｼ医が繝励す繝ｧ繝ｳ�ｼ峨�
# False 縺ｮ縺ｨ縺肴ｮｵ髫�2縺ｯ繝槭せ繧ｿ蜍､諤�縺ｮ譌･莉倡ｯ�蝗ｲ縺ｮ縺ｿ縺ｧ蜑ｲ莉倥＠縲∵ｮ九ｊ縺ｯ驟榊床谿九�ｻ驟榊床荳榊庄縺ｮ縺ｾ縺ｾ縺ｨ縺吶ｋ縲�
STAGE2_EXTEND_ATTENDANCE_CALENDAR = False
SCHEDULE_EXTEND_MAX_EXTRA_DAYS = 366

# 邏肴悄蝓ｺ貅匁律繧帝℃縺弱※繧ょｽ楢ｩｲ萓晞�ｼ縺ｫ谿矩㍼縺後≠繧九→縺阪�**縺昴�ｮ萓晞�ｼNO縺�縺�** due_basis 繧� +1 縺励�
# 蠖楢ｩｲ萓晞�ｼ縺ｮ蜑ｲ蠖薙�ｻ繧ｿ繧､繝�繝ｩ繧､繝ｳ繧貞ｷｻ縺肴綾縺励※**繧ｫ繝ｬ繝ｳ繝繝ｼ蜈磯�ｭ縺九ｉ**蜀阪す繝溘Η繝ｬ繝ｼ繧ｷ繝ｧ繝ｳ縺吶ｋ�ｼ井ｻ紋ｾ晞�ｼ縺ｮ蜑ｲ蠖薙�ｯ邯ｭ謖��ｼ峨�
# 繝槭せ繧ｿ蜍､諤�縺ｮ譛邨よ律繧定ｶ�縺医※蠕後ｍ蛟偵＠縺ｧ縺阪↑縺�萓晞�ｼ縺ｯ縲碁�榊床谿�(蜍､蜍吶き繝ｬ繝ｳ繝繝ｼ荳崎ｶｳ)縲阪→縺吶ｋ縲ょ推蜀崎ｩｦ陦悟燕縺ｫ蜍､諤�諡｡蠑ｵ蛻�縺ｯ繝槭せ繧ｿ譌･莉倥∈謌ｻ縺吶�
# 譌｢螳� **False**�ｼ磯�榊床隧ｦ陦碁��繧呈ｭ｣縺ｨ縺励∬ｨ育判蝓ｺ貅冶ｶ�驕弱〒繧ゅ％縺ｮ蟾ｻ縺肴綾縺怜�崎ｩｦ陦後�ｯ陦後ｏ縺ｪ縺��ｼ峨ょｾ捺擂謖吝虚縺悟ｿ�隕√↑縺ｨ縺阪□縺� True縲�
STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING = False
# 邏肴悄蝓ｺ貅悶�ｮ +1 譌･縺ｫ繧医ｋ蟾ｻ縺肴綾縺怜�阪す繝溘Η縺ｯ萓晞�ｼNO縺斐→縺ｫ譛螟ｧ縺薙�ｮ蝗樊焚�ｼ�6 蝗樒岼莉･髯阪�ｯ蠖楢ｩｲ萓晞�ｼ縺ｮ縺ｿ繧ｷ繝輔ヨ縺帙★縲∵悴螳御ｺ�陦後↓邏肴悄隕狗峩縺怜ｿ�隕√ｒ莉倅ｸ弱＠蠕励ｋ�ｼ峨�
STAGE2_RETRY_SHIFT_DUE_MAX_ROUNDS = 5

# True 縺ｮ縺ｨ縺阪�驟榊床險育判繧ｷ繝ｼ繝医�ｮ隱ｭ縺ｿ霎ｼ縺ｿ陦碁���ｼ亥推萓晞�ｼNO縺ｮ蛻晏�ｺ陦後′譌ｩ縺�縺ｻ縺ｩ蜈茨ｼ峨〒 1 萓晞�ｼ縺�縺代ｒ
# 蠖捺律蛟呵｣懊↓谿九＠縲∝ｮ瑚ｵｰ縺励※縺九ｉ谺｡萓晞�ｼ縺ｸ騾ｲ繧縲�**莉紋ｾ晞�ｼ縺ｯ荳蛻�縺昴�ｮ譌･驟榊床縺輔ｌ縺ｪ縺�**縺溘ａ縲�
# 繧｢繧ｯ繝�繧｣繝紋ｾ晞�ｼ縺ｮ1陦後〒繧りｩｰ縺ｾ繧九→蜈ｨ菴薙′驟榊床荳榊庄縺ｫ隕九∴繧具ｼ医Ο繧ｰ縲御ｾ晞�ｼNO逶ｴ蛻鈴�榊床 逶ｴ蛻怜ｾ�=1縲搾ｼ峨�
# 譌｢螳� False縲ょ宍蟇�縺ｪ萓晞�ｼNO逶ｴ蛻励′蠢�隕√↑縺ｨ縺阪□縺� STAGE2_SERIAL_DISPATCH_BY_TASK_ID=1 繧定ｨｭ螳壹☆繧九�
STAGE2_SERIAL_DISPATCH_BY_TASK_ID = (
    os.environ.get("STAGE2_SERIAL_DISPATCH_BY_TASK_ID", "0")
    .strip()
    .lower()
    in ("1", "true", "yes", "on", "縺ｯ縺�")
)

# True: 竭�谿九ち繧ｹ繧ｯ縺ｮ縺�縺｡驟榊床隧ｦ陦碁��縺梧怙蟆上�ｮ1繧ｿ繧ｹ繧ｯ縺�縺代ｒ驕ｸ縺ｳ縲�1繝ｭ繝ｼ繝ｫ縺壹▽蜑ｲ莉倥�
# 竭｡蜴溷渚謚募�･譌･縺ｨ蜷御ｸ譌･縺ｫ髢句ｧ九☆繧句�ｴ蜷医�ｯ 13:00 莉･髯搾ｼ�same_day_raw_start_limit 繧� 13:00�ｼ峨�
# 竭｢竭｣險ｭ蛯咏ｩｺ縺阪ｒ max 縺ｧ郢ｰ繧贋ｸ翫￡�ｼ域律蜀�縲らｿ梧律縺ｯ譌･莉倥Ν繝ｼ繝励〒繧ｿ繧､繝�繝ｩ繧､繝ｳ繧ｷ繝ｼ繝会ｼ峨�
# 竭､竭･竭ｦ竭ｧ莠ｺ縺ｮ遨ｺ縺阪〒繝√�ｼ繝�繧呈ｱｺ繧√√Ο繝ｼ繝ｫ縺斐→縺ｫ avail 繧呈峩譁ｰ�ｼ亥酔譌･縺ｯ蜑阪Ο繝ｼ繝ｫ縺ｨ蜷御ｸ繝√�ｼ繝�繧貞━蜈茨ｼ峨�
# 辟｡蜉ｹ蛹�: 迺ｰ蠅�螟画焚 STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST=0
STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST = os.environ.get(
    "STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST", "1"
).strip().lower() not in ("0", "false", "no", "off", "縺�縺�縺�", "辟｡蜉ｹ")

# True�ｼ域里螳夲ｼ�: start_date_req<=蠖捺律 縺九▽谿九≠繧翫�ｮ繧ｿ繧ｹ繧ｯ縺ｮ縺�縺｡縲�驟榊床隧ｦ陦碁��縺ｮ譛蟆上梧棧縲阪□縺代′蜑ｲ莉伜ｯｾ雎｡縲�
# 繧医ｊ螟ｧ縺阪＞隧ｦ陦碁��縺ｯ縲√ｈ繧雁ｰ上＆縺�隧ｦ陦碁��縺ｫ譛ｪ螳御ｺ�縺梧ｮ九ｋ髯舌ｊ繝悶Ο繝�繧ｯ�ｼ育ｴ肴悄縺瑚ｿ代￥縺ｦ繧ょ牡繧願ｾｼ縺ｾ縺ｪ縺��ｼ峨�
STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT = os.environ.get(
    "STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT", "1"
).strip().lower() not in ("0", "false", "no", "off", "縺�縺�縺�", "辟｡蜉ｹ")

# True�ｼ域里螳夲ｼ�: 蜑ｲ莉伜呵｣懊ｒ縲瑚ｨｭ蛯吶�ｻ莠ｺ縺ｮ螢∵凾險亥頃譛牙玄髢薙阪〒莠碁㍾讀懈渊縺励√ち繧､繝�繝ｩ繧､繝ｳ霑ｽ險倥→蜷梧悄逋ｻ骭ｲ縺吶ｋ
# �ｼ医ヶ繝ｭ繝�繧ｯ繝�繝ｼ繝悶Ν縺ｨ蜷瑚ｶ｣譌ｨ縲�Excel 繧ｻ繝ｫ騾先ｬ｡ I/O 縺ｯ陦後ｏ縺ｪ縺��ｼ峨�
# False: 蠕捺擂縺ｩ縺翫ｊ avail_dt / machine_avail_dt 縺ｮ縺ｿ縲�
DISPATCH_INTERVAL_MIRROR_ENFORCE = os.environ.get(
    "DISPATCH_INTERVAL_MIRROR_ENFORCE", "1"
).strip().lower() not in ("0", "false", "no", "off", "縺�縺�縺�", "辟｡蜉ｹ")


def _clone_attendance_day_shifted(source_day: dict, old_date: date, new_date: date) -> dict:
    """繝｡繝ｳ繝舌�ｼ蛻･蜍､諤�繝悶Ο繝�繧ｯ繧� new_date 縺ｫ繧ｷ繝輔ヨ縺励◆豬�縺�繧ｳ繝斐�ｼ繧定ｿ斐☆縲�"""
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
    """驟榊床蜿ｯ閭ｽ縺ｪ繝｡繝ｳ繝舌�ｼ縺�1莠ｺ縺ｧ繧ゅ＞繧狗峩霑代�ｮ譌･繧偵ユ繝ｳ繝励Ξ縺ｫ謗｡逕ｨ�ｼ域怙邨よ律縺悟�ｨ莨代〒繧よ怏蜉ｹ縺ｪ蝙九ｒ菴ｿ縺��ｼ峨�"""
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
    """繧ｫ繝ｬ繝ｳ繝繝ｼ荳�1譌･蜈医ｒ plan_dates 縺ｫ霑ｽ蜉�縺励√ユ繝ｳ繝励Ξ譌･縺ｮ繧ｷ繝輔ヨ隍�陬ｽ縺ｧ attendance 繧貞沂繧√ｋ縲ょ､ｱ謨玲凾 False縲�"""
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
        "驟榊床螳御ｺ�縺ｾ縺ｧ蜍､諤�繧定�ｪ蜍墓僑蠑ｵ: %s 繧定ｿｽ蜉��ｼ医ユ繝ｳ繝励Ξ=%s縲√Γ繝ｳ繝舌�ｼ謨ｰ=%s�ｼ�",
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
    plan_dates 繧貞�磯�ｭ縺九ｉ鬆�縺ｫ yield縲よ忰蟆ｾ縺ｾ縺ｧ譚･縺ｦ繧よｮ九ち繧ｹ繧ｯ縺後≠繧後�ｰ蜍､諤�繧�1譌･縺壹▽諡｡蠑ｵ縺励※邯咏ｶ壹�
    plan_dates / attendance_data 縺ｯ繧､繝ｳ繝励Ξ繝ｼ繧ｹ譖ｴ譁ｰ縺輔ｌ繧九�
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
                "谿九ち繧ｹ繧ｯ縺後≠繧翫∪縺吶′蜍､諤�縺ｮ閾ｪ蜍墓僑蠑ｵ縺御ｸ企剞�ｼ�%s 譌･�ｼ峨↓驕斐＠縺ｾ縺励◆縲る�榊床谿九�ｻ驟榊床荳榊庄縺梧ｮ九ｋ蜿ｯ閭ｽ諤ｧ縺後≠繧翫∪縺吶�",
                SCHEDULE_EXTEND_MAX_EXTRA_DAYS,
            )
            return
        if not _extend_attendance_one_calendar_day(attendance_data, plan_dates):
            logging.warning(
                "蜍､諤�繧�1譌･諡｡蠑ｵ縺ｧ縺阪∪縺帙ｓ縺ｧ縺励◆�ｼ医ユ繝ｳ繝励Ξ譌･縺ｮ繝�繝ｼ繧ｿ谺�關ｽ�ｼ峨よｮ九ち繧ｹ繧ｯ縺ｯ譛ｪ蜑ｲ蠖薙�ｮ縺ｾ縺ｾ縺ｧ縺吶�"
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
    """萓晞�ｼNO 竊� 蜉�蟾･蜀�螳ｹ縺ｮ蟾･遞句錐繝ｪ繧ｹ繝茨ｼ郁｡ｨ縺ｮ荳翫�ｮ譁ｹ縺ｧ譛蛻昴↓迴ｾ繧後◆髱樒ｩｺ縺ｮ陦後ｒ謗｡逕ｨ�ｼ峨�"""
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
    蟾･遞句錐�ｼ磯�榊床險育判縺ｮ縲悟ｷ･遞句錐縲榊�暦ｼ峨′縲∝��繝�繝ｼ繧ｿ縺ｮ縲悟刈蟾･蜀�螳ｹ縲阪き繝ｳ繝槫玄蛻�繧翫ヨ繝ｼ繧ｯ繝ｳ縺ｮ縺�縺壹ｌ縺九→
    豁｣隕丞喧荳閾ｴ縺吶ｋ縺九ゅヨ繝ｼ繧ｯ繝ｳ縺檎┌縺��ｼ亥刈蟾･蜀�螳ｹ譛ｪ險伜�･縺ｮ萓晞�ｼ�ｼ峨�ｯ辣ｧ蜷亥ｯｾ雎｡螟悶→縺励※ True縲�
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
    """驟榊床險育判繧ｷ繝ｼ繝医�ｮ縲悟━蜈亥ｺｦ縲阪ょｰ上＆縺�縺ｻ縺ｩ蜈医よ悴蜈･蜉帙�ｻ荳肴ｭ｣縺ｯ 999縲�"""
    p = t.get("priority", 999)
    try:
        return int(p)
    except (TypeError, ValueError):
        return 999


def _task_blocked_by_same_request_dependency(task, task_queue) -> bool:
    """
    蜷御ｸ萓晞�ｼNO縺ｮ逡ｰ縺ｪ繧句ｷ･遞九ｒ蜷梧凾蛻ｻ縺ｫ蝗槭＆縺ｪ縺��ｼ磯�榊床繝ｫ繝ｼ繝ｫ ﾂｧA-1繝ｻﾂｧA-2�ｼ峨�
    - 荳｡陦後↓蜉�蟾･蜀�螳ｹ逕ｱ譚･縺ｮ rank 縺後≠繧九→縺阪�ｯ rank 縺ｮ縺ｿ縺ｧ蜑榊ｾ鯉ｼ按ｧA-1�ｼ峨�
    - 縺ｩ縺｡繧峨°縺ｫ rank 縺檎┌縺�縺ｨ縺阪�ｯ縲�驟榊床險育判繧ｷ繝ｼ繝医�ｮ陦碁�� same_request_line_seq 縺ｧ蜑榊ｾ鯉ｼ按ｧA-2�ｼ峨�
    ﾂｧB-2 / ﾂｧB-3: ``roll_pipeline_inspection`` 縺ｾ縺溘�ｯ ``roll_pipeline_rewind`` 陦後′
    ``roll_pipeline_ec`` 蜈郁｡後↓繧医ｊ ﾂｧA-1 縺ｧ豁｢縺ｾ繧句�ｴ蜷医�
    ``_roll_pipeline_inspection_assign_room`` > 0 縺ｪ繧牙ｽ楢ｩｲ繝壹い縺�縺代ヶ繝ｭ繝�繧ｯ縺励↑縺�縲�
    蜑埼ｲ驟榊床縺ｧ縺ｯ ``_trial_order_flow_eligible_tasks`` 縺� EC 螳瑚ｵｰ縺ｾ縺ｧ讀懈渊繧貞､悶☆縺溘ａ縲�
    EC 谿九′縺ゅｋ髢薙�ｯ譛ｬ蛻�蟯舌↓蛻ｰ驕斐＠縺ｪ縺�縲ゅΜ繝ｯ繧､繝ｳ繝臥ｭ峨〒讀懈渊縺瑚ｼ峨ｋ螻髱｢縺ｨ縺ｮ謨ｴ蜷育畑縲�
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
    繧ｭ繝･繝ｼ迥ｶ諷倶ｸ翫√％縺ｮ陦後�ｯ縺ｾ縺�譌･谺｡驟榊床縺ｧ騾ｲ繧√ｉ繧後↑縺��ｼ按ｧA 蜷御ｸ萓晞�ｼ縺ｮ蜑榊ｷ･遞区ｮ九√∪縺溘�ｯ ﾂｧB-2/ﾂｧB-3 縺ｮ譫�繧ｼ繝ｭ�ｼ峨�
    `_min_pending_dispatch_trial_order_for_date` 縺ｨ `_equipment_line_lower_dispatch_trial_still_pending`
    縺ｧ蜷後§蝓ｺ貅悶ｒ蜈ｱ譛峨☆繧九ら援譁ｹ縺�縺醍峩縺吶→縲∝酔荳險ｭ蛯吶く繝ｼ縺ｧ蜈ｨ莉ｶ譛ｪ蜑ｲ蠖薙′谿九ｋ繝�繝�繝峨Ο繝�繧ｯ縺瑚ｵｷ縺榊ｾ励ｋ縲�
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
    """辭ｱ陞咲捩讀懈渊陦後�ｮ縺ｿ縺ｮ邏ｯ險亥ｮ御ｺ�繝ｭ繝ｼ繝ｫ�ｼ医ヨ繝ｬ繝ｼ繧ｹ逕ｨ�ｼ峨�"""
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
    """ﾂｧB-2 讀懈渊陦鯉ｼ仰ｧB-3 蟾ｻ霑斐＠陦後�ｮ縲∝酔荳萓晞�ｼ蜀�縺ｮ蠕檎ｶ壹ヱ繧､繝励Λ繧､繝ｳ邏ｯ險亥ｮ御ｺ�繝ｭ繝ｼ繝ｫ縲�"""
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
    """蜷御ｸ萓晞�ｼNO縺ｫ EC�ｼ医Ο繝ｼ繝ｫ繝代う繝励Λ繧､繝ｳ蜈郁｡鯉ｼ峨ち繧ｹ繧ｯ縺後く繝･繝ｼ縺ｫ蜷ｫ縺ｾ繧後ｋ縺九�"""
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
    """蜷御ｸ萓晞�ｼNO縺ｮ EC 繝ｭ繝ｼ繝ｫ繝代う繝励Λ繧､繝ｳ陦後′縺吶∋縺ｦ谿矩㍼繧ｼ繝ｭ�ｼ亥ｮ瑚ｵｰ�ｼ峨°縲�"""
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
    # EC 蜈ｨ繝ｭ繝ｼ繝ｫ螳御ｺ�蠕後�ｯ縲窪C 蜈郁｡後�ｻ繝舌ャ繝輔ぃ縲阪�ｯ譌｢縺ｫ貅縺溘＆繧後※縺�繧九ゅ％縺薙〒 max_insp 繧� ec_done 縺ｫ
    # 謠�縺医ｋ縺ｨ縲√す繝ｼ繝井ｸ翫�ｮ讀懈渊�ｼ医�ｻ蟾ｻ霑斐＠�ｼ画ｮ九Ο繝ｼ繝ｫ謨ｰ縺� EC 螳御ｺ�繝ｭ繝ｼ繝ｫ謨ｰ繧剃ｸ雁屓繧九ョ繝ｼ繧ｿ縺ｧ
    # max_insp - insp_done 縺� 0 縺ｮ縺ｾ縺ｾ谿九ｊ縲∵､懈渊陦後′ eligible 縺九ｉ螟悶ｌ驟榊床隧ｦ陦碁��縺梧ｰｸ荵�縺ｫ隧ｰ縺ｾ繧�
    # �ｼ亥�咲樟繝ｭ繧ｰ: ec_fully_done 縺九▽ insp_done==max_insp==ec_done 縺ｧ room=0 竊� 蠕檎ｶ夊ｩｦ陦碁��縺碁�榊床荳榊庄�ｼ峨�
    if _pipeline_ec_fully_done_for_tid(task_queue, task_id):
        return float(ROLL_PIPELINE_INSP_UNCAPPED_ROOM)
    # EC 遞ｼ蜒堺ｸｭ: 蜈郁｡後ヰ繝�繝輔ぃ B 縺ｫ繧医ｊ讀懈渊繝ｭ繝ｼ繝ｫ荳企剞繧� ec_done 縺九ｉ驕�蟒ｶ縺輔○繧具ｼ�B=2 縺ｮ蠑上�ｯ繧ｳ繝｡繝ｳ繝亥盾辣ｧ�ｼ峨�
    max_insp = max(0.0, ec_done - float(ROLL_PIPELINE_INITIAL_BUFFER_ROLLS) + 1.0)
    _room = max(0.0, max_insp - insp_done)
    return _room


def _roll_pipeline_inspection_task_row_for_tid(
    task_queue: list, task_id: str
) -> dict | None:
    """蜷御ｸ萓晞�ｼNO縺ｮ ﾂｧB-2 讀懈渊陦後∪縺溘�ｯ ﾂｧB-3 蟾ｻ霑斐＠陦後ｒ1莉ｶ霑斐☆縲ら┌縺代ｌ縺ｰ None縲�"""
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
    """蜷御ｸ萓晞�ｼ縺ｮ EC 繝ｭ繝ｼ繝ｫ遒ｺ螳壹＃縺ｨ縺ｮ邨ゆｺ�譎ょ綾繧呈凾邉ｻ蛻励〒霑斐☆�ｼ�assigned_history 縺ｮ end_dt�ｼ峨�"""
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
    谺｡縺ｮ讀懈渊繝ｭ繝ｼ繝ｫ繧帝幕蟋九＠縺ｦ繧医＞譛譌ｩ譎ょ綾縲�
    邏ｯ險域､懈渊螳御ｺ�繝ｭ繝ｼ繝ｫ謨ｰ繧� K縲√ヰ繝�繝輔ぃ繧� B�ｼ�=ROLL_PIPELINE_INITIAL_BUFFER_ROLLS�ｼ峨→縺吶ｋ縺ｨ縲�
    EC 螳御ｺ�繝ｭ繝ｼ繝ｫ縺梧凾邉ｻ蛻励〒 (K+B) 譛ｬ逶ｮ縺ｫ蛻ｰ驕斐＠縺滓凾蛻ｻ�ｼ医◎縺ｮ繝ｭ繝ｼ繝ｫ縺ｮ end_dt�ｼ画悴貅縺ｫ縺ｯ髢句ｧ九＠縺ｪ縺�縲�
    �ｼ域･ｭ蜍吶Ν繝ｼ繝ｫ: 莉ｻ諢上�ｮ譎らせ縺ｧ EC_RollEndCount - KENSA_RollEndCount >= B 繧呈ｺ縺溘☆縺ｾ縺ｧ讀懈渊繧帝ｲ繧√↑縺�縲�
    縺ｮ縲後Ο繝ｼ繝ｫ邨ゆｺ�譎ょ綾蝓ｺ貅悶阪�ｮ螳溯｣�縲ゑｼ�
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
    """assigned_history 縺ｮ team 譁�蟄怜�暦ｼ井ｸｻ繝ｻ陬懊ｒ縲�,縲阪後√榊玄蛻�繧奇ｼ峨°繧画球蠖楢�蜷阪ｒ謚ｽ蜃ｺ�ｼ�NFKC�ｼ峨�"""
    if team_cell is None:
        return set()
    s = str(team_cell).strip()
    if not s:
        return set()
    out: set[str] = set()
    for part in re.split(r"[,縲‐", s):
        t = part.strip()
        if t:
            out.add(unicodedata.normalize("NFKC", t))
    return out


def _pipeline_b2_assigned_member_names_nfkc_for_side(
    task_queue: list, task_id: str, *, ec_side: bool
) -> set[str]:
    """蜷御ｸ萓晞�ｼ縺ｮ EC 陦後∪縺溘�ｯ讀懈渊陦後�ｮ assigned_history 縺ｫ蜃ｺ縺滓球蠖楢�蜷搾ｼ�NFKC 髮�蜷茨ｼ峨�"""
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
    """蜷御ｸ萓晞�ｼNO縺ｫ ﾂｧB-2/ﾂｧB-3 縺ｮ EC 陦後→蠕檎ｶ夊｡鯉ｼ域､懈渊縺ｾ縺溘�ｯ蟾ｻ霑斐＠�ｼ峨�ｮ荳｡譁ｹ縺後く繝･繝ｼ縺ｫ縺ゅｋ縺九�"""
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
    ﾂｧB-2 / ﾂｧB-3 蜷御ｸ萓晞�ｼ縺ｧ縺ｯ縲・C 陦後↓荳蠎ｦ縺ｧ繧ょ�･縺｣縺溯�縺ｯ蠕檎ｶ夲ｼ域､懈渊�ｼ丞ｷｻ霑斐＠�ｼ峨�ｮ蛟呵｣懊°繧牙､悶＠縲�
    蠕檎ｶ壹↓蜈･縺｣縺溯�縺ｯ EC 縺ｮ蛟呵｣懊°繧牙､悶☆縲�
    �ｼ育､ｾ蜀�繝ｫ繝ｼ繝ｫ: 諡�蠖楢�髮�蜷医ｒ蠢�縺壼��縺代ｋ縲ＡPLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS` 縺ｧ辟｡蜉ｹ蛹門庄�ｼ�
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
            _side = "蟾ｻ霑斐＠"
        else:
            _side = "讀懈渊"
        _log_dispatch_trace_schedule(
            tid,
            "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繝悶Ο繝�繧ｯ蛻､螳�: B-2諡�蠖楢�蛻�髮｢ side=%s machine=%s "
            "蛟呵｣憺勁螟�=%s 谿句呵｣�=%s(%s)",
            tid,
            _side,
            task.get("machine"),
            ",".join(str(x) for x in removed),
            len(filtered),
            ",".join(str(x) for x in filtered) if filtered else "縺ｪ縺�",
        )
    return filtered


def _exclusive_b1_inspection_holder_for_machine(task_queue, occupant_key: str):
    """
    蜷御ｸ迚ｩ逅�讖滓｢ｰ�ｼ域ｩ滓｢ｰ蜷阪�吶�ｼ繧ｹ縺ｮ蜊�譛峨く繝ｼ�ｼ我ｸ翫〒縲�ﾂｧB-2 辭ｱ陞咲捩讀懈渊縺ｾ縺溘�ｯ ﾂｧB-3 蟾ｻ霑斐＠縺� **譌｢縺ｫ蜑ｲ莉倥ｒ髢句ｧ�** 縺玲ｮ九Ο繝ｼ繝ｫ縺梧ｮ九ｋ陦後′縺ゅｌ縺ｰ
    縺昴�ｮ繧ｿ繧ｹ繧ｯ dict 繧�1莉ｶ霑斐☆�ｼ医↑縺代ｌ縺ｰ None�ｼ峨�

    繝代う繝励Λ繧､繝ｳ譫�縺ｧ蠕檎ｶ壹ｒ謨ｰ繝ｭ繝ｼ繝ｫ縺壹▽縺励°蜈･繧後↑縺�險ｭ險医�ｮ縺溘ａ縲∵棧繧ｼ繝ｭ縺ｮ髫咎俣縺ｫ **蛻･萓晞�ｼ** 縺悟酔縺倩ｨｭ蛯吶↓蜈･繧翫�
    邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ縺ｧ繧ｿ繧ｹ繧ｯ陦ｨ遉ｺ縺碁比ｸｭ縺ｧ蛻�繧頑崛繧上ｋ莠玖ｱ｡繧帝亟縺舌ょ頃譛我ｸｭ縺ｯ蠖楢ｩｲ迚ｩ逅�讖滓｢ｰ縺ｧ縺ｯ莉悶ち繧ｹ繧ｯ繧定ｩｦ陦後＠縺ｪ縺�縲�
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
    """need 繧ｷ繝ｼ繝医〒蟾ｦ縺ｫ縺ゅｋ縲悟ｷ･遞句錐+讖滓｢ｰ蜷阪榊�励⊇縺ｩ蟆上＆縺�蛟､�ｼ医く繝･繝ｼ縺ｧ蜈茨ｼ峨�"""
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
    generate_plan 蜀帝�ｭ縺翫ｈ縺ｳ邏肴悄繧ｷ繝輔ヨ蜀崎ｩｦ陦梧凾縺ｮ task_queue.sort 逕ｨ繧ｭ繝ｼ縲�

    1. 蜉�蟾･騾比ｸｭ�ｼ�in_progress�ｼ峨ｒ蜈�
    2. 邏肴悄蝓ｺ貅� due_basis_date�ｼ亥屓遲皮ｴ肴悄竊呈欠螳夂ｴ肴悄縲よ掠縺�縺ｻ縺ｩ蜈茨ｼ�
    3. ﾂｧB-1 竊� ﾂｧB-2/ﾂｧB-3 蟶ｯ 竊� 縺昴�ｮ莉厄ｼ�b_tier�ｼ�
    4. ﾂｧB-2/ﾂｧB-3 蟶ｯ蜀�縺ｮ縺ｿ EC 繧呈悴逹謇九�ｮ讀懈渊�ｼ丞ｷｻ霑斐＠繧医ｊ蜈茨ｼ�b2_queue_sub�ｼ�
    5. need 繧ｷ繝ｼ繝亥ｷｦ蛻励⊇縺ｩ蜈茨ｼ亥ｷ･遞句錐+讖滓｢ｰ蜷榊�励�ｮ菴咲ｽｮ�ｼ�
    6. 萓晞�ｼNO繧ｿ繧､繝悶Ξ繝ｼ繧ｯ�ｼ�_task_id_same_machine_due_tiebreak_key�ｼ�

    _req_map / _need_rules 縺ｯ蜻ｼ縺ｳ蜃ｺ縺嶺ｺ呈鋤縺ｮ縺溘ａ谿九☆縲�
    """
    insp = bool(task.get("roll_pipeline_inspection"))
    rw = bool(task.get("roll_pipeline_rewind"))
    ip = bool(task.get("in_progress"))
    ec = bool(task.get("roll_pipeline_ec"))
    if insp and ip:
        b_tier = 0  # ﾂｧB-1
    elif ec or (insp and not ip) or (rw and not ip):
        b_tier = 1  # ﾂｧB-2 / ﾂｧB-3 蟶ｯ
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
    ﾂｧB-2 / ﾂｧB-3: 蜷御ｸ task_id 縺ｮ `roll_pipeline_ec` 陦後�ｮ逶ｴ蠕後↓縲∵悴逹謇九�ｮ蠕檎ｶ夊｡�
    �ｼ�`roll_pipeline_inspection` 縺ｾ縺溘�ｯ `roll_pipeline_rewind`�ｼ峨ｒ陦碁��縺ｧ髫｣謗･縺輔○繧九�
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
            "ﾂｧB-2/ﾂｧB-3 驟榊床隧ｦ陦碁��: EC 縺ｨ譛ｪ逹謇句ｾ檎ｶ夲ｼ域､懈渊�ｼ丞ｷｻ霑斐＠�ｼ峨ｒ髫｣謗･縺励◆萓晞�ｼNO: %s",
            ",".join(moved_tids),
        )


def _assign_sequential_dispatch_trial_order(task_queue: list) -> None:
    """
    `task_queue` 縺ｮ繝ｪ繧ｹ繝磯��縺ｫ蜷医ｏ縺帙※ `dispatch_trial_order` 繧� 1..n 縺ｸ莉倥￠逶ｴ縺吶�
    `_reorder_task_queue_b2_ec_inspection_consecutive` 縺ｮ逶ｴ蠕鯉ｼ医♀繧医�ｳ繧ｭ繝･繝ｼ蜀阪た繝ｼ繝医�ｮ逶ｴ蠕鯉ｼ峨↓蜻ｼ縺ｳ縲�
    EC 縺ｨ蠕檎ｶ夲ｼ域､懈渊�ｼ丞ｷｻ霑斐＠�ｼ峨�ｮ騾｣邯夂分蜿ｷ繧剃ｿ晁ｨｼ縺吶ｋ縲�
    """
    for i, t in enumerate(task_queue, start=1):
        t["dispatch_trial_order"] = i


def _task_queue_all_have_sheet_dispatch_trial_order(task_queue: list) -> bool:
    """驟榊床險育判繧ｷ繝ｼ繝医�ｮ縲碁�榊床隧ｦ陦碁��逡ｪ縲阪′繧ｭ繝･繝ｼ蜈ｨ陦後↓豁｣縺ｮ謨ｴ謨ｰ縺ｧ蜈･縺｣縺ｦ縺�繧九°縲�"""
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
    驟榊床隧ｦ陦碁��縺ｮ遒ｺ螳壹ゅす繝ｼ繝医↓蜈ｨ陦悟��縺ｮ隧ｦ陦碁��縺後≠繧後�ｰ縺昴ｌ繧呈治逕ｨ�ｼ按ｧB-2/3 縺ｮ髫｣謗･郢ｰ繧贋ｸ翫￡縺ｯ陦後ｏ縺ｪ縺��ｼ峨�
    谺�謳阪′縺ゅｌ縺ｰ蠕捺擂縺ｩ縺翫ｊ繝槭せ繧ｿ繝ｻ邏肴悄繝ｻneed 蛻鈴��縺ｪ縺ｩ縺ｧ繧ｽ繝ｼ繝医＠縲・C 髫｣謗･蠕後↓ 1..n 繧剃ｻ倅ｸ弱�
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
            "驟榊床隧ｦ陦碁��逡ｪ: 縲�%s縲榊�励�ｮ蛟､繧偵◎縺ｮ縺ｾ縺ｾ菴ｿ逕ｨ縺励∪縺励◆�ｼ亥�ｨ %s 陦鯉ｼ峨�",
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
        "驟榊床隧ｦ陦碁��逡ｪ: 繝槭せ繧ｿ繝ｻ繧ｿ繧ｹ繧ｯ蜈･蜉帙°繧芽�ｪ蜍戊ｨ育ｮ励＠ 1..%s 繧剃ｻ倅ｸ弱＠縺ｾ縺励◆縲�",
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
    谿ｵ髫�1蜃ｺ蜉� DataFrame 縺ｮ縲碁�榊床隧ｦ陦碁��逡ｪ縲阪ｒ縲∵ｮｵ髫�2 蜀帝�ｭ縺ｨ蜷後§謇矩���ｼ医た繝ｼ繝医�ｻﾂｧB-2/3 髫｣謗･繝ｻ騾｣逡ｪ�ｼ峨〒蝓九ａ繧九�
    驟榊床蟇ｾ雎｡螟悶�ｮ陦後�ｯ遨ｺ縺ｮ縺ｾ縺ｾ縲�
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
            # Excel 荳翫�ｯ謨ｰ蛟､繧ｻ繝ｫ縺ｫ縺励√ヵ繧｣繝ｫ繧ｿ繝ｼ繝ｻ荳ｦ縺ｹ譖ｿ縺医ｒ縺励ｄ縺吶￥縺吶ｋ�ｼ域枚蟄怜�励□縺ｨ謨ｰ蛟､縺ｨ蛻･繧ｰ繝ｫ繝ｼ繝励↓縺ｪ繧具ｼ�
            plan_df.iat[iloc, col_idx] = int(dto)
        except (TypeError, ValueError):
            if pd.api.types.is_numeric_dtype(plan_df.iloc[:, col_idx]):
                plan_df.iat[iloc, col_idx] = float("nan")
            else:
                plan_df.iat[iloc, col_idx] = ""


def _equipment_schedule_unified_sub_string_map(timeline_for_eq_grid: list) -> dict:
    """
    蜷御ｸ譌･繝ｻ蜷御ｸ險ｭ蛯吝�励く繝ｼ繝ｻ蜷御ｸ萓晞�ｼNO 縺ｮ蜉�蟾･縺ｫ縺､縺�縺ｦ縲∬ｨｭ蛯呎凾髢灘牡繧ｻ繝ｫ逕ｨ縺ｮ縲瑚｣懊崎｡ｨ遉ｺ譁�蟄怜�励�
    繧ｿ繧､繝�繝ｩ繧､繝ｳ荳翫�ｮ蜷�繝悶Ο繝�繧ｯ縺ｮ `sub` 縺ｫ迴ｾ繧後◆陬懷勧閠�蜷阪ｒ蜥碁寔蜷医＠縲∵��鬆�縺ｧ ", " 騾｣邨舌☆繧九�
    繝｡繝ｳ繝舌�ｼ譌･遞九�ｻ蜊�譛芽ｨ育ｮ励↓菴ｿ縺�繧ｿ繧､繝�繝ｩ繧､繝ｳ縺ｮ `sub` 縺ｯ螟画峩縺励↑縺��ｼ郁｡ｨ遉ｺ蟆ら畑�ｼ峨�
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
    邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ縺ｨ蜷悟ｽ｢蠑上�ｮ DataFrame�ｼ�10 蛻�譫�繝ｻ險ｭ蛯吝�暦ｼ矩ｲ蠎ｦ蛻暦ｼ峨�
    first_eq_schedule_cell_by_task_id 繧呈ｸ｡縺励◆縺ｨ縺阪�ｮ縺ｿ縲∝�晏�ｺ繧ｻ繝ｫ蠎ｧ讓吶ｒ險倬鹸�ｼ育ｵ先棡繝上う繝代�ｼ繝ｪ繝ｳ繧ｯ逕ｨ�ｼ峨�
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
        eq_empty_cols[f"{eq}騾ｲ蠎ｦ"] = ""

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

        all_eq_rows.append({"譌･譎ょｸｯ": f"笆� {d.strftime('%Y/%m/%d (%a)')} 笆�", **eq_empty_cols})

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
                "譌･譎ょｸｯ": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"
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
                        eq_text = "莨第�ｩ"
                    elif not _use_prog:
                        _ek_disp = _timeline_event_kind(active_ev)
                        _tag = {
                            TIMELINE_EVENT_MACHINE_DAILY_STARTUP: "譌･谺｡蟋区･ｭ貅門ｙ",
                            TIMELINE_EVENT_CHANGEOVER_CLEANUP: "萓晞�ｼ蛻�譖ｿ蠕悟ｧ区忰",
                            TIMELINE_EVENT_CHANGEOVER_PREP: "蜉�蟾･蜑肴ｺ門ｙ",
                        }.get(
                            _ek_disp,
                            "繧ｻ繝�繝医い繝�繝�",
                        )
                        _sub_n = _eq_cell_display_sub(active_ev, d)
                        _sub_text = f" 陬�:{_sub_n}" if _sub_n else ""
                        _tid_d = str(active_ev.get("task_id") or "").strip()
                        eq_text = (
                            f"[{_tid_d}] 荳ｻ:{active_ev.get('op', '')}{_sub_text} ({_tag})"
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
                        sub_text = f" 陬�:{_sub_s}" if _sub_s else ""
                        eq_text = f"[{active_ev['task_id']}] 荳ｻ:{active_ev['op']}{sub_text}"
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
                row_data[f"{eq}騾ｲ蠎ｦ"] = progress_text

            all_eq_rows.append(row_data)
            curr_grid = next_grid
        all_eq_rows.append({"譌･譎ょｸｯ": "", **eq_empty_cols})

    df_eq = pd.DataFrame(all_eq_rows)
    _eq_hdr = _equipment_schedule_header_labels(equipment_list)
    _eq_rename = {}
    for _eq, _lab in zip(equipment_list, _eq_hdr):
        if _eq in df_eq.columns:
            _eq_rename[_eq] = _lab
        _pqc = f"{_eq}騾ｲ蠎ｦ"
        if _pqc in df_eq.columns:
            _eq_rename[_pqc] = f"{_lab}騾ｲ蠎ｦ"
    if _eq_rename:
        df_eq = df_eq.rename(columns=_eq_rename)
    return df_eq


def _machine_display_key_for_equipment(eq: str) -> str:
    """skills 蛻励く繝ｼ縲悟ｷ･遞�+讖滓｢ｰ縲阪°繧画ｩ滓｢ｰ蜷崎｡ｨ遉ｺ繧ｭ繝ｼ繧貞ｾ励ｋ�ｼ磯㍾隍�譎ゅ�ｯ隍�蜷医く繝ｼ縺斐→縺ｫ蛻･蛻暦ｼ峨�"""
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
    讖滓｢ｰ蜷榊腰菴阪↓蛻励ｒ縺ｾ縺ｨ繧√∝推 10 蛻�譫�縺ｧ蜊�譛我ｸｭ縺ｮ萓晞�ｼNO�ｼ郁､�謨ｰ譎ゅ�ｯ縲鯉ｼ上搾ｼ峨ｒ陦ｨ遉ｺ縺吶ｋ縲�
    蛻苓ｦ句�ｺ縺励�ｯ讖滓｢ｰ蜷阪�ｮ縺ｿ�ｼ亥ｷ･遞�+讖滓｢ｰ縺ｮ隍�蜷医く繝ｼ縺ｯ莉倥￠縺ｪ縺��ｼ峨ょ酔荳迚ｩ逅�讖滓｢ｰ縺ｯ蜊�譛峨く繝ｼ縺ｧ1蛻励↓髮�邏�縺吶ｋ縲�
    """
    timeline_for_eq_grid = _expand_timeline_events_for_equipment_grid(timeline_events)
    events_by_date = defaultdict(list)
    for e in timeline_for_eq_grid:
        events_by_date[e["date"]].append(e)

    # 蜊�譛峨く繝ｼ�ｼ域ｩ滓｢ｰ蜷榊�ｴ繝ｻ豁｣隕丞喧�ｼ峨＃縺ｨ縺ｫ1蛻励りｦ句�ｺ縺励�ｯ equipment_list 蛻晏�ｺ縺ｮ讖滓｢ｰ蜷崎｡ｨ遉ｺ縺ｮ縺ｿ縲�
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

        all_rows.append({"譌･譎ょｸｯ": f"笆� {d.strftime('%Y/%m/%d (%a)')} 笆�", **empty_tail})

        curr_grid = d_start
        while curr_grid < d_end:
            next_grid = curr_grid + timedelta(minutes=10)
            if next_grid > d_end:
                next_grid = d_end
            mid_t = curr_grid + (next_grid - curr_grid) / 2
            row_data = {
                "譌･譎ょｸｯ": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"
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
                    tids_by_mcol[mcol].add("�ｼ井ｼ第�ｩ�ｼ�")
                else:
                    tid = str(active_ev.get("task_id") or "").strip()
                    if tid:
                        tids_by_mcol[mcol].add(tid)
            for mcol in machine_cols:
                parts = sorted(tids_by_mcol.get(mcol, ()))
                row_data[mcol] = "�ｼ�".join(parts) if parts else ""
            all_rows.append(row_data)
            curr_grid = next_grid
        all_rows.append({"譌･譎ょｸｯ": "", **empty_tail})

    return pd.DataFrame(all_rows)


def _build_block_table_dataframe(
    sorted_dates: list,
    equipment_list: list,
    members: list,
    attendance_data: dict,
    timeline_events: list,
) -> "pd.DataFrame":
    """
    險ｭ蛯吝�暦ｼ亥頃譛我ｸｭ縺ｮ萓晞�ｼNO�ｼ会ｼ九Γ繝ｳ繝舌�ｼ蛻暦ｼ亥酔�ｼ峨ｒ 10 蛻�譫�縺ｧ荳ｦ縺ｹ縺溘ヶ繝ｭ繝�繧ｯ蜿ｯ隕門喧逕ｨ繧ｷ繝ｼ繝医�
    """
    timeline_for_eq_grid = _expand_timeline_events_for_equipment_grid(timeline_events)
    events_by_date = defaultdict(list)
    for e in timeline_for_eq_grid:
        events_by_date[e["date"]].append(e)

    _eq_hdr = _equipment_schedule_header_labels(equipment_list)
    eq_disp_to_key: dict[str, str] = {}
    for eq, lab in zip(equipment_list, _eq_hdr):
        eq_disp_to_key[f"險ｭ蛯�:{lab}"] = eq

    mem_cols = [f"莠ｺ:{m}" for m in members]
    eq_cols = [f"險ｭ蛯�:{lab}" for lab in _eq_hdr]
    all_cols = ["譌･譎ょｸｯ"] + eq_cols + mem_cols
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

        banner = {"譌･譎ょｸｯ": f"笆� {d.strftime('%Y/%m/%d (%a)')} 笆�"}
        banner.update({c: "" for c in all_cols if c != "譌･譎ょｸｯ"})
        rows_out.append(banner)

        curr_grid = d_start
        while curr_grid < d_end:
            next_grid = curr_grid + timedelta(minutes=10)
            if next_grid > d_end:
                next_grid = d_end
            mid_t = curr_grid + (next_grid - curr_grid) / 2
            row_data: dict = {
                "譌･譎ょｸｯ": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"
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
                    row_data[col_eq] = "莨第�ｩ"
                else:
                    tid = str(active_ev.get("task_id") or "").strip()
                    row_data[col_eq] = tid if tid else "蜊�譛�"

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
                        busy_member_task[op].add("莨第�ｩ" if tid else "莨第�ｩ")
                    elif tid:
                        busy_member_task[op].add(tid)
                for s in str(ev.get("sub") or "").split(","):
                    s = s.strip()
                    if not s:
                        continue
                    if any(
                        b_s <= mid_t < b_e for b_s, b_e in ev.get("breaks") or ()
                    ):
                        busy_member_task[s].add("莨第�ｩ")
                    elif tid:
                        busy_member_task[s].add(tid)

            for m in members:
                col_m = f"莠ｺ:{m}"
                parts = sorted(busy_member_task.get(m, ()))
                row_data[col_m] = "�ｼ�".join(parts) if parts else ""

            rows_out.append(row_data)
            curr_grid = next_grid

        tail = {"譌･譎ょｸｯ": ""}
        tail.update({c: "" for c in all_cols if c != "譌･譎ょｸｯ"})
        rows_out.append(tail)

    return pd.DataFrame(rows_out, columns=all_cols)


def _day_schedule_task_sort_key(
    task: dict,
    _task_queue: list | None = None,
    need_combo_col_index: dict | None = None,
):
    """
    蜷御ｸ譌･蜀�縺ｮ蜑ｲ莉倩ｩｦ陦碁���ｼ�STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST=0 縺ｮ荳ｻ繝ｫ繝ｼ繝礼畑�ｼ峨�
    蜈磯�ｭ繧ｭ繝ｼ縺ｯ _generate_plan_task_queue_sort_key 縺ｨ蜷瑚ｶ｣譌ｨ�ｼ亥刈蟾･騾比ｸｭ繝ｻ邏肴悄蝓ｺ貅� due_basis_date繝ｻﾂｧB 谿ｵ繝ｻb2_queue_sub繝ｻneed 蛻鈴��繝ｻ萓晞�ｼNO�ｼ峨�
    邯壹￠縺ｦ ﾂｧB-1 縺ｮ驟榊床隧ｦ陦碁��郢ｰ繧贋ｸ翫￡縲∝ｷ･遞� rank縲‥ispatch_trial_order縲�ﾂｧB-2 谿ｵ蜀� EC 蜈郁｡後∝━蜈亥ｺｦ縲∫ｵ先棡逕ｨ繧ｭ繝ｼ縲�
    蜷御ｸ迚ｩ逅�讖滓｢ｰ荳翫�ｮ髫咎俣蜑ｲ繧願ｾｼ縺ｿ縺ｯ _equipment_line_lower_dispatch_trial_still_pending 縺ｧ隧ｦ陦碁��繧貞ｼｷ蛻ｶ縺吶ｋ縲�
    STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT=1 縺ｮ縺ｨ縺阪�ｯ _task_blocked_by_global_dispatch_trial_order 縺�
    繧医ｊ蟆上＆縺�隧ｦ陦碁��縺ｮ譛ｪ螳御ｺ�繧定ｷｨ縺�縺�蜑ｲ繧願ｾｼ縺ｿ繧貞挨騾斐ヶ繝ｭ繝�繧ｯ縺吶ｋ縲�
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
    蜷御ｸ迚ｩ逅�讖滓｢ｰ�ｼ�machine 蜊�譛峨く繝ｼ�ｼ我ｸ翫〒縲√ｈ繧雁ｰ上＆縺�驟榊床隧ｦ陦碁��縺ｮ陦後′縺ｾ縺�谿矩㍼繧呈戟縺､縺九�
    machine_avail_dt 縺ｯ繝√Ε繝ｳ繧ｯ髢薙�ｮ髫咎俣縺ｫ蠕檎ｶ夊ｩｦ陦碁��縺悟�･繧願ｾｼ繧√ｋ縺溘ａ縲√％縺薙〒鬆�蠎上ｒ蠑ｷ蛻ｶ縺吶ｋ縲�
    險ｭ蛯吶ｒ霍ｨ縺�縺�隧ｦ陦碁��縺ｮ蜑榊ｾ後�ｯ _task_blocked_by_global_dispatch_trial_order 縺ｧ蛻･騾泌宛蠕｡縺吶ｋ縲�

    繧ｭ繝･繝ｼ蜈磯�ｭ縺ｫ谿矩㍼縺後≠繧九□縺代〒縺ｯ繝悶Ο繝�繧ｯ縺励↑縺�縲Ｕasks_today 縺ｨ蜷梧ｧ倥↓
    start_date_req <= current_date 縺ｮ陦後□縺代ｒ縲悟�郁ｩｦ陦碁��縺ｮ遶ｶ蜷医阪→縺ｿ縺ｪ縺吶�
    �ｼ医∪縺�髢句ｧ区律縺ｫ驕斐＠縺ｦ縺�縺ｪ縺�陦後′蜈ｨ譌･繝悶Ο繝�繧ｫ繝ｼ縺ｫ縺ｪ繧翫∝ｾ檎ｶ壹′縺ｻ縺ｼ驟榊床荳榊庄縺ｫ縺ｪ繧九�ｮ繧帝亟縺舌ゑｼ�

    繧医ｊ蟆上＆縺�隧ｦ陦碁��縺ｮ陦後′ **蜷御ｸ萓晞�ｼ縺ｮ蜑榊ｷ･遞句ｾ�縺｡遲峨〒縺ｾ縺�蜑ｲ莉倅ｸ崎�ｽ**縺ｪ縺ｨ縺阪�ｯ縲檎ｫｶ蜷医�ｮ谿九阪→縺ｿ縺ｪ縺輔↑縺�縲�
    �ｼ亥ｽ楢ｩｲ陦後�ｯ eligible 縺ｫ繧ょ�･繧峨↑縺�縺溘ａ縲√％縺薙〒蠕�縺溘○繧九→蠕檎ｶ夊ｩｦ陦碁��縺悟酔荳險ｭ蛯吶〒豌ｸ荵�蛛懈ｭ｢縺怜ｾ励ｋ縲ゑｼ�

    繧医ｊ蟆上＆縺�隧ｦ陦碁��縺ｮ陦後′ **蠖捺律縺ｮ讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ縺�縺代〒險育判遯薙ｒ蜈ｨ譌･蜊�譛�**�ｼ医◎縺ｮ險ｭ蛯吶�ｯ蠖捺律繧ｹ繝ｭ繝�繝医ぞ繝ｭ�ｼ峨↑繧�
    縲檎ｫｶ蜷医�ｮ谿九阪→縺ｿ縺ｪ縺輔↑縺��ｼ医げ繝ｭ繝ｼ繝舌Ν隧ｦ陦碁��縺ｨ縺ゅｏ縺帙※莉冶ｨｭ蛯吶′蜈ｨ譌･豁｢縺ｾ繧九�ｮ繧帝亟縺撰ｼ峨�
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
    start_date_req <= current_date 縺九▽谿矩㍼縺ゅｊ縺ｮ繧ｿ繧ｹ繧ｯ縺ｮ驟榊床隧ｦ陦碁��縺ｮ譛蟆丞､縲�
    _equipment_line_lower_dispatch_trial_still_pending 縺ｨ蜷梧ｧ倥√∪縺�髢句ｧ区律縺ｫ驕斐＠縺ｦ縺�縺ｪ縺�陦後�ｯ
    縲悟�郁｡瑚ｩｦ陦碁��縺ｮ遶ｶ蜷医阪↓蜷ｫ繧√↑縺�縲�

    **繧ｰ繝ｭ繝ｼ繝舌Ν隧ｦ陦碁��繝悶Ο繝�繧ｯ**�ｼ�STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT�ｼ臥畑縺ｫ縲�
    縲後％縺ｮ譌･縺ｾ縺�蜑ｲ莉伜呵｣懊↓縺ｪ繧雁ｾ励↑縺�縲崎｡後�ｯ譛蟆丞､縺九ｉ髯､螟悶☆繧九ゅ＆繧ゅ↑縺�縺ｨ蜷御ｸ萓晞�ｼ縺ｮ
    ﾂｧA-1/ﾂｧA-2 蜑榊ｷ･遞具ｼ郁ｩｦ陦碁��縺ｯ蠕後ｍ縺�縺瑚｡碁��縺ｯ蜈茨ｼ峨′蠢�隕√↑陦後′縲√ｈ繧雁ｰ上＆縺�隧ｦ陦碁��縺ｮ陦後→
    蠕ｪ迺ｰ縺励※豌ｸ荵�縺ｫ蜍輔￠縺ｪ縺�縲�
    - `_task_not_yet_schedulable_due_to_dependency_or_b2_room` 縺� True 縺ｮ陦�
    - �ｼ�daily_status繝ｻmembers 縺梧ｸ｡繧九→縺搾ｼ牙ｽ捺律讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ縺�縺代〒險育判遯灘�ｨ譌･蜊�譛峨�ｮ陦�
    - �ｼ�machine_avail_dt 遲峨′貂｡繧九→縺搾ｼ芽ｨｭ蛯吝｣∵凾險医′險育判邨らｫｯ莉･荳翫〒蠖捺律繧ｹ繝ｭ繝�繝医↑縺励�ｮ陦�

    1 繝ｭ繝ｼ繝ｫ蜑ｲ蠖薙�励Ο繝ｼ繝悶↓繧医ｋ髯､螟悶�ｯ陦後ｏ縺ｪ縺��ｼ�`_effective_min_dispatch_trial_order_from_pool` 蛛ｴ縺ｧ螻､縺斐→縺ｫ蛻､螳夲ｼ峨�
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
    繧医ｊ蟆上＆縺�驟榊床隧ｦ陦碁��縺ｫ縲∝ｽ捺律蜑ｲ莉伜庄閭ｽ縺ｪ譛ｪ螳御ｺ�縺後≠繧九→縺阪∝ｽ楢ｩｲ繧ｿ繧ｹ繧ｯ繧偵ヶ繝ｭ繝�繧ｯ縺吶ｋ縲�
    min_dispatch_effective: 繝励�ｼ繝ｫ�ｼ九�励Ο繝ｼ繝悶〒豎ゅａ縺溷ｮ溷柑譛蟆剰ｩｦ陦碁���ｼ域悴謖�螳壽凾縺ｯ螳我ｾ｡繝輔ぅ繝ｫ繧ｿ縺ｮ縺ｿ縺ｮ譛蟆擾ｼ峨�
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
    """蜍､諤�霎樊嶌縺九ｉ繝槭せ繧ｿ縺ｫ辟｡縺�譌･莉倥く繝ｼ繧貞炎髯､縺吶ｋ�ｼ郁�ｪ蜍墓僑蠑ｵ蛻�縺ｮ蟾ｻ縺肴綾縺暦ｼ峨�"""
    for dk in list(attendance_data.keys()):
        if dk not in keep_dates:
            del attendance_data[dk]


def _partial_task_id_due_shift_outcome(
    task_queue: list, task_id: str, calendar_last: date
) -> tuple[bool, bool]:
    """
    驟榊床谿九�ｮ萓晞�ｼNO縺ｫ縺､縺�縺ｦ邏肴悄+1譌･繝ｪ繝医Λ繧､縺ｮ蛻�鬘槭�
    謌ｻ繧雁､: (shift_ok, calendar_shortfall)
    - shift_ok: 邏肴悄蝓ｺ貅厄ｼ�due_basis_date�ｼ峨ｒ謖√▽陦後′縺ゅｊ縲√◎繧後ｉ縺吶∋縺ｦ縺ｧ +1 譌･縺後�槭せ繧ｿ譛邨りｨ育判譌･莉･荳�
    - calendar_shortfall: 邏肴悄蝓ｺ貅悶ｒ謖√▽陦後′縺ゅｊ縲√＞縺壹ｌ縺九〒 +1 譌･縺後�槭せ繧ｿ譛邨りｨ育判譌･繧定ｶ�縺医ｋ
    蝓ｺ貅也ｴ肴悄縺御ｸ陦後ｂ辟｡縺�萓晞�ｼ縺ｯ (False, False)�ｼ磯壼ｸｸ縺ｮ驟榊床谿九�ｮ縺ｾ縺ｾ�ｼ峨�
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
    驟榊床谿九Μ繝医Λ繧､逕ｨ: **蜀�驛ｨ縺ｮ邏肴悄蝓ｺ貅厄ｼ�due_basis_date�ｼ峨□縺�**繧� +1 譌･縺吶ｋ縲�
    邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ逕ｨ縺ｮ ``due_basis_date_result_sheet`` 縺ｯ螟画峩縺励↑縺��ｼ�+1 蜑阪�ｮ譌･莉倥ｒ菫晄戟�ｼ峨�
    蝗樒ｭ皮ｴ肴悄繝ｻ謖�螳夂ｴ肴悄繧る�榊床險育判繧ｷ繝ｼ繝育罰譚･縺ｮ縺ｾ縺ｾ縲�
    due_urgent 縺ｯ縺壹ｉ縺励◆ due_basis_date 縺ｧ蜀崎ｨ育ｮ励☆繧九�
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
    """蜷御ｸ譌･蜀�縺ｮ譌｢蟄� timeline 縺九ｉ險ｭ蛯咏ｩｺ縺阪�ｻ繝｡繝ｳ繝舌�ｼ遨ｺ縺阪�ｮ荳矩剞繧貞渚譏�縺吶ｋ�ｼ磯Κ蛻�蜀埼�榊床逕ｨ�ｼ峨�"""
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
    """蜊企幕蛹ｺ髢薙ヶ繝ｭ繝�繧ｯ [start,end) 縺ｫ t 縺悟�･繧矩俣縲∫ｵらｫｯ縺ｸ郢ｰ繧贋ｸ翫￡繧九�"""
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
    # Excel 縺ｧ 0 繧偵檎ｩｺ縲阪→縺励※縺�繧句�励ｄ縲∵焚蠑上�ｮ邨先棡 0 縺ｯ蜊�譛峨＠縺ｪ縺��ｼ亥ｾ捺擂 True 縺�縺ｨ蜈ｨ譌･蜊�譛画桶縺�縺ｫ縺ｪ繧雁ｾ励ｋ�ｼ�
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
    讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ1繧ｹ繝ｭ繝�繝� [slot_start, slot_end) 繧貞ｷ･蝣ｴ遞ｼ蜒肴棧縺ｫ繧ｯ繝ｪ繝�繝励☆繧九�
    譫�螟悶�ｮ縺ｿ縺ｮ繧ｹ繝ｭ繝�繝医�ｯ None�ｼ磯�榊床縺ｧ縺ｯ辟｡隕厄ｼ峨よｮｵ髫�2縺ｧ縺ｯ master 繝｡繧､繝ｳ A12/B12 縺ｧ
    DEFAULT_START_TIME / DEFAULT_END_TIME 縺御ｸ頑嶌縺肴ｸ医∩�ｼ�generate_plan 縺ｮ繧ｳ繝ｳ繝�繧ｭ繧ｹ繝亥��縺ｧ隱ｭ霎ｼ�ｼ峨�
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
    讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ蜊�譛峨�ｮ蜿ｳ遶ｯ繧貞��繧倶ｸ企剞縲ょｷ･蝣ｴ繝槭せ繧ｿ邨よ･ｭ�ｼ�DEFAULT_END_TIME�ｼ峨→縲�
    蠖捺律驟榊床蟇ｾ雎｡繝｡繝ｳ繝舌�ｼ縺ｮ蜍､蜍咏ｵゆｺ�譎ょ綾縺ｮ譛蟆上�ｮ蟆上＆縺�譁ｹ�ｼ井ｺｺ縺後＞縺ｪ縺�譎る俣蟶ｯ縺ｮ縲悟頃譛峨阪〒
    險ｭ蛯吝ｺ翫□縺代′邨よ･ｭ繧定ｶ�縺医↑縺�繧医≧縺ｫ縺吶ｋ�ｼ峨�
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
    """蜊�譛牙濠髢句玄髢薙ｒ [w0, w1) 縺ｫ繧ｯ繝ｪ繝�繝励＠縺ｦ縺九ｉ繝槭�ｼ繧ｸ縺吶ｋ縲�"""
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
    master.xlsm縲梧ｩ滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ縲阪ｒ隱ｭ縺ｿ縲∬ｨｭ蛯吝�励�ｮ髱樒ｩｺ繧ｻ繝ｫ�ｼ晏ｽ楢ｩｲ 1 譎る俣繧ｹ繝ｭ繝�繝亥頃譛峨→縺ｿ縺ｪ縺吶�
    謌ｻ繧�: 譌･莉� -> equipment_list 縺ｮ繧ｭ繝ｼ -> 蜊企幕蛹ｺ髢� [start, end) 縺ｮ繝ｪ繧ｹ繝茨ｼ医�槭�ｼ繧ｸ貂医∩�ｼ峨�
    """
    if not master_path or not os.path.isfile(master_path):
        return {}
    try:
        xls = pd.ExcelFile(master_path)
        if SHEET_MACHINE_CALENDAR not in xls.sheet_names:
            return {}
        raw = pd.read_excel(master_path, sheet_name=SHEET_MACHINE_CALENDAR, header=None)
    except Exception as e:
        logging.warning("讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ: 繧ｷ繝ｼ繝郁ｪｭ霎ｼ繧偵せ繧ｭ繝�繝励＠縺ｾ縺励◆ (%s)", e)
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
    """蠖捺律縺ｮ繧ｿ繧､繝�繝ｩ繧､繝ｳ繧ｷ繝ｼ繝牙ｾ後∵ｩ滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ蜊�譛峨〒險ｭ蛯咏ｩｺ縺堺ｸ矩剞繧堤ｹｰ繧贋ｸ翫￡繧九�"""
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
    """day_blocks 縺九ｉ蜊�譛峨く繝ｼ�ｼ郁｡ｨ險倥ｆ繧峨℃險ｱ螳ｹ�ｼ峨↓荳閾ｴ縺吶ｋ蛹ｺ髢薙Μ繧ｹ繝医ｒ蠕励ｋ縲�"""
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
    蠖捺律縺ｮ讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ蜊�譛峨′險育判遯� [蟋区･ｭ, min(邨よ･ｭ,遞ｼ蜒阪Γ繝ｳ繝舌�ｼ邨ゆｺ�) ) 蜈ｨ菴薙ｒ蝪槭℃縲�
    縺昴�ｮ險ｭ蛯吶〒縺ｯ蠖捺律 1 譛ｬ繧ょ刈蟾･繧貞�･繧後ｉ繧後↑縺�縺ｨ縺� True縲�
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
    蠖楢ｩｲ繧ｿ繧ｹ繧ｯ縺ｮ蜊�譛芽ｨｭ蛯吶′縲∝ｽ捺律縺ｮ讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ縺�縺代〒險育判遯薙ｒ蜈ｨ譌･蝪槭′繧後※縺�繧九�
    繧ｰ繝ｭ繝ｼ繝舌Ν隧ｦ陦碁��繝悶Ο繝�繧ｯ逕ｨ縺ｮ縲梧怙蟆剰ｩｦ陦碁��縲阪°繧牙､悶☆�ｼ井ｻ冶ｨｭ蛯吶�ｮ驟榊床繝�繝�繝峨Ο繝�繧ｯ髦ｲ豁｢�ｼ峨�
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
    machine_avail_dt�ｼ医す繝ｼ繝峨�ｻ讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ蠎翫�ｻ蠖捺律遒ｺ螳壹Ο繝ｼ繝ｫ蜿肴丐蠕鯉ｼ峨〒縲�
    蜊�譛芽ｨｭ蛯吶�ｮ遨ｺ縺堺ｸ矩剞縺瑚ｨ育判遯鍋ｵらｫｯ莉･荳翫↑繧牙ｽ捺律縺ｯ蠖楢ｨｭ蛯吶↓繧ｹ繝ｭ繝�繝医↑縺励�
    `machine_handoff` 遲峨′貂｡繧九→縺阪�ｯ `_resolve_machine_changeover_floor_segments` 縺ｫ繧医ｊ
    `_assign_one_roll_trial_order_flow` 縺ｨ蜷後§ **螳溷柑蜉�蟾･髢句ｧ倶ｸ矩剞** 縺ｧ蛻､螳壹☆繧�
    �ｼ育函縺ｮ machine_avail 縺�縺代〒縺ｯ繝√Ε繝ｳ繧ｸ繧ｪ繝ｼ繝舌�ｼ蠕後�ｮ荳矩剞縺梧ｬ�縺代∝呵｣懊ｄ min_dto 縺檎汲縺�縺ｮ繧帝亟縺撰ｼ峨�
    縺ｾ縺溽ｩｺ縺堺ｸ矩剞縺檎ｵらｫｯ繧医ｊ蜑阪〒繧ゅ∬ｨ育判遯薙〒縺ｮ **谿九ｊ騾｣邯壹′ 1 繝ｭ繝ｼ繝ｫ蛻�縺ｫ雜ｳ繧翫↑縺�**
    縺ｨ蛻､譁ｭ縺ｧ縺阪ｋ蝣ｴ蜷医�ｯ True�ｼ亥ｮ溷ロ荳崎ｶｳ繝�繝�繝峨Ο繝�繧ｯ髦ｲ豁｢�ｼ峨�
    繧ｫ繝ｬ繝ｳ繝繝ｼ蛹ｺ髢鍋�ｧ蜷医�ｮ繧ｭ繝ｼ蜿悶ｊ縺薙⊂縺励ｒ髦ｲ縺舌�
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
    # eff_time_per_unit 竕� base / avg_eff / t_eff ﾃ� 菴吝鴨菫よ焚縲Ｂvg_eff 縺ｯ繝√�ｼ繝�谺｡隨ｬ縺ｧ荳九′繧九�
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
    """繝ｭ繝ｼ繝ｫ遒ｺ螳夂峩蠕�: 邨ゆｺ�譎ょ綾縺後き繝ｬ繝ｳ繝繝ｼ蜊�譛峨せ繝ｭ繝�繝亥��縺ｪ繧臥ｵらｫｯ縺ｾ縺ｧ郢ｰ繧贋ｸ翫￡縲�"""
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
    master.xlsm 縺ｮ莉ｻ諢上す繝ｼ繝�:
      - 縲瑚ｨｭ螳喟萓晞�ｼ蛻�譖ｿ蜑榊ｾ梧凾髢薙坂ｦ 蟾･遞句錐繝ｻ讖滓｢ｰ蜷阪�ｻ貅門ｙ蛻�繝ｻ蠕悟ｧ区忰蛻��ｼ�1 陦檎岼隕句�ｺ縺励�2 陦檎岼莉･髯阪ョ繝ｼ繧ｿ�ｼ�
      - 縲瑚ｨｭ螳喟讖滓｢ｰ_譌･谺｡蟋区･ｭ貅門ｙ縲坂ｦ 讖滓｢ｰ蜷阪�ｻ譌･谺｡蟋区･ｭ貅門ｙ蛻�

    萓晞�ｼNO�ｼ医ち繧ｹ繧ｯ�ｼ峨′蜷御ｸ迚ｩ逅�讖滓｢ｰ荳翫〒蛻�繧頑崛繧上ｋ縺ｨ縺阪∫峩蜑阪ヶ繝ｭ繝�繧ｯ縺ｮ蠕悟ｧ区忰竊貞ｽ楢ｩｲ繝悶Ο繝�繧ｯ縺ｮ貅門ｙ繧�
    險ｭ蛯咏ｩｺ縺堺ｸ矩剞縺ｫ蜉�邂励☆繧九ょ酔荳萓晞�ｼNO縺ｮ騾｣邯壹Ο繝ｼ繝ｫ縺ｮ髢薙↓縺ｯ蜉�邂励＠縺ｪ縺�縲�
    譌･谺｡蟋区･ｭ貅門ｙ縺ｯ縲∝酔荳繧ｫ繝ｬ繝ｳ繝譌･縺ｧ蠖楢ｩｲ讖滓｢ｰ縺ｮ蜈磯�ｭ繝ｭ繝ｼ繝ｫ縺ｫ縺ｮ縺ｿ蜉�邂励☆繧九�

    謌ｻ繧�: (險ｭ蛯呵｡後く繝ｼ縲悟ｷ･遞�+讖滓｢ｰ縲阪♀繧医�ｳ豁｣隕丞喧繧ｭ繝ｼ -> (貅門ｙ蛻�, 蠕悟ｧ区忰蛻�),
          讖滓｢ｰ蜷阪♀繧医�ｳ豁｣隕丞喧繧ｭ繝ｼ -> 蟋区･ｭ貅門ｙ蛻�)
    """
    changeover: dict[str, tuple[int, int]] = {}
    startup: dict[str, int] = {}
    if not master_path or not os.path.isfile(master_path):
        return changeover, startup
    try:
        xls = pd.ExcelFile(master_path)
    except Exception as e:
        logging.warning("讖滓｢ｰ貅門ｙ/蛻�譖ｿ險ｭ螳�: 繝悶ャ繧ｯ繧帝幕縺代∪縺帙ｓ (%s)", e)
        return changeover, startup

    if SHEET_MACHINE_CHANGEOVER in xls.sheet_names:
        try:
            df = pd.read_excel(
                master_path, sheet_name=SHEET_MACHINE_CHANGEOVER, header=0
            )
            df.columns = [str(c).strip() for c in df.columns]
            c_proc = _df_pick_column(df, "蟾･遞句錐", "蟾･遞�")
            c_mac = _df_pick_column(df, "讖滓｢ｰ蜷�", "讖滓｢ｰ")
            c_prep = _df_pick_column(
                df,
                "貅門ｙ譎る俣_蛻�",
                "貅門ｙ蛻�",
                "蜉�蟾･蜑肴ｺ門ｙ_蛻�",
                "蜉�蟾･髢句ｧ句燕貅門ｙ_蛻�",
            )
            c_clean = _df_pick_column(
                df,
                "蠕悟ｧ区忰譎る俣_蛻�",
                "蠕悟ｧ区忰蛻�",
                "蜉�蟾･蠕悟ｾ悟ｧ区忰_蛻�",
                "蜉�蟾･邨ゆｺ�蠕悟ｾ悟ｧ区忰_蛻�",
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
                        "繝槭せ繧ｿ縲�%s縲�: 蟾･遞�+讖滓｢ｰ %s 陦後�ｮ貅門ｙ/蠕悟ｧ区忰�ｼ亥���ｼ峨ｒ隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�",
                        SHEET_MACHINE_CHANGEOVER,
                        n_ent,
                    )
        except Exception as e:
            logging.warning(
                "繝槭せ繧ｿ縲�%s縲崎ｪｭ霎ｼ螟ｱ謨暦ｼ育┌隕厄ｼ�: %s", SHEET_MACHINE_CHANGEOVER, e
            )

    if SHEET_MACHINE_DAILY_STARTUP in xls.sheet_names:
        try:
            df2 = pd.read_excel(
                master_path, sheet_name=SHEET_MACHINE_DAILY_STARTUP, header=0
            )
            df2.columns = [str(c).strip() for c in df2.columns]
            c_mn = _df_pick_column(df2, "讖滓｢ｰ蜷�", "讖滓｢ｰ")
            c_su = _df_pick_column(
                df2, "譌･谺｡蟋区･ｭ貅門ｙ_蛻�", "蟋区･ｭ貅門ｙ_蛻�", "譌･蟋区･ｭ貅門ｙ_蛻�"
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
                        "繝槭せ繧ｿ縲�%s縲�: 讖滓｢ｰ %s 莉ｶ縺ｮ譌･谺｡蟋区･ｭ貅門ｙ�ｼ亥���ｼ峨ｒ隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�",
                        SHEET_MACHINE_DAILY_STARTUP,
                        len({k for k in startup if "+" not in str(k)}),
                    )
        except Exception as e:
            logging.warning(
                "繝槭せ繧ｿ縲�%s縲崎ｪｭ霎ｼ螟ｱ謨暦ｼ育┌隕厄ｼ�: %s", SHEET_MACHINE_DAILY_STARTUP, e
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
    蠖捺律 eligible 縺ｮ縺�縺｡縲∝ｽ楢ｩｲ蟾･遞�+讖滓｢ｰ縺ｧ OP 繧ｹ繧ｭ繝ｫ繧呈戟縺､閠�縺ｮ縺�縺｡蜆ｪ蜈亥ｺｦ縺梧怙蟆上�ｮ1蜷阪�
    貅門ｙ繝ｻ譌･谺｡蟋区･ｭ縺ｮ莨第�ｩ繧ｹ繧ｭ繝�繝励↓逕ｨ縺�繧具ｼ�avail_dt 縺ｯ隕九↑縺��ｼ峨�
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
    """繧ｹ繧ｭ繝ｫ OP 縺梧鏡縺医↑縺�縺ｨ縺阪�ｮ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ�ｼ亥｣∵凾險医↓蛻�繧定ｶｳ縺呻ｼ丞ｮ壼ｸｸ髢句ｧ句渕貅悶�ｮ譌･谺｡蟋区･ｭ縺ｯ邨ゆｺ�譎ょ綾縺ｧ max�ｼ峨�"""
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
    蜑阪Ο繝ｼ繝ｫ蜉�蟾･邨ゆｺ� prev_machining_end_dt 縺九ｉ縲∵律谺｡蟋区･ｭ�ｼ亥ｽ捺律蜈磯�ｭ縺ｮ縺ｿ�ｼ峨�ｻ蜷梧律萓晞�ｼ蛻�譖ｿ縺ｮ蠕悟ｧ区忰繝ｻ貅門ｙ繧�
    邨�縺ｿ遶九※縲�(蜉�蟾･髢句ｧ区怙譌ｩ譎ょ綾, 繧ｿ繧､繝�繝ｩ繧､繝ｳ逕ｨ繧ｻ繧ｰ繝｡繝ｳ繝磯屁蠖｢) 繧定ｿ斐☆縲�
    譌･谺｡蟋区･ｭ縺ｯ master 繝｡繧､繝ｳ A15�ｼ亥ｮ壼ｸｸ髢句ｧ具ｼ峨′隱ｭ繧√ｌ縺ｰ [髢句ｧ�, 髢句ｧ�+N蛻�) 縺ｮ螢∵凾險茨ｼ亥共諤� forward 縺励↑縺��ｼ峨�
    A15 縺瑚ｪｭ繧√↑縺�縺ｨ縺阪�ｮ縺ｿ縲∝ｾ捺擂縺ｩ縺翫ｊ莉｣陦ｨ繧ｹ繧ｭ繝ｫ OP 縺ｮ蜍､蜍吶�ｻ莨第�ｩ縺ｫ豐ｿ縺｣縺ｦ forward 縺吶ｋ縲�
    蜷御ｸ蜊�譛峨く繝ｼ縺ｧ逶ｴ蜑榊刈蟾･縺ｨ蜷御ｸ萓晞�ｼNO縺ｮ縺ｨ縺阪�ｯ蜉�蟾･蜑肴ｺ門ｙ繧剃ｻ倥￠縺ｪ縺��ｼ磯｣邯壹Ο繝ｼ繝ｫ�ｼ峨�
    譌･谺｡蟋区･ｭ繧ｻ繧ｰ繝｡繝ｳ繝医�ｮ op 縺ｯ遨ｺ�ｼ医ち繧､繝�繝ｩ繧､繝ｳ縺ｧ縺ｯ莠ｺ繧定ｼ峨○縺夊ｨｭ蛯吶�ｮ縺ｿ�ｼ峨よｺ門ｙ繝ｻ蠕悟ｧ区忰縺ｮ op 縺ｯ forward 逕ｨ縺ｮ莉｣陦ｨ�ｼ冗峩蜑堺ｸｻ縲�
    繧ｻ繧ｰ繝｡繝ｳ繝� dict 縺ｯ start_dt, end_dt, op, event_kind, machine, machine_occupancy_key 繧呈戟縺､縲�
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
    險ｭ蛯吶�ｮ螢∵凾險医↓縺翫￠繧九悟ｽ楢ｩｲ繝ｭ繝ｼ繝ｫ縺ｮ蜉�蟾･髢句ｧ九堺ｻ･蜑阪�ｮ荳矩剞縲�
    daily_status繝ｻskills_dict繝ｻcurrent_date 縺梧純縺�縺ｨ縺阪�ｯ縲《kills 驕ｩ蜷� OP 縺ｮ蜍､蜍吶�ｻ莨第�ｩ縺ｫ豐ｿ縺｣縺ｦ
    譌･谺｡蟋区･ｭ繝ｻ蠕悟ｧ区忰繝ｻ貅門ｙ繧� forward 縺励◆譛譌ｩ蜉�蟾･髢句ｧ九よ純繧上↑縺�縺ｨ縺阪�ｯ蛻�縺ｮ螢∵凾險亥刈邂励↓繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ縲�
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
    險ｭ蛯吶�ｮ蜉�蟾･髢句ｧ倶ｸ矩剞縺ｨ縲√ち繧､繝�繝ｩ繧､繝ｳ霑ｽ險倡畑繧ｻ繝�繝医い繝�繝怜玄髢薙�
    謌ｻ繧雁､ (floor_dt, segments, abort)縲Ｂbort 縺� True 縺ｮ縺ｨ縺阪�ｯ蠖楢ｩｲ繝ｭ繝ｼ繝ｫ蜑ｲ蠖薙ｒ蜈ｨ菴薙→縺励※譽�蜊ｴ縺吶ｋ縲�
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
    """繧ｿ繧､繝�繝ｩ繧､繝ｳ逕ｨ縺ｮ荳ｻ�ｼ剰｣懊よ律谺｡蟋区･ｭ縺ｯ莠ｺ縺ｪ縺励よｺ門ｙ縺ｯ逶ｴ蠕後Ο繝ｼ繝ｫ縲∝ｾ悟ｧ区忰縺ｯ handoff 縺ｮ逶ｴ蜑阪Ο繝ｼ繝ｫ縲�"""
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
    """繧ｻ繝�繝医い繝�繝礼ｳｻ繧ｻ繧ｰ繝｡繝ｳ繝医ｒ繧ｿ繧､繝�繝ｩ繧､繝ｳ繝ｻ繝溘Λ繝ｼ繝ｻ諡�蠖楢� avail 縺ｫ蜿肴丐縲�"""
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
    蠖楢ｩｲ譌･縺ｮ邨ゆｺ�譎らせ縺ｧ縲∫ｴ肴悄蝓ｺ貅匁律�ｼ亥ｽ捺律蜷ｫ繧�ｼ我ｻ･蜑阪↑縺ｮ縺ｫ谿矩㍼縺梧ｮ九ｋ萓晞�ｼNO縲�
    縲檎ｴ肴悄譌･蜀�縺ｫ螳碁≠縺ｧ縺阪↑縺九▲縺溘�= 蠕後ｍ蛟偵＠蜀崎ｩｦ陦後�ｮ蛟呵｣懊�
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
    繧ｿ繧､繝�繝ｩ繧､繝ｳ縺九ｉ縲∝推 machine_occupancy_key 縺ｫ縺､縺�縺ｦ
    險育判譌･ current_date 莉･蜑阪�ｮ **蜉�蟾･ (machining)** 繧､繝吶Φ繝医�ｮ譛邨らｵゆｺ�繧貞ｾｩ蜈�縺吶ｋ縲�
    繧ｻ繝�繝医い繝�繝礼ｳｻ event_kind 縺ｯ last_tid / 蠕悟ｧ区忰蛻､螳壹↓蜷ｫ繧√↑縺�縲�
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
    """蜴溷渚謚募�･譌･繧定ｵｷ轤ｹ縺ｫ縲√◎縺ｮ譌･縺ｮ蜉�蟾･髢句ｧ九�ｮ荳矩剞譎ょ綾�ｼ亥酔譌･縺ｯ 13:00 莉･髯阪ｒ蜷ｫ繧�ｼ峨�"""
    floor = datetime.combine(current_date, DEFAULT_START_TIME)
    # ﾂｧB-2 讀懈渊 / ﾂｧB-3 蟾ｻ霑斐＠縺ｯ EC 螳御ｺ�繧貞ｾ�縺｣縺ｦ髢句ｧ九〒縺阪ｋ縺溘ａ縲�
    # 蜴溷渚謚募�･譌･�ｼ�=蜷梧律13:00莉･髯搾ｼ峨�ｮ蛻ｶ邏�繧偵◎縺ｮ縺ｾ縺ｾ驕ｩ逕ｨ縺吶ｋ縺ｨ蠕檎ｶ壹′荳榊ｿ�隕√↓蠕後ｍ縺ｸ蛟偵ｌ繧九�
    # EC螳御ｺ�譎ょ綾荳矩剞�ｼ�_roll_pipeline_b2_inspection_ec_completion_floor_dt�ｼ峨〒謨ｴ蜷医ｒ蜿悶ｋ縲�
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
        # min_dto 縺九ｉ蜈ｨ譌･繧ｫ繝ｬ繝ｳ繝繝ｼ蜊�譛峨�ｯ髯､螟匁ｸ医∩縺ｧ繧ゅ∝酔譌･隧ｦ陦碁��縺ｮ縲後ヶ繝ｭ繝�繧ｯ縲阪�ｯ my_o>m 縺ｮ縺ｿ縺ｮ縺溘ａ
        # 隧ｦ陦碁��=min 縺ｮ蜊�譛芽｡後′谿九ｊ縲∽ｻ冶ｩｦ陦碁��縺梧ｰｸ荵�蛛懈ｭ｢縺怜ｾ励ｋ縲ょｽ捺律繧ｹ繝ｭ繝�繝医ぞ繝ｭ縺ｮ陦後�ｯ蛟呵｣懷､悶↓縺吶ｋ縲�
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
    邨�縺ｿ蜷医ｏ縺幄｡ｨ繝励Μ繧ｻ繝�繝�1陦後�ｮ莠ｺ謨ｰ遽�蝗ｲ (lo, hi)縲Ｏeed 縺ｮ蝓ｺ譛ｬ莠ｺ謨ｰ繧医ｊ繧ｷ繝ｼ繝亥�ｴ繧貞━蜈医☆繧九�
    - 蠢�隕∽ｺｺ謨ｰ蛻励′豁｣縺ｮ縺ｨ縺阪�ｯ繝｡繝ｳ繝舌�ｼ蛻励�ｮ莠ｺ謨ｰ縺ｨ荳閾ｴ縺吶ｋ縺薙→縲�
    - hi 縺ｯ need 縺ｮ荳企剞縺ｨ螳滉ｺｺ謨ｰ縺ｮ螟ｧ縺阪＞譁ｹ�ｼ医�励Μ繧ｻ繝�繝医′ need 繧医ｊ蟆台ｺｺ謨ｰ縺ｧ繧よ治逕ｨ蜿ｯ閭ｽ�ｼ峨�
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
    """蜉�蟾･險育判縺ｮ蠢�隕∽ｺｺ謨ｰ蛻励′豁｣縺ｮ謨ｴ謨ｰ縺ｪ繧峨◎縺ｮ蛟､縲ら┌蜉ｹ縺ｪ繧� None縲�"""
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
    """繝ｬ繧ｬ繧ｷ繝ｼ譌･谺｡驟榊床繝ｫ繝ｼ繝礼畑: 蜊倅ｸ繝√�ｼ繝�縺梧�千ｫ九☆繧後�ｰ team_candidates 縺ｫ 1 莉ｶ霑ｽ蜉�縺励※ True縲�"""
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
    1繝ｭ繝ｼ繝ｫ蛻�縺ｮ譛濶ｯ繝√�ｼ繝�繧呈ｱｺ螳壹☆繧九りｨｭ蛯咏ｩｺ縺阪�ｻ譌･髢句ｧ倶ｸ矩剞繧� team_start 縺ｫ郢斐ｊ霎ｼ繧縲�
    preferred_team 縺御ｸ弱∴繧峨ｌ縲√°縺､縲悟酔荳譌･蜀�縺ｮ逶ｴ蜑阪Ο繝ｼ繝ｫ縲阪→縺励※謌千ｫ九☆繧後�ｰ縲�
    邨�蜷医○謗｢邏｢繧医ｊ蜆ｪ蜈医＠縺ｦ謗｡逕ｨ縺吶ｋ�ｼ育ｿ梧律縺ｫ縺ｯ謖√■雜翫＆縺ｪ縺��ｼ峨�
    謌ｻ繧雁､: team(tuple), start_dt, end_dt, breaks, eff, op, eff_time_per_unit, extra_max, rq_base, need_src_line, extra_src_line, machine, machine_name, eq_line, req_num, max_team_size
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
            need_src_line = (need_src_line + "�ｼ�") if need_src_line else ""
            need_src_line += f"險育判繧ｷ繝ｼ繝亥ｿ�隕∽ｺｺ謨ｰ{plan_ro}縺ｯ譛ｪ菴ｿ逕ｨ�ｼ�need蝓ｺ貅�={req_num}�ｼ�"
    else:
        if plan_ro is not None:
            req_num = plan_ro
            need_src_line = f"險育判繧ｷ繝ｼ繝医悟ｿ�隕＾P(荳頑嶌)縲�={req_num}"
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
            (need_src_line + " 竊� ") if need_src_line else ""
        ) + "繝｡繧､繝ｳ荳頑嶌ignore_need_minimum縺ｧreq=1"

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
            "繝｡繧､繝ｳ繧ｰ繝ｭ繝ｼ繝舌Ν(譌･莉佚怜ｷ･遞�): task=%s date=%s 蟾･遞�=%r 繝√�ｼ繝�蠢�鬆�=%s",
            task.get("task_id"),
            current_date,
            machine,
            ",".join(_gdp_must),
        )
    if fixed_team_anchor:
        _nfix = len(fixed_team_anchor)
        if _nfix > req_num:
            need_src_line = (need_src_line + " 竊� ") if need_src_line else ""
            need_src_line += f"繧ｰ繝ｭ繝ｼ繝舌Ν(譌･莉佚怜ｷ･遞�)謖�蜷阪〒譛菴施_nfix}莠ｺ"
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
            (extra_src_line + " 竊� ") if extra_src_line else ""
        ) + "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW縺ｧ0"
    extra_max = (
        extra_max_sheet if TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS else 0
    )
    if (
        extra_max_sheet > 0
        and not TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS
        and not TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW
    ):
        extra_src_line = (
            (extra_src_line + " 竊� ") if extra_src_line else ""
        ) + "繝｡繧､繝ｳ縺ｯ蝓ｺ譛ｬ莠ｺ謨ｰ縺ｮ縺ｿ�ｼ井ｽ吝鴨譫�縺ｯ蜈ｨ驟榊床蠕後↓譛ｪ蜑ｲ蠖禿励せ繧ｭ繝ｫ縺ｧ霑ｽ險假ｼ�"
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
            "need莠ｺ謨ｰ(隧ｦ陦碁��蜆ｪ蜈医ヵ繝ｭ繝ｼ) order=%s task=%s 蟾･遞�/讖滓｢ｰ=%s/%s "
            "req_num=%s [%s] extra_max=%s [%s] max_team蛟呵｣�=%s capable=%s莠ｺ",
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
            "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] " + msg,
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
                "蛟呵｣懷唆荳�: 繝√�ｼ繝�莠ｺ謨ｰ螟� team=%s size=%s req=%s max=%s",
                ",".join(str(x) for x in team),
                len(team),
                lo,
                hi,
            )
            return None
        op_list = [m for m in team if skill_role_priority(m)[0] == "OP"]
        if not op_list:
            _trace_assign(
                "蛟呵｣懷唆荳�: OP荳榊惠 team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        if not all(m in daily_status for m in team):
            _trace_assign(
                "蛟呵｣懷唆荳�: 蠖捺律蜍､諤�繧ｭ繝ｼ縺ｪ縺� team=%s",
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
                "蛟呵｣懷唆荳�: 髢句ｧ�>=邨よ･ｭ team=%s start=%s end_limit=%s",
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
                "蛟呵｣懷唆荳�: 莨第�ｩ蟶ｯ蜀�繝ｻ邨よ･ｭ逶ｴ蜑�(蟆乗ｮ�)縺ｧ蠖捺律荳榊庄 team=%s",
                ",".join(str(x) for x in team),
            )
            return None
        team_start = team_start_d
        if team_start >= team_end_limit:
            _trace_assign(
                "蛟呵｣懷唆荳�: 繝�繝輔ぃ繝ｼ蠕後↓髢句ｧ�>=邨よ･ｭ team=%s start=%s end_limit=%s",
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
                "蛟呵｣懷唆荳�: 螳溷ロ荳崎ｶｳ team=%s start=%s avail_mins=%s need_mins=%.2f",
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
                "蛟呵｣懷唆荳�: 邨よ･ｭ逶ｴ蜑阪〒蠖捺律蜿主ｮｹ繝ｭ繝ｼ繝ｫ謨ｰ縺碁明蛟､譛ｪ貅 team=%s cap=%s th=%s start=%s",
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
                "蛟呵｣懷唆荳�: 莨第�ｩ縺ｾ縺溘℃縺ｮ縺溘ａ騾｣邯壼ｮ溷ロ荳崎ｶｳ team=%s contiguous_min=%s need_mins=%s start=%s",
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
                "蛹ｺ髢薙Α繝ｩ繝ｼ蜊ｴ荳�: team=%s start=%s end=%s eq=%s",
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

    # 迚ｹ蛻･謖�螳�: 蜷御ｸ譌･繝ｻ騾｣邯壹Ο繝ｼ繝ｫ縺ｯ蜑榊屓繝√�ｼ繝�繧貞━蜈茨ｼ育ｿ梧律縺ｸ縺ｯ謖√■雜翫＆縺ｪ縺��ｼ峨�
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
    # 邨�縺ｿ蜷医ｏ縺幄｡ｨ繝励Μ繧ｻ繝�繝医�ｯ縲梧�千ｫ九＠縺溘ｉ蜊ｳ return縲阪○縺壹∫ｵ�蜷医○謗｢邏｢縺ｨ縺ｾ縺ｨ繧√※
    # team_start / 繧ｹ繝ｩ繝�繧ｯ莉倥″繧ｿ繝励Ν縺ｧ譛濶ｯ繧帝∈縺ｶ�ｼ医す繝ｼ繝井ｸ翫�ｮ蜆ｪ蜈亥ｺｦ鬆�縺ｯ隧ｦ陦碁��縺ｮ縺ｿ�ｼ峨�
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
                    "蛟呵｣懷崋螳�: 諡�蠖徹P謖�蜷�=%s 縺ｮ縺溘ａ 1莠ｺ繝√�ｼ繝�縺ｯ蠖謎ｺｺ縺ｮ縺ｿ隧ｦ陦�",
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
                "谿ｵ髫�2: 萓晞�ｼNO=%s 譌･莉�=%s 蟾･遞�/讖滓｢ｰ=%s/%s 縺ｧ繝√�ｼ繝�蛟呵｣懊′0莉ｶ縲�"
                "繧ｹ繧ｭ繝ｫ驕ｩ蜷�(OP/AS)縺ｯ %s 莠ｺ縺�縺ｾ縺吶′縲∬ｨｭ蛯吶�ｮ蜉�蟾･髢句ｧ倶ｸ矩剞=%s 縺�"
                "蠖捺律縺ｮ諡�蠖灘呵｣懊�ｮ騾蜍､(%s)莉･髯阪�ｮ縺溘ａ縺薙�ｮ譌･縺ｯ蜑ｲ蠖薙〒縺阪∪縺帙ｓ縲�"
                "master縲梧ｩ滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ縲阪〒蠖楢ｩｲ譌･繝ｻ蠖楢ｩｲ讖滓｢ｰ蛻励↓荳崎ｦ√↑險伜�･縺後↑縺�縺九�"
                "縺ｾ縺溘�ｯ蜑榊ｷ･遞九�ｮ蜊�譛峨〒險ｭ蛯吩ｸ矩剞縺檎ｵよ･ｭ縺ｾ縺ｧ郢ｰ繧贋ｸ翫′縺｣縺ｦ縺�縺ｪ縺�縺狗｢ｺ隱阪＠縺ｦ縺上□縺輔＞"
                "�ｼ磯�榊床繝ｫ繝ｼ繝ｫ 3.2.1 讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ繝ｻ繝医Λ繝悶Ν繧ｷ繝･繝ｼ繝茨ｼ峨�"
                "蜿り�: changeover蜑阪�ｮ險ｭ蛯咏ｩｺ縺堺ｸ矩剞=%s 蜊�譛峨く繝ｼ=%s",
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
    迴ｾ蝨ｨ縺ｮ avail_dt / machine_avail_dt / machine_handoff 縺ｮ繧ｹ繝翫ャ繝励す繝ｧ繝�繝医〒
    `_assign_one_roll_trial_order_flow` 縺� None 縺ｫ縺ｪ繧九↑繧� True縲�
    讖滓｢ｰ譫�縺ｯ蜊∝��縺ｧ繧ゆｺｺ繝ｻ莨第�ｩ繝ｻ繝溘Λ繝ｼ遲峨〒隧ｰ縺ｾ繧翫√げ繝ｭ繝ｼ繝舌Ν隧ｦ陦碁��縺�縺代′蜈磯�ｭ陦後↓蠑ｵ繧贋ｻ倥￥縺ｮ繧帝亟縺舌�
    蜑ｯ菴懃畑縺ｪ縺暦ｼ�need 莠ｺ謨ｰ繝ｭ繧ｰ逕ｨ set 縺ｯ豈主屓遨ｺ�ｼ峨�
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
            "trial_order_assign_probe 萓句､悶�ｮ縺溘ａ蠖楢ｩｲ陦後�ｯ髯､螟悶＠縺ｪ縺�: task=%s err=%s",
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
    """`_min_pending_dispatch_trial_order_for_date` 縺ｨ蜷御ｸ縺ｮ螳我ｾ｡繝輔ぅ繝ｫ繧ｿ繧帝夐℃縺励◆繧ｿ繧ｹ繧ｯ縺ｮ繝ｪ繧ｹ繝医�"""
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
    pool 繧呈��鬆� dto 縺ｧ隕九※縲�**縺昴�ｮ dto 縺ｫ螻槭☆繧玖｡後�ｮ縺�縺｡ 1 莉ｶ縺ｧ繧�** 1 繝ｭ繝ｼ繝ｫ蜑ｲ蠖薙�励Ο繝ｼ繝悶′騾壹ｌ縺ｰ
    縺昴�ｮ dto 繧偵悟ｮ溷柑縺ｮ譛蟆剰ｩｦ陦碁��縲阪→縺吶ｋ縲�
    蜈磯�ｭ dto 螻､縺悟�ｨ貊��ｼ域ｩ滓｢ｰ縺ｯ遨ｺ縺�縺ｦ縺�繧九′莠ｺ縺ｧ遨阪ａ縺ｪ縺�遲会ｼ峨�ｮ縺ｨ縺阪∵ｬ｡縺ｮ dto 縺ｫ騾ｲ縺ｿ繧ｰ繝ｭ繝ｼ繝舌Ν蛛懈ｭ｢繧帝亟縺舌�
    繝励Ο繝ｼ繝也┌縺励�ｮ縺ｨ縺阪�ｯ pool 縺ｮ譛蟆� dto 繧定ｿ斐☆縲�
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
    竭�蠖捺律蛟呵｣懊ｒ驟榊床隧ｦ陦碁��縺ｮ譏�鬆�縺ｫ荳ｦ縺ｹ繧具ｼ�1 繝代せ蛻��ｼ峨�
    **螳悟�ｨ莠檎嶌�ｼ按ｧB-2 / ﾂｧB-3�ｼ�**: **繝輔ぉ繝ｼ繧ｺ1**縺ｧ **蠕檎ｶ壹ヱ繧､繝励Λ繧､繝ｳ陦�**�ｼ育�ｱ陞咲捩讀懈渊繝ｻ蟾ｻ霑斐＠�ｼ�**繧帝勁縺�**蛟呵｣懶ｼ�EC繝ｻ莉紋ｾ晞�ｼ繝ｻ莉門ｷ･遞具ｼ峨ｒ隧ｦ陦碁��縺ｩ縺翫ｊ
    **`_drain_rolls_for_task`** 縺励�**繝輔ぉ繝ｼ繧ｺ2**縺ｯ ﾂｧB-2 讀懈渊�ｼ渉ｧB-3 蟾ｻ霑斐＠陦後�ｮ縺ｿ�ｼ�**蜷御ｸ萓晞�ｼ縺ｮ EC 縺悟�ｨ譌･縺ｧ螳瑚ｵｰ縺励◆蠕�**縺ｫ髯舌ｊ蛟呵｣懷喧縲�
    EC 谿九′縺ゅｋ譌･縺ｯ `_trial_order_flow_eligible_tasks` 縺ｧ蠕檎ｶ壹ｒ螟悶＠縲∫ｿ檎ｨｼ蜒肴律莉･髯阪ｂ EC 縺ｮ縺ｿ蜑埼ｲ縺吶ｋ縲�
    繧ｫ繝ｬ繝ｳ繝繝ｼ騾夂ｮ励〒 EC 螳瑚ｵｰ蠕後～_run_b2_inspection_rewind_pass` 縺梧律莉伜�磯�ｭ縺九ｉ蠕檎ｶ壹□縺大�崎ｵｰ譟ｻ縺吶ｋ�ｼ峨�
    EC 縺ｨ蠕檎ｶ壹ｒ **蜷御ｸ諡�蠖楢�縺ｧ** 莠､莠偵↓隧ｰ繧√ｋ縺ｨ EC 縺後ヶ繝ｭ繝�繧ｯ縺輔ｌ繧九◆繧√∝ｾ捺擂縺ｯ繝輔ぉ繝ｼ繧ｺ1繧貞�医↓隧ｰ繧√◆縲�
    縺溘□縺怜ｾ檎ｶ壹′蛟呵｣懷喧縺励◆譎らせ縺ｧ **讀懈渊縺ｨ蜷後§迚ｩ逅�讖滓｢ｰ**縺ｮ繝輔ぉ繝ｼ繧ｺ1繧� **蜷御ｸ萓晞�ｼ縺ｮ EC** 縺悟�ｨ譌･蜈医↓騾ｲ繧縺ｨ縲�
    讀懈渊縺ｯ `start_ge_end_initial`�ｼ郁ｨｭ蛯咏ｩｺ縺阪′邨よ･ｭ繧医ｊ蠕鯉ｼ峨〒蜈ｨ譌･螟ｱ謨励☆繧九つｧB-2/ﾂｧB-3 蠕檎ｶ壹′縺ゅｋ縺ｨ縺阪�ｯ
    縲悟酔荳萓晞�ｼEC繝ｻ讀懈渊讖溘→讖滓｢ｰ蜈ｱ譛峨☆繧九ヵ繧ｧ繝ｼ繧ｺ1繝ｻ蠕檎ｶ壹阪ｒ **驟榊床隧ｦ陦碁��**縺ｧ繝槭�ｼ繧ｸ縺励�
    蜷碁��縺ｧ縺ｯ **蠕檎ｶ壹ｒ EC 繧医ｊ蜈医↓**縲�**縺昴�ｮ莉悶�ｮ繝輔ぉ繝ｼ繧ｺ1** 縺ｨ縺ゅｏ縺帙※ **驟榊床隧ｦ陦碁��**縺ｧ謨ｴ蛻励＠
    **譛螟ｧ1繝ｭ繝ｼ繝ｫ縺壹▽**縺�縺大捉蝗槭☆繧具ｼ医�槭�ｼ繧ｸ繝ｻrest 縺ｨ繧ゆｸ諡ｬ繝峨Ξ繧､繝ｳ縺励↑縺�縲よ､懈渊OP縺御ｻ門ｷ･遞九↓
    蜷梧律蜿悶ｊ蛻�繧峨ｌ start_ge_end_initial 縺ｫ縺ｪ繧九�ｮ繧帝亟縺撰ｼ峨�
    繝ｪ繝ｯ繧､繝ｳ繝牙�ｴ縺ｮ蠕檎ｶ夊｡後�ｯ蜷�繝ｭ繝ｼ繝ｫ縺ｫ縺､縺�縺ｦ `_roll_pipeline_inspection_assign_room` 縺翫ｈ縺ｳ
    `_roll_pipeline_b2_inspection_ec_completion_floor_dt`�ｼ�EC 繝ｭ繝ｼ繝ｫ邨ゆｺ�譎ょ綾荳矩剞�ｼ峨〒謨ｴ蜷医☆繧九�
    隧ｦ陦碁��譛蟆上�ｮ陦後□縺代′蠖捺律蜈･繧峨↑縺�蝣ｴ蜷医〒繧ゅ�**蜷後§繝輔ぉ繝ｼ繧ｺ蜀�縺ｧ谺｡縺ｮ隧ｦ陦碁��縺ｸ騾ｲ縺ｿ**莉冶ｨｭ蛯吶ｒ蝓九ａ繧九�
    讖滓｢ｰ繝ｻ莠ｺ縺ｮ遨ｺ縺阪�ｯ繝ｭ繝ｼ繝ｫ縺斐→縺ｫ譖ｴ譁ｰ縺吶ｋ�ｼ遺即竭ｧ�ｼ峨�
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
                    "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繝ｭ繝ｼ繝ｫ遒ｺ螳� 繝｡繧､繝ｳ day=%s machine=%s machine_name=%s "
                    "start=%s end=%s 謗｡逕ｨ莠ｺ謨ｰ=%s req_num=%s 繝｡繧､繝ｳ謗｢邏｢extra_max=%s "
                    "菴吝臆莠ｺ謨ｰ驕ｩ逕ｨ(繝｡繧､繝ｳ)=%s team=%s",
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
        # 蜷後§驟榊床隧ｦ陦碁��縺ｧ縺ｯ蠕檎ｶ夲ｼ域､懈渊繝ｻ蟾ｻ霑斐＠�ｼ峨ｒ EC 繧医ｊ蜈医↓蝗槭＠縲∫�ｱ陞咲捩縺ｮ螢∵凾險医ｒ
        # 蜷梧律譌ｩ縺�谿ｵ髫弱〒蜿悶ｊ縺ｫ陦後￥�ｼ按ｧB-2 諡�蠖楢�蛻�髮｢縺ｧ EC 縺ｨ讀懈渊縺ｯ蛻･繝｡繝ｳ繝舌�ｼ諠ｳ螳夲ｼ峨�
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
    ﾂｧB-2 / ﾂｧB-3: EC 蛛ｴ繧貞�医↓蜈ｨ譌･縺ｧ騾ｲ繧√◆蠕後∵､懈渊�ｼ丞ｷｻ霑斐＠蛛ｴ縺ｮ縺ｿ繧呈律莉伜�磯�ｭ縺九ｉ蜀崎ｵｰ譟ｻ縺励※驟榊床縺吶ｋ縲�
    timeline_events 繧剃ｺｺ繝ｻ險ｭ蛯吶�ｮ繝悶Ο繝�繧ｯ繝�繝ｼ繝悶Ν縺ｨ縺励※菴ｿ縺�縲∵律霍ｨ縺弱�ｮ蜊�譛峨ｒ菫晄戟縺吶ｋ縲�
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
    need縲碁�榊床譎りｿｽ蜉�莠ｺ謨ｰ�ｼ丈ｽ吝鴨譎りｿｽ蜉�莠ｺ謨ｰ縲崎｡後�ｮ荳企剞縺ｾ縺ｧ縲√Γ繧､繝ｳ蜑ｲ莉倥〒謗｡逕ｨ縺励″繧後↑縺九▲縺滓棧繧定ｿｽ險倥☆繧九�
    蜷�繧ｿ繧､繝�繝ｩ繧､繝ｳ繝悶Ο繝�繧ｯ縺ｫ縺､縺�縺ｦ縲√◎縺ｮ譎る俣蟶ｯ縺ｫ莉悶ヶ繝ｭ繝�繧ｯ縺ｸ譛ｪ蜿ょ刈�ｼ亥玄髢馴㍾縺ｪ繧翫↑縺暦ｼ峨〒
    eligible 縺九▽ OP/AS 繧ｹ繧ｭ繝ｫ縺ｮ閠�繧偵し繝悶↓霑ｽ蜉�縺吶ｋ縲�
    譌･谺｡蟋区･ｭ繝ｻ萓晞�ｼ蛻�譖ｿ蠕悟ｧ区忰繝ｻ蜉�蟾･蜑肴ｺ門ｙ�ｼ�event_kind 縺悟刈蟾･莉･螟厄ｼ峨�ｯ譛ｬ蜃ｦ逅�縺ｮ蟇ｾ雎｡螟厄ｼ井ｽ吝臆繧ｵ繝悶�ｯ蜉�蟾･縺ｫ縺ｮ縺ｿ霑ｽ險假ｼ峨�
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
                "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 菴吝鴨霑ｽ險�(繝｡繧､繝ｳ螳御ｺ�蠕�) day=%s machine=%s machine_name=%s "
                "start=%s end=%s 霑ｽ險倅ｺｺ謨ｰ=%s 霑ｽ險伜燕莠ｺ謨ｰ=%s 霑ｽ險伜ｾ御ｺｺ謨ｰ=%s req_num=%s "
                "need霑ｽ蜉�譫�(繧ｷ繝ｼ繝�)=%s 螻･豁ｴ鮟�(菴吝臆莠ｺ謨ｰ雜�驕�)=%s 霑ｽ險倥Γ繝ｳ繝舌�ｼ=%s",
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
# 3. 繝｡繧､繝ｳ險育判逕滓�� (譌･豈弱Ν繝ｼ繝励�ｻ謖√■雜翫＠蟇ｾ蠢�)
#    谿ｵ髫�2縺ｮ譛ｬ菴薙Ｑlan_simulation_stage2 縺九ｉ縺ｮ縺ｿ蜻ｼ縺ｰ繧後ｋ諠ｳ螳壹�
#    驟榊床險育判繧ｷ繝ｼ繝郁ｪｭ霎ｼ 竊� 繧ｿ繧ｹ繧ｯ繧ｭ繝･繝ｼ 竊� 譌･莉倥＃縺ｨ縺ｫ險ｭ蛯吶�ｻOP蜑ｲ莉� 竊� 邨先棡繝悶ャ繧ｯ蜃ｺ蜉帙�
# =========================================================
def generate_plan():
    """
    谿ｵ髫�2縺ｮ繝｡繧､繝ｳ蜃ｦ逅�縲よ綾繧雁､縺ｪ縺暦ｼ医Ο繧ｰ繝ｻExcel 蜃ｺ蜉帙〒螳檎ｵ撰ｼ峨�

    蜑肴署: 迺ｰ蠅�螟画焚 TASK_INPUT_WORKBOOK縲√き繝ｬ繝ｳ繝医ョ繧｣繝ｬ繧ｯ繝医Μ縺後せ繧ｯ繝ｪ繝励ヨ繝輔か繝ｫ繝縲�
    蜃ｺ蜉�: ``output_dir`` 逶ｴ荳九�ｮ ``production_plan_multi_day_*.xlsx`` / ``member_schedule_*.xlsx``�ｼ域怙譁ｰ1邨�縺ｮ縺ｿ�ｼ峨√♀繧医�ｳ log/execution_log.txt縲�
    """
    master_abs = os.path.abspath(os.path.join(os.getcwd(), MASTER_FILE))
    with _override_default_factory_hours_from_master(master_abs):
        _generate_plan_impl()


def _generate_plan_impl():
    # 驟榊床繝医Ξ繝ｼ繧ｹ�ｼ郁ｨｭ螳壹す繝ｼ繝� A3 莉･髯阪�ｮ縺ｿ�ｼ峨�ｯ縲√Γ繝ｳ繝舌�ｼ0莠ｺ遲峨〒譌ｩ譛� return 縺励※繧�
    # execution_log 縺ｫ谿九ｋ繧医≧ skills 隱ｭ霎ｼ繧医ｊ蜑阪〒遒ｺ螳壹�ｻ繝ｭ繧ｰ縺吶ｋ縲�
    global TRACE_SCHEDULE_TASK_IDS, DEBUG_DISPATCH_ONLY_TASK_IDS
    _wb_trace = (os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK)
    _ids_from_sheet = _read_trace_schedule_task_ids_from_config_sheet(_wb_trace)
    TRACE_SCHEDULE_TASK_IDS = frozenset(
        str(x).strip() for x in _ids_from_sheet if str(x).strip()
    )
    if _ids_from_sheet:
        _preview = _ids_from_sheet[:25]
        _suffix = " 窶ｦ" if len(_ids_from_sheet) > 25 else ""
        logging.info(
            "險ｭ螳壹す繝ｼ繝医�%s縲喉3 莉･髯�: 繝医Ξ繝ｼ繧ｹ逕ｨ萓晞�ｼNO繧� %s 莉ｶ隱ｭ縺ｿ霎ｼ縺ｿ�ｼ�%s%s�ｼ�",
            APP_CONFIG_SHEET_NAME,
            len(_ids_from_sheet),
            ", ".join(_preview),
            _suffix,
        )
    else:
        logging.info(
            "險ｭ螳壹す繝ｼ繝医�%s縲喉3 莉･髯�: 繝医Ξ繝ｼ繧ｹ逕ｨ萓晞�ｼNO縺ｯ辟｡縺暦ｼ育ｩｺ縺ｾ縺溘�ｯ繧ｷ繝ｼ繝育┌縺暦ｼ�",
            APP_CONFIG_SHEET_NAME,
        )
    if TRACE_SCHEDULE_TASK_IDS:
        logging.info(
            "驟榊床繝医Ξ繝ｼ繧ｹ: 譛牙柑 task_id = %s�ｼ郁ｨｭ螳壹す繝ｼ繝� A3 莉･髯搾ｼ�",
            ", ".join(sorted(TRACE_SCHEDULE_TASK_IDS)),
        )
    else:
        logging.info(
            "驟榊床繝医Ξ繝ｼ繧ｹ: 蟇ｾ雎｡縺ｪ縺暦ｼ�[驟榊床繝医Ξ繝ｼ繧ｹ 窶ｦ] 繝ｭ繧ｰ縺ｯ蜃ｺ縺ｾ縺帙ｓ�ｼ�"
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
            "繝�繝舌ャ繧ｰ驟榊床: 縲�%s縲坑3莉･髯阪↓繧医ｊ驟榊床蟇ｾ雎｡繧� %s 莉ｶ縺ｮ萓晞�ｼNO縺ｫ髯仙ｮ壹＠縺ｾ縺�: %s",
            APP_CONFIG_SHEET_NAME,
            len(DEBUG_DISPATCH_ONLY_TASK_IDS),
            ", ".join(sorted(DEBUG_DISPATCH_ONLY_TASK_IDS)),
        )
        _show_stage2_debug_dispatch_mode_dialog(sorted(DEBUG_DISPATCH_ONLY_TASK_IDS))
    if TRACE_TEAM_ASSIGN_TASK_ID:
        logging.info(
            "迺ｰ蠅�螟画焚 TRACE_TEAM_ASSIGN_TASK_ID=%r 竊� 繝√�ｼ繝�蜑ｲ蠖薙ヨ繝ｬ繝ｼ繧ｹ譛牙柑",
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
            "邨�縺ｿ蜷医ｏ縺幄｡ｨ: 蟾･遞�+讖滓｢ｰ繧ｭ繝ｼ %s 遞ｮ鬘槭�ｻ邱ｨ謌占｡� %s 繧帝�榊床繝励Μ繧ｻ繝�繝医→縺励※隱ｭ縺ｿ霎ｼ縺ｿ縺ｾ縺励◆縲�",
            len(team_combo_presets),
            _nrules,
        )
    elif TEAM_ASSIGN_USE_MASTER_COMBO_SHEET:
        logging.info(
            "邨�縺ｿ蜷医ｏ縺幄｡ｨ: 繝励Μ繧ｻ繝�繝育┌縺暦ｼ医す繝ｼ繝域ｬ�螯ゅ�ｻ遨ｺ繝ｻ縺ｾ縺溘�ｯ隱ｭ霎ｼ螟ｱ謨暦ｼ峨ょｾ捺擂縺ｮ繝√�ｼ繝�謗｢邏｢縺ｮ縺ｿ縲�"
        )
    if not members:
        master_abs = os.path.abspath(MASTER_FILE)
        logging.error(
            "谿ｵ髫�2繧剃ｸｭ譁ｭ縺励∪縺励◆: 繝｡繝ｳ繝舌�ｼ縺�0莠ｺ縺ｧ縺呻ｼ医�槭せ繧ｿ縺ｮ skills 縺檎ｩｺ縲√∪縺溘�ｯ隱ｭ縺ｿ霎ｼ縺ｿ螟ｱ謨暦ｼ峨�"
            " 譛溷ｾ�繝代せ: %s �ｼ医き繝ｬ繝ｳ繝�: %s�ｼ峨ゅユ繧ｹ繝医さ繝ｼ繝臥峩荳九↓ master.xlsm 繧堤ｽｮ縺阪�"
            "planning_core 縺ｮ繧ｫ繝ｬ繝ｳ繝医′縺昴�ｮ繝輔か繝ｫ繝縺ｫ縺ｪ繧九ｈ縺� python\\ 驟咲ｽｮ繧堤｢ｺ隱阪＠縺ｦ縺上□縺輔＞縲�"
            " 縺薙�ｮ迥ｶ諷九〒縺ｯ production_plan / member_schedule 縺ｯ蜃ｺ蜉帙＆繧後∪縺帙ｓ縲�",
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
            "讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ: 隱ｭ霎ｼ萓句､悶�ｮ縺溘ａ蜊�譛峨↑縺励→縺励※邯夊｡後＠縺ｾ縺� (%s)", e
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
            "讖滓｢ｰ貅門ｙ/萓晞�ｼ蛻�譖ｿ繝ｻ譌･谺｡蟋区･ｭ險ｭ螳�: 隱ｭ霎ｼ萓句､悶�ｮ縺溘ａ辟｡隕悶＠縺ｾ縺� (%s)", e
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
                "譌･谺｡蟋区･ｭ貅門ｙ: 螳壼ｸｸ髢句ｧ� master 繝｡繧､繝ｳ A15=%s 繧呈治逕ｨ�ｼ�[髢句ｧ�, 髢句ｧ�+蛻�) 繧貞｣∵凾險医〒蜊�譛峨�A15 辟｡蜉ｹ譎ゅ�ｯ蠕捺擂縺ｮ蜍､諤� forward�ｼ�",
                _rs_a15.strftime("%H:%M"),
            )
    except Exception as e:
        logging.warning("螳壼ｸｸ髢句ｧ�(A15) 隱ｭ霎ｼ螟ｱ謨�: 譌･谺｡蟋区･ｭ縺ｯ蠕捺擂縺ｮ蜍､諤� forward 縺ｫ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ (%s)", e)
        _STAGE2_REGULAR_SHIFT_START = None
    if _MACHINE_CALENDAR_BLOCKS_BY_DATE:
        _n_iv = sum(
            len(ivs)
            for _dm in _MACHINE_CALENDAR_BLOCKS_BY_DATE.values()
            for ivs in _dm.values()
        )
        logging.info(
            "讖滓｢ｰ繧ｫ繝ｬ繝ｳ繝繝ｼ: %s 譌･蛻�繝ｻ險ｭ蛯吝頃譛峨ヶ繝ｭ繝�繧ｯ險� %s 繧帝�榊床縺ｫ蜿肴丐縺励∪縺吶�",
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
            "need驟榊床譎りｿｽ蜉�莠ｺ謨ｰ: 繝｡繧､繝ｳ蜑ｲ莉倥�ｯ蝓ｺ譛ｬ蠢�隕∽ｺｺ謨ｰ縺ｮ縺ｿ縲�"
            "菴吝鴨縺ｯ蜈ｨ繧ｷ繝溘Η繝ｬ繝ｼ繧ｷ繝ｧ繝ｳ蠕後∵凾髢馴㍾縺ｪ繧翫�ｮ縺ｪ縺�譛ｪ蜑ｲ蠖薙°縺､繧ｹ繧ｭ繝ｫ驕ｩ蜷郁�繧偵し繝悶↓霑ｽ險倥＠縺ｾ縺吶�"
            "�ｼ医Γ繧､繝ｳ縺ｧ蠅怜藤謗｢邏｢縺吶ｋ蠕捺擂謖吝虚: TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS=1�ｼ�"
        )

    # 谿ｵ髫�2縺ｮ蝓ｺ貅匁律譎ゅ�ｯ縲後�槭け繝ｭ螳溯｡梧凾蛻ｻ縲阪〒縺ｯ縺ｪ縺上後ョ繝ｼ繧ｿ謚ｽ蜃ｺ譌･縲阪ｒ菴ｿ逕ｨ
    data_extract_dt = _extract_data_extraction_datetime()
    base_now_dt = data_extract_dt if data_extract_dt is not None else datetime.now()
    run_date = base_now_dt.date()
    data_extract_dt_str = (
        base_now_dt.strftime("%Y/%m/%d %H:%M:%S") if data_extract_dt is not None else "窶�"
    )
    logging.info(
        "險育判蝓ｺ貅匁律譎�: %s�ｼ�%s�ｼ�",
        base_now_dt.strftime("%Y/%m/%d %H:%M:%S"),
        "繝�繝ｼ繧ｿ謚ｽ蜃ｺ譌･" if data_extract_dt is not None else "迴ｾ蝨ｨ譎ょ綾繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ",
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
            "繝｡繧､繝ｳ繝ｻ繧ｰ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝�: 蟾･蝣ｴ莨第･ｭ謇ｱ縺�縺ｮ譌･莉� 竊� %s",
            ", ".join(str(x) for x in sorted(_factory_closure_dates)),
        )
    ai_log_data["繝｡繧､繝ｳ_繧ｰ繝ｭ繝ｼ繝舌Ν_蟾･蝣ｴ莨第･ｭ譌･(隗｣譫�)"] = (
        ", ".join(str(x) for x in sorted(_factory_closure_dates))
        if _factory_closure_dates
        else "�ｼ医↑縺暦ｼ�"
    )
    _sn = str(global_priority_override.get("scheduler_notes_ja") or "").strip()
    if _sn:
        ai_log_data["繝｡繧､繝ｳ_繧ｰ繝ｭ繝ｼ繝舌Ν_譛ｪ驕ｩ逕ｨ繝｡繝｢(AI)"] = _sn[:2000]

    sorted_dates = sorted(list(attendance_data.keys()))
    # 邨先棡繧ｷ繝ｼ繝医�ｯ縲悟渕貅匁律�ｼ医ョ繝ｼ繧ｿ謚ｽ蜃ｺ譌･�ｼ峨堺ｻ･髯阪�ｮ縺ｿ陦ｨ遉ｺ繝ｻ險育判蟇ｾ雎｡縺ｨ縺吶ｋ
    sorted_dates = [d for d in sorted_dates if d >= run_date]
    if not sorted_dates:
        logging.error("蠖捺律莉･髯阪�ｮ蜃ｦ逅�蟇ｾ雎｡譌･莉倥′縺ゅｊ縺ｾ縺帙ｓ縲�")
        _try_write_main_sheet_gemini_usage_summary("谿ｵ髫�2")
        return

    # 繧ｿ繧ｹ繧ｯ蜈･蜉�: 繝悶ャ繧ｯ蜀�縲碁�榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙搾ｼ域ｮｵ髫�1縺ｧ蜃ｺ蜉帚�貞叙繧願ｾｼ縺ｿ蠕後↓邱ｨ髮��ｼ�
    try:
        tasks_df = load_planning_tasks_df()
    except Exception as e:
        logging.error(f"驟榊床險育判繧ｿ繧ｹ繧ｯ繧ｷ繝ｼ繝郁ｪｭ縺ｿ霎ｼ縺ｿ繧ｨ繝ｩ繝ｼ: {e}")
        _try_write_main_sheet_gemini_usage_summary("谿ｵ髫�2")
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
            "繝�繝舌ャ繧ｰ驟榊床: 縲�%s縲阪�ｮ陦後ｒ %s 竊� %s 縺ｫ邨槭ｊ霎ｼ縺ｿ縺ｾ縺励◆縲�",
            PLAN_INPUT_SHEET_NAME,
            _n_tasks_before,
            _n_tasks_after,
        )
        if _n_tasks_after == 0:
            logging.error(
                "繝�繝舌ャ繧ｰ驟榊床: B3莉･髯阪�ｮ萓晞�ｼNO縺ｫ荳閾ｴ縺吶ｋ陦後′縺ゅｊ縺ｾ縺帙ｓ縲よｮｵ髫�2繧剃ｸｭ譁ｭ縺励∪縺吶�"
            )
            _try_write_main_sheet_gemini_usage_summary("谿ｵ髫�2")
            return

    if global_priority_raw.strip():
        snip = global_priority_raw[:2500]
        if len(global_priority_raw) > 2500:
            snip += "窶ｦ"
        ai_log_data["繝｡繧､繝ｳ_蜀榊━蜈育音蛻･險倩ｼ�(蜴滓枚)"] = snip
    else:
        ai_log_data["繝｡繧､繝ｳ_蜀榊━蜈育音蛻･險倩ｼ�(蜴滓枚)"] = (
            "�ｼ育ｩｺ縲√∪縺溘�ｯ繝｡繧､繝ｳ繧ｷ繝ｼ繝医↓縲後げ繝ｭ繝ｼ繝舌Ν繧ｳ繝｡繝ｳ繝医崎ｦ句�ｺ縺励′隕九▽縺九ｊ縺ｾ縺帙ｓ�ｼ�"
        )
    ai_log_data["繝｡繧､繝ｳ_蜀榊━蜈育音蛻･險倩ｼ�(AI)"] = json.dumps(
        global_priority_override, ensure_ascii=False
    )
    if global_priority_override.get("ignore_skill_requirements"):
        logging.warning(
            "繝｡繧､繝ｳ蜀榊━蜈育音險�: 繧ｹ繧ｭ繝ｫ隕∽ｻｶ繧堤┌隕悶＠縺ｦ驟榊床縺励∪縺吶�%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("ignore_need_minimum"):
        logging.warning(
            "繝｡繧､繝ｳ蜀榊━蜈育音險�: 繝√�ｼ繝�莠ｺ謨ｰ繧�1蜷阪↓蝗ｺ螳壹＠縺ｾ縺呻ｼ�need繝ｻ陦後�ｮ蠢�隕＾P荳頑嶌縺阪ｈ繧雁━蜈茨ｼ峨�%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("abolish_all_scheduling_limits"):
        logging.warning(
            "繝｡繧､繝ｳ蜀榊━蜈育音險�: 險ｭ蛯吝ｰよ怏繝ｻ蜴溷渚蜷梧律髢句ｧ九�ｻ謖�螳夐幕蟋区凾蛻ｻ繝ｻ繝槭け繝ｭ螳溯｡梧凾蛻ｻ荳矩剞繧帝←逕ｨ縺励∪縺帙ｓ縲�%s",
            global_priority_override.get("interpretation_ja", ""),
        )

    # 縲悟ｽ捺律縲榊愛螳壹→譛譌ｩ髢句ｧ区凾蛻ｻ縺ｫ縺ｯ蝓ｺ貅匁律譎ゑｼ医ョ繝ｼ繧ｿ謚ｽ蜃ｺ譌･�ｼ峨ｒ菴ｿ縺�
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
    # 髢句ｧ区律縺碁撼遞ｼ蜒肴律縺ｮ蝣ｴ蜷医�ｯ縲∫峩蜑阪�ｮ遞ｼ蜒肴律縺ｸ陬懈ｭ｣�ｼ井ｾ�: 4/4, 4/5 縺碁撼遞ｼ蜒阪↑繧� 4/3 縺ｸ�ｼ�
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
                        "DEBUG[task=%s] start_date_req 繧帝撼遞ｼ蜒肴律陬懈ｭ｣: %s -> %s",
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
            f"譛牙柑縺ｪ繧ｿ繧ｹ繧ｯ縺後≠繧翫∪縺帙ｓ縲ゅ鶏PLAN_INPUT_SHEET_NAME}縲阪�ｮ縲御ｾ晞�ｼNO縲阪悟ｷ･遞句錐縲阪梧鋤邂玲焚驥上阪�"
            "縺ｾ縺溘�ｯ螳御ｺ�蛹ｺ蛻�繝ｻ螳溷�ｺ譚･鬮俶鋤邂励↓繧医ｊ谿矩㍼縺檎┌縺�陦後�ｮ縺ｿ縺ｮ蜿ｯ閭ｽ諤ｧ縺後≠繧翫∪縺吶�"
        )

    # 驟榊床隧ｦ陦碁��: 繧ｷ繝ｼ繝亥�励′謠�縺｣縺ｦ縺�繧後�ｰ縺昴ｌ繧呈治逕ｨ縲よｬ�謳肴凾縺ｯ ﾂｧB 蟶ｯ繝ｻ邏肴悄繝ｻneed 蛻鈴��縺ｧ繧ｽ繝ｼ繝医＠ EC 髫｣謗･蠕後↓ 1..n
    _apply_dispatch_trial_order_for_generate_plan(
        task_queue, req_map, need_rules, need_combo_col_index
    )
    if DEBUG_TASK_ID:
        dbg_items = [t for t in task_queue if str(t.get("task_id", "")).strip() == DEBUG_TASK_ID]
        if dbg_items:
            t0 = dbg_items[0]
            logging.info(
                "DEBUG[task=%s] queue蝓ｺ貅�: start_date_req=%s due_basis=%s answer_due=%s specified_due=%s specified_due_ov=%s due_source=%s priority=%s in_progress=%s remark=%s",
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
            logging.info("DEBUG[task=%s] task_queue縺ｫ蟄伜惠縺励∪縺帙ｓ�ｼ亥ｮ御ｺ�/谿矩㍼0/萓晞�ｼNO荳堺ｸ閾ｴ縺ｮ蜿ｯ閭ｽ諤ｧ�ｼ峨�", DEBUG_TASK_ID)
    timeline_events = []

    # ---------------------------------------------------------
    # 譌･豈弱�ｮ繧ｹ繧ｱ繧ｸ繝･繝ｼ繝ｪ繝ｳ繧ｰ繝ｫ繝ｼ繝�
    # STAGE2_EXTEND_ATTENDANCE_CALENDAR 縺� True 縺ｮ縺ｨ縺阪�ｮ縺ｿ縲∵ｮ九ち繧ｹ繧ｯ縺後≠繧後�ｰ蜍､諤�繧呈律莉倩､�陬ｽ縺ｧ諡｡蠑ｵ縲�
    # STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING 縺� True 縺ｮ縺ｨ縺阪�ｮ縺ｿ: 邏肴悄蝓ｺ貅悶ｒ驕弱℃縺ｦ繧よｮ九′縺ゅｋ萓晞�ｼ縺ｫ縺､縺�縺ｦ
    # due_basis +1繝ｻ蠖楢ｩｲ萓晞�ｼ縺ｮ蜑ｲ蠖捺綾縺励�ｻ蜈磯�ｭ縺九ｉ蜀榊ｮ溯｡後ょ推蜀崎ｩｦ陦悟燕縺ｫ蜍､諤�諡｡蠑ｵ蛻�縺ｯ繝槭せ繧ｿ譌･莉倥∈蟾ｻ縺肴綾縺吶�
    # 譌｢螳� False 縺ｮ縺溘ａ騾壼ｸｸ縺ｯ 1 繝代せ�ｼ医き繝ｬ繝ｳ繝繝ｼ騾壹＠ 1 蝗橸ｼ峨�ｮ縺ｿ縲�
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
            "DISPATCH_INTERVAL_MIRROR_ENFORCE: 險ｭ蛯吶�ｻ莠ｺ縺ｮ蜊�譛峨ｒ蛹ｺ髢薙Α繝ｩ繝ｼ縺ｧ霑ｽ霍｡縺励∪縺�"
            "�ｼ育┌蜉ｹ蛹悶�ｯ 險ｭ螳喟迺ｰ蠅�螟画焚 遲峨〒 DISPATCH_INTERVAL_MIRROR_ENFORCE=0�ｼ峨�"
        )

    if STAGE2_SERIAL_DISPATCH_BY_TASK_ID:
        logging.info(
            "萓晞�ｼNO逶ｴ蛻鈴�榊床: 譛牙柑�ｼ�STAGE2_SERIAL_DISPATCH_BY_TASK_ID�ｼ峨�"
            " 蜷�譌･縺ｯ繧｢繧ｯ繝�繧｣繝悶↑萓晞�ｼNO縺ｮ陦後□縺代′蛟呵｣懊�ｮ縺溘ａ縲∝ｽ楢ｩｲ萓晞�ｼ縺瑚ｩｰ縺ｾ繧九→莉紋ｾ晞�ｼ縺ｯ荳蛻�騾ｲ縺ｿ縺ｾ縺帙ｓ縲�"
        )
    else:
        logging.info(
            "萓晞�ｼNO逶ｴ蛻鈴�榊床: 辟｡蜉ｹ縲Ｔtart_date 繧呈ｺ縺溘☆蜈ｨ陦後′蠖捺律蛟呵｣懊↓縺ｪ繧翫�驟榊床隧ｦ陦碁��繝ｻ險ｭ蛯吶Ν繝ｼ繝ｫ縺ｧ鬆�蠎丈ｻ倥￠縺励∪縺吶�"
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
            # 險ｭ蛯吶＃縺ｨ縺ｮ遨ｺ縺肴凾蛻ｻ�ｼ亥酔荳險ｭ蛯吶�ｮ蜷梧凾荳ｦ陦悟牡蠖薙ｒ髦ｲ豁｢�ｼ�
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
                logging.info("DEBUG[day=%s] 遞ｼ蜒阪Γ繝ｳ繝舌�ｼ0縺ｮ縺溘ａ蜑ｲ莉倥せ繧ｭ繝�繝�", current_date)
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
                    "萓晞�ｼNO逶ｴ蛻鈴�榊床 day=%s 繧｢繧ｯ繝�繧｣繝紋ｾ晞�ｼNO=%s 逶ｴ蛻励Μ繧ｹ繝井ｽ咲ｽｮ=%s/%s "
                    "蠖捺律蛟呵｣懆｡梧焚(逶ｴ蛻怜燕)=%s 逶ｴ蛻怜ｾ�=%s 繧ｭ繝･繝ｼ谿玖｡�(蜈ｨ譌･)=%s",
                    current_date,
                    _active_serial_tid if _active_serial_tid is not None else "窶�",
                    _serial_pos if _serial_pos else "窶�",
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
                    "DEBUG[day=%s] 蜑ｲ莉伜ｯｾ雎｡繧ｿ繧ｹ繧ｯ0莉ｶ pending_total=%s earliest_start_date_req=%s",
                    current_date,
                    pending_total,
                    earliest_wait,
                )
            elif DEBUG_TASK_ID:
                has_dbg_today = any(str(t.get("task_id", "")).strip() == DEBUG_TASK_ID for t in tasks_today)
                if current_date.isoformat() == "2026-04-03" or has_dbg_today:
                    logging.info(
                        "DEBUG[day=%s] avail_members=%s tasks_today=%s (task=%s 蜷ｫ繧=%s)",
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
                                    "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繧ｹ繧ｭ繝�繝�: 蜷御ｸ萓晞�ｼNO縺ｮ蜈郁｡悟ｷ･遞句ｾ�縺｡ day=%s machine=%s rem=%.4f",
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
                                    "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繧ｹ繧ｭ繝�繝�: ﾂｧB-2/ﾂｧB-3 蠕檎ｶ壹Ο繝ｼ繝ｫ譫�繧ｼ繝ｭ day=%s machine=%s "
                                    "ec邏ｯ險亥ｮ御ｺ�R=%.4f 蠕檎ｶ夂ｴｯ險亥ｮ御ｺ�R=%.4f rem_follower=%.4f",
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
                                        "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繧ｹ繧ｭ繝�繝�: 蜷御ｸ險ｭ蛯吶�ｮ讀懈渊蜊�譛我ｸｭ day=%s "
                                        "蜊�譛芽�萓晞�ｼNO=%s 蜊�譛芽�隧ｦ陦碁��=%s",
                                        task.get("task_id"),
                                        current_date,
                                        _b1_holder.get("task_id"),
                                        _b1_holder.get("dispatch_trial_order"),
                                    )
                                continue
                        if DEBUG_TASK_ID and str(task.get("task_id", "")).strip() == DEBUG_TASK_ID:
                            logging.info(
                                "DEBUG[task=%s] day=%s 髢句ｧ句愛螳�: start_date_req=%s remaining_units=%s machine=%s",
                                DEBUG_TASK_ID,
                                current_date,
                                task.get("start_date_req"),
                                task.get("remaining_units"),
                                task.get("machine"),
                            )
                        if task.get("has_done_deadline_override"):
                            logging.info(
                                "DEBUG[螳御ｺ�譌･謖�螳咯 萓晞�ｼNO=%s 譌･莉�=%s start_date_req=%s due_basis=%s 謖�螳夂ｴ肴悄(荳頑嶌縺�)=%s 騾ｲ謐�=%s/%s",
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
                                    "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繧ｹ繧ｭ繝�繝�: 繧医ｊ蟆上＆縺�驟榊床隧ｦ陦碁��縺ｫ譛ｪ螳御ｺ�縺ゅｊ "
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
                                    "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繧ｹ繧ｭ繝�繝�: 蜷御ｸ險ｭ蛯吶〒驟榊床隧ｦ陦碁��縺悟�医�ｮ陦後′譛ｪ螳御ｺ� "
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
                                    (need_src_line + "�ｼ�") if need_src_line else ""
                                )
                                need_src_line += (
                                    f"險育判繧ｷ繝ｼ繝亥ｿ�隕∽ｺｺ謨ｰ{plan_ro}縺ｯ譛ｪ菴ｿ逕ｨ�ｼ�need蝓ｺ貅�={req_num}�ｼ�"
                                )
                        else:
                            if plan_ro is not None:
                                req_num = plan_ro
                                need_src_line = f"險育判繧ｷ繝ｼ繝医悟ｿ�隕＾P(荳頑嶌)縲�={req_num}"
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
                                (need_src_line + " 竊� ")
                                if need_src_line
                                else ""
                            ) + "繝｡繧､繝ｳ荳頑嶌ignore_need_minimum縺ｧreq=1"
    
                        # 繝｡繝ｳ繝舌�ｼﾃ苓ｨｭ蛯吶せ繧ｭ繝ｫ�ｼ�parse_op_as_skill_cell: 蟆上＆縺�蜆ｪ蜈亥ｺｦ縺ｻ縺ｩ蜈医↓繝√�ｼ繝�蛟呵｣懊∈謗｡逕ｨ�ｼ�
                        # skills 隱ｭ霎ｼ譎ゅ↓縲梧ｩ滓｢ｰ蜷阪榊腰迢ｬ繧ｭ繝ｼ縺ｸ繧ｨ繧､繝ｪ繧｢繧ｹ縺吶ｋ縺溘ａ縲∝ｷ･遞句錐+讖滓｢ｰ蜷阪′荳｡譁ｹ縺ゅｋ陦後〒縺ｯ
                        # 隍�蜷医く繝ｼ縲悟ｷ･遞句錐+讖滓｢ｰ蜷阪阪�ｮ縺ｿ繧定ｦ九ｋ�ｼ亥挨蟾･遞九�ｮ蜷悟錐讖滓｢ｰ縺ｮ OP 縺梧ｵ√ｌ霎ｼ縺ｾ縺ｪ縺�繧医≧縺ｫ縺吶ｋ�ｼ峨�
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
                                "DEBUG[螳御ｺ�譌･謖�螳咯 萓晞�ｼNO=%s 險ｭ蛯�=%s req_num=%s capable_members=%s machine_free=%s",
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
                                "諡�蠖徹P謖�蜷�: 蠖捺律縺ｮOP蛟呵｣懊↓荳閾ｴ縺帙★蛻ｶ邏�縺ｪ縺� task=%s raw=%r",
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
                                "繝｡繧､繝ｳ繧ｰ繝ｭ繝ｼ繝舌Ν(譌･莉佚怜ｷ･遞�): task=%s date=%s 蟾･遞�=%r 繝√�ｼ繝�蠢�鬆�=%s",
                                task.get("task_id"),
                                current_date,
                                machine,
                                ",".join(_gdp_must),
                            )
                        if fixed_team_anchor:
                            _nfix = len(fixed_team_anchor)
                            if _nfix > req_num:
                                need_src_line = (
                                    (need_src_line + " 竊� ")
                                    if need_src_line
                                    else ""
                                )
                                need_src_line += (
                                    f"繧ｰ繝ｭ繝ｼ繝舌Ν(譌･莉佚怜ｷ･遞�)謖�蜷阪〒譛菴施_nfix}莠ｺ"
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
                                (extra_src_line + " 竊� ")
                                if extra_src_line
                                else ""
                            ) + "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW縺ｧ0"
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
                                (extra_src_line + " 竊� ")
                                if extra_src_line
                                else ""
                            ) + "繝｡繧､繝ｳ縺ｯ蝓ｺ譛ｬ莠ｺ謨ｰ縺ｮ縺ｿ�ｼ井ｽ吝鴨譫�縺ｯ蜈ｨ驟榊床蠕後↓譛ｪ蜑ｲ蠖禿励せ繧ｭ繝ｫ縺ｧ霑ｽ險假ｼ�"
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
                                "need莠ｺ謨ｰ(驟榊床隧ｦ陦碁��蛻晏屓) order=%s task=%s 蟾･遞�/讖滓｢ｰ=%s/%s "
                                "req_num=%s [%s] extra_max=%s [%s] max_team蛟呵｣�=%s capable=%s莠ｺ",
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
                                "TRACE驟榊床[%s] %s 蟾･遞�/讖滓｢ｰ=%s / %s req_num=%s extra_max=%s 竊� max_team=%s "
                                "capable(n=%s)=%s ignore_need1=%s ignore_skill=%s abolish=%s 諡�蠖徹P謖�螳�=%r竊�%s",
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
                        # 繝励Μ繧ｻ繝�繝医�ｯ謌千ｫ句��繧偵☆縺ｹ縺ｦ蛟呵｣懊↓霈峨○縲∽ｸ九�ｮ邨�蜷医○謗｢邏｢縺ｨ縺ｾ縺ｨ繧√※譛濶ｯ繧帝∈縺ｶ縲�
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
                                        "諡�蠖徹P謖�蜷�: 繝√�ｼ繝�莠ｺ謨ｰ繧呈ｺ縺溘○縺ｪ縺�縺溘ａ謖�蜷阪ｒ辟｡隕� task=%s size=%s raw=%r",
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
                                    # 蜷御ｸ險ｭ蛯吶�ｯ1譎らせ縺ｧ1繧ｿ繧ｹ繧ｯ縺ｮ縺ｿ�ｼ郁ｨｭ蛯咏ｩｺ縺搾ｼ区律谺｡蟋区･ｭ/萓晞�ｼ蛻�譖ｿ縺ｮ貅門ｙ繝ｻ蠕悟ｧ区忰�ｼ�
                                    machine_free_dt = _mach_floor_legacy
                                    if team_start < machine_free_dt:
                                        team_start = machine_free_dt
                                    # 蜴溷渚謚募�･譌･縺ｨ蜷梧律縺ｮ髢句ｧ九�ｯ 13:00 莉･髯搾ｼ郁ｩｦ陦碁��蜆ｪ蜈医ヵ繝ｭ繝ｼ縺ｨ荳閾ｴ�ｼ�
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
                                    # 蠖捺律縺ｯ縲後�槭け繝ｭ螳溯｡後＠縺滓凾蛻ｻ縲阪ｈ繧雁燕縺ｫ髢句ｧ九〒縺阪↑縺�
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
                                        "TRACE驟榊床[%s] %s tsize=%s 竊� 縺薙�ｮ莠ｺ謨ｰ縺ｧ謌千ｫ九☆繧九メ繝ｼ繝�縺ｪ縺�",
                                        tid,
                                        current_date,
                                        tsize,
                                    )
                                else:
                                    sm = min(sub, key=_team_cand_key)
                                    logging.info(
                                        "TRACE驟榊床[%s] %s tsize=%s 莠ｺ謨ｰ蜀�譛濶ｯ: members=%s "
                                        "start=%s units_today=%s prio_sum=%s eff_t/unit=%.6f "
                                        "豈碑ｼ�繝ｫ繝ｼ繝ｫ=%s 窶ｻ蜈ｨ譌･譛譌ｩ髢句ｧ�=%s 繧貞渕貅悶↓霎樊嶌蠑上〒蟆上＆縺�譁ｹ縺梧治逕ｨ",
                                        tid,
                                        current_date,
                                        tsize,
                                        sm["team"],
                                        sm["team_start"],
                                        sm["units_today"],
                                        sm["prio_sum"],
                                        sm["eff_time_per_unit"],
                                        _tk,
                                        t_min.isoformat(sep=" ") if t_min else "窶�",
                                    )
    
                        if trace_assign and best_team is not None:
                            logging.info(
                                "TRACE驟榊床[%s] %s 笘�謗｡逕ｨ n=%s members=%s start=%s units_today=%s prio_sum=%s",
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
                                        "TRACE驟榊床[%s] %s 1莠ｺ謗｡逕ｨ�ｼ�TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF�ｼ�: "
                                        "繧医ｊ螟ｧ縺阪＞莠ｺ謨ｰ縺ｧ譛牙柑縺ｪ繝√�ｼ繝�縺ｪ縺暦ｼ�OP荳崎ｶｳ繝ｻ0蜊倅ｽ阪�ｻ髢句ｧ�>=邨ゆｺ�遲会ｼ峨�",
                                        task["task_id"],
                                        current_date,
                                    )
                                else:
                                    logging.info(
                                        "TRACE驟榊床[%s] %s 1莠ｺ謗｡逕ｨ: 莠ｺ謨ｰ繧貞｢励ｄ縺吶→髢句ｧ九′驕�繧後�"
                                        "繧ｹ繝ｩ繝�繧ｯ螟悶〒縺ｯ髢句ｧ句━蜈医〒1莠ｺ縺碁∈縺ｰ繧後◆蜿ｯ閭ｽ諤ｧ縲�"
                                        "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES=%s縲√∪縺溘�ｯ蠕捺擂縺ｮ莠ｺ謨ｰ譛蜆ｪ蜈医�ｯ迺ｰ蠅�螟画焚蜿ら�ｧ縲�",
                                        task["task_id"],
                                        current_date,
                                        TEAM_ASSIGN_START_SLACK_WAIT_MINUTES,
                                    )
    
                        if best_team:
                            if len(best_team) > req_num:
                                logging.info(
                                    "驟榊床謗｡逕ｨ莠ｺ謨ｰ>req_num task=%s day=%s order=%s 蟾･遞�/讖滓｢ｰ=%s/%s "
                                    "謗｡逕ｨ=%s莠ｺ req_num=%s extra_max=%s max_team=%s [%s] [%s]",
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
                                        "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繧ｹ繧ｭ繝�繝�: 繝√�ｼ繝�謗｡逕ｨ蠕後�ｮ螳溷柑繝ｦ繝九ャ繝�0 "
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
                            
                            # 縲後�槭け繝ｭ螳溯｡梧凾轤ｹ縲阪�ｮ螳御ｺ�邇��ｼ井ｺ亥ｮ壹�ｮ騾ｲ謐励〒縺ｯ縺ｪ縺上∝ｮ溷刈蟾･謨ｰ繝吶�ｼ繧ｹ�ｼ�
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
                                    "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繧ｿ繧､繝�繝ｩ繧､繝ｳ霑ｽ險� chunk day=%s machine=%s "
                                    "done_units=%s already_done=%s total_u=%s rem_after=%.4f "
                                    "start=%s end=%s eff_t/unit=%.4f rp_room(蠖捺凾)=%s",
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
                                    "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繝ｭ繝ｼ繝ｫ遒ｺ螳� 繝｡繧､繝ｳ day=%s machine=%s machine_name=%s "
                                    "start=%s end=%s 謗｡逕ｨ莠ｺ謨ｰ=%s req_num=%s 繝｡繧､繝ｳ謗｢邏｢extra_max=%s "
                                    "菴吝臆莠ｺ謨ｰ驕ｩ逕ｨ(繝｡繧､繝ｳ)=%s team=%s",
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
                                    "DEBUG[螳御ｺ�譌･謖�螳咯 萓晞�ｼNO=%s 譌･莉�=%s 縺ｯ蜑ｲ蠖謎ｸ榊庄�ｼ郁ｦ∝藤/險ｭ蛯咏ｩｺ縺肴擅莉ｶ縺ｧ繝√�ｼ繝�荳肴�千ｫ具ｼ峨Ｓemaining_units=%s",
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
                            "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 譌･谺｡邨ゆｺ�譎らせ縺ｮ谿� day=%s machine=%s "
                            "machine_name=%s rem=%.4f roll_b2_follower=%s 隧ｦ陦碁��=%s",
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
                                "邏肴悄雜�驕弱Μ繝医Λ繧､: 險育判蝓ｺ貅�+1譌･縺励※蠖楢ｩｲ萓晞�ｼ縺ｮ縺ｿ蜀埼�榊床�ｼ域､懷�ｺ譌･=%s 萓晞�ｼNO=%s 蠖楢ｩｲ萓晞�ｼ縺ｮ邏ｯ險郁ｩｦ陦�=%s�ｼ�",
                                current_date.isoformat(),
                                ",".join(sorted(allowed_shift_tids)),
                                _trials_detail,
                            )
                            _full_calendar_without_deadline_restart = False
                            break
                        else:
                            # 萓晞�ｼ縺斐→荳企剞縺ｧ繧ｷ繝輔ヨ縺ｧ縺阪↑縺�縺�縺代�ｮ縺ｨ縺阪�ｯ譌･莉倥Ν繝ｼ繝励ｒ邯咏ｶ壹☆繧具ｼ�break 縺吶ｋ縺ｨ譛ｪ蜃ｦ逅�譌･縺梧ｮ九ｊ驟榊床荳榊庄縺悟､ｧ驥冗匱逕溘☆繧具ｼ峨�
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
                                    "邏肴悄蠕後ｍ蛟偵＠蜀埼�榊床: 谺｡縺ｮ萓晞�ｼNO縺ｯ萓晞�ｼ縺斐→縺ｮ荳企剞�ｼ亥推 %s 蝗橸ｼ峨�ｮ縺溘ａ縺薙�ｮ讀懷�ｺ縺ｧ縺ｯ +1 縺励∪縺帙ｓ縲�"
                                    " 繧ｫ繝ｬ繝ｳ繝繝ｼ縺ｯ邯咏ｶ壹＠縺ｾ縺呻ｼ域悴螳御ｺ�縺ｯ邨ゆｺ�譎ゅ↓邏肴悄隕狗峩縺怜ｿ�隕√ｒ莉倥￠蠕励∪縺呻ｼ�: %s",
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
                    "ﾂｧB-2/ﾂｧB-3 繝ｪ繝ｯ繧､繝ｳ繝�: EC 螳瑚ｵｰ蠕後↓讀懈渊�ｼ丞ｷｻ霑斐＠縺ｮ縺ｿ譌･莉伜�磯�ｭ縺九ｉ蜀埼�榊床縺励∪縺励◆�ｼ�timeline_events 繧貞頃譛峨ユ繝ｼ繝悶Ν縺ｨ縺励※蛻ｩ逕ｨ�ｼ峨�"
                )
            break

    if TRACE_SCHEDULE_TASK_IDS:
        for _tt in TRACE_SCHEDULE_TASK_IDS:
            for _t in task_queue:
                if str(_t.get("task_id", "")).strip() != _tt:
                    continue
                _log_dispatch_trace_schedule(
                    _tt,
                    "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繧ｷ繝溘Η繝ｬ繝ｼ繧ｷ繝ｧ繝ｳ邨ゆｺ�譎� machine=%s machine_name=%s "
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
                    "[驟榊床繝医Ξ繝ｼ繧ｹ task=%s] 繧ｿ繧､繝�繝ｩ繧､繝ｳ譛邨ょ｡�(蟾･遞句�励＃縺ｨ) machine=%s "
                    "already_done+units_done=%s+%s=%s total_units=%s end_dt=%s",
                    _tt,
                    _mk,
                    _ad,
                    _ud,
                    _ad + _ud,
                    _ev.get("total_units"),
                    _ev.get("end_dt"),
                )

    # 繝｡繧､繝ｳ蜑ｲ莉倥∪縺ｧ縺ｮ繧ｿ繧､繝�繝ｩ繧､繝ｳ�ｼ�need 菴吝鴨霑ｽ險伜燕�ｼ峨５EMP_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ逕ｨ縲�
    timeline_before_need_surplus = copy.deepcopy(timeline_events)

    # need縲碁�榊床譎りｿｽ蜉�莠ｺ謨ｰ縲�: 繝｡繧､繝ｳ蜑ｲ莉伜ｾ後↓縲∵悴蜿ょ刈ﾃ励せ繧ｭ繝ｫ驕ｩ蜷郁�繧偵し繝悶∈霑ｽ險假ｼ域里螳夲ｼ�
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
                "need菴吝鴨: 繝｡繧､繝ｳ蜑ｲ莉伜ｮ御ｺ�蠕後↓繧ｵ繝� %s 蜷阪ｒ霑ｽ險假ｼ域悴蜑ｲ蠖禿励せ繧ｭ繝ｫ繝ｻ譎る俣驥阪↑繧翫↑縺暦ｼ�",
                _n_sur,
            )

    if _dispatch_interval_mirror is not None:
        _dispatch_interval_mirror.rebuild_from_timeline(timeline_events)

    # 繧ｿ繧､繝�繝ｩ繧､繝ｳ繧呈律莉伜挨縺ｫ繧､繝ｳ繝�繝�繧ｯ繧ｹ蛹悶＠縲√し繝悶Γ繝ｳ繝舌�ｼ荳隕ｧ繧剃ｺ句燕隗｣譫撰ｼ井ｻ･髯阪�ｮ蜃ｺ蜉帙Ν繝ｼ繝励ｒ鬮倬溷喧�ｼ�
    for e in timeline_events:
        e["subs_list"] = [s.strip() for s in e["sub"].split(",")] if e.get("sub") else []

    events_by_date = defaultdict(list)
    for e in timeline_events:
        events_by_date[e["date"]].append(e)

    # =========================================================
    # 4. Excel蜃ｺ蜉� (繝｡繧､繝ｳ險育判)
    # =========================================================
    _remove_prior_stage2_workbooks_and_prune_empty_dirs(output_dir)
    # 蜷御ｸ遘貞��縺ｮ蜀榊ｮ溯｡後〒繝輔ぃ繧､繝ｫ蜷阪′陦晉ｪ√＠縺ｪ縺�繧医≧繝槭う繧ｯ繝ｭ遘偵∪縺ｧ蜷ｫ繧√ｋ
    _stage2_out_stamp = base_now_dt.strftime("%Y%m%d_%H%M%S_%f")
    output_filename = os.path.join(
        output_dir, f"production_plan_multi_day_{_stage2_out_stamp}.xlsx"
    )
    # 繧ｿ繧ｹ繧ｯID 竊� 邨先棡_險ｭ蛯呎ｯ弱�ｮ譎る俣蜑ｲ縺ｧ蠖楢ｩｲ繧ｿ繧ｹ繧ｯ縺梧怙蛻昴↓迴ｾ繧後ｋ繧ｻ繝ｫ�ｼ井ｾ� B12�ｼ峨らｵ先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ繝ｪ繝ｳ繧ｯ逕ｨ縲�
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

    # 邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ逕ｨ: 繧ｷ繝溘Η繝ｬ繝ｼ繧ｷ繝ｧ繝ｳ荳翫�ｮ蠖楢ｩｲ繧ｿ繧ｹ繧ｯ縺ｮ譛譌ｩ髢句ｧ九�ｻ譛驕�邨ゆｺ��ｼ�timeline_events 髮�邏��ｼ�
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

    # 邨先棡_繧ｿ繧ｹ繧ｯ荳隕ｧ縺ｮ縲悟屓遲皮ｴ肴悄縲阪梧欠螳夂ｴ肴悄縲阪�ｯ驟榊床險育判_繧ｿ繧ｹ繧ｯ蜈･蜉帙�ｮ蠖楢ｩｲ陦後そ繝ｫ縺ｮ縺ｿ縲�
    # 縲悟次蜿肴兜蜈･譌･縲阪�ｯ荳頑嶌縺榊�励↓譌･莉倥′縺ゅｋ縺ｨ縺阪◎縺ｮ蛟､縲∫┌縺�縺ｨ縺榊�励悟次蜿肴兜蜈･譌･縲搾ｼ郁ｨ育判蝓ｺ貅也ｴ肴悄縺ｨ豺ｷ蜷後＠縺ｪ縺��ｼ�
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
    
    # 繧ｹ繝�繝ｼ繧ｿ繧ｹ�ｼ磯�榊床縺ｮ蜿ｯ蜷ｦ繝ｻ谿具ｼ会ｼ壼ｮ御ｺ�逶ｸ蠖�=驟榊床蜿ｯ�ｼ乗悴蜑ｲ蠖�=驟榊床荳榊庄�ｼ丈ｸ驛ｨ縺ｮ縺ｿ=驟榊床谿�
    # 險育判蝓ｺ貅�+1 縺ｮ蜀崎ｩｦ陦後′萓晞�ｼNO縺斐→縺ｮ荳企剞縺ｫ驕斐＠縺滉ｾ晞�ｼ縺ｮ譛ｪ螳御ｺ�陦後↓縺ｯ�ｼ育ｴ肴悄隕狗峩縺怜ｿ�隕��ｼ峨ｒ莉倅ｸ弱☆繧九�
    sorted_tasks_for_result = sorted(task_queue, key=_result_task_sheet_sort_key)
    for t in sorted_tasks_for_result:
        rem_u = float(t.get("remaining_units") or 0)
        hist = bool(t.get("assigned_history"))
        if rem_u <= 1e-9:
            status = "驟榊床蜿ｯ"
        elif hist and t.get("_partial_retry_calendar_blocked"):
            status = "驟榊床谿�(蜍､蜍吶き繝ｬ繝ｳ繝繝ｼ荳崎ｶｳ)"
        elif not hist and rem_u > 1e-9:
            status = "驟榊床荳榊庄"
        else:
            status = "驟榊床谿�"
        _tid_res = str(t.get("task_id", "") or "").strip()
        if (
            _tid_res in _due_shift_exhausted_requests
            and rem_u > 1e-9
            and "邏肴悄隕狗峩縺怜ｿ�隕�" not in status
        ):
            status = f"{status}�ｼ育ｴ肴悄隕狗峩縺怜ｿ�隕��ｼ�"
        
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
        # 蛻鈴��: A=繧ｹ繝�繝ｼ繧ｿ繧ｹ 竊� 繧ｿ繧ｹ繧ｯID/蟾･遞�/讖滓｢ｰ/蜆ｪ蜈亥ｺｦ 竊� 螻･豁ｴ1..n 竊� 縺昴�ｮ莉� 竊� 譛蠕後↓迚ｹ蛻･謖�螳喟AI
        row_status = {"繧ｹ繝�繝ｼ繧ｿ繧ｹ": status}
        _dto = t.get("dispatch_trial_order")
        row_core = {
            "繧ｿ繧ｹ繧ｯID": t['task_id'],
            "蟾･遞句錐": t['machine'],
            "讖滓｢ｰ蜷�": t.get("machine_name", ""),
            "蜆ｪ蜈亥ｺｦ": t.get("priority", 999),
            RESULT_TASK_COL_DISPATCH_TRIAL_ORDER: _dto if _dto is not None else "",
        }
        row_history = {}
        for i in range(max_history_len):
            if i < len(t['assigned_history']):
                h = t['assigned_history'][i]
                row_history[f"螻･豁ｴ{i+1}"] = _format_result_task_history_cell(t, h)
            else:
                row_history[f"螻･豁ｴ{i+1}"] = ""

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
            "蠢�隕＾P(荳頑嶌)": rov if rov is not None else "",
            "繧ｿ繧ｹ繧ｯ蜉ｹ邇�": parse_float_safe(t.get("task_eff_factor"), 1.0),
            "蜉�蟾･騾比ｸｭ": "縺ｯ縺�" if t.get("in_progress") else "縺�縺�縺�",
            "迚ｹ蛻･謖�螳壹≠繧�": "縺ｯ縺�" if t.get("has_special_remark") else "縺�縺�縺�",
            "諡�蠖徹P謖�蜷�": (t.get("preferred_operator_raw") or "")[:120],
            "蝗樒ｭ皮ｴ肴悄": ans_s,
            "謖�螳夂ｴ肴悄": spec_s,
            "險育判蝓ｺ貅也ｴ肴悄": basis_s,
            TASK_COL_RAW_INPUT_DATE: kenhan_s,
            "邏肴悄邱頑･": "縺ｯ縺�" if t.get("due_urgent") else "縺�縺�縺�",
            "蜉�蟾･髢句ｧ区律": start_req_s,
            "驟榊ｮ契蜉�蟾･髢句ｧ�": plan_assign_start_s,
            "驟榊ｮ契蜉�蟾･邨ゆｺ�": plan_assign_end_s,
            RESULT_TASK_COL_PLAN_END_BY_ANSWER_OR_SPEC_16: _plan_end_ans_spec16,
            "邱丞刈蟾･驥�": f"{total_r}R ({t['total_qty_m']}m)",
            "谿句刈蟾･驥�": f"{rem_r}R ({int(t['remaining_units'] * t['unit_m'])}m)",
            "螳御ｺ�邇�(螳溯｡梧凾轤ｹ)": f"{pct_macro}%",
        }
        row_ai_last = {"迚ｹ蛻･謖�螳喟AI": (t.get("task_special_ai_note") or "")[:300]}
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
                    clock_out_s = "莨�"
                cal_rows.append({
                    "譌･莉�": d,
                    "繝｡繝ｳ繝舌�ｼ": m,
                    "蜃ｺ蜍､": data['start_dt'].strftime("%H:%M") if data['is_working'] else "莨�",
                    "騾蜍､": clock_out_s,
                    "蜉ｹ邇�": data['efficiency'],
                    "蛯呵�": data['reason'],
                })

    utilization_data = []
    for d in sorted_dates:
        row_data = {"蟷ｴ譛域律": d.strftime("%Y/%m/%d (%a)")}
        # 縺昴�ｮ譌･縺ｮ繧､繝吶Φ繝医°繧峨Γ繝ｳ繝舌�ｼ蛻･菴懈･ｭ蛻�繧剃ｸ諡ｬ髮�險茨ｼ亥�ｨ繝｡繝ｳ繝舌�ｼﾃ怜�ｨ繧､繝吶Φ繝医�ｮ莠碁㍾繝ｫ繝ｼ繝励ｒ驕ｿ縺代ｋ�ｼ�
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
                row_data[m] = f"{ratio:.1f}% ({worked_mins}/{total_avail_mins}蛻�)"
            else:
                row_data[m] = "莨�"
        utilization_data.append(row_data)
        
    df_utilization = pd.DataFrame(utilization_data)

    df_mprio_legend, df_mprio_tbl = build_member_assignment_priority_reference(
        skills_dict, members
    )
    if df_mprio_tbl.empty:
        df_mprio_tbl = pd.DataFrame(
            [
                {
                    "蟾･遞句錐": "",
                    "讖滓｢ｰ蜷�": "",
                    "繧ｹ繧ｭ繝ｫ蛻励く繝ｼ": "",
                    "蜆ｪ蜈磯��菴�": "",
                    "繝｡繝ｳ繝舌�ｼ": "",
                    "繝ｭ繝ｼ繝ｫ": "",
                    "蜆ｪ蜈亥ｺｦ蛟､_蟆上＆縺�縺ｻ縺ｩ蜈�": "",
                    "skills繧ｻ繝ｫ蛟､": "",
                    "蛯呵�": "繝槭せ繧ｿ skills 縺ｫ縲悟ｷ･遞句錐+讖滓｢ｰ蜷阪榊ｽ｢蠑上�ｮ蛻励′隕九▽縺九ｉ縺ｪ縺�縺九√ョ繝ｼ繧ｿ縺後≠繧翫∪縺帙ｓ縲�",
                }
            ]
        )

    _usage_txt = build_gemini_usage_summary_text()
    if _usage_txt:
        ai_log_data["Gemini_繝医�ｼ繧ｯ繝ｳ繝ｻ譁咎≡繧ｵ繝槭Μ"] = _usage_txt[:50000]

    _master_abs_for_result_fmt = os.path.abspath(os.path.join(os.getcwd(), MASTER_FILE))
    _reg_shift_start, _reg_shift_end = _read_master_main_regular_shift_times(
        _master_abs_for_result_fmt
    )
    if _reg_shift_start is not None and _reg_shift_end is not None:
        logging.info(
            "螳壼ｸｸ譫�: master 繝｡繧､繝ｳ A15/B15 竊� %s �ｽ� %s�ｼ育ｵ先棡縺ｮ螳壼ｸｸ螟悶梧律譎ょｸｯ縲咲捩濶ｲ�ｼ�",
            _reg_shift_start.strftime("%H:%M"),
            _reg_shift_end.strftime("%H:%M"),
        )

    logging.info(
        "谿ｵ髫�2: 邨先棡繝悶ャ繧ｯ繧剃ｽ懈�舌＠縺ｾ縺� 竊� %s",
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
            pd.DataFrame(cal_rows).to_excel(writer, sheet_name='邨先棡_繧ｫ繝ｬ繝ｳ繝繝ｼ(蜃ｺ蜍､邁ｿ)', index=False)
            df_utilization.to_excel(writer, sheet_name='邨先棡_繝｡繝ｳ繝舌�ｼ蛻･菴懈･ｭ蜑ｲ蜷�', index=False)
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
                    "蛻怜錐": task_column_order_dedup,
                    "陦ｨ遉ｺ": vis_list_dedup,
                }
            ).to_excel(writer, sheet_name=COLUMN_CONFIG_SHEET_NAME, index=False)
            df_tasks.to_excel(writer, sheet_name=RESULT_TASK_SHEET_NAME, index=False)
            pd.DataFrame(list(ai_log_data.items()), columns=["鬆�逶ｮ", "蜀�螳ｹ"]).to_excel(writer, sheet_name='邨先棡_AI繝ｭ繧ｰ', index=False)

            _mprio_sheet = RESULT_MEMBER_PRIORITY_SHEET_NAME
            df_mprio_legend.to_excel(writer, sheet_name=_mprio_sheet, index=False)
            _mprio_gap = len(df_mprio_legend) + 2
            df_mprio_tbl.to_excel(
                writer, sheet_name=_mprio_sheet, index=False, startrow=_mprio_gap
            )

            logging.info(
                "谿ｵ髫�2: 險ｭ蛯吶ぎ繝ｳ繝医ｒ逕滓�舌＠縺ｦ縺�縺ｾ縺呻ｼ医ョ繝ｼ繧ｿ驥上↓繧医ｊ謨ｰ蛻�縺九°繧九％縺ｨ縺後≠繧翫∪縺呻ｼ�"
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

            # 譛ｪ繧ｹ繧ｱ繧ｸ繝･繝ｼ繝ｫ陦鯉ｼ磯�榊床荳榊庄繝ｻ驟榊床谿具ｼ峨ｒ逶ｮ遶九◆縺帙ｋ
            status_col_idx = None
            for col_idx, col_name in enumerate(df_tasks.columns, 1):
                if str(col_name) == "繧ｹ繝�繝ｼ繧ｿ繧ｹ":
                    status_col_idx = col_idx
                    break
            if status_col_idx is not None:
                unscheduled_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                for r in range(2, worksheet_tasks.max_row + 1):
                    st_val = worksheet_tasks.cell(row=r, column=status_col_idx).value
                    st = str(st_val).strip() if st_val is not None else ""
                    if st in ("驟榊床荳榊庄", "驟榊床谿�"):
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
            "谿ｵ髫�2: 邨先棡繝悶ャ繧ｯ縺ｮ菴懈�舌�ｻ菫晏ｭ倥↓螟ｱ謨励＠縺ｾ縺励◆: %s�ｼ�%s�ｼ峨�"
            "output 蜀�縺ｮ production_plan_multi_day_*.xlsx 繧� Excel 縺ｧ髢九＞縺ｦ縺�縺ｪ縺�縺狗｢ｺ隱阪＠縺ｦ縺上□縺輔＞縲�",
            output_filename,
            e,
        )
        raise

    try:
        _apply_excel_date_columns_date_only_display(
            output_filename, "邨先棡_繧ｫ繝ｬ繝ｳ繝繝ｼ(蜃ｺ蜍､邁ｿ)", frozenset({"譌･莉�"})
        )
    except Exception as e:
        logging.warning(f"邨先棡_繧ｫ繝ｬ繝ｳ繝繝ｼ(蜃ｺ蜍､邁ｿ)縺ｮ譌･莉伜�苓｡ｨ遉ｺ謨ｴ蠖｢: {e}")

    _stage2_try_copy_column_config_shapes_from_input(
        output_filename,
        (os.environ.get("TASK_INPUT_WORKBOOK", "").strip() or TASKS_INPUT_WORKBOOK),
    )

    logging.info(f"螳御ｺ�: '{output_filename}' 繧堤函謌舌＠縺ｾ縺励◆縲�")

    # =========================================================
    # 5. 笘�霑ｽ蜉�: 繝｡繝ｳ繝舌�ｼ豈弱�ｮ陦悟虚繧ｹ繧ｱ繧ｸ繝･繝ｼ繝ｫ (蛻･繝輔ぃ繧､繝ｫ) 蜃ｺ蜉�
    # =========================================================
    member_output_filename = os.path.join(
        output_dir, f"member_schedule_{_stage2_out_stamp}.xlsx"
    )
    
    # 譎る俣蟶ｯ縺ｯ蜈ｨ繝｡繝ｳ繝舌�ｼ蜈ｱ騾壹〒1蝗槭□縺醍函謌撰ｼ医Γ繝ｳ繝舌�ｼ謨ｰ蛻�縺ｮ驥崎､�險育ｮ励ｒ驕ｿ縺代ｋ�ｼ�
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
        "谿ｵ髫�2: 繝｡繝ｳ繝舌�ｼ蛻･繧ｹ繧ｱ繧ｸ繝･繝ｼ繝ｫ繧剃ｽ懈�舌＠縺ｾ縺� 竊� %s",
        os.path.basename(member_output_filename),
    )
    try:
        with pd.ExcelWriter(member_output_filename, engine="openpyxl") as member_writer:
            for m in members:
                # 蜷�陦後�ｮ霎樊嶌繧貞�晄悄蛹�
                m_schedule = {t_label: {"譎る俣蟶ｯ": t_label} for t_label in time_labels}
            
                # 蜷�譌･莉倥�ｮ繧ｹ繧ｱ繧ｸ繝･繝ｼ繝ｫ繧貞�励→縺励※蝓九ａ縺ｦ縺�縺�
                for d in sorted_dates:
                    d_str = d.strftime("%m/%d (%a)")
                
                    # 蜈ｨ譌･髱槫共蜍�: 蟷ｴ莨托ｼ医き繝ｬ繝ｳ繝繝ｼ *�ｼ峨�ｯ縲主ｹｴ莨代上∝ｷ･蝣ｴ莨第律縺ｪ縺ｩ縺ｯ縲惹ｼ代�
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
                    
                        # 蛻､螳夂畑縺ｮ荳ｭ髢捺凾蛻ｻ繧定ｨ育ｮ�
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
                            # 隧ｲ蠖薙☆繧九ち繧ｹ繧ｯ繧呈爾縺呻ｼ�subs_list 縺ｯ莠句燕隗｣譫先ｸ医∩�ｼ�
                            active_ev = next((e for e in events_today if e['start_dt'] <= grid_mid_dt < e['end_dt'] and (e['op'] == m or m in e.get('subs_list', []))), None)
                            if active_ev:
                                role = "荳ｻ" if active_ev['op'] == m else "陬�"
                                text = f"[{active_ev['task_id']}] {active_ev['machine']}({role})"
                            else:
                                text = "" # 菴輔ｂ蜑ｲ繧雁ｽ薙※繧峨ｌ縺ｦ縺�縺ｪ縺�遨ｺ縺肴凾髢�
                    
                        m_schedule[t_label][d_str] = text
                    
                # 繝�繝ｼ繧ｿ繝輔Ξ繝ｼ繝�蛹悶＠縺ｦ繧ｷ繝ｼ繝医↓譖ｸ縺崎ｾｼ縺ｿ
                df_m = pd.DataFrame(list(m_schedule.values()))
                cols = ["譎る俣蟶ｯ"] + [d.strftime("%m/%d (%a)") for d in sorted_dates]
                df_m = df_m[[c for c in cols if c in df_m.columns]]
                df_m.to_excel(member_writer, sheet_name=m, index=False)
            
                # --- 譌｢螳壹ヵ繧ｩ繝ｳ繝医�ｻ鄂ｫ邱壹�ｻ隕句�ｺ縺苓レ譎ｯ�ｼ亥�怜ｹ�縺ｯ VBA 蜿悶ｊ霎ｼ縺ｿ譎ゅ�ｮ AutoFit�ｼ� ---
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
            "谿ｵ髫�2: 繝｡繝ｳ繝舌�ｼ蛻･繧ｹ繧ｱ繧ｸ繝･繝ｼ繝ｫ縺ｮ菫晏ｭ倥↓螟ｱ謨励＠縺ｾ縺励◆: %s�ｼ�%s�ｼ峨�"
            "member_schedule_*.xlsx 繧� Excel 縺ｧ髢九＞縺ｦ縺�縺ｪ縺�縺狗｢ｺ隱阪＠縺ｦ縺上□縺輔＞縲�",
            member_output_filename,
            e,
        )
        raise

    logging.info(f"螳御ｺ�: 蛟倶ｺｺ蛻･繧ｹ繧ｱ繧ｸ繝･繝ｼ繝ｫ繧� '{member_output_filename}' 縺ｫ蜃ｺ蜉帙＠縺ｾ縺励◆縲�")
    _try_write_main_sheet_gemini_usage_summary("谿ｵ髫�2")