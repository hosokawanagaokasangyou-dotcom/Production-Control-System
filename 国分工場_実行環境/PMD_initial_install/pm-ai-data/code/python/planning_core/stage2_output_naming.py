# -*- coding: utf-8 -*-
"""
段階2の output 成果物ファイル名。

本体（.xlsx / 同名 .json ミラー）のベース名は Unicode で 20 文字以内。
タイムスタンプは yyMMddHHmmss（12 桁）＋マイクロ秒下 4 桁で同一秒内の衝突を緩和する。
"""

from __future__ import annotations

from datetime import datetime

# 計画ブック・個人別ブックのファイル名接頭辞（日本語）
PLAN_BOOK_PREFIX = "計画"
MEMBER_BOOK_PREFIX = "人員"

# JSON バリアント（拡張子の直前に 1〜2 文字を付す）
JSON_VARIANT_TABULAR = "表"
JSON_VARIANT_LOGICAL = "論"
JSON_VARIANT_EQUIP_GANTT_CONTRACT = "設"
JSON_VARIANT_ACTUAL_DETAIL_GANTT_CONTRACT = "実"
# 結果_タスク一覧のサイドカー（2 文字）
JSON_VARIANT_RESULT_TASK_LIST = "一覧"


def format_stage2_stamp(base_now_dt: datetime, run_dt: datetime) -> str:
    """
    データ抽出基準の壁時計 yyMMddHHmmss と、実行時刻のマイクロ秒下 4 桁を連結した 16 桁。
    """
    core = base_now_dt.strftime("%y%m%d%H%M%S")
    frac = int(run_dt.microsecond) % 10000
    return f"{core}{frac:04d}"


def plan_workbook_basename(stamp: str) -> str:
    """拡張子なしの計画ブック名（例: 計画2605060836241234）。"""
    return f"{PLAN_BOOK_PREFIX}{stamp}"


def member_workbook_basename(stamp: str) -> str:
    """拡張子なしの個人別ブック名（例: 人員2605060836241234）。"""
    return f"{MEMBER_BOOK_PREFIX}{stamp}"


def plan_workbook_filename(stamp: str) -> str:
    return f"{plan_workbook_basename(stamp)}.xlsx"


def member_workbook_filename(stamp: str) -> str:
    return f"{member_workbook_basename(stamp)}.xlsx"
