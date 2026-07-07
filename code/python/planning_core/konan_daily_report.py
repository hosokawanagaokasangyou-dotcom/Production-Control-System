# -*- coding: utf-8 -*-
"""
湖南工場「加工日報発行問合せ」CSV の読込・集約ユーティリティ。

正本 CSV: \\\\192.168.0.101\\...\\●DATA\\加工日報\\加工日報発行問合せ_*.csv
UTF-8 BOM、先頭3行メタ、4行目ヘッダー、5行目以降データ。

配台・タスク入力パイプラインには接続しない。Java「原本転記・計画確認」タブの完了表示、
解析スクリプト、単体テスト向け。
"""

from __future__ import annotations

import glob
import logging
import os
import unicodedata
from typing import Any

import pandas as pd

_LOG = logging.getLogger(__name__)

ENV_DAILY_REPORT_SOURCE_DIR = "PM_AI_DAILY_REPORT_SOURCE_DIR"
ENV_DAILY_REPORT_MERGE = "PM_AI_DAILY_REPORT_MERGE"
ENV_DAILY_REPORT_CSV_PATH = "PM_AI_DAILY_REPORT_CSV_PATH"
ENV_DAILY_REPORT_APPEND_MISSING = "PM_AI_DAILY_REPORT_APPEND_MISSING"

DAILY_REPORT_FILENAME_PREFIX = "加工日報発行問合せ_"

DEFAULT_DAILY_REPORT_SOURCE_DIR = (
    "\\\\192.168.0.101\\"
    "\u5171\u6709\u30d5\u30a9\u30eb\u30c0\\"
    "\u6e56\u5357\u5de5\u5834\\"
    "\u6e56\u5357\u5171\u6709\\"
    "\u751f\u7523\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\\"
    "\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\\"
    "\u25cfDATA\\"
    "\u52a0\u5de5\u65e5\u5831"
)

COL_TASK_ID = "依頼NO"
COL_ORDER_NO = "受注NO"
COL_PROCESS = "工程名"
COL_MACHINE_NAME = "機械名"
COL_QTY = "換算数量"
COL_UNPROCESSED = "未加工"
COL_COMPLETION_FLAG = "加工完了区分"
COL_ACTUAL_DONE = "実加工数"
COL_ACTUAL_OUTPUT = "実出来高"
COL_ANSWER_DUE = "回答納期"
COL_RAW_INPUT_DATE = "原反投入日"
COL_PRODUCT = "製品名"

DR_COL_DAY = "加工日付"
DR_COL_QTY = "換算数量"
DR_COL_CUM_ACTUAL = "加工実績累計"
DR_COL_CUM_OUTPUT = "実製品出来高累計"
DR_COL_COMPLETION = "完了区分"
DR_COL_ORDER_COMPLETION = "注文単位加工完了区分"
DR_COL_ANSWER_DUE = "回答納期"
DR_COL_RAW_ARRIVAL = "原反着予定"
DR_COL_PRODUCT = "製品"
DR_COL_ORDER_NO = "受注NO"

_AGG_COLUMNS = (
    COL_TASK_ID,
    COL_PROCESS,
    COL_MACHINE_NAME,
    DR_COL_DAY,
    DR_COL_QTY,
    DR_COL_CUM_ACTUAL,
    DR_COL_CUM_OUTPUT,
    DR_COL_COMPLETION,
    DR_COL_ORDER_COMPLETION,
    DR_COL_ANSWER_DUE,
    DR_COL_RAW_ARRIVAL,
    DR_COL_PRODUCT,
    DR_COL_ORDER_NO,
)


def _truthy_env(name: str, *, default: bool = True) -> bool:
    raw = (os.environ.get(name) or "").strip().lower()
    if not raw:
        return default
    return raw not in ("0", "false", "off", "no", "いいえ")


def _parse_qty(val: Any) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    s = str(val).strip().replace(",", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except (TypeError, ValueError):
        return 0.0


def _norm_key_part(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return unicodedata.normalize("NFKC", str(val).strip())


def _task_key(task_id: str, process: str, machine_name: str) -> tuple[str, str, str]:
    return (
        _norm_key_part(task_id),
        _norm_key_part(process),
        _norm_key_part(machine_name),
    )


def pick_newest_daily_report_csv(dir_path: str) -> str | None:
    """加工日報フォルダ内の最新 ``加工日報発行問合せ_*.csv`` を返す。"""
    if not dir_path or not os.path.isdir(dir_path):
        return None
    pattern = os.path.join(dir_path, f"{DAILY_REPORT_FILENAME_PREFIX}*.csv")
    best: str | None = None
    best_key = -1.0
    for fp in glob.glob(pattern):
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
    return best


def resolve_daily_report_csv_path() -> str | None:
    """``PM_AI_DAILY_REPORT_CSV_PATH`` → ディレクトリ内最新 の順で解決。"""
    explicit = (os.environ.get(ENV_DAILY_REPORT_CSV_PATH) or "").strip()
    if explicit and os.path.isfile(explicit):
        return os.path.normpath(os.path.abspath(explicit))
    src = (os.environ.get(ENV_DAILY_REPORT_SOURCE_DIR) or "").strip()
    if not src:
        src = DEFAULT_DAILY_REPORT_SOURCE_DIR
    picked = pick_newest_daily_report_csv(src)
    if picked:
        return os.path.normpath(os.path.abspath(picked))
    return None


def read_konan_daily_report_csv(path: str) -> pd.DataFrame:
    """加工日報 CSV を DataFrame 化（メタ3行スキップ）。"""
    with open(path, encoding="utf-8-sig", newline="") as fp:
        df = pd.read_csv(fp, skiprows=3, dtype=str, keep_default_na=False)
    df.columns = df.columns.astype(str).str.strip()
    return df


def aggregate_daily_report_latest(df: pd.DataFrame) -> pd.DataFrame:
    """(依頼NO, 工程名, 機械名) ごとに加工日付が最新の行を1件に畳む。"""
    if df is None or df.empty:
        return pd.DataFrame(columns=list(_AGG_COLUMNS))
    need = {COL_TASK_ID, COL_PROCESS, COL_MACHINE_NAME, DR_COL_DAY}
    if not need.issubset(set(df.columns)):
        missing = need - set(df.columns)
        raise ValueError(f"加工日報に必須列がありません: {sorted(missing)}")

    work = df.copy()
    work["_sort_day"] = work[DR_COL_DAY].map(_norm_key_part)
    work = work.sort_values("_sort_day")
    grouped = work.groupby(
        [COL_TASK_ID, COL_PROCESS, COL_MACHINE_NAME], sort=False, as_index=False
    ).tail(1)
    keep = [c for c in _AGG_COLUMNS if c in grouped.columns]
    return grouped[keep].reset_index(drop=True)


def _recalc_unprocessed(qty: float, actual_done: float, current_unp: float | None) -> float:
    """
    未加工 = max(0, 換算数量 - 実加工数)。
    換算>0・実加工=0・未加工=0 の行は全数未加工ルールに従う。
    """
    if qty <= 1e-12:
        return max(0.0, current_unp or 0.0)
    if actual_done <= 1e-12 and current_unp is not None and abs(current_unp) <= 1e-12:
        return qty
    return max(0.0, qty - actual_done)


def _format_qty_cell(value: float) -> str:
    if abs(value - round(value)) <= 1e-9:
        return str(int(round(value)))
    return str(value)


def _daily_report_row_needs_dispatch(dr_row: pd.Series) -> bool:
    """計画外タスクの合成行追加対象: 工程未完、または累計実績が換算数量未満。"""
    comp = str(dr_row.get(DR_COL_COMPLETION) or "")
    if "未完" in comp:
        return True
    cum = _parse_qty(dr_row.get(DR_COL_CUM_ACTUAL))
    qty = _parse_qty(dr_row.get(DR_COL_QTY))
    return qty > 1e-12 and cum + 1e-9 < qty


def _apply_daily_report_row_to_plan(
    out: pd.DataFrame,
    i: int,
    dr_row: pd.Series,
    *,
    qty_override_from_dr: bool,
) -> None:
    cum_actual = _parse_qty(dr_row.get(DR_COL_CUM_ACTUAL))
    cum_output = _parse_qty(dr_row.get(DR_COL_CUM_OUTPUT))
    qty_plan = _parse_qty(out.iloc[i].get(COL_QTY))
    qty_dr = _parse_qty(dr_row.get(DR_COL_QTY))
    qty = qty_dr if qty_override_from_dr and qty_dr > 1e-12 else qty_plan
    if qty <= 1e-12 and qty_dr > 1e-12:
        qty = qty_dr
    cur_unp = _parse_qty(out.iloc[i].get(COL_UNPROCESSED)) if COL_UNPROCESSED in out.columns else None
    unp = _recalc_unprocessed(
        qty,
        cum_actual,
        cur_unp if qty_plan > 1e-12 else None,
    )

    out.at[out.index[i], COL_ACTUAL_DONE] = _format_qty_cell(cum_actual)
    if cum_output > 1e-12:
        out.at[out.index[i], COL_ACTUAL_OUTPUT] = _format_qty_cell(cum_output)
    if qty_dr > 1e-12 and qty_plan <= 1e-12:
        out.at[out.index[i], COL_QTY] = _format_qty_cell(qty_dr)
    out.at[out.index[i], COL_UNPROCESSED] = _format_qty_cell(unp)
    comp = str(dr_row.get(DR_COL_COMPLETION) or "").strip()
    if comp:
        out.at[out.index[i], COL_COMPLETION_FLAG] = comp


def merge_daily_report_into_plan_df(plan_df: pd.DataFrame) -> pd.DataFrame:
    """
    加工計画DATA DataFrame に加工日報の累計実績・完了区分を反映する。

    - 一致行: 実加工数・実出来高・加工完了区分・未加工を上書き
    - 計画に無い (依頼NO, 工程名, 機械名): 行を追加（合成タスク）
    """
    if not _truthy_env(ENV_DAILY_REPORT_MERGE, default=True):
        return plan_df
    if plan_df is None:
        return plan_df

    path = resolve_daily_report_csv_path()
    if not path:
        _LOG.info("加工日報: ソースが解決できないためマージをスキップします。")
        return plan_df

    try:
        dr_raw = read_konan_daily_report_csv(path)
        dr_latest = aggregate_daily_report_latest(dr_raw)
    except Exception as ex:
        _LOG.warning("加工日報: 読込に失敗したためマージをスキップ (%s): %s", path, ex)
        return plan_df

    if dr_latest.empty:
        _LOG.info("加工日報: データ行が空のためマージをスキップ (%s)。", path)
        return plan_df

    out = plan_df.copy()
    for col in (
        COL_TASK_ID,
        COL_ORDER_NO,
        COL_PROCESS,
        COL_MACHINE_NAME,
        COL_QTY,
        COL_UNPROCESSED,
        COL_COMPLETION_FLAG,
        COL_ACTUAL_DONE,
        COL_ACTUAL_OUTPUT,
        COL_ANSWER_DUE,
        COL_RAW_INPUT_DATE,
        COL_PRODUCT,
    ):
        if col not in out.columns:
            out[col] = ""

    lookup: dict[tuple[str, str, str], pd.Series] = {}
    lookup_proc_only: dict[tuple[str, str], list[tuple[str, str, str]]] = {}
    for _, row in dr_latest.iterrows():
        k = _task_key(row.get(COL_TASK_ID), row.get(COL_PROCESS), row.get(COL_MACHINE_NAME))
        if not k[0] or not k[1]:
            continue
        lookup[k] = row
        lookup_proc_only.setdefault((k[0], k[1]), []).append(k)

    matched_keys: set[tuple[str, str, str]] = set()
    updated = 0

    for i in range(len(out)):
        plan_row = out.iloc[i]
        tid = _norm_key_part(plan_row.get(COL_TASK_ID))
        proc = _norm_key_part(plan_row.get(COL_PROCESS))
        mach = _norm_key_part(plan_row.get(COL_MACHINE_NAME))
        if not tid or not proc:
            continue
        k3 = (tid, proc, mach)
        dr_row = lookup.get(k3)
        if dr_row is None:
            cands = lookup_proc_only.get((tid, proc), [])
            if len(cands) == 1:
                dr_row = lookup.get(cands[0])
        if dr_row is None:
            continue

        matched_keys.add(
            _task_key(dr_row.get(COL_TASK_ID), dr_row.get(COL_PROCESS), dr_row.get(COL_MACHINE_NAME))
        )
        _apply_daily_report_row_to_plan(out, i, dr_row, qty_override_from_dr=False)
        updated += 1

    appended = 0
    new_rows: list[dict[str, Any]] = []
    append_missing = _truthy_env(ENV_DAILY_REPORT_APPEND_MISSING, default=True)
    for _, dr_row in dr_latest.iterrows():
        k = _task_key(dr_row.get(COL_TASK_ID), dr_row.get(COL_PROCESS), dr_row.get(COL_MACHINE_NAME))
        if not k[0] or not k[1] or k in matched_keys:
            continue
        if not append_missing or not _daily_report_row_needs_dispatch(dr_row):
            continue
        cum_actual = _parse_qty(dr_row.get(DR_COL_CUM_ACTUAL))
        cum_output = _parse_qty(dr_row.get(DR_COL_CUM_OUTPUT))
        qty = _parse_qty(dr_row.get(DR_COL_QTY))
        unp = _recalc_unprocessed(qty, cum_actual, 0.0 if qty > 1e-12 else None)
        rec = {c: "" for c in out.columns}
        rec[COL_TASK_ID] = str(dr_row.get(COL_TASK_ID) or "").strip()
        rec[COL_ORDER_NO] = str(dr_row.get(DR_COL_ORDER_NO) or "").strip()
        rec[COL_PROCESS] = str(dr_row.get(COL_PROCESS) or "").strip()
        rec[COL_MACHINE_NAME] = str(dr_row.get(COL_MACHINE_NAME) or "").strip()
        rec[COL_QTY] = _format_qty_cell(qty)
        rec[COL_UNPROCESSED] = _format_qty_cell(unp)
        rec[COL_ACTUAL_DONE] = _format_qty_cell(cum_actual)
        if cum_output > 1e-12:
            rec[COL_ACTUAL_OUTPUT] = _format_qty_cell(cum_output)
        comp = str(dr_row.get(DR_COL_COMPLETION) or "").strip()
        if comp:
            rec[COL_COMPLETION_FLAG] = comp
        rec[COL_ANSWER_DUE] = str(dr_row.get(DR_COL_ANSWER_DUE) or "").strip()
        rec[COL_RAW_INPUT_DATE] = str(dr_row.get(DR_COL_RAW_ARRIVAL) or "").strip()
        rec[COL_PRODUCT] = str(dr_row.get(DR_COL_PRODUCT) or "").strip()
        new_rows.append(rec)
        appended += 1

    if new_rows:
        out = pd.concat([out, pd.DataFrame(new_rows)], ignore_index=True)

    _LOG.info(
        "加工日報: %s から %s キーを読込、計画 %s 行中 %s 行更新、%s 行追加。",
        os.path.basename(path),
        len(dr_latest),
        len(plan_df),
        updated,
        appended,
    )
    return out
