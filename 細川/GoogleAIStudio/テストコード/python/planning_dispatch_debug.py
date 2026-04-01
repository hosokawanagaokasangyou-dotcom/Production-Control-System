# -*- coding: utf-8 -*-
"""
配台のロール単位トレース（planning_core が出力する JSONL）を Excel シートへ展開する。

環境変数（段階2の前に設定）::
    DISPATCH_ROLL_TRACE_JSONL … 既定 log/dispatch_roll_trace.jsonl（ブック相対または絶対）
    DISPATCH_DEBUG_STOP_AFTER_ROLLS … 整数 N のとき N ロールで割付シミュレーションを打ち切り

xlwings の RunPython から::

    import planning_dispatch_debug as _dd
    _dd.refresh_roll_trace_sheet_from_jsonl()
"""
from __future__ import annotations

import json
import os
import sys

SHEET_ROLL_TRACE = "配台_ロールトレース"


def _caller_workbook_dir() -> str:
    import xlwings as xw

    wb = xw.Book.caller()
    return os.path.dirname(os.path.abspath(str(wb.fullname)))


def _resolve_jsonl_path(explicit: str | None) -> str:
    if explicit and explicit.strip():
        p = explicit.strip()
        if os.path.isabs(p):
            return p
        root = _caller_workbook_dir()
        return os.path.normpath(os.path.join(root, p.replace("/", os.sep)))
    root = _caller_workbook_dir()
    return os.path.join(root, "log", "dispatch_roll_trace.jsonl")


def refresh_roll_trace_sheet_from_jsonl(jsonl_path: str | None = None) -> str:
    """
    JSONL を読み、ブック内に SHEET_ROLL_TRACE シートを作成／上書きして一覧表示する。
    戻り値: ユーザー向け短文メッセージ。
    """
    import xlwings as xw

    path = _resolve_jsonl_path(jsonl_path)
    if not os.path.isfile(path):
        return f"ロールトレースがありません: {path}（先に段階2で DISPATCH_ROLL_TRACE_JSONL を有効にしてください）"

    rows: list[list] = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                o = json.loads(line)
            except json.JSONDecodeError:
                continue
            rows.append(
                [
                    o.get("seq"),
                    o.get("date"),
                    o.get("task_id"),
                    o.get("machine_line"),
                    o.get("dispatch_trial_order"),
                    o.get("start_dt"),
                    o.get("end_dt"),
                    o.get("units_done"),
                    o.get("remaining_rolls_after"),
                    o.get("lead_op"),
                    o.get("subs"),
                ]
            )

    wb = xw.Book.caller()
    try:
        ws = wb.sheets[SHEET_ROLL_TRACE]
    except Exception:
        ws = wb.sheets.add(SHEET_ROLL_TRACE, after=wb.sheets[0])

    hdr = [
        "seq",
        "日付",
        "task_id",
        "設備列",
        "配台試行順",
        "開始",
        "終了",
        "本ロール単位数",
        "残ロール数(後)",
        "主OP",
        "サブ",
    ]
    try:
        ws.clear()
    except Exception:
        ws.cells.clear()
    ws.range("A1").value = [hdr] + rows
    ws.range("A1").columns.autofit()
    return f"配台_ロールトレース: {len(rows)} 行を {path} から反映しました。"


def run_refresh_roll_trace_sheet_for_xlwings() -> int:
    """xlwings RunPython 用。0=成功, 1=失敗"""
    try:
        msg = refresh_roll_trace_sheet_from_jsonl()
        print(msg, file=sys.stdout)
        return 0
    except Exception:
        import logging

        logging.exception("配台_ロールトレースの更新に失敗しました")
        return 1
