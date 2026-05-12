# -*- coding: utf-8 -*-
"""
設備ガント系シートの JSON 表に、同ブック内「結果_タスク一覧」由来の担当ヒントを付与する。

``production_plan_multi_day*.json`` および ``*_logical_view.json`` は
:func:`planning_core.workbook_payload.workbook_payload_from_final_xlsx_file`
経由でビューアへ渡るため、当モジュールはその出口で一括適用する。

無効化: 環境変数 ``PM_AI_ENRICH_EQUIPMENT_GANTT_MEMBERS=0``（false/no/off/none 同様）。
"""

from __future__ import annotations

import logging
import os
import re
import unicodedata
from typing import Any

logger = logging.getLogger(__name__)

ENV_ENRICH_GANTT_MEMBERS = "PM_AI_ENRICH_EQUIPMENT_GANTT_MEMBERS"

RESULT_TASK_SHEET_NAME = "結果_タスク一覧"

_TASK_ID_KEYS = ("タスクID", "依頼NO")

_SKIPPED_CELL_MARKERS = (
    "休憩",
    "日次始業準備",
    "アラジン",
)

_TASK_BRACKET_RE = re.compile(r"\[([^\]\r\n]+)\]")
_TIME_HEADER_RE = re.compile(r"^\s*\d{1,2}:\d{2}\s*$")


def _enrich_enabled() -> bool:
    v = (os.environ.get(ENV_ENRICH_GANTT_MEMBERS) or "").strip().lower()
    if v in ("0", "false", "no", "off", "none"):
        return False
    return True


def _norm_tid(v: Any) -> str:
    return unicodedata.normalize("NFKC", str(v or "").strip())


def _history_keys_sorted(columns: list[str]) -> list[str]:
    out: list[tuple[int, str]] = []
    for c in columns:
        cs = str(c).strip()
        if cs.startswith("履歴"):
            m = re.search(r"(\d+)", cs)
            if m:
                out.append((int(m.group(1)), cs))
    out.sort(key=lambda x: x[0])
    return [x[1] for x in out]


def _extract_main_team_from_history_text(s: str) -> str:
    """履歴セル文字列から「主:」に続くメインチーム文字列を返す。"""
    if not s or "主:" not in s:
        return ""
    after = s.split("主:", 1)[1]
    for sep in (" +", " 余:", " 余："):
        if sep in after:
            after = after.split(sep)[0]
    return after.strip()


def _member_hint_from_result_task_row(row: dict, hist_cols: list[str]) -> str:
    """結果_タスク一覧の1行から担当表示用の短いヒントを返す。"""
    for hk in reversed(hist_cols):
        hint = _extract_main_team_from_history_text(str(row.get(hk) or ""))
        if hint:
            return hint
    pref = _norm_tid(row.get("担当OP指定"))
    if pref:
        return pref[:120]
    return ""


def build_task_id_to_member_hint(sheet: dict) -> dict[str, str]:
    """「結果_タスク一覧」シートのペイロードから 依頼NO -> 担当ヒント（後勝ち）。"""
    cols = sheet.get("columns") or []
    rows = sheet.get("rows") or []
    if not cols or not rows:
        return {}
    hist_cols = _history_keys_sorted([str(c) for c in cols])
    out: dict[str, str] = {}
    colset = {str(c) for c in cols}
    for row in rows:
        if not isinstance(row, dict):
            continue
        tid = ""
        for k in _TASK_ID_KEYS:
            if k in colset:
                tid = _norm_tid(row.get(k))
                if tid:
                    break
        if not tid:
            continue
        hint = _member_hint_from_result_task_row(row, hist_cols)
        if hint:
            out[tid] = hint
    return out


def _is_time_column_header(h: str) -> bool:
    return bool(_TIME_HEADER_RE.match(str(h)))


def _should_skip_cell(text: str) -> bool:
    t = text.strip()
    if not t:
        return True
    for m in _SKIPPED_CELL_MARKERS:
        if m in t:
            return True
    return False


def _already_has_name_paren(text: str) -> bool:
    """既に「（…）」で注記があるセルは二重付与しない（日付以外の注記想定）。"""
    if "（" in text and "）" in text:
        return True
    return False


def enrich_equipment_gantt_sheet(
    gantt: dict,
    tid_to_hint: dict[str, str],
) -> dict:
    """単一シート dict の rows を in-place で更新する。"""
    cols = list(gantt.get("columns") or [])
    rows = gantt.get("rows") or []
    if not cols or not tid_to_hint:
        return gantt
    time_indices = [i for i, h in enumerate(cols) if _is_time_column_header(str(h))]
    if not time_indices:
        return gantt
    new_rows: list[dict] = []
    for row in rows:
        if not isinstance(row, dict):
            new_rows.append(row)
            continue
        r = dict(row)
        for ii in time_indices:
            if ii >= len(cols):
                continue
            key = cols[ii]
            if key not in r:
                continue
            raw = r.get(key)
            if raw is None:
                continue
            text = str(raw).replace("\r\n", "\n")
            if _should_skip_cell(text):
                continue
            if _already_has_name_paren(text):
                continue
            ids = _TASK_BRACKET_RE.findall(text)
            if not ids:
                continue
            hints: list[str] = []
            seen: set[str] = set()
            for tid_raw in ids:
                tid = _norm_tid(tid_raw)
                hint = tid_to_hint.get(tid)
                if not hint:
                    continue
                nk = unicodedata.normalize("NFKC", hint)
                if nk not in seen:
                    seen.add(nk)
                    hints.append(hint)
            if not hints:
                continue
            suffix = "（" + "・".join(hints) + "）"
            if suffix in text:
                continue
            r[key] = text.rstrip() + suffix
        new_rows.append(r)
    gantt["rows"] = new_rows
    return gantt


def enrich_workbook_payload_equipment_gantt_members(payload: dict) -> dict:
    """
    ブック表形式ペイロードの ``sheets`` 内、名前に「設備ガント」を含むシートの
    タイム列セルへ、``結果_タスク一覧`` から得た担当ヒントを追記する。
    """
    if not _enrich_enabled():
        return payload
    sheets = payload.get("sheets")
    if not isinstance(sheets, dict):
        return payload
    rt = sheets.get(RESULT_TASK_SHEET_NAME)
    if not isinstance(rt, dict):
        return payload
    tid_map = build_task_id_to_member_hint(rt)
    if not tid_map:
        logger.debug(
            "equipment_gantt_json_enrich: 結果_タスク一覧から担当ヒントが得られずスキップ"
        )
        return payload
    n_applied = 0
    for name, sheet in sheets.items():
        if not isinstance(name, str) or not isinstance(sheet, dict):
            continue
        if "設備ガント" not in name:
            continue
        enrich_equipment_gantt_sheet(sheet, tid_map)
        n_applied += 1
    if n_applied:
        payload["equipment_gantt_member_enrich"] = True
        logger.info(
            "equipment_gantt_json_enrich: 設備ガント系シート %s 件に担当ヒントをジョイン（タスクキー=%s）",
            n_applied,
            len(tid_map),
        )
    return payload
