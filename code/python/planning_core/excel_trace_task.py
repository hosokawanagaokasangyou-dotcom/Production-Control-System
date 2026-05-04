# -*- coding: utf-8 -*-
"""
段階2の Excel 生成経路を、1 件の依頼NOで NDJSON 追跡する（オプトイン）。

**設定（JavaFX 専用）**:
子プロセスの環境に ``PM_AI_EXCEL_TRACE_TASK_ID``（例: ``Y5-14``）を入れるのは
**本アプリの「環境変数」タブ**のみ。``workbook_env_bootstrap`` や OS の
``PM_AI_*`` は本ランチャー経路では使わない想定（``PythonProcessRunner`` が
UI に無い ``PM_AI_*`` を子へ渡さない）。

**ログファイル**:
``CURSOR_DEBUG_LOG`` / ``PM_AI_DEBUG_LOG`` が子へ渡っていればそのパス。なければ
``PM_AI_REPO_ROOT/.cursor/`` または本モジュール位置から解決したリポジトリ根の
``.cursor/debug-excel-trace.log``（親は自動作成）。

未設定（空）のときは一切書き込まない。
"""
from __future__ import annotations

import json
import os
import pathlib
import time

ENV_TRACE_TASK_ID = "PM_AI_EXCEL_TRACE_TASK_ID"
TASK_COL = "依頼NO"


def trace_task_id() -> str:
    return (os.environ.get(ENV_TRACE_TASK_ID) or "").strip()


def _log_path() -> str | None:
    p = (os.environ.get("CURSOR_DEBUG_LOG") or os.environ.get("PM_AI_DEBUG_LOG") or "").strip()
    if p:
        return p
    rr = (os.environ.get("PM_AI_REPO_ROOT") or "").strip()
    if rr:
        return str(pathlib.Path(rr) / ".cursor" / "debug-excel-trace.log")
    try:
        here = pathlib.Path(__file__).resolve()
        # code/python/planning_core/excel_trace_task.py -> parents[3] = リポジトリルート想定
        root = here.parents[3]
        return str(root / ".cursor" / "debug-excel-trace.log")
    except Exception:
        return None


def append(payload: dict) -> None:
    tid = trace_task_id()
    if not tid:
        return
    path = _log_path()
    if not path:
        return
    line = {
        "timestamp": int(time.time() * 1000),
        "traceTaskId": tid,
        **payload,
    }
    try:
        pathlib.Path(path).parent.mkdir(parents=True, exist_ok=True)
        with open(path, "a", encoding="utf-8") as fp:
            fp.write(json.dumps(line, ensure_ascii=False) + "\n")
    except Exception:
        pass


def log_df_tasks(df, stage: str, *, output_basename: str = "") -> None:
    """結果_タスク一覧相当の DataFrame から、追跡依頼の行をスナップショットする。"""
    if not trace_task_id() or df is None or getattr(df, "empty", True):
        return
    if TASK_COL not in df.columns:
        append(
            {
                "stage": stage,
                "hypothesisId": "EX1",
                "message": "df_tasks missing 依頼NO column",
                "outputBasename": output_basename,
                "sampleColumns": list(df.columns)[:30],
            }
        )
        return
    t = trace_task_id()
    sub = df[df[TASK_COL].astype(str).str.strip() == t]
    if sub.empty:
        append(
            {
                "stage": stage,
                "hypothesisId": "EX1",
                "message": "no matching row in df_tasks",
                "outputBasename": output_basename,
                "dfRowCount": int(len(df)),
            }
        )
        return
    row = {}
    for k in sub.columns:
        if str(k).startswith("履歴"):
            continue
        row[str(k)] = sub.iloc[0][k]
    if len(str(row)) > 12000:
        keys = list(row.keys())[:35]
        row = {k: row[k] for k in keys}
    append(
        {
            "stage": stage,
            "hypothesisId": "EX1",
            "message": "df_tasks row for trace id",
            "outputBasename": output_basename,
            "matchRows": int(len(sub)),
            "row": row,
        }
    )


def log_timeline_events(timeline_events: list | None, stage: str) -> None:
    if not trace_task_id() or not timeline_events:
        return
    t = trace_task_id()
    hit: list[dict] = []
    for ev in timeline_events:
        if not isinstance(ev, dict):
            continue
        if str(ev.get("task_id") or "").strip() != t:
            continue
        hit.append(ev)
    sample: list[dict] = []
    for ev in hit[:8]:
        sample.append(
            {
                "date": str(ev.get("date")),
                "machine": str(ev.get("machine") or "")[:120],
                "process": str(ev.get("process") or "")[:80],
            }
        )
    append(
        {
            "stage": stage,
            "hypothesisId": "EX2",
            "message": "timeline_events for trace id",
            "eventCount": len(hit),
            "totalTimelineEvents": len(timeline_events),
            "sample": sample,
        }
    )


def log_gantt_label_specs(label_specs: list | None, stage: str) -> None:
    """角丸ラベル spec リストから、テキストに依頼NOが含まれる件数（ヒューリスティック）。"""
    if not trace_task_id() or not label_specs:
        return
    t = trace_task_id()
    n_text = 0
    for sp in label_specs:
        if not isinstance(sp, dict):
            continue
        txt = str(sp.get("text") or sp.get("label") or "")
        if t in txt:
            n_text += 1
    append(
        {
            "stage": stage,
            "hypothesisId": "EX3",
            "message": "gantt timeline label specs (substring match on task id)",
            "specCountWithTaskSubstring": n_text,
            "totalSpecs": len(label_specs),
        }
    )


def log_sidecar_result_task_row(sidecar_path: str | None) -> None:
    """書き出した ``*_結果_タスク一覧.json`` から追跡行を読み戻す。"""
    if not trace_task_id():
        return
    t = trace_task_id()
    if not sidecar_path or not os.path.isfile(sidecar_path):
        append(
            {
                "stage": "sidecar_after_write",
                "hypothesisId": "EX4",
                "message": "sidecar path missing",
                "path": sidecar_path or "",
            }
        )
        return
    try:
        with open(sidecar_path, encoding="utf-8-sig") as fp:
            data = json.load(fp)
    except Exception as e:
        append(
            {
                "stage": "sidecar_after_write",
                "hypothesisId": "EX4",
                "message": "sidecar json read failed",
                "path": sidecar_path,
                "error": str(e)[:300],
            }
        )
        return
    rows = data.get("rows") if isinstance(data, dict) else None
    if not isinstance(rows, list):
        append(
            {
                "stage": "sidecar_after_write",
                "hypothesisId": "EX4",
                "message": "invalid sidecar structure",
                "path": sidecar_path,
            }
        )
        return
    found = None
    for r in rows:
        if isinstance(r, dict) and str(r.get(TASK_COL) or "").strip() == t:
            found = r
            break
    if not found:
        append(
            {
                "stage": "sidecar_after_write",
                "hypothesisId": "EX4",
                "message": "no row in sidecar json",
                "path": sidecar_path,
                "jsonRowCount": len(rows),
            }
        )
        return
    row = {k: v for k, v in found.items() if not str(k).startswith("履歴")}
    if len(str(row)) > 12000:
        row = dict(list(row.items())[:35])
    append(
        {
            "stage": "sidecar_after_write",
            "hypothesisId": "EX4",
            "message": "sidecar row for trace id",
            "path": sidecar_path,
            "row": row,
        }
    )
