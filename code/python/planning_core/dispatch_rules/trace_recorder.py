"""Trace sidecar recording."""

from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from planning_core.dispatch_rules.paths import work_dir_from_summary_workbook

_active_run_snapshot_id: str | None = None
_sequence = 0
_events: list[dict[str, Any]] = []


def set_active_run_snapshot(run_snapshot_id: str | None) -> None:
    global _active_run_snapshot_id
    _active_run_snapshot_id = run_snapshot_id


def reset_trace() -> None:
    global _sequence, _events
    _sequence = 0
    _events = []


def append_event(
    *,
    task_id: str,
    day: str | None,
    rule_id: str,
    apply_order: int,
    execution_source: str,
    phase: str,
    effect: str,
    reason_code: str = "",
    summary_ja: str = "",
    graph_node_ids: list[str] | None = None,
) -> dict[str, Any]:
    global _sequence
    _sequence += 1
    event = {
        "task_id": task_id,
        "day": day or "",
        "rule_id": rule_id,
        "apply_order": apply_order,
        "sequence_in_run": _sequence,
        "run_snapshot_id": _active_run_snapshot_id or "",
        "execution_source": execution_source,
        "phase": phase,
        "effect": effect,
        "reason_code": reason_code,
        "summary_ja": summary_ja,
        "graph_node_ids": graph_node_ids or [],
        "recorded_at": datetime.now(timezone.utc).isoformat(),
    }
    _events.append(event)
    return event


def events() -> list[dict[str, Any]]:
    return list(_events)


def write_sidecar(path: Path | None = None) -> Path | None:
    work = work_dir_from_summary_workbook()
    if work is None and path is None:
        return None
    target = path or (work / "dispatch_rule_applications.json")
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = {"events": _events, "count": len(_events)}
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return target
