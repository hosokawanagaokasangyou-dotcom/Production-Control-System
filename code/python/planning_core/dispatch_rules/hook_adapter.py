"""Thin entry from planning_core._core hook points."""

from __future__ import annotations

import os
from typing import Any, Callable

from planning_core.dispatch_rules.context import RuleContext
from planning_core.dispatch_rules.engine import get_plan, load_rules, run_phase
from planning_core.dispatch_rules.execution_planner import engine_globally_enabled
from planning_core.dispatch_rules.phases import RulePhase
from planning_core.dispatch_rules.paths import ensure_work_json_from_repo_template
from planning_core.dispatch_rules import trace_recorder


def ensure_rules_env() -> None:
    ensure_work_json_from_repo_template()


def apply_speed_special_rules(
    row: Any,
    task_id: str,
    machine: str,
    machine_name: str,
    speed: float,
    legacy_fn: Callable[..., float],
) -> float:
    """QUEUE_BUILD speed hook (L4 etc.)."""
    ensure_rules_env()
    legacy = lambda: legacy_fn(
        row=row, task_id=task_id, machine=machine, machine_name=machine_name, speed=speed
    )
    if not engine_globally_enabled():
        return legacy()
    load_rules()
    plan = get_plan()
    entry = next((e for e in plan.entries if e.rule_id == "L4"), None)
    if entry is None or entry.source != "dsl":
        return legacy()
    task_row = {
        "task_id": task_id,
        "工程名": machine,
        "機械名": machine_name,
    }
    try:
        from planning_core._core import PLAN_COL_PRODUCT_WIDTH, _planning_df_cell_scalar

        pw = _planning_df_cell_scalar(row, PLAN_COL_PRODUCT_WIDTH)
        task_row["製品幅"] = int(float(pw)) if pw is not None else None
    except Exception:
        pass
    ctx = RuleContext(phase=RulePhase.QUEUE_BUILD.value, task=task_row)
    run_phase(RulePhase.QUEUE_BUILD, ctx, plan)
    if ctx.speed_mpm is not None:
        return float(ctx.speed_mpm)
    return legacy()


def build_eligible_l13_context() -> dict:
    """eligible 1 回あたり 1 度だけ rules/plan を読む（タスクループ内の load_rules 回避）。"""
    ensure_rules_env()
    load_rules()
    plan = get_plan()
    entry = next((e for e in plan.entries if e.rule_id == "L13"), None)
    use_dsl = bool(
        entry is not None
        and entry.source == "dsl"
        and engine_globally_enabled()
    )
    return {"plan": plan, "use_dsl": use_dsl}


def eligible_l13_connection_skip_with_context(
    task: dict,
    wip_connection_before_sec: float | None,
    task_queue: list,
    l13_ctx: dict,
) -> bool:
    """ELIGIBLE_FILTER L13 — True => exclude task (continue in eligible loop)."""
    if l13_ctx.get("use_dsl"):
        metrics = {"wip_connection_sec": float(wip_connection_before_sec or 0)}
        ctx = RuleContext(
            phase=RulePhase.ELIGIBLE_FILTER.value,
            task=task,
            task_queue=task_queue,
            metrics=metrics,
        )
        result = run_phase(RulePhase.ELIGIBLE_FILTER, ctx, l13_ctx["plan"])
        if ctx.blocked:
            return True
        if result.modified:
            return ctx.blocked
        return False
    legacy_skip = _legacy_l13_connection_skip(task, wip_connection_before_sec)
    if legacy_skip:
        trace_recorder.append_event(
            task_id=str(task.get("task_id", "")),
            day=None,
            rule_id="L13",
            apply_order=40,
            execution_source="legacy",
            phase=RulePhase.ELIGIBLE_FILTER.value,
            effect="block_candidate",
            reason_code="H-SR-B6",
            summary_ja="WIP上限により接続を当日候補から除外（従来）",
        )
    return legacy_skip


def eligible_l13_connection_skip(
    task: dict,
    wip_connection_before_sec: float | None,
    task_queue: list,
) -> bool:
    """ELIGIBLE_FILTER L13 — True => exclude task (continue in eligible loop)."""
    return eligible_l13_connection_skip_with_context(
        task,
        wip_connection_before_sec,
        task_queue,
        build_eligible_l13_context(),
    )


def _legacy_l13_connection_skip(
    task: dict,
    wip_connection_before_sec: float | None,
) -> bool:
    limit_raw = os.environ.get("WIP_LIMIT_CONNECTION_BEFORE_SEC_ROLLS", "20")
    try:
        limit = float(limit_raw)
    except ValueError:
        limit = 20.0
    if limit <= 0 or wip_connection_before_sec is None:
        return False
    if wip_connection_before_sec < limit:
        return False
    try:
        from planning_core._core import (
            SPECIAL_WIP_CONNECTION_MACHINE,
            SPECIAL_WIP_CONNECTION_PROCESS,
            _normalize_equipment_match_key,
            _normalize_process_name_for_rule_match,
        )
    except ImportError:
        return False
    proc = _normalize_process_name_for_rule_match(task.get("machine"))
    mach = _normalize_equipment_match_key(task.get("machine_name"))
    return (
        proc == _normalize_process_name_for_rule_match(SPECIAL_WIP_CONNECTION_PROCESS)
        and mach == _normalize_equipment_match_key(SPECIAL_WIP_CONNECTION_MACHINE)
    )


def preview_task(document: dict, task_row: dict, *, rule_id: str | None = None) -> str:
    from planning_core.dispatch_rules.preview import preview_one_line

    return preview_one_line(document, task_row, rule_id=rule_id)
