"""One-line preview for Inspector."""

from __future__ import annotations

from typing import Any

from planning_core.dispatch_rules.engine import run_phase
from planning_core.dispatch_rules.context import RuleContext
from planning_core.dispatch_rules.execution_planner import plan_execution
from planning_core.dispatch_rules.phases import RulePhase
from planning_core.dispatch_rules.schema import RuleDocument, RuleSet


def preview_one_line(document: dict, task_row: dict, *, rule_id: str | None = None) -> str:
    doc = RuleDocument.from_dict(document)
    rule_set = RuleSet(document=doc)
    plan = plan_execution(rule_set)
    ctx = RuleContext(
        phase=RulePhase.ELIGIBLE_FILTER.value,
        task=task_row,
        metrics={"wip_connection_sec": 21, "request_roll_diff": 3},
    )
    targets = [e for e in plan.entries if e.source == "dsl" and (rule_id is None or e.rule_id == rule_id)]
    if not targets:
        return "（試走タブで詳細確認）"
    for entry in targets:
        run_phase(RulePhase.ELIGIBLE_FILTER, ctx, plan)
        if ctx.blocked:
            return f"{entry.rule_id}: {ctx.block_reason or '候補から除外'}"
        if ctx.speed_mpm is not None:
            return f"{entry.rule_id}: 速度→{ctx.speed_mpm}m/分"
    return "特別ルールによる変更なし"
