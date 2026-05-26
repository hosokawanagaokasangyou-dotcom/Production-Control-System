"""Simulation steps for rule test lab."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from planning_core.dispatch_rules.context import RuleContext
from planning_core.dispatch_rules.engine import _topo_nodes, get_plan, load_rules, run_phase
from planning_core.dispatch_rules.execution_planner import plan_execution
from planning_core.dispatch_rules.nodes.registry import get
from planning_core.dispatch_rules.phases import RulePhase
from planning_core.dispatch_rules.schema import RuleDocument, RuleSet


@dataclass
class SimulationStep:
    sequence: int
    phase: str
    rule_id: str
    node_id: str
    node_type: str
    edge_from: str | None
    edge_to: str | None
    effect: str | None
    summary_ja: str
    metrics: dict[str, Any] = field(default_factory=dict)
    task_snapshot: dict[str, Any] = field(default_factory=dict)


@dataclass
class SimulationResult:
    steps: list[SimulationStep]
    final_blocked: bool
    summary_ja: str


def simulate_task(
    document: dict,
    task_row: dict,
    *,
    rule_id: str | None = None,
    day: str | None = None,
    context_overrides: dict | None = None,
) -> SimulationResult:
    doc = RuleDocument.from_dict(document)
    rule_set = RuleSet(document=doc)
    plan = plan_execution(rule_set)
    metrics = {"wip_connection_sec": 21, "request_roll_diff": 3}
    if context_overrides:
        metrics.update(context_overrides.get("metrics") or {})
    ctx = RuleContext(
        phase=RulePhase.ELIGIBLE_FILTER.value,
        task=dict(task_row),
        day=day,
        metrics=metrics,
    )
    steps: list[SimulationStep] = []
    seq = 0
    entries = [e for e in plan.entries if e.source == "dsl" and (rule_id is None or e.rule_id == rule_id)]
    for entry in entries:
        state: dict[str, Any] = {}
        prev_node = None
        for node in _topo_nodes(entry.rule):
            seq += 1
            executor = get(node.type)
            effect = None
            if executor:
                executor(node, ctx, state)
                effect = state.get("effect")
            edge_from = prev_node
            edge_to = node.id
            prev_node = node.id
            steps.append(
                SimulationStep(
                    sequence=seq,
                    phase=RulePhase.ELIGIBLE_FILTER.value,
                    rule_id=entry.rule_id,
                    node_id=node.id,
                    node_type=node.type,
                    edge_from=edge_from,
                    edge_to=edge_to,
                    effect=effect,
                    summary_ja=node.label or node.type,
                    metrics=dict(metrics),
                    task_snapshot=dict(ctx.task or {}),
                )
            )
    summary = ctx.block_reason or ("候補から除外" if ctx.blocked else "通過")
    return SimulationResult(steps=steps, final_blocked=ctx.blocked, summary_ja=summary)
