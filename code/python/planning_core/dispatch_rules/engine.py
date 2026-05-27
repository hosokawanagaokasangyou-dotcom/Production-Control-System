"""DSL rule engine: load, plan, run phases."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from planning_core.dispatch_rules.context import RuleContext
from planning_core.dispatch_rules.execution_planner import RuleDispatchPlan, plan_execution
from planning_core.dispatch_rules.migrations import migrate_document
from planning_core.dispatch_rules.nodes.registry import get
from planning_core.dispatch_rules.phases import RulePhase
from planning_core.dispatch_rules.schema import RuleDocument, RuleSet
from planning_core.dispatch_rules import trace_recorder


@dataclass
class RulePhaseResult:
    ctx: RuleContext
    modified: bool = False
    events: list[dict[str, Any]] = field(default_factory=list)


_plan_cache: RuleDispatchPlan | None = None
_rule_set_cache: RuleSet | None = None


def load_rules(path: str | None = None) -> RuleSet:
    global _plan_cache, _rule_set_cache
    from planning_core.dispatch_rules.paths import ensure_work_json_from_repo_template, resolve_dispatch_special_rules_json

    resolved = path or resolve_dispatch_special_rules_json()
    if not resolved:
        ensured = ensure_work_json_from_repo_template()
        resolved = ensured
    if not resolved:
        empty = RuleSet(document=RuleDocument(schema_version=1), source_path=None)
        _rule_set_cache = empty
        _plan_cache = plan_execution(empty)
        return empty
    raw = json.loads(Path(resolved).read_text(encoding="utf-8"))
    migrated = migrate_document(raw)
    doc = RuleDocument.from_dict(migrated)
    rule_set = RuleSet(document=doc, source_path=resolved)
    _rule_set_cache = rule_set
    _plan_cache = plan_execution(rule_set)
    return rule_set


def get_plan() -> RuleDispatchPlan:
    global _plan_cache
    if _plan_cache is None:
        load_rules()
    return _plan_cache  # type: ignore[return-value]


def _topo_nodes(rule) -> list:
    nodes = {n.id: n for n in rule.graph.nodes}
    if not nodes:
        return []
    incoming = {nid: 0 for nid in nodes}
    adjacency: dict[str, list[str]] = {nid: [] for nid in nodes}
    for edge in rule.graph.edges:
        if edge.from_node in nodes and edge.to_node in incoming:
            incoming[edge.to_node] += 1
            adjacency[edge.from_node].append(edge.to_node)
    queue = [nodes[nid] for nid, count in incoming.items() if count == 0]
    ordered = []
    while queue:
        node = queue.pop(0)
        ordered.append(node)
        for dest in adjacency.get(node.id, []):
            incoming[dest] -= 1
            if incoming[dest] == 0:
                queue.append(nodes[dest])
    seen = {n.id for n in ordered}
    for node in rule.graph.nodes:
        if node.id not in seen:
            ordered.append(node)
    return ordered


def _run_dsl_rule(rule, ctx: RuleContext, phase: RulePhase) -> RulePhaseResult:
    state: dict[str, Any] = {
        "scope_match": True,
        "condition_pass": True,
        "__edge_to_from__": {edge.to_node: edge.from_node for edge in rule.graph.edges},
        "__edges__": list(rule.graph.edges),
    }
    events: list[dict[str, Any]] = []
    for node in _topo_nodes(rule):
        executor = get(node.type)
        if executor is None:
            continue
        executor(node, ctx, state)
        if state.get("effect"):
            events.append(
                trace_recorder.append_event(
                    task_id=str((ctx.task or {}).get("task_id", "")),
                    day=ctx.day,
                    rule_id=rule.id,
                    apply_order=rule.apply_order,
                    execution_source="dsl",
                    phase=phase.value,
                    effect=str(state.get("effect")),
                    reason_code=str(node.params.get("reason_code", "")),
                    summary_ja=str(node.params.get("summary_ja", node.label)),
                    graph_node_ids=[node.id],
                )
            )
            state.pop("effect", None)
    modified = ctx.blocked or ctx.speed_mpm is not None
    return RulePhaseResult(ctx=ctx, modified=modified, events=events)


def run_phase(phase: RulePhase, ctx: RuleContext, plan: RuleDispatchPlan | None = None) -> RulePhaseResult:
    plan = plan or get_plan()
    merged = RulePhaseResult(ctx=ctx)
    for entry in plan.entries:
        if entry.source == "skip":
            continue
        if entry.source == "legacy":
            continue
        result = _run_dsl_rule(entry.rule, ctx, phase)
        merged.modified = merged.modified or result.modified
        merged.events.extend(result.events)
        if ctx.blocked:
            break
    return merged


def invalidate_cache() -> None:
    global _plan_cache, _rule_set_cache
    _plan_cache = None
    _rule_set_cache = None
