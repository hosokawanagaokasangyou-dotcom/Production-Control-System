"""Resolve enabled rules and execution source (dsl / legacy / skip)."""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Literal

from planning_core.dispatch_rules.schema import DispatchRule, RuleSet

ExecutionSource = Literal["dsl", "legacy", "skip"]


def engine_globally_enabled() -> bool:
    return os.environ.get("PM_AI_DISPATCH_RULE_ENGINE", "0").strip() in {
        "1",
        "true",
        "True",
        "yes",
        "on",
    }


def legacy_fallback_enabled() -> bool:
    return os.environ.get("PM_AI_DISPATCH_RULE_LEGACY_FALLBACK", "1").strip() not in {
        "0",
        "false",
        "False",
        "no",
        "off",
    }


@dataclass
class RuleDispatchPlanEntry:
    rule_id: str
    apply_order: int
    source: ExecutionSource
    rule: DispatchRule


@dataclass
class RuleDispatchPlan:
    entries: list[RuleDispatchPlanEntry]

    def source_for(self, rule_id: str) -> ExecutionSource | None:
        for e in self.entries:
            if e.rule_id == rule_id:
                return e.source
        return None


def _validate_graph(rule: DispatchRule) -> bool:
    return bool(rule.graph.nodes)


def resolve_source(rule: DispatchRule) -> ExecutionSource:
    if not rule.enabled:
        return "skip"
    mode = (rule.execution_mode or "auto").lower()
    if mode == "legacy":
        return "legacy"
    if mode == "dsl":
        return "dsl" if _validate_graph(rule) else "skip"
    # auto
    if not engine_globally_enabled():
        return "legacy"
    if _validate_graph(rule):
        return "dsl"
    if rule.legacy_fallback and legacy_fallback_enabled():
        return "legacy"
    return "skip"


def plan_execution(rule_set: RuleSet) -> RuleDispatchPlan:
    enabled = [r for r in rule_set.document.rules if r.enabled]
    enabled.sort(key=lambda r: (r.apply_order, r.id))
    entries: list[RuleDispatchPlanEntry] = []
    for rule in enabled:
        entries.append(
            RuleDispatchPlanEntry(
                rule_id=rule.id,
                apply_order=rule.apply_order,
                source=resolve_source(rule),
                rule=rule,
            )
        )
    return RuleDispatchPlan(entries=entries)
