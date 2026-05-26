"""Execution planner tests."""

import os

from planning_core.dispatch_rules.execution_planner import plan_execution, resolve_source
from planning_core.dispatch_rules.schema import DispatchRule, RuleDocument, RuleGraph, RuleSet


def test_disabled_rule_skipped():
    rule = DispatchRule(id="L4", enabled=False, apply_order=10)
    assert resolve_source(rule) == "skip"


def test_legacy_mode_when_engine_off():
    os.environ.pop("PM_AI_DISPATCH_RULE_ENGINE", None)
    rule = DispatchRule(id="L13", enabled=True, execution_mode="auto", graph=RuleGraph(nodes=[]))
    assert resolve_source(rule) == "legacy"
