"""Phase 2/3 node stubs (legacy bridge when graph present)."""

from __future__ import annotations

from planning_core.dispatch_rules.context import RuleContext
from planning_core.dispatch_rules.nodes.registry import register
from planning_core.dispatch_rules.schema import RuleNode


def _noop(node: RuleNode, ctx: RuleContext, state: dict) -> dict:
    state["effect"] = state.get("effect")
    return state


def _register_phase23_stubs() -> None:
    for node_type in (
        "action.set_min_team",
        "action.set_required_team",
        "action.reorder_trial_adjacent",
        "action.set_sort_tier",
        "action.trial_order_priority",
        "action.timeline_start_floor",
        "scope.roll_pipeline_b2",
        "action.two_phase_inspection",
    ):
        register(node_type, _noop)


_register_phase23_stubs()
