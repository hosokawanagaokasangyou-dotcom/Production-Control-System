"""Node executor registry."""

from __future__ import annotations

from typing import Any, Callable

from planning_core.dispatch_rules.context import RuleContext
from planning_core.dispatch_rules.schema import RuleNode

Executor = Callable[[RuleNode, RuleContext, dict[str, Any]], dict[str, Any] | None]

_REGISTRY: dict[str, Executor] = {}


def register(node_type: str, fn: Executor) -> None:
    _REGISTRY[node_type] = fn


def get(node_type: str) -> Executor | None:
    return _REGISTRY.get(node_type)


def _scope_matches(node: RuleNode, ctx: RuleContext) -> bool:
    task = ctx.task or {}
    params = node.params
    process = str(params.get("process_name", "")).strip()
    machine = str(params.get("machine_name", "")).strip()
    if process and str(task.get("工程名", task.get("process", ""))).strip() != process:
        return False
    if machine and str(task.get("機械名", task.get("machine", ""))).strip() != machine:
        return False
    return True


def _exec_scope(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    state["scope_match"] = _scope_matches(node, ctx)
    return state


def _exec_filter_row(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    if not state.get("scope_match", True):
        state["condition_pass"] = False
        return state
    task = ctx.task or {}
    conditions = node.params.get("conditions") or []
    require_all = bool(node.params.get("require_all", True))
    results: list[bool] = []
    for cond in conditions:
        col = str(cond.get("column", ""))
        op = str(cond.get("op", "eq"))
        expected = cond.get("value")
        actual = task.get(col)
        if op == "eq":
            results.append(str(actual) == str(expected))
        elif op == "ne":
            results.append(str(actual) != str(expected))
        else:
            results.append(False)
    if not results:
        state["condition_pass"] = True
    elif require_all:
        state["condition_pass"] = all(results)
    else:
        state["condition_pass"] = any(results)
    return state


def _exec_compare(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    left = state.get("metric_value")
    right = state.get("threshold_value")
    if left is None or right is None:
        state["compare_pass"] = False
        return state
    op = str(node.params.get("operator", ">="))
    if op == ">=":
        state["compare_pass"] = float(left) >= float(right)
    elif op == ">":
        state["compare_pass"] = float(left) > float(right)
    elif op == "<=":
        state["compare_pass"] = float(left) <= float(right)
    elif op == "<":
        state["compare_pass"] = float(left) < float(right)
    elif op == "==":
        state["compare_pass"] = float(left) == float(right)
    else:
        state["compare_pass"] = False
    return state


def _exec_const_number(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    import os

    env_key = str(node.params.get("env_key", "")).strip()
    if env_key:
        raw = os.environ.get(env_key, str(node.params.get("value", 0)))
        try:
            state["threshold_value"] = float(raw)
        except ValueError:
            state["threshold_value"] = float(node.params.get("value", 0) or 0)
    else:
        state["threshold_value"] = float(node.params.get("value", 0) or 0)
    return state


def _exec_metric_wip(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    metrics = ctx.metrics
    key = str(node.params.get("metric_key", "wip_connection_sec"))
    state["metric_value"] = float(metrics.get(key, metrics.get("wip_total", 0)) or 0)
    return state


def _exec_metric_roll_diff(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    metrics = ctx.metrics
    key = str(node.params.get("metric_key", "request_roll_diff"))
    state["metric_value"] = float(metrics.get(key, 0) or 0)
    return state


def _exec_block_candidate(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    if state.get("compare_pass"):
        ctx.blocked = True
        ctx.block_reason = str(
            node.params.get("summary_ja", node.params.get("reason_code", "blocked"))
        )
        state["effect"] = "block_candidate"
    return state


def _exec_block_downstream(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    if state.get("compare_pass"):
        downstream = str(node.params.get("downstream_process", "SEC"))
        task = ctx.task or {}
        proc = str(task.get("工程名", task.get("process", ""))).strip()
        if proc == downstream:
            ctx.blocked = True
            ctx.block_reason = str(
                node.params.get("summary_ja", node.params.get("reason_code", "blocked"))
            )
            state["effect"] = "block_downstream"
    return state


def _exec_set_speed(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    if state.get("scope_match", True) and state.get("condition_pass", True):
        ctx.speed_mpm = float(node.params.get("speed_mpm", 20))
        state["effect"] = "set_speed_mpm"
    return state


def _register_defaults() -> None:
    register("scope.process_machine", _exec_scope)
    register("scope.process_pipeline", _exec_scope)
    register("filter.row_conditions", _exec_filter_row)
    register("compare.threshold", _exec_compare)
    register("const.number", _exec_const_number)
    register("metric.wip_total_rolls", _exec_metric_wip)
    register("metric.request_roll_diff", _exec_metric_roll_diff)
    register("action.block_candidate", _exec_block_candidate)
    register("action.block_downstream", _exec_block_downstream)
    register("action.set_speed_mpm", _exec_set_speed)


_register_defaults()

from planning_core.dispatch_rules.nodes import phase23_stubs as _phase23_stubs  # noqa: F401
