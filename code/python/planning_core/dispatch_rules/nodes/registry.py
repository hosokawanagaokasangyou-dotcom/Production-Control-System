"""Node executor registry."""

from __future__ import annotations

import re
import unicodedata
from typing import Any, Callable

from planning_core.dispatch_rules.context import RuleContext
from planning_core.dispatch_rules.schema import RuleNode

Executor = Callable[[RuleNode, RuleContext, dict[str, Any]], dict[str, Any] | None]

_REGISTRY: dict[str, Executor] = {}


def register(node_type: str, fn: Executor) -> None:
    _REGISTRY[node_type] = fn


def get(node_type: str) -> Executor | None:
    return _REGISTRY.get(node_type)


def _normalize_machine_name(val: object) -> str:
    if val is None:
        return ""
    t = unicodedata.normalize("NFKC", str(val))
    t = t.replace("\u00a0", " ").replace("\u3000", " ")
    t = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", t)
    return re.sub(r"\s+", " ", t).strip()


def _machine_matches(rule_machine: str, task_machine: str) -> bool:
    rule_m = _normalize_machine_name(rule_machine)
    task_m = _normalize_machine_name(task_machine)
    if not rule_m:
        return True
    if not task_m:
        return False
    if task_m == rule_m:
        return True
    return task_m.startswith(rule_m + " ") or task_m.startswith(rule_m)


def _scope_matches(node: RuleNode, ctx: RuleContext) -> bool:
    task = ctx.task or {}
    params = node.params
    process = str(params.get("process_name", "")).strip()
    machine = str(params.get("machine_name", "")).strip()
    task_proc = str(task.get("工程名", task.get("process", ""))).strip()
    task_mach = str(task.get("機械名", task.get("machine_name", task.get("machine", "")))).strip()
    if process and task_proc != process:
        return False
    if machine and not _machine_matches(machine, task_mach):
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
    edges = state.get("__edges__") or []
    for edge in edges:
        if edge.to_node != node.id:
            continue
        if edge.to_port == "threshold":
            right = state.get(f"threshold:{edge.from_node}", right)
        else:
            left = state.get(f"metric:{edge.from_node}", left)
    if left is None or right is None:
        passed = False
    else:
        op = str(node.params.get("operator", ">="))
        if op == ">=":
            passed = float(left) >= float(right)
        elif op == ">":
            passed = float(left) > float(right)
        elif op == "<=":
            passed = float(left) <= float(right)
        elif op == "<":
            passed = float(left) < float(right)
        elif op == "==":
            passed = float(left) == float(right)
        else:
            passed = False
    compares = state.setdefault("compare_by_node", {})
    compares[node.id] = passed
    state["compare_pass"] = passed
    return state


def _inbound_compare_pass(node: RuleNode, state: dict[str, Any]) -> bool:
    edge_from = (state.get("__edge_to_from__") or {}).get(node.id)
    compares = state.get("compare_by_node") or {}
    if edge_from and edge_from in compares:
        return bool(compares[edge_from])
    return bool(state.get("compare_pass"))


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
    state[f"threshold:{node.id}"] = state["threshold_value"]
    return state


def _exec_metric_wip(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    metrics = ctx.metrics
    key = str(node.params.get("metric_key", "wip_connection_sec"))
    value = float(metrics.get(key, metrics.get("wip_total", 0)) or 0)
    state["metric_value"] = value
    state[f"metric:{node.id}"] = value
    return state


def _exec_metric_roll_diff(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    metrics = ctx.metrics
    key = str(node.params.get("metric_key", "request_roll_diff"))
    value = float(metrics.get(key, 0) or 0)
    state["metric_value"] = value
    state[f"metric:{node.id}"] = value
    return state


def _exec_block_candidate(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    if _inbound_compare_pass(node, state):
        ctx.blocked = True
        ctx.block_reason = str(
            node.params.get("summary_ja", node.params.get("reason_code", "blocked"))
        )
        state["effect"] = "block_candidate"
    return state


def _exec_block_downstream(node: RuleNode, ctx: RuleContext, state: dict[str, Any]) -> dict[str, Any] | None:
    if _inbound_compare_pass(node, state):
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
