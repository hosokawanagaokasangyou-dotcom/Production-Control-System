"""Simulation steps for rule test lab."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from planning_core.dispatch_rules.context import RuleContext
from planning_core.dispatch_rules.engine import _topo_nodes
from planning_core.dispatch_rules.execution_planner import RuleDispatchPlanEntry, plan_execution
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
    roll_index: int = 0
    roll_total: int = 0
    wip_count: float = 0.0
    animation_kind: str = ""
    flow_phase: str = ""


@dataclass
class SimulationResult:
    steps: list[SimulationStep]
    final_blocked: bool
    summary_ja: str
    roll_total: int = 0
    blocked_at_roll: int = 0


def _parse_roll_count(task_row: dict) -> int:
    raw = task_row.get("配台ロール数") or task_row.get("dispatch_roll_count") or "1"
    try:
        return max(1, int(float(str(raw).strip())))
    except (TypeError, ValueError):
        return 1


def _initial_wip(context_overrides: dict | None) -> float:
    metrics = (context_overrides or {}).get("metrics") or {}
    for key in ("wip_connection_sec_initial", "initial_wip"):
        if key in metrics and metrics[key] is not None and str(metrics[key]).strip() != "":
            try:
                return float(metrics[key])
            except (TypeError, ValueError):
                pass
    return 0.0


def _task_process(task_row: dict) -> str:
    return str(task_row.get("工程名", task_row.get("process", ""))).strip()


def roll_stage_snapshot_pipeline(
    *,
    roll_total: int,
    connection_done: int,
    sec_done: int,
    on_connection: bool = False,
    on_sec: bool = False,
) -> dict[str, int]:
    """Four-column roll buckets for 接続→SEC pipeline (sums to roll_total)."""
    total = max(1, roll_total)
    cd = max(0, min(connection_done, total))
    sd = max(0, min(sec_done, total))
    oc = 1 if on_connection else 0
    os_ = 1 if on_sec else 0
    return {
        "pre_input_raw_rolls": max(0, total - cd - oc),
        "connection_rolls": oc,
        "sec_before_wip_rolls": max(0, cd - sd),
        "sec_complete_rolls": sd,
    }


def roll_stage_snapshot(
    task_row: dict,
    *,
    roll_index: int,
    roll_total: int,
    wip: float,
) -> dict[str, int]:
    """Roll-count buckets (always within 配台ロール数). wip is rule metric only."""
    proc = _task_process(task_row)
    total = max(1, roll_total)
    idx = max(0, roll_index)
    if proc == "SEC":
        return {
            "pre_input_raw_rolls": 0,
            "connection_rolls": 0,
            "sec_before_wip_rolls": max(0, total - idx + (1 if idx > 0 else 0)),
            "sec_complete_rolls": max(0, idx - 1),
        }
    return {
        "pre_input_raw_rolls": max(0, total - idx),
        "connection_rolls": 1 if idx > 0 else 0,
        "sec_before_wip_rolls": max(0, idx - 1),
        "sec_complete_rolls": 0,
    }


def _metrics_with_roll_stages(
    metrics: dict[str, Any],
    task_row: dict,
    *,
    roll_index: int,
    roll_total: int,
    wip: float,
) -> dict[str, Any]:
    out = dict(metrics)
    out.update(
        roll_stage_snapshot(
            task_row,
            roll_index=roll_index,
            roll_total=roll_total,
            wip=wip,
        )
    )
    return out

def _resolve_dsl_entries(
    doc: RuleDocument,
    plan,
    rule_id: str | None,
) -> list[RuleDispatchPlanEntry]:
    entries = [e for e in plan.entries if e.source == "dsl" and (rule_id is None or e.rule_id == rule_id)]
    if not entries and rule_id:
        target = next((r for r in doc.rules if r.id == rule_id and r.enabled), None)
        if target is not None and target.graph.nodes:
            entries = [
                RuleDispatchPlanEntry(
                    rule_id=target.id,
                    apply_order=target.apply_order,
                    source="dsl",
                    rule=target,
                )
            ]
    return entries


def _walk_graph(
    *,
    entries: list[RuleDispatchPlanEntry],
    ctx: RuleContext,
    metrics: dict[str, Any],
    seq_start: int,
    roll_index: int,
    roll_total: int,
) -> tuple[list[SimulationStep], int, RuleContext]:
    steps: list[SimulationStep] = []
    seq = seq_start
    task_row = dict(ctx.task or {})
    wip = float(metrics.get("wip_connection_sec", 0) or 0)
    staged_metrics = _metrics_with_roll_stages(
        metrics,
        task_row,
        roll_index=roll_index,
        roll_total=roll_total if roll_total > 0 else _parse_roll_count(task_row),
        wip=wip,
    )
    for entry in entries:
        state: dict[str, Any] = {
            "__edge_to_from__": {
                edge.to_node: edge.from_node for edge in entry.rule.graph.edges
            },
            "__edges__": list(entry.rule.graph.edges),
        }
        prev_node = None
        for node in _topo_nodes(entry.rule):
            seq += 1
            executor = get(node.type)
            effect = None
            if executor:
                executor(node, ctx, state)
                effect = state.get("effect")
            steps.append(
                SimulationStep(
                    sequence=seq,
                    phase=RulePhase.ELIGIBLE_FILTER.value,
                    rule_id=entry.rule_id,
                    node_id=node.id,
                    node_type=node.type,
                    edge_from=prev_node,
                    edge_to=node.id,
                    effect=effect,
                    summary_ja=node.label or node.type,
                    metrics=dict(staged_metrics),
                    task_snapshot=dict(ctx.task or {}),
                    roll_index=roll_index,
                    roll_total=roll_total,
                    wip_count=float(metrics.get("wip_connection_sec", 0) or 0),
                    animation_kind="graph_node",
                )
            )
            prev_node = node.id
    return steps, seq, ctx


def _merge_roll_metrics(
    metrics: dict[str, Any],
    *,
    roll_total: int,
    roll_index: int,
    connection_done: int,
    sec_done: int,
    on_connection: bool = False,
    on_sec: bool = False,
) -> dict[str, Any]:
    out = dict(metrics)
    out.update(
        roll_stage_snapshot_pipeline(
            roll_total=roll_total,
            connection_done=connection_done,
            sec_done=sec_done,
            on_connection=on_connection,
            on_sec=on_sec,
        )
    )
    return out


def _metrics_for_roll(
    *,
    roll_index: int,
    initial_wip: float,
    connection_done: int,
    sec_done: int,
    context_overrides: dict | None,
) -> dict[str, Any]:
    wip = initial_wip + float(connection_done - sec_done)
    metrics: dict[str, Any] = {"wip_connection_sec": wip}
    if context_overrides:
        extra = context_overrides.get("metrics") or {}
        if "request_roll_diff" in extra:
            metrics["request_roll_diff"] = float(extra["request_roll_diff"])
        else:
            metrics["request_roll_diff"] = max(6.0, float(roll_index))
        for key, val in extra.items():
            if key not in (
                "wip_connection_sec",
                "request_roll_diff",
                "wip_connection_sec_initial",
                "initial_wip",
            ):
                metrics[key] = val
    else:
        metrics["request_roll_diff"] = max(6.0, float(roll_index))
    return metrics


def _append_accumulate_step(
    all_steps: list[SimulationStep],
    *,
    seq: int,
    roll_idx: int,
    roll_total: int,
    flow_phase: str,
    task_row: dict,
    metrics: dict[str, Any],
    rule_id: str,
    summary_ja: str,
) -> None:
    all_steps.append(
        SimulationStep(
            sequence=seq,
            phase=RulePhase.ELIGIBLE_FILTER.value,
            rule_id=rule_id,
            node_id="__roll_accumulate__",
            node_type="animation.roll_accumulate",
            edge_from=None,
            edge_to=None,
            effect=None,
            summary_ja=summary_ja,
            metrics=dict(metrics),
            task_snapshot=dict(task_row),
            roll_index=roll_idx,
            roll_total=roll_total,
            wip_count=float(metrics.get("wip_connection_sec", 0) or 0),
            animation_kind="roll_accumulate",
            flow_phase=flow_phase,
        )
    )


def simulate_connection_sec_pipeline(
    document: dict,
    connection_row: dict,
    sec_row: dict,
    *,
    rule_id: str | None = None,
    day: str | None = None,
    context_overrides: dict | None = None,
) -> SimulationResult:
    """接続→SEC: 先に全ロール接続（WIP 蓄積）→ 続けて全ロール SEC まで試走。"""
    doc = RuleDocument.from_dict(document)
    rule_set = RuleSet(document=doc)
    plan = plan_execution(rule_set)
    entries = _resolve_dsl_entries(doc, plan, rule_id)
    if not entries:
        return SimulationResult(
            steps=[],
            final_blocked=False,
            summary_ja="試走対象ルールがありません",
        )

    roll_total = min(_parse_roll_count(connection_row), _parse_roll_count(sec_row))
    initial_wip = _initial_wip(context_overrides)
    rid = rule_id or (entries[0].rule_id if entries else "")
    all_steps: list[SimulationStep] = []
    seq = 0
    final_blocked = False
    blocked_at_roll = 0
    summary = f"接続→SEC 全 {roll_total} ロール完了"

    for roll_idx in range(1, roll_total + 1):
        conn_done_before = roll_idx - 1
        wip = initial_wip + float(roll_idx - 1)
        metrics = _metrics_for_roll(
            roll_index=roll_idx,
            initial_wip=initial_wip,
            connection_done=conn_done_before,
            sec_done=0,
            context_overrides=context_overrides,
        )
        metrics["wip_connection_sec"] = wip
        staged = _merge_roll_metrics(
            metrics,
            roll_total=roll_total,
            roll_index=roll_idx,
            connection_done=conn_done_before,
            sec_done=0,
            on_connection=True,
        )
        seq += 1
        _append_accumulate_step(
            all_steps,
            seq=seq,
            roll_idx=roll_idx,
            roll_total=roll_total,
            flow_phase="connection",
            task_row=connection_row,
            metrics=staged,
            rule_id=rid,
            summary_ja=(
                f"接続 ロール {roll_idx}/{roll_total} — "
                f"WIP={wip:.0f}（原反残{max(0, roll_total - roll_idx)}）"
            ),
        )

        ctx = RuleContext(
            phase=RulePhase.ELIGIBLE_FILTER.value,
            task=dict(connection_row),
            day=day,
            metrics=dict(metrics),
        )
        roll_steps, seq, ctx = _walk_graph(
            entries=entries,
            ctx=ctx,
            metrics=dict(staged),
            seq_start=seq,
            roll_index=roll_idx,
            roll_total=roll_total,
        )
        for step in roll_steps:
            step.flow_phase = "connection"
        all_steps.extend(roll_steps)

        if ctx.blocked:
            final_blocked = True
            blocked_at_roll = roll_idx
            summary = f"接続 ロール {roll_idx}/{roll_total}: {ctx.block_reason or '候補から除外'}"
            break

    if not final_blocked:
        conn_done = roll_total
        for roll_idx in range(1, roll_total + 1):
            sec_done_before = roll_idx - 1
            wip = initial_wip + float(roll_total - sec_done_before)
            metrics = _metrics_for_roll(
                roll_index=roll_idx,
                initial_wip=initial_wip,
                connection_done=conn_done,
                sec_done=sec_done_before,
                context_overrides=context_overrides,
            )
            metrics["wip_connection_sec"] = wip
            staged = _merge_roll_metrics(
                metrics,
                roll_total=roll_total,
                roll_index=roll_idx,
                connection_done=conn_done,
                sec_done=sec_done_before,
                on_sec=True,
            )
            seq += 1
            _append_accumulate_step(
                all_steps,
                seq=seq,
                roll_idx=roll_idx,
                roll_total=roll_total,
                flow_phase="sec",
                task_row=sec_row,
                metrics=staged,
                rule_id=rid,
                summary_ja=(
                    f"SEC ロール {roll_idx}/{roll_total} — "
                    f"WIP={wip:.0f}（SEC完了 {sec_done_before}/{roll_total}）"
                ),
            )

            ctx = RuleContext(
                phase=RulePhase.ELIGIBLE_FILTER.value,
                task=dict(sec_row),
                day=day,
                metrics=dict(metrics),
            )
            roll_steps, seq, ctx = _walk_graph(
                entries=entries,
                ctx=ctx,
                metrics=dict(staged),
                seq_start=seq,
                roll_index=roll_idx,
                roll_total=roll_total,
            )
            for step in roll_steps:
                step.flow_phase = "sec"
            all_steps.extend(roll_steps)

            if ctx.blocked:
                final_blocked = True
                blocked_at_roll = roll_idx
                summary = f"SEC ロール {roll_idx}/{roll_total}: {ctx.block_reason or '候補から除外'}"
                break

            sec_done = roll_idx
            metrics = _metrics_for_roll(
                roll_index=roll_idx,
                initial_wip=initial_wip,
                connection_done=conn_done,
                sec_done=sec_done,
                context_overrides=context_overrides,
            )
            metrics["wip_connection_sec"] = initial_wip + float(roll_total - sec_done)
            staged = _merge_roll_metrics(
                metrics,
                roll_total=roll_total,
                roll_index=roll_idx,
                connection_done=conn_done,
                sec_done=sec_done,
            )
            seq += 1
            _append_accumulate_step(
                all_steps,
                seq=seq,
                roll_idx=roll_idx,
                roll_total=roll_total,
                flow_phase="sec",
                task_row=sec_row,
                metrics=staged,
                rule_id=rid,
                summary_ja=f"SEC ロール {roll_idx}/{roll_total} 完了（SEC完了 {sec_done}/{roll_total}）",
            )

    return SimulationResult(
        steps=all_steps,
        final_blocked=final_blocked,
        summary_ja=summary,
        roll_total=roll_total,
        blocked_at_roll=blocked_at_roll,
    )


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
    metrics: dict[str, Any] = {"wip_connection_sec": 21, "request_roll_diff": 3}
    if context_overrides:
        metrics.update(context_overrides.get("metrics") or {})
    ctx = RuleContext(
        phase=RulePhase.ELIGIBLE_FILTER.value,
        task=dict(task_row),
        day=day,
        metrics=metrics,
    )
    entries = _resolve_dsl_entries(doc, plan, rule_id)
    steps, _, ctx = _walk_graph(
        entries=entries,
        ctx=ctx,
        metrics=metrics,
        seq_start=0,
        roll_index=0,
        roll_total=0,
    )
    summary = ctx.block_reason or ("候補から除外" if ctx.blocked else "通過")
    return SimulationResult(steps=steps, final_blocked=ctx.blocked, summary_ja=summary)


def simulate_task_all_rolls(
    document: dict,
    task_row: dict,
    *,
    rule_id: str | None = None,
    day: str | None = None,
    context_overrides: dict | None = None,
) -> SimulationResult:
    """Run one graph trial per roll; WIP increments between 接続 and SEC."""
    doc = RuleDocument.from_dict(document)
    rule_set = RuleSet(document=doc)
    plan = plan_execution(rule_set)
    entries = _resolve_dsl_entries(doc, plan, rule_id)
    if not entries:
        return SimulationResult(
            steps=[],
            final_blocked=False,
            summary_ja="試走対象ルールがありません",
        )

    roll_total = _parse_roll_count(task_row)
    initial_wip = _initial_wip(context_overrides)
    all_steps: list[SimulationStep] = []
    seq = 0
    final_blocked = False
    blocked_at_roll = 0
    summary = f"全 {roll_total} ロール通過"

    for roll_idx in range(1, roll_total + 1):
        proc = _task_process(task_row)
        if proc == "SEC":
            wip = max(0.0, initial_wip + float(roll_total - roll_idx))
            summary_wip = f"SEC前={wip:.0f}"
        else:
            wip = initial_wip + float(roll_idx - 1)
            summary_wip = f"接続WIP={wip:.0f}"
        metrics: dict[str, Any] = {
            "wip_connection_sec": wip,
        }
        if context_overrides:
            extra = context_overrides.get("metrics") or {}
            if "request_roll_diff" in extra:
                metrics["request_roll_diff"] = float(extra["request_roll_diff"])
            else:
                metrics["request_roll_diff"] = max(6.0, float(roll_idx))
            for key, val in extra.items():
                if key not in (
                    "wip_connection_sec",
                    "request_roll_diff",
                    "wip_connection_sec_initial",
                    "initial_wip",
                ):
                    metrics[key] = val
        else:
            metrics["request_roll_diff"] = max(6.0, float(roll_idx))

        seq += 1
        rid = rule_id or (entries[0].rule_id if entries else "")
        staged = _metrics_with_roll_stages(
            metrics,
            task_row,
            roll_index=roll_idx,
            roll_total=roll_total,
            wip=wip,
        )
        all_steps.append(
            SimulationStep(
                sequence=seq,
                phase=RulePhase.ELIGIBLE_FILTER.value,
                rule_id=rid,
                node_id="__roll_accumulate__",
                node_type="animation.roll_accumulate",
                edge_from=None,
                edge_to=None,
                effect=None,
                summary_ja=f"ロール {roll_idx}/{roll_total} — 接続WIP={wip:.0f}（原反残{max(0, roll_total - roll_idx)}）",
                metrics=dict(staged),
                task_snapshot=dict(task_row),
                roll_index=roll_idx,
                roll_total=roll_total,
                wip_count=wip,
                animation_kind="roll_accumulate",
            )
        )

        ctx = RuleContext(
            phase=RulePhase.ELIGIBLE_FILTER.value,
            task=dict(task_row),
            day=day,
            metrics=metrics,
        )
        roll_steps, seq, ctx = _walk_graph(
            entries=entries,
            ctx=ctx,
            metrics=metrics,
            seq_start=seq,
            roll_index=roll_idx,
            roll_total=roll_total,
        )
        all_steps.extend(roll_steps)

        if ctx.blocked:
            final_blocked = True
            blocked_at_roll = roll_idx
            summary = f"ロール {roll_idx}/{roll_total}: {ctx.block_reason or '候補から除外'}"
            break

    return SimulationResult(
        steps=all_steps,
        final_blocked=final_blocked,
        summary_ja=summary,
        roll_total=roll_total,
        blocked_at_roll=blocked_at_roll,
    )
