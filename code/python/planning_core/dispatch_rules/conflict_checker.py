"""Conflict detection between enabled rules."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

ConflictKind = Literal[
    "effect_contradiction",
    "block_vs_require",
    "duplicate_scope_action",
    "apply_order_tie",
    "dsl_legacy_divergence",
    "pipeline_incompatible",
]
ConflictSeverity = Literal["error", "warning"]


@dataclass
class ConflictItem:
    kind: ConflictKind
    severity: ConflictSeverity
    rule_ids: list[str]
    message: str
    node_ids: list[str] = field(default_factory=list)


@dataclass
class ConflictReport:
    conflicts: list[ConflictItem] = field(default_factory=list)

    @property
    def error_count(self) -> int:
        return sum(1 for c in self.conflicts if c.severity == "error")

    @property
    def warning_count(self) -> int:
        return sum(1 for c in self.conflicts if c.severity == "warning")


def _speed_rules(rules: list[dict]) -> list[dict]:
    out = []
    for r in rules:
        if not r.get("enabled", True):
            continue
        graph = r.get("graph") or {}
        for n in graph.get("nodes") or []:
            if n.get("type") == "action.set_speed_mpm":
                out.append(r)
                break
    return out


def check_rule_conflicts(document: dict[str, Any]) -> ConflictReport:
    rules = [r for r in document.get("rules") or [] if isinstance(r, dict)]
    enabled = [r for r in rules if r.get("enabled", True)]
    report = ConflictReport()
    order_map: dict[int, list[str]] = {}
    for r in enabled:
        order = int(r.get("applyOrder", 100))
        order_map.setdefault(order, []).append(str(r.get("id", "")))
    for order, ids in order_map.items():
        if len(ids) > 1:
            report.conflicts.append(
                ConflictItem(
                    kind="apply_order_tie",
                    severity="warning",
                    rule_ids=ids,
                    message=f"applyOrder={order} が重複: {', '.join(ids)}",
                )
            )
    speed_rules = _speed_rules(enabled)
    if len(speed_rules) >= 2:
        ids = [str(r.get("id")) for r in speed_rules]
        report.conflicts.append(
            ConflictItem(
                kind="effect_contradiction",
                severity="warning",
                rule_ids=ids,
                message="複数の速度上書きルールが有効です（applyOrder で順序確認）",
            )
        )
    l10 = next((r for r in enabled if r.get("id") == "L10"), None)
    l13 = next((r for r in enabled if r.get("id") == "L13"), None)
    if l10 and l13:
        report.conflicts.append(
            ConflictItem(
                kind="pipeline_incompatible",
                severity="warning",
                rule_ids=["L10", "L13"],
                message="L10 と L13 が同時有効 — SEC ゲート経路の競合に注意",
            )
        )
    return report
