"""DSL document schema constants and dataclasses."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

CURRENT_SCHEMA_VERSION = 1
SUPPORTED_SCHEMA_MAX = 1
ENGINE_IMPLEMENTATION_VERSION = "1.0.0"

EXECUTION_MODES = frozenset({"auto", "dsl", "legacy"})


@dataclass
class RuleEdge:
    id: str
    from_node: str
    to_node: str
    from_port: str = "out"
    to_port: str = "in"

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> RuleEdge:
        return cls(
            id=str(raw.get("id", "")),
            from_node=str(raw.get("from", "")),
            to_node=str(raw.get("to", "")),
            from_port=str(raw.get("fromPort", "out")),
            to_port=str(raw.get("toPort", "in")),
        )


@dataclass
class RuleNode:
    id: str
    type: str
    label: str = ""
    x: float = 0.0
    y: float = 0.0
    params: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> RuleNode:
        return cls(
            id=str(raw.get("id", "")),
            type=str(raw.get("type", "")),
            label=str(raw.get("label", "")),
            x=float(raw.get("x", 0) or 0),
            y=float(raw.get("y", 0) or 0),
            params=dict(raw.get("params") or {}),
        )


@dataclass
class RuleGraph:
    nodes: list[RuleNode] = field(default_factory=list)
    edges: list[RuleEdge] = field(default_factory=list)

    @classmethod
    def from_dict(cls, raw: dict[str, Any] | None) -> RuleGraph:
        if not raw:
            return cls()
        nodes = [RuleNode.from_dict(n) for n in raw.get("nodes") or []]
        edges = [RuleEdge.from_dict(e) for e in raw.get("edges") or []]
        return cls(nodes=nodes, edges=edges)


@dataclass
class DispatchRule:
    id: str
    name: str = ""
    enabled: bool = True
    apply_order: int = 100
    execution_mode: str = "auto"
    legacy_fallback: bool = True
    graph: RuleGraph = field(default_factory=RuleGraph)

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> DispatchRule:
        return cls(
            id=str(raw.get("id", "")),
            name=str(raw.get("name", "")),
            enabled=bool(raw.get("enabled", True)),
            apply_order=int(raw.get("applyOrder", raw.get("priority", 100)) or 100),
            execution_mode=str(raw.get("executionMode", "auto")),
            legacy_fallback=bool(raw.get("legacyFallback", True)),
            graph=RuleGraph.from_dict(raw.get("graph")),
        )


@dataclass
class RuleDocument:
    schema_version: int
    engine_min_version: str = ENGINE_IMPLEMENTATION_VERSION
    saved_at: str = ""
    rules: list[DispatchRule] = field(default_factory=list)
    raw: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> RuleDocument:
        ver = raw.get("schemaVersion")
        if ver is None and "version" in raw:
            ver = raw.get("version")
        schema_version = int(ver or 0)
        rules_raw = raw.get("rules") or []
        return cls(
            schema_version=schema_version,
            engine_min_version=str(raw.get("engineMinVersion", ENGINE_IMPLEMENTATION_VERSION)),
            saved_at=str(raw.get("savedAt", "")),
            rules=[DispatchRule.from_dict(r) for r in rules_raw if isinstance(r, dict)],
            raw=raw,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "schemaVersion": self.schema_version,
            "engineMinVersion": self.engine_min_version,
            "savedAt": self.saved_at,
            "rules": [
                {
                    "id": r.id,
                    "name": r.name,
                    "enabled": r.enabled,
                    "applyOrder": r.apply_order,
                    "executionMode": r.execution_mode,
                    "legacyFallback": r.legacy_fallback,
                    "graph": {
                        "nodes": [
                            {
                                "id": n.id,
                                "type": n.type,
                                "label": n.label,
                                "x": n.x,
                                "y": n.y,
                                "params": n.params,
                            }
                            for n in r.graph.nodes
                        ],
                        "edges": [
                            {
                                "id": e.id,
                                "from": e.from_node,
                                "to": e.to_node,
                                "fromPort": e.from_port,
                                "toPort": e.to_port,
                            }
                            for e in r.graph.edges
                        ],
                    },
                }
                for r in self.rules
            ],
        }


@dataclass
class RuleSet:
    document: RuleDocument
    source_path: str | None = None
