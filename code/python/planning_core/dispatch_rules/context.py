"""Rule evaluation context wrapper."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class RuleContext:
    phase: str
    task: dict[str, Any] | None = None
    task_queue: list[dict[str, Any]] = field(default_factory=list)
    day: str | None = None
    metrics: dict[str, Any] = field(default_factory=dict)
    extras: dict[str, Any] = field(default_factory=dict)
    blocked: bool = False
    block_reason: str | None = None
    speed_mpm: float | None = None
    trace_events: list[dict[str, Any]] = field(default_factory=list)

    @classmethod
    def from_legacy_dict(cls, phase: str, legacy_ctx: dict[str, Any]) -> RuleContext:
        return cls(
            phase=phase,
            task=legacy_ctx.get("task"),
            task_queue=list(legacy_ctx.get("task_queue") or []),
            day=legacy_ctx.get("day"),
            metrics=dict(legacy_ctx.get("metrics") or {}),
            extras=legacy_ctx,
        )

    def to_legacy_dict(self) -> dict[str, Any]:
        out = dict(self.extras)
        out.update(
            {
                "task": self.task,
                "task_queue": self.task_queue,
                "day": self.day,
                "metrics": self.metrics,
                "blocked": self.blocked,
                "block_reason": self.block_reason,
                "speed_mpm": self.speed_mpm,
            }
        )
        return out
