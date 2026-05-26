"""Thin bridge to legacy _core special rule functions."""

from __future__ import annotations

from typing import Any, Callable

from planning_core.dispatch_rules.context import RuleContext


def call_legacy(legacy_fn: Callable[..., Any], ctx: RuleContext, **kwargs: Any) -> Any:
    """Invoke legacy callable with ctx extras merged."""
    merged = ctx.to_legacy_dict()
    merged.update(kwargs)
    return legacy_fn(**merged) if _accepts_kwargs(legacy_fn) else legacy_fn(ctx.task, ctx.task_queue)


def _accepts_kwargs(fn: Callable[..., Any]) -> bool:
    import inspect

    try:
        sig = inspect.signature(fn)
    except (TypeError, ValueError):
        return False
    params = sig.parameters
    return any(
        p.kind in (inspect.Parameter.VAR_KEYWORD, inspect.Parameter.KEYWORD_ONLY)
        or p.name in ("ctx", "context", "task", "task_queue")
        for p in params.values()
    )
