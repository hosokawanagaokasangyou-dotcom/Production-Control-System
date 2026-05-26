"""Dispatch special rules DSL engine (§B)."""

from planning_core.dispatch_rules.engine import load_rules, plan_execution, run_phase
from planning_core.dispatch_rules.phases import RulePhase
from planning_core.dispatch_rules.schema import (
    CURRENT_SCHEMA_VERSION,
    ENGINE_IMPLEMENTATION_VERSION,
    SUPPORTED_SCHEMA_MAX,
)

__all__ = [
    "CURRENT_SCHEMA_VERSION",
    "ENGINE_IMPLEMENTATION_VERSION",
    "SUPPORTED_SCHEMA_MAX",
    "RulePhase",
    "load_rules",
    "plan_execution",
    "run_phase",
]
