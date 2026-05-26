"""Rule evaluation phases aligned with _core.py hook points."""

from enum import Enum


class RulePhase(str, Enum):
    QUEUE_BUILD = "queue_build"
    TRIAL_ORDER = "trial_order"
    SORT_KEY = "sort_key"
    ELIGIBLE_FILTER = "eligible_filter"
    ASSIGN_PROBE = "assign_probe"
    NEED_EXPLORE = "need_explore"
    TIMELINE = "timeline"

    @classmethod
    def from_str(cls, value: str) -> RulePhase | None:
        try:
            return cls(str(value))
        except ValueError:
            return None
