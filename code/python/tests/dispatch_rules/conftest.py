"""Pytest setup for dispatch_rules (avoid full planning_core bootstrap)."""

import sys

# Minimal path for isolated dispatch_rules tests
if "code/python" not in sys.path[0]:
    pass
