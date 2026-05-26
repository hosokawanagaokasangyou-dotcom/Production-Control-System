"""Conflict checker tests."""

from planning_core.dispatch_rules.conflict_checker import check_rule_conflicts


def test_apply_order_tie_warning():
    doc = {
        "schemaVersion": 1,
        "rules": [
            {"id": "L4", "enabled": True, "applyOrder": 10, "graph": {"nodes": [], "edges": []}},
            {"id": "L6", "enabled": True, "applyOrder": 10, "graph": {"nodes": [], "edges": []}},
        ],
    }
    report = check_rule_conflicts(doc)
    assert report.warning_count >= 1
