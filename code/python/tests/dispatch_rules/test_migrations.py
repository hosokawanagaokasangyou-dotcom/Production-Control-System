"""Tests for dispatch_rules migrations."""

import json
from pathlib import Path

from planning_core.dispatch_rules.migrations import migrate_document
from planning_core.dispatch_rules.schema import CURRENT_SCHEMA_VERSION

FIXTURES = Path(__file__).parent / "fixtures"


def test_v0_to_v1_priority_becomes_apply_order():
    raw = {
        "version": 0,
        "rules": [{"id": "L13", "priority": 40, "graph": {"nodes": [], "edges": []}}],
    }
    out = migrate_document(raw)
    assert out["schemaVersion"] == CURRENT_SCHEMA_VERSION
    assert out["rules"][0]["applyOrder"] == 40
    assert "priority" not in out["rules"][0]


def test_template_loads():
    template = FIXTURES / "dispatch_special_rules_v1.json"
    if not template.is_file():
        template = Path(__file__).resolve().parents[3] / "json" / "dispatch_special_rules" / "dispatch_special_rules.json"
    raw = json.loads(template.read_text(encoding="utf-8"))
    out = migrate_document(raw)
    assert out["schemaVersion"] == 1
    assert any(r["id"] == "L13" for r in out["rules"])
