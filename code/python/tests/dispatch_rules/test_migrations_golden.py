"""Golden migration tests."""

import json
from pathlib import Path

from planning_core.dispatch_rules.migrations import migrate_document
from planning_core.dispatch_rules.schema import CURRENT_SCHEMA_VERSION

FIXTURES = Path(__file__).parent / "fixtures"


def test_v0_fixture_golden():
    raw = json.loads((FIXTURES / "dispatch_special_rules_v0.json").read_text(encoding="utf-8"))
    out = migrate_document(raw)
    golden = json.loads((FIXTURES / "dispatch_special_rules_v1_golden.json").read_text(encoding="utf-8"))
    assert out["schemaVersion"] == CURRENT_SCHEMA_VERSION
    assert out["rules"][0]["applyOrder"] == golden["rules"][0]["applyOrder"]
    assert out["rules"][0]["enabled"] is True
    assert out["rules"][0]["executionMode"] == "auto"
