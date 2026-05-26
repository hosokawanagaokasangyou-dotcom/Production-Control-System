"""Simulation golden tests."""

import json
import os
from pathlib import Path

from planning_core.dispatch_rules.simulation import simulate_task

TEMPLATE = Path(__file__).resolve().parents[3] / "json" / "dispatch_special_rules" / "dispatch_special_rules.json"


def test_l13_simulation_blocks_at_high_wip():
    os.environ["PM_AI_DISPATCH_RULE_ENGINE"] = "1"
    doc = json.loads(TEMPLATE.read_text(encoding="utf-8"))
    for r in doc["rules"]:
        if r["id"] == "L13":
            r["executionMode"] = "dsl"
    task = {
        "task_id": "Y6-3-接続-01",
        "工程名": "接続",
        "機械名": "熱融着機",
        "依頼NO": "Y6-3",
    }
    result = simulate_task(
        doc,
        task,
        rule_id="L13",
        context_overrides={"metrics": {"wip_connection_sec": 21, "request_roll_diff": 3}},
    )
    assert result.steps
    assert result.final_blocked
