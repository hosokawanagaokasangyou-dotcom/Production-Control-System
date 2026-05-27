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


def test_l13_all_rolls_blocks_when_wip_reaches_threshold():
    prev = os.environ.pop("PM_AI_DISPATCH_RULE_ENGINE", None)
    try:
        doc = json.loads(TEMPLATE.read_text(encoding="utf-8"))
        for r in doc["rules"]:
            if r["id"] == "L13":
                r["executionMode"] = "auto"
        task = {
            "task_id": "Y6-3-接続-01",
            "工程名": "接続",
            "機械名": "熱融着機 湖南",
            "依頼NO": "Y6-3",
            "配台ロール数": "16",
        }
        from planning_core.dispatch_rules.simulation import simulate_task_all_rolls

        result = simulate_task_all_rolls(
            doc,
            task,
            rule_id="L13",
            context_overrides={"metrics": {"initial_wip": 5}},
        )
        assert result.roll_total == 16
        assert result.final_blocked
        assert result.blocked_at_roll == 16
        assert any(s.animation_kind == "roll_accumulate" for s in result.steps)
        assert len(result.steps) > 16
    finally:
        if prev is None:
            os.environ.pop("PM_AI_DISPATCH_RULE_ENGINE", None)
        else:
            os.environ["PM_AI_DISPATCH_RULE_ENGINE"] = prev


def test_roll_stage_snapshot_connection_and_sec():
    conn_row = {"工程名": "接続", "配台ロール数": "16"}
    sec_row = {"工程名": "SEC", "配台ロール数": "16"}
    from planning_core.dispatch_rules.simulation import roll_stage_snapshot

    s3 = roll_stage_snapshot(conn_row, roll_index=3, roll_total=16, wip=7.0)
    assert s3["pre_input_raw_rolls"] == 13
    assert s3["connection_rolls"] == 1
    assert s3["sec_before_wip_rolls"] == 2
    assert s3["sec_complete_rolls"] == 0

    s7 = roll_stage_snapshot(sec_row, roll_index=7, roll_total=16, wip=14.0)
    assert s7["pre_input_raw_rolls"] == 0
    assert s7["connection_rolls"] == 0
    assert s7["sec_before_wip_rolls"] == 10
    assert s7["sec_complete_rolls"] == 6


def test_roll_stage_snapshot_pipeline():
    from planning_core.dispatch_rules.simulation import roll_stage_snapshot_pipeline

    s = roll_stage_snapshot_pipeline(
        roll_total=16,
        connection_done=3,
        sec_done=2,
        on_connection=True,
    )
    assert s["pre_input_raw_rolls"] == 12
    assert s["connection_rolls"] == 1
    assert s["sec_before_wip_rolls"] == 1
    assert s["sec_complete_rolls"] == 2
    assert sum(s.values()) == 16

    done = roll_stage_snapshot_pipeline(roll_total=16, connection_done=16, sec_done=16)
    assert done["sec_complete_rolls"] == 16
    assert sum(done.values()) == 16


def test_connection_sec_pipeline_completes_all_rolls():
    prev = os.environ.pop("PM_AI_DISPATCH_RULE_ENGINE", None)
    try:
        doc = json.loads(TEMPLATE.read_text(encoding="utf-8"))
        for r in doc["rules"]:
            if r["id"] == "L13":
                r["executionMode"] = "auto"
        conn = {
            "task_id": "Y6-3-接続-01",
            "工程名": "接続",
            "機械名": "熱融着機 湖南",
            "依頼NO": "Y6-3",
            "配台ロール数": "16",
        }
        sec = {
            "task_id": "Y6-3-SEC-01",
            "工程名": "SEC",
            "機械名": "SEC機 湖南",
            "依頼NO": "Y6-3",
            "配台ロール数": "16",
        }
        from planning_core.dispatch_rules.simulation import simulate_connection_sec_pipeline

        result = simulate_connection_sec_pipeline(
            doc,
            conn,
            sec,
            rule_id="L13",
            context_overrides={"metrics": {"initial_wip": 0}},
        )
        assert result.roll_total == 16
        assert not result.final_blocked
        assert result.summary_ja.startswith("接続→SEC")
        assert any(s.flow_phase == "connection" for s in result.steps)
        assert any(s.flow_phase == "sec" for s in result.steps)
        last = result.steps[-1]
        assert last.metrics["sec_complete_rolls"] == 16
        assert last.metrics["pre_input_raw_rolls"] == 0
    finally:
        if prev is None:
            os.environ.pop("PM_AI_DISPATCH_RULE_ENGINE", None)
        else:
            os.environ["PM_AI_DISPATCH_RULE_ENGINE"] = prev


def test_connection_sec_pipeline_blocks_at_l13():
    prev = os.environ.pop("PM_AI_DISPATCH_RULE_ENGINE", None)
    try:
        doc = json.loads(TEMPLATE.read_text(encoding="utf-8"))
        for r in doc["rules"]:
            if r["id"] == "L13":
                r["executionMode"] = "auto"
        conn = {
            "task_id": "Y6-3-接続-01",
            "工程名": "接続",
            "機械名": "熱融着機 湖南",
            "依頼NO": "Y6-3",
            "配台ロール数": "16",
        }
        sec = {
            "task_id": "Y6-3-SEC-01",
            "工程名": "SEC",
            "機械名": "SEC機 湖南",
            "依頼NO": "Y6-3",
            "配台ロール数": "16",
        }
        from planning_core.dispatch_rules.simulation import simulate_connection_sec_pipeline

        result = simulate_connection_sec_pipeline(
            doc,
            conn,
            sec,
            rule_id="L13",
            context_overrides={"metrics": {"initial_wip": 5}},
        )
        assert result.final_blocked
        assert result.blocked_at_roll == 16
    finally:
        if prev is None:
            os.environ.pop("PM_AI_DISPATCH_RULE_ENGINE", None)
        else:
            os.environ["PM_AI_DISPATCH_RULE_ENGINE"] = prev


def test_l13_simulation_auto_mode_without_engine_still_walks_graph():
    prev = os.environ.pop("PM_AI_DISPATCH_RULE_ENGINE", None)
    try:
        doc = json.loads(TEMPLATE.read_text(encoding="utf-8"))
        for r in doc["rules"]:
            if r["id"] == "L13":
                r["executionMode"] = "auto"
        task = {
            "task_id": "Y6-3-接続-01",
            "工程名": "接続",
            "機械名": "熱融着機 湖南",
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
    finally:
        if prev is None:
            os.environ.pop("PM_AI_DISPATCH_RULE_ENGINE", None)
        else:
            os.environ["PM_AI_DISPATCH_RULE_ENGINE"] = prev
