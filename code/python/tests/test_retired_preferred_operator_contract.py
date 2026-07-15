# -*- coding: utf-8 -*-
"""廃止した単一担当OP指定をAI契約・キャッシュから除去するテスト。"""

from __future__ import annotations

from datetime import date

import pandas as pd

from planning_core import _core as pc


def test_task_special_ai_entries_drop_legacy_preferred_operator_without_mutation():
    raw = {
        "REQ-1": {
            "preferred_operator": "甲",
            "priority": 1,
        },
        "REQ-2": [
            {"preferred_operator": "乙", "start_date": "2026-07-20"},
        ],
    }

    sanitized = pc._drop_retired_preferred_operator_ai_keys(raw)

    assert sanitized == {
        "REQ-1": {"priority": 1},
        "REQ-2": [{"start_date": "2026-07-20"}],
    }
    assert raw["REQ-1"]["preferred_operator"] == "甲"


def test_global_ai_contract_drops_legacy_task_preferred_operators():
    raw = {
        "task_preferred_operators": {"REQ-1": "甲"},
        "global_day_process_operator_rules": [
            {
                "date": "2026-07-20",
                "process_contains": "工程A",
                "operator_names": ["甲", "乙"],
            }
        ],
    }

    coerced = pc._coerce_global_priority_override_dict(raw, 2026)

    assert "task_preferred_operators" not in coerced
    assert coerced["global_day_process_operator_rules"][0]["operator_names"] == [
        "甲",
        "乙",
    ]


def test_ai_cache_namespaces_change_after_preferred_operator_retirement():
    assert pc.TASK_SPECIAL_CACHE_KEY_PREFIX.startswith("TASK_SPECIAL_v4|")
    assert pc.GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX.startswith("GLOBAL_PRIO_v9|")


def test_task_queue_keeps_global_speed_rules_after_preferred_operator_retirement():
    tasks = pd.DataFrame(
        [
            {
                pc.TASK_COL_TASK_ID: "REQ-1",
                pc.TASK_COL_MACHINE: "工程A",
                pc.TASK_COL_MACHINE_NAME: "設備1",
                pc.TASK_COL_QTY: 100,
                pc.TASK_COL_UNPROCESSED: 100,
                pc.TASK_COL_SPEED: 10,
            }
        ]
    )

    queue = pc.build_task_queue_from_planning_df(
        tasks,
        date(2026, 7, 15),
        {},
        ai_by_tid={},
        global_priority_override={
            "global_speed_rules": [
                {
                    "process_contains": "工程A",
                    "machine_contains": "",
                    "speed_multiplier": 0.5,
                }
            ]
        },
        equipment_list=[],
    )

    assert queue
    assert queue[0][pc.TASK_COL_SPEED] == 5.0
