# -*- coding: utf-8 -*-
"""工程名+機械名の配台優先（エンボス連続）。試行順の連番塊は使わない。"""

from planning_core.core.process_machine_priority import (
    PRIORITY_HIGH,
    PRIORITY_HIGHEST,
    PRIORITY_LOW,
    PRIORITY_NORMAL,
    consecutive_prefer_sort_penalty,
    default_priority_rules,
    last_process_for_occupancy,
    match_rule,
    may_skip_lower_trial_blocker,
    parse_priority_label,
    process_from_eq_line,
    priority_rank_for_task,
    rules_from_document,
    should_defer_blocker_behind_consecutive_peers,
)


def test_default_is_emboss_normal_consecutive_any_machine():
    rules = default_priority_rules()
    assert len(rules) == 1
    r = rules[0]
    assert r.process_name_key == "エンボス"
    assert r.machine_name_key == ""
    assert r.priority == PRIORITY_NORMAL
    assert r.consecutive is True


def test_parse_priority_labels():
    assert parse_priority_label("最優先") == PRIORITY_HIGHEST
    assert parse_priority_label("優先") == PRIORITY_HIGH
    assert parse_priority_label("通常") == PRIORITY_NORMAL
    assert parse_priority_label("優先度低") == PRIORITY_LOW
    assert parse_priority_label("") == PRIORITY_NORMAL


def test_match_emboss_ignores_space_variants():
    rules = default_priority_rules()
    task = {"machine": "エンボス", "machine_name": "エンボス　湖南"}
    assert match_rule(task, rules) is not None
    split_task = {"machine": "分割", "machine_name": "エンボス　湖南"}
    assert match_rule(split_task, rules) is None


def test_machine_specific_rule_does_not_match_other_machine():
    doc = {
        "processMachinePriorities": [
            {
                "processName": "エンボス",
                "machineName": "エンボス　湖南",
                "priority": "通常",
                "consecutive": True,
            }
        ]
    }
    rules = rules_from_document(doc)
    assert match_rule({"machine": "エンボス", "machine_name": "エンボス　湖南"}, rules)
    assert match_rule({"machine": "エンボス", "machine_name": "エンボス　国分"}, rules) is None


def test_missing_key_uses_default_empty_array_disables():
    assert len(rules_from_document({})) == 1
    assert rules_from_document({"processMachinePriorities": []}) == []


def test_normal_skips_other_process_only_when_last_was_emboss():
    rules = default_priority_rules()
    emboss = {"machine": "エンボス", "machine_name": "エンボス　湖南"}
    split = {"machine": "分割", "machine_name": "エンボス　湖南"}
    assert not may_skip_lower_trial_blocker(
        emboss, split, last_process="", rules=rules
    )
    assert may_skip_lower_trial_blocker(
        emboss, split, last_process="エンボス", rules=rules
    )
    assert not may_skip_lower_trial_blocker(
        emboss, emboss, last_process="エンボス", rules=rules
    )


def test_highest_skips_other_process_even_on_first_job():
    rules = rules_from_document(
        {
            "processMachinePriorities": [
                {
                    "processName": "エンボス",
                    "machineName": "",
                    "priority": "最優先",
                    "consecutive": True,
                }
            ]
        }
    )
    emboss = {"machine": "エンボス", "machine_name": "エンボス　湖南"}
    split = {"machine": "分割", "machine_name": "エンボス　湖南"}
    assert may_skip_lower_trial_blocker(
        emboss, split, last_process="", rules=rules
    )


def test_low_priority_never_skips_lower_trial():
    rules = rules_from_document(
        {
            "processMachinePriorities": [
                {
                    "processName": "エンボス",
                    "machineName": "",
                    "priority": "優先度低",
                    "consecutive": True,
                }
            ]
        }
    )
    emboss = {"machine": "エンボス", "machine_name": "エンボス　湖南"}
    split = {"machine": "分割", "machine_name": "エンボス　湖南"}
    assert not may_skip_lower_trial_blocker(
        emboss, split, last_process="エンボス", rules=rules
    )


def test_sort_penalty_prefers_same_process_after_last_emboss():
    rules = default_priority_rules()
    handoff = {"last_eq": {"occ": "エンボス"}}
    emboss = {
        "machine": "エンボス",
        "machine_name": "エンボス　湖南",
        "equipment_line_key": "occ",
    }
    split = {
        "machine": "分割",
        "machine_name": "エンボス　湖南",
        "equipment_line_key": "occ",
    }
    assert consecutive_prefer_sort_penalty(emboss, handoff, occupancy="occ", rules=rules) == 0
    assert consecutive_prefer_sort_penalty(split, handoff, occupancy="occ", rules=rules) == 1


def test_unlisted_process_has_normal_rank():
    rules = default_priority_rules()
    assert priority_rank_for_task({"machine": "スリット"}, rules) == PRIORITY_NORMAL
    assert priority_rank_for_task({"machine": "エンボス"}, rules) == PRIORITY_NORMAL


def test_process_from_eq_line_extracts_process_name():
    assert process_from_eq_line("エンボス+エンボス　湖南") == "エンボス"
    assert process_from_eq_line("エンボス") == "エンボス"


def test_sort_penalty_reads_process_from_equipment_line_last_eq():
    rules = default_priority_rules()
    handoff = {"last_eq": {"occ": "エンボス+エンボス　湖南"}}
    emboss = {"machine": "エンボス", "machine_name": "エンボス　湖南"}
    split = {"machine": "分割", "machine_name": "エンボス　湖南"}
    assert consecutive_prefer_sort_penalty(emboss, handoff, occupancy="occ", rules=rules) == 0
    assert consecutive_prefer_sort_penalty(split, handoff, occupancy="occ", rules=rules) == 1
    assert last_process_for_occupancy(handoff, "occ") == "エンボス"


def test_specific_machine_rule_beats_empty_machine():
    rules = rules_from_document(
        {
            "processMachinePriorities": [
                {
                    "processName": "エンボス",
                    "machineName": "",
                    "priority": "通常",
                    "consecutive": True,
                },
                {
                    "processName": "エンボス",
                    "machineName": "エンボス　湖南",
                    "priority": "最優先",
                    "consecutive": True,
                },
            ]
        }
    )
    honan = {"machine": "エンボス", "machine_name": "エンボス　湖南"}
    kokubu = {"machine": "エンボス", "machine_name": "エンボス　国分"}
    assert match_rule(honan, rules).priority == PRIORITY_HIGHEST
    assert match_rule(kokubu, rules).priority == PRIORITY_NORMAL


def test_same_request_never_skips_predecessor_even_if_highest():
    rules = rules_from_document(
        {
            "processMachinePriorities": [
                {
                    "processName": "エンボス",
                    "machineName": "",
                    "priority": "最優先",
                    "consecutive": True,
                }
            ]
        }
    )
    emboss = {
        "machine": "エンボス",
        "machine_name": "エンボス　湖南",
        "task_id": "REQ-1",
        "process_sequence_rank": 1,
    }
    split = {
        "machine": "分割",
        "machine_name": "エンボス　湖南",
        "task_id": "REQ-1",
        "process_sequence_rank": 0,
    }
    other_split = {
        "machine": "分割",
        "machine_name": "エンボス　湖南",
        "task_id": "REQ-2",
        "process_sequence_rank": 0,
    }
    assert not may_skip_lower_trial_blocker(
        emboss, split, last_process="", rules=rules
    )
    assert may_skip_lower_trial_blocker(
        emboss, other_split, last_process="", rules=rules
    )
    branched_emboss = {
        "machine": "エンボス",
        "machine_name": "エンボス　湖南",
        "task_id": "REQ-1-B",
        "rule_task_id": "REQ-1",
    }
    branched_split = {
        "machine": "分割",
        "machine_name": "エンボス　湖南",
        "task_id": "REQ-1-A",
        "rule_task_id": "REQ-1",
    }
    assert not may_skip_lower_trial_blocker(
        branched_emboss, branched_split, last_process="", rules=rules
    )


def test_same_request_does_not_defer_predecessor_behind_own_emboss():
    rules = default_priority_rules()
    emboss = {
        "machine": "エンボス",
        "machine_name": "エンボス　湖南",
        "occ": "m1",
        "task_id": "REQ-1",
    }
    split = {
        "machine": "分割",
        "machine_name": "エンボス　湖南",
        "occ": "m1",
        "task_id": "REQ-1",
    }

    def occ_of(t):
        return t["occ"]

    assert not should_defer_blocker_behind_consecutive_peers(
        split,
        [emboss, split],
        occupancy="m1",
        last_process="エンボス",
        occupancy_of=occ_of,
        rules=rules,
    )


def test_defer_split_when_consecutive_emboss_peer_can_skip():
    rules = default_priority_rules()
    emboss = {"machine": "エンボス", "machine_name": "エンボス　湖南", "occ": "m1"}
    split = {"machine": "分割", "machine_name": "エンボス　湖南", "occ": "m1"}
    other = {"machine": "検査", "machine_name": "検査　湖南", "occ": "m2"}

    def occ_of(t):
        return t["occ"]

    assert should_defer_blocker_behind_consecutive_peers(
        split,
        [emboss, split, other],
        occupancy="m1",
        last_process="エンボス",
        occupancy_of=occ_of,
        rules=rules,
    )
    assert not should_defer_blocker_behind_consecutive_peers(
        split,
        [emboss, split, other],
        occupancy="m1",
        last_process="",
        occupancy_of=occ_of,
        rules=rules,
    )
