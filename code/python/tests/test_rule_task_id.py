# -*- coding: utf-8 -*-
"""_rule_task_id（特別ルール・WIP・工程依存の集計キー）の単体テスト。"""
from planning_core import _core as pc


def test_rule_task_id_branch_uses_parent():
    task = {"task_id": "Y3-24-01", "rule_task_id": "Y3-24"}
    assert pc._rule_task_id(task) == "Y3-24"


def test_rule_task_id_defaults_to_task_id_when_no_rule_id():
    assert pc._rule_task_id({"task_id": "Y3-24"}) == "Y3-24"
    assert pc._rule_task_id({"task_id": "Y3-24", "rule_task_id": ""}) == "Y3-24"


def test_rule_task_id_non_dict_or_empty():
    assert pc._rule_task_id(None) == ""
    assert pc._rule_task_id({}) == ""
