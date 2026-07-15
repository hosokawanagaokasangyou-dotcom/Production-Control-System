# -*- coding: utf-8 -*-
"""段階2「担当OP_限定」の解析・チーム制約テスト。"""

from __future__ import annotations

from datetime import date, datetime, timedelta

import pytest

from planning_core import _core as pc


def _task(**overrides):
    task = {
        "task_id": "REQ-42",
        "planning_excel_row": 7,
        "machine": "工程A",
        "machine_name": "機械1",
        "limited_operator_names": (),
        "preferred_operator_raw": "",
    }
    task.update(overrides)
    return task


def _role_map(mapping):
    return lambda member: mapping.get(member, (None, 999))


def test_limited_operator_blank_is_backward_compatible():
    assert pc._parse_limited_operator_json_cell("", 7, "REQ-42") == ()
    assert pc._parse_limited_operator_json_cell(None, 7, "REQ-42") == ()

    result = pc._prepare_limited_operator_team_constraints(
        _task(), ["甲", "乙"], _role_map({"甲": ("OP", 1)}), ["甲"], None
    )

    assert result is None


@pytest.mark.parametrize(
    ("raw", "reason"),
    [
        ("{", "不正JSON"),
        ('{"name":"甲"}', "配列"),
        ("[]", "1名以上"),
        ('["甲", 1]', "文字列"),
        ('["甲", " "]', "空名"),
        ('["甲", "甲"]', "重複名"),
        ('["富田 太郎", "冨田太郎"]', "重複名"),
    ],
)
def test_limited_operator_json_errors_include_excel_row_request_and_reason(raw, reason):
    with pytest.raises(pc.PlanningValidationError) as exc_info:
        pc._parse_limited_operator_json_cell(raw, 7, "REQ-42")

    message = str(exc_info.value)
    assert "Excel行7" in message
    assert "依頼NO=REQ-42" in message
    assert reason in message


def test_limited_operator_resolves_in_input_order_with_existing_normalization():
    resolved = pc._resolve_limited_operator_names(
        ("富田 太郎", "乙 花子"),
        ["冨田太郎", "乙 花子"],
        excel_row_number=7,
        task_id="REQ-42",
    )

    assert resolved == ("冨田太郎", "乙 花子")


@pytest.mark.parametrize(
    ("selected", "members", "reason"),
    [
        (("不存在",), ["甲"], "未知名"),
        (("佐藤",), ["佐藤 太郎", "佐藤 花子"], "曖昧名"),
    ],
)
def test_limited_operator_rejects_unknown_and_ambiguous_names(
    selected, members, reason
):
    with pytest.raises(pc.PlanningValidationError) as exc_info:
        pc._resolve_limited_operator_names(
            selected,
            members,
            excel_row_number=7,
            task_id="REQ-42",
        )

    assert reason in str(exc_info.value)


def test_limited_operator_overrides_need_members_and_preset_for_only_that_row():
    task = _task(limited_operator_names=("乙", "甲"), preferred_operator_raw="甲")

    result = pc._prepare_limited_operator_team_constraints(
        task,
        ["甲", "乙", "丙"],
        _role_map({"甲": ("OP", 1), "乙": ("AS", 2), "丙": ("OP", 3)}),
        ["甲", "乙", "丙"],
        "甲",
    )

    assert result["required_count"] == 2
    assert result["max_count"] == 2
    assert result["capable_members"] == ["乙", "甲"]
    assert result["fixed_team"] == ["乙", "甲"]
    assert result["preset_rows"] == [(0, 2, ("乙", "甲"), None)]
    assert result["preferred_member"] == "甲"


def test_limited_operator_preferred_outside_team_is_conflict():
    with pytest.raises(pc.PlanningValidationError, match="矛盾"):
        pc._prepare_limited_operator_team_constraints(
            _task(limited_operator_names=("甲",), preferred_operator_raw="乙"),
            ["甲", "乙"],
            _role_map({"甲": ("OP", 1), "乙": ("OP", 2)}),
            ["甲", "乙"],
            "乙",
        )


@pytest.mark.parametrize(
    ("roles", "reason"),
    [
        ({"甲": (None, 999), "乙": ("OP", 1)}, "資格外"),
        ({"甲": ("AS", 1), "乙": ("AS", 2)}, "最低1名OP"),
    ],
)
def test_limited_operator_rejects_unqualified_or_as_only_team(roles, reason):
    with pytest.raises(pc.PlanningValidationError) as exc_info:
        pc._prepare_limited_operator_team_constraints(
            _task(limited_operator_names=("甲", "乙")),
            ["甲", "乙"],
            _role_map(roles),
            ["甲", "乙"],
            None,
        )

    assert reason in str(exc_info.value)


def test_limited_operator_nonattendance_does_not_fallback_to_other_member():
    task = _task(limited_operator_names=("甲", "乙"))

    result = pc._prepare_limited_operator_team_constraints(
        task,
        ["甲", "乙", "丙"],
        _role_map({"甲": ("OP", 1), "乙": ("AS", 2), "丙": ("OP", 3)}),
        ["甲", "丙"],
        None,
    )

    assert result["capable_members"] == []
    assert "非出勤" in task["_limited_operator_last_rejection_reason"]


@pytest.mark.parametrize(
    "reason",
    ["休憩", "終業", "二重配台", "設備占有"],
)
def test_limited_operator_dynamic_rejection_is_reported_at_calendar_end(reason):
    task = _task(
        limited_operator_names=("甲",),
        remaining_units=1,
        initial_remaining_units=1,
        unit_m=100,
    )
    pc._record_limited_operator_rejection(task, reason)

    with pytest.raises(pc.PlanningValidationError) as exc_info:
        pc._raise_limited_operator_remaining_tasks(
            [task], date(2026, 7, 31), context_label="段階2"
        )

    message = str(exc_info.value)
    assert "Excel行7" in message
    assert "依頼NO=REQ-42" in message
    assert reason in message


def test_new_and_legacy_paths_share_same_limited_operator_helper():
    assert (
        pc._new_dispatch_limited_operator_constraints
        is pc._legacy_dispatch_limited_operator_constraints
        is pc._prepare_limited_operator_team_constraints
    )


def test_limited_row_never_abolishes_equipment_occupancy_in_either_path():
    override = {"abolish_all_scheduling_limits": True}
    limited_task = _task(limited_operator_names=("甲",))
    ordinary_task = _task()

    assert pc._scheduling_limits_abolished_for_task(override, limited_task) is False
    assert pc._scheduling_limits_abolished_for_task(override, ordinary_task) is True
    assert (
        pc._new_dispatch_scheduling_limits_abolished
        is pc._legacy_dispatch_scheduling_limits_abolished
        is pc._scheduling_limits_abolished_for_task
    )
    assert pc._machine_occupancy_tracking_required(
        override, [ordinary_task, limited_task]
    )
    assert not pc._machine_occupancy_tracking_required(
        override, [ordinary_task]
    )


def test_l2_fallback_keeps_limited_selected_count_and_team():
    limited = {
        "required_count": 2,
        "capable_members": ["乙", "甲"],
        "fixed_team": ["乙", "甲"],
    }

    assert pc._l2_fallback_required_count(1, limited) == 2
    assert pc._l2_fallback_required_count(1, None) == 1
    assert limited["capable_members"] == ["乙", "甲"]
    assert limited["fixed_team"] == ["乙", "甲"]


def test_equipment_protection_is_bidirectional_for_limited_and_ordinary_rows():
    base = datetime(2026, 7, 15, 9, 0)
    ordinary = _task()
    limited = _task(limited_operator_names=("甲",))
    override = {"abolish_all_scheduling_limits": True}
    protected = pc._LimitedEquipmentProtection()

    # 未指定→限定: 限定行は通常の設備空き時刻を無視できない。
    assert pc._limited_equipment_interval_blocked(
        limited,
        override,
        {"機械1": base + timedelta(hours=1)},
        protected,
        "機械1",
        base,
        base + timedelta(minutes=30),
    )

    # 限定→未指定: 限定由来の保護区間は全制約撤廃中の未指定行も無視できない。
    pc._register_limited_equipment_interval(
        protected,
        limited,
        "機械1",
        base,
        base + timedelta(hours=1),
    )
    assert pc._limited_equipment_earliest_start(
        protected,
        "機械1",
        base + timedelta(minutes=30),
        timedelta(minutes=30),
    ) == base + timedelta(hours=1)
    assert pc._limited_equipment_interval_blocked(
        ordinary,
        override,
        {},
        protected,
        "機械1",
        base + timedelta(minutes=30),
        base + timedelta(hours=2),
    )
    assert (
        pc._new_dispatch_limited_equipment_interval_blocked
        is pc._legacy_dispatch_limited_equipment_interval_blocked
        is pc._limited_equipment_interval_blocked
    )


def test_limited_row_never_ignores_real_skills_in_either_path():
    override = {"ignore_skill_requirements": True}
    limited = _task(limited_operator_names=("甲", "乙"))
    ordinary = _task()

    assert pc._skill_requirements_ignored_for_task(override, limited) is False
    assert pc._skill_requirements_ignored_for_task(override, ordinary) is True
    assert (
        pc._new_dispatch_skill_requirements_ignored
        is pc._legacy_dispatch_skill_requirements_ignored
        is pc._skill_requirements_ignored_for_task
    )


def test_legacy_candidate_computes_actual_end_before_protection_check():
    start = datetime(2026, 7, 15, 9, 0)
    end_limit = datetime(2026, 7, 15, 17, 0)
    task = _task(limited_operator_names=("甲",))

    actual_end, blocked = pc._legacy_candidate_end_and_protection(
        task,
        {},
        {},
        pc._LimitedEquipmentProtection(),
        "機械1",
        start,
        30,
        [],
        end_limit,
    )

    assert actual_end == datetime(2026, 7, 15, 9, 30)
    assert blocked is False


def test_legacy_path_reaches_successful_candidate_with_limited_equipment_protection():
    day = date(2026, 7, 15)
    start = datetime(2026, 7, 15, 9, 0)
    candidates = []
    task = _task(
        limited_operator_names=("甲",),
        start_date_req=day,
        base_time_per_unit=30.0,
        task_eff_factor=1.0,
        remaining_units=1.0,
    )

    added = pc._append_legacy_dispatch_candidate_for_team(
        task,
        ("甲",),
        {"甲": start},
        {},
        {
            "甲": {
                "end_dt": datetime(2026, 7, 15, 17, 0),
                "breaks_dt": [],
                "efficiency": 1.0,
            }
        },
        day,
        day,
        start,
        _role_map({"甲": ("OP", 1)}),
        "",
        1,
        0,
        {"abolish_all_scheduling_limits": True},
        candidates,
        limited_equipment_mirror=pc._LimitedEquipmentProtection(),
        machine_floor_cached=start,
    )

    assert added is True
    assert len(candidates) == 1
    assert candidates[0]["actual_end_dt"] == start + timedelta(minutes=30)


def test_recalculate_capacity_after_protected_start_keeps_small_assignment():
    start = datetime(2026, 7, 15, 16, 0)

    capacity = pc._candidate_capacity_at_start(
        start,
        30.0,
        10.0,
        [],
        datetime(2026, 7, 15, 17, 0),
    )

    assert capacity == {
        "units_can_do": 2,
        "units_today": 2,
        "work_mins_needed": 60,
    }


@pytest.mark.parametrize(
    ("protected_start", "expected_start"),
    [
        (datetime(2026, 7, 15, 10, 0), datetime(2026, 7, 15, 9, 0)),
        (
            datetime(2026, 7, 15, 9, 59, 59),
            datetime(2026, 7, 15, 10, 30),
        ),
    ],
)
def test_protected_interval_boundary_keeps_exact_roll_or_defers_one_second_short(
    protected_start, expected_start
):
    start = datetime(2026, 7, 15, 9, 0)
    protected = pc._LimitedEquipmentProtection()
    protected.register(
        "機械1",
        protected_start,
        datetime(2026, 7, 15, 10, 30),
    )

    result = pc._candidate_capacity_after_equipment_protection(
        protected,
        "機械1",
        start,
        60.0,
        1.0,
        [],
        datetime(2026, 7, 15, 12, 0),
    )

    assert result is not None
    actual_start, capacity = result
    assert actual_start == expected_start
    assert capacity["units_today"] == 1
    assert capacity["work_mins_needed"] == 60


def test_future_protected_interval_trims_multiple_rolls_before_it():
    start = datetime(2026, 7, 15, 9, 0)
    protected = pc._LimitedEquipmentProtection()
    protected.register(
        "機械1",
        datetime(2026, 7, 15, 11, 0),
        datetime(2026, 7, 15, 12, 0),
    )

    result = pc._candidate_capacity_after_equipment_protection(
        protected,
        "機械1",
        start,
        30.0,
        10.0,
        [],
        datetime(2026, 7, 15, 17, 0),
    )

    assert result is not None
    actual_start, capacity = result
    assert actual_start == start
    assert capacity["units_today"] == 4
    assert capacity["work_mins_needed"] == 120


def test_future_protected_interval_respects_break_contiguous_capacity():
    start = datetime(2026, 7, 15, 9, 0)
    protected = pc._LimitedEquipmentProtection()
    protected.register(
        "機械1",
        datetime(2026, 7, 15, 11, 0),
        datetime(2026, 7, 15, 12, 0),
    )

    result = pc._candidate_capacity_after_equipment_protection(
        protected,
        "機械1",
        start,
        30.0,
        10.0,
        [
            (
                datetime(2026, 7, 15, 10, 30),
                datetime(2026, 7, 15, 10, 45),
            )
        ],
        datetime(2026, 7, 15, 17, 0),
    )

    assert result is not None
    actual_start, capacity = result
    assert actual_start == start
    assert capacity["units_today"] == 3
    assert capacity["work_mins_needed"] == 90


def test_legacy_roll_candidate_recalculates_capacity_after_protected_delay():
    day = date(2026, 7, 15)
    start = datetime(2026, 7, 15, 9, 0)
    candidates = []
    protected = pc._LimitedEquipmentProtection()
    protected.register(
        "機械1",
        start,
        datetime(2026, 7, 15, 16, 0),
    )
    task = _task(
        limited_operator_names=("甲",),
        start_date_req=day,
        base_time_per_unit=30.0,
        task_eff_factor=1.0,
        remaining_units=10.0,
    )

    added = pc._append_legacy_dispatch_candidate_for_team(
        task,
        ("甲",),
        {"甲": start},
        {},
        {
            "甲": {
                "end_dt": datetime(2026, 7, 15, 17, 0),
                "breaks_dt": [],
                "efficiency": 1.0,
            }
        },
        day,
        day,
        start,
        _role_map({"甲": ("OP", 1)}),
        "",
        1,
        0,
        {"abolish_all_scheduling_limits": True},
        candidates,
        limited_equipment_mirror=protected,
        machine_floor_cached=start,
    )

    assert added is True
    assert candidates[0]["team_start"] == datetime(2026, 7, 15, 16, 0)
    assert candidates[0]["units_today"] == 2
    assert candidates[0]["actual_end_dt"] == datetime(2026, 7, 15, 17, 0)


def test_limited_equipment_intervals_defer_only_when_candidate_overlaps():
    base = datetime(2026, 7, 15, 9, 0)
    protected = pc._LimitedEquipmentProtection()
    protected.register("機械1", base + timedelta(hours=1), base + timedelta(hours=2))

    assert pc._limited_equipment_earliest_start(
        protected, "機械1", base, timedelta(minutes=30)
    ) == base
    assert pc._limited_equipment_earliest_start(
        protected,
        "機械1",
        base + timedelta(minutes=45),
        timedelta(minutes=30),
    ) == base + timedelta(hours=2)
    assert pc._limited_equipment_earliest_start(
        protected,
        "機械1",
        base + timedelta(hours=2, minutes=15),
        timedelta(minutes=30),
    ) == base + timedelta(hours=2, minutes=15)


def test_limited_equipment_intervals_handle_consecutive_ranges_in_start_order():
    base = datetime(2026, 7, 15, 9, 0)
    protected = pc._LimitedEquipmentProtection()
    # 逆順登録でも開始順に管理される。
    protected.register("機械1", base + timedelta(hours=2), base + timedelta(hours=3))
    protected.register("機械1", base + timedelta(hours=1), base + timedelta(hours=2))

    assert pc._limited_equipment_earliest_start(
        protected,
        "機械1",
        base + timedelta(hours=1, minutes=30),
        timedelta(minutes=45),
    ) == base + timedelta(hours=3)


def test_limited_equipment_interval_scan_breaks_before_far_future_ranges():
    base = datetime(2026, 7, 15, 9, 0)
    protected = pc._LimitedEquipmentProtection()
    for offset in reversed(range(100)):
        start = base + timedelta(days=1, hours=offset)
        protected.register("機械1", start, start + timedelta(minutes=30))

    assert pc._limited_equipment_earliest_start(
        protected, "機械1", base, timedelta(minutes=15)
    ) == base
    assert protected.last_scan_count <= 1


def test_limited_equipment_intervals_merge_overlaps_and_touching_ranges():
    base = datetime(2026, 7, 15, 9, 0)
    protected = pc._LimitedEquipmentProtection()
    protected.register(
        "機械1", base + timedelta(hours=2), base + timedelta(hours=3)
    )
    protected.register(
        "機械1", base + timedelta(hours=1), base + timedelta(hours=2)
    )
    protected.register(
        "機械1",
        base + timedelta(hours=1, minutes=30),
        base + timedelta(hours=4),
    )

    assert protected.interval_count("機械1") == 1
    assert pc._limited_equipment_earliest_start(
        protected,
        "機械1",
        base + timedelta(hours=1, minutes=45),
        timedelta(minutes=30),
    ) == base + timedelta(hours=4)


def test_limited_equipment_search_skips_many_past_ranges_with_bisect():
    base = datetime(2026, 1, 1, 0, 0)
    protected = pc._LimitedEquipmentProtection()
    for offset in range(5000):
        start = base + timedelta(minutes=offset * 20)
        protected.register("機械1", start, start + timedelta(minutes=5))
    candidate = base + timedelta(minutes=5000 * 20 + 10)

    assert pc._limited_equipment_earliest_start(
        protected, "機械1", candidate, timedelta(minutes=10)
    ) == candidate
    assert protected.last_scan_count <= 2
    assert protected.would_block_equipment(
        "機械1", candidate, candidate + timedelta(minutes=10)
    ) is False
    assert protected.last_scan_count <= 2
