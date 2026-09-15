# -*- coding: utf-8 -*-
"""Tests for attendance_member_roster."""

from __future__ import annotations

from datetime import date

from planning_core.core.attendance_member_roster import (
    DEFAULT_MEMBER_ROSTER,
    KOKUBU_DEFAULT_MEMBER_ROSTER,
    apply_member_roster_patch,
    attendance_grid_member_names,
    default_member_roster_for_factory,
    ensure_member_roster,
    member_active_on,
    member_visible_in_month,
    members_for_attendance_analysis,
)
from planning_core.core.attendance_store import apply_member_attendance_patch, empty_store


def test_default_roster_has_fourteen_members_in_order(monkeypatch):
    monkeypatch.delenv("PM_AI_FACTORY_SITE", raising=False)
    store = empty_store(2026)
    names = attendance_grid_member_names(store)
    assert len(names) == 14
    assert names[0] == DEFAULT_MEMBER_ROSTER[0]["name"]
    assert names[-1] == DEFAULT_MEMBER_ROSTER[-1]["name"]


def test_kokubu_default_roster_has_twenty_one_members(monkeypatch):
    monkeypatch.setenv("PM_AI_FACTORY_SITE", "KOKUBU")
    store = empty_store(2026)
    names = attendance_grid_member_names(store)
    assert len(names) == 21
    assert names[0] == KOKUBU_DEFAULT_MEMBER_ROSTER[0]["name"]
    assert names[-1] == KOKUBU_DEFAULT_MEMBER_ROSTER[-1]["name"]
    assert all(
        e["primary_role"] == "後加工" for e in default_member_roster_for_factory()
    )


def test_members_for_attendance_analysis_merges_skills():
    store = empty_store(2026)
    merged = members_for_attendance_analysis(["技能者A", "細川　守"], store)
    assert merged[0] == "細川　守"
    assert "技能者A" in merged
    assert merged.index("技能者A") > merged.index("東出　繫利")


def test_apply_member_roster_patch_renames_attendance_keys():
    store = empty_store(2026)
    ensure_member_roster(store)
    old = attendance_grid_member_names(store)[0]
    new = "テスト　太郎"
    roster = [{"name": new, "primary_role": "物流"}]
    for ent in ensure_member_roster(store)[1:]:
        roster.append(dict(ent))
    store["member_attendance"]["2026-08-01"] = {old: {"day_preset": "WORK", "manual_edit": True}}
    apply_member_roster_patch(store, roster)
    assert new in store["member_attendance"]["2026-08-01"]
    assert old not in store["member_attendance"]["2026-08-01"]


def test_merge_patch_includes_roster():
    store = empty_store(2026)
    ensure_member_roster(store)
    roster = [{"name": "新規　メンバー", "primary_role": "後加工"}]
    result = apply_member_attendance_patch(
        store,
        {
            "year": 2026,
            "month": 8,
            "cells": {},
            "member_roster": roster,
        },
    )
    assert result["roster_count"] == 1
    assert attendance_grid_member_names(store) == ["新規　メンバー"]


def test_ensure_member_roster_keeps_inactive_from():
    store = empty_store(2026)
    store["member_roster"] = [
        {"name": "菅沼　めぐみ", "primary_role": "後加工", "inactive_from": "2026-09-15"},
        {"name": "細川　守", "primary_role": "後加工", "inactive_from": ""},
    ]
    roster = ensure_member_roster(store)
    by_name = {e["name"]: e for e in roster}
    assert by_name["菅沼　めぐみ"]["inactive_from"] == "2026-09-15"
    assert "inactive_from" not in by_name["細川　守"]


def test_member_active_on_is_false_from_transfer_date():
    entry = {"name": "A", "primary_role": "後加工", "inactive_from": "2026-09-15"}
    assert member_active_on(entry, date(2026, 9, 14))
    assert not member_active_on(entry, date(2026, 9, 15))
    assert member_active_on({"name": "B", "primary_role": "後加工"}, date(2026, 9, 15))


def test_member_visible_in_month_uses_first_of_month():
    mid = {"name": "A", "primary_role": "後加工", "inactive_from": "2026-09-15"}
    first = {"name": "B", "primary_role": "後加工", "inactive_from": "2026-09-01"}
    assert member_visible_in_month(mid, 2026, 8)
    assert member_visible_in_month(mid, 2026, 9)
    assert not member_visible_in_month(mid, 2026, 10)
    assert member_visible_in_month(first, 2026, 8)
    assert not member_visible_in_month(first, 2026, 9)


def test_attendance_grid_member_names_filters_by_month():
    store = empty_store(2026)
    store["member_roster"] = [
        {"name": "在籍", "primary_role": "後加工"},
        {"name": "異動", "primary_role": "物流", "inactive_from": "2026-09-15"},
    ]
    assert attendance_grid_member_names(store, 2026, 9) == ["在籍", "異動"]
    assert attendance_grid_member_names(store, 2026, 10) == ["在籍"]
    assert attendance_grid_member_names(store) == ["在籍", "異動"]


def test_apply_member_roster_patch_keeps_attendance_when_inactive_from_set():
    store = empty_store(2026)
    store["member_roster"] = [{"name": "菅沼　めぐみ", "primary_role": "後加工"}]
    store["member_attendance"]["2026-08-01"] = {
        "菅沼　めぐみ": {"day_preset": "WORK", "manual_edit": True}
    }
    apply_member_roster_patch(
        store,
        [
            {
                "name": "菅沼　めぐみ",
                "primary_role": "後加工",
                "inactive_from": "2026-09-15",
            }
        ],
    )
    assert store["member_attendance"]["2026-08-01"]["菅沼　めぐみ"]["day_preset"] == "WORK"
    assert ensure_member_roster(store)[0]["inactive_from"] == "2026-09-15"


def test_apply_member_roster_patch_clears_inactive_from_on_cancel():
    store = empty_store(2026)
    store["member_roster"] = [
        {"name": "菅沼　めぐみ", "primary_role": "後加工", "inactive_from": "2026-09-15"}
    ]
    apply_member_roster_patch(
        store, [{"name": "菅沼　めぐみ", "primary_role": "後加工"}]
    )
    assert "inactive_from" not in ensure_member_roster(store)[0]


def test_ensure_member_roster_keeps_returned_on():
    store = empty_store(2026)
    store["member_roster"] = [
        {
            "name": "菅沼　めぐみ",
            "primary_role": "後加工",
            "inactive_from": "2026-09-15",
            "returned_on": "2026-11-01",
        }
    ]
    roster = ensure_member_roster(store)
    assert roster[0]["inactive_from"] == "2026-09-15"
    assert roster[0]["returned_on"] == "2026-11-01"


def test_member_active_on_after_returned_on():
    entry = {
        "name": "A",
        "primary_role": "後加工",
        "inactive_from": "2026-09-15",
        "returned_on": "2026-11-01",
    }
    assert member_active_on(entry, date(2026, 9, 14))
    assert not member_active_on(entry, date(2026, 9, 15))
    assert not member_active_on(entry, date(2026, 10, 31))
    assert member_active_on(entry, date(2026, 11, 1))


def test_member_visible_in_month_after_return():
    entry = {
        "name": "A",
        "primary_role": "後加工",
        "inactive_from": "2026-09-15",
        "returned_on": "2026-11-01",
    }
    assert member_visible_in_month(entry, 2026, 9)
    assert not member_visible_in_month(entry, 2026, 10)
    assert member_visible_in_month(entry, 2026, 11)


def test_attendance_grid_member_names_includes_returned_month():
    store = empty_store(2026)
    store["member_roster"] = [
        {
            "name": "異動",
            "primary_role": "物流",
            "inactive_from": "2026-09-15",
            "returned_on": "2026-11-01",
        }
    ]
    assert attendance_grid_member_names(store, 2026, 10) == []
    assert attendance_grid_member_names(store, 2026, 11) == ["異動"]
