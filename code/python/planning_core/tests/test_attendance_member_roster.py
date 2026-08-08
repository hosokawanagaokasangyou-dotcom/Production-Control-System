# -*- coding: utf-8 -*-
"""Tests for attendance_member_roster."""

from __future__ import annotations

from planning_core.core.attendance_member_roster import (
    DEFAULT_MEMBER_ROSTER,
    KOKUBU_DEFAULT_MEMBER_ROSTER,
    apply_member_roster_patch,
    attendance_grid_member_names,
    default_member_roster_for_factory,
    ensure_member_roster,
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
