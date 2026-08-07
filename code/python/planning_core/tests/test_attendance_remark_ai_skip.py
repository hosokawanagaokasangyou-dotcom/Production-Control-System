# -*- coding: utf-8 -*-
"""勤怠備考 AI: プリセット行は Gemini に載せない。"""

from __future__ import annotations

from planning_core.core.master_data import (
    _attendance_skip_remark_ai,
)


def test_skip_preset_public_holiday_remark():
    assert _attendance_skip_remark_ai("公休", "公休") is True


def test_skip_preset_paid_leave():
    assert _attendance_skip_remark_ai("年休", "年休") is True


def test_skip_no_dispatch_marker():
    assert _attendance_skip_remark_ai("-", "-") is True


def test_keep_custom_remark():
    assert _attendance_skip_remark_ai("午後は会議", "通常") is False


def test_keep_empty_leave_with_custom_context():
    assert _attendance_skip_remark_ai("", "配台不可") is False
