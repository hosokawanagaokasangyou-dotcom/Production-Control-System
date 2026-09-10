# -*- coding: utf-8 -*-
"""master.xlsm speed シートの基本速度/実稼働比率行解決。"""
from __future__ import annotations

import pandas as pd

from planning_core import _core as core


def test_resolve_speed_rows_current_layout_labels_at_excel_3_4():
    """現行: 工程/機械/基本速度/実稼働比率 が 1〜4 行目。"""
    raw = pd.DataFrame(
        [
            ["工程名", None, None, "SEC", "EC"],
            ["機械名", None, None, "SEC機　湖南", "EC機　湖南"],
            ["基本速度", None, None, 27, 20],
            ["実稼働比率", None, None, 1, 1],
            [None, "ロジック式", None, None, None],
        ]
    )
    got = core._resolve_master_speed_base_ratio_row_indices(raw)
    assert got == (2, 3)


def test_resolve_speed_rows_legacy_layout_excel_4_5():
    """旧: 基本速度=4行目・実稼働比率=5行目（ラベル無しでも数値で解決）。"""
    raw = pd.DataFrame(
        [
            ["工程名", None, None, "SEC"],
            ["機械名", None, None, "SEC機　湖南"],
            [None, None, None, None],
            [None, None, None, 27],
            [None, None, None, 1],
        ]
    )
    got = core._resolve_master_speed_base_ratio_row_indices(raw)
    assert got == (3, 4)


def test_resolve_ignores_annotation_labels_without_metrics():
    """「基本速度メモ」等は完全一致でないので無視し、数値のある (2,3) へフォールバック。"""
    raw = pd.DataFrame(
        [
            ["工程名", None, None, "SEC"],
            ["機械名", None, None, "SEC機　湖南"],
            ["基本速度", None, None, 27],
            ["実稼働比率", None, None, 1],
            ["基本速度メモ", None, None, None],
            ["実稼働時間の比率メモ", None, None, None],
        ]
    )
    got = core._resolve_master_speed_base_ratio_row_indices(raw)
    assert got == (2, 3)


def test_resolve_label_rows_without_positive_metrics_falls_back():
    """ラベル行の数値が空で、旧レイアウト (3,4) に数値がある場合はフォールバック。"""
    raw = pd.DataFrame(
        [
            ["工程名", None, None, "SEC"],
            ["機械名", None, None, "SEC機　湖南"],
            ["基本速度", None, None, None],
            [None, None, None, 27],
            [None, None, None, 1],
        ]
    )
    got = core._resolve_master_speed_base_ratio_row_indices(raw)
    assert got == (3, 4)


def test_load_lookup_from_current_layout_dataframe(monkeypatch):
    """現行レイアウト DataFrame から SEC 速度 27 が取れること。"""
    raw = pd.DataFrame(
        [
            ["工程名", None, None, "SEC"],
            ["機械名", None, None, "SEC機　湖南"],
            ["基本速度", None, None, 27],
            ["実稼働比率", None, None, 1],
        ]
    )

    def _fake_read_excel(*_a, **_k):
        return raw

    monkeypatch.setattr(core, "_master_speed_sheet_apply_enabled", lambda: True)
    monkeypatch.setattr(
        core, "_require_master_workbook_path_exists", lambda: "dummy.xlsm"
    )
    monkeypatch.setattr(core.pd, "read_excel", _fake_read_excel)

    lu = core._load_master_speed_lookup_from_master_workbook()
    key = (
        core._normalize_process_name_for_rule_match("SEC"),
        core._normalize_equipment_match_key("SEC機　湖南"),
    )
    assert lu.get(key) == 27.0


def test_load_lookup_legacy_layout_and_fullwidth_digits(monkeypatch):
    """旧レイアウト + 全角数字でも 27.0 になること。"""
    raw = pd.DataFrame(
        [
            ["工程名", None, None, "SEC"],
            ["機械名", None, None, "SEC機　湖南"],
            [None, None, None, None],
            [None, None, None, "２７"],
            [None, None, None, "１"],
        ]
    )

    def _fake_read_excel(*_a, **_k):
        return raw

    monkeypatch.setattr(core, "_master_speed_sheet_apply_enabled", lambda: True)
    monkeypatch.setattr(
        core, "_require_master_workbook_path_exists", lambda: "dummy.xlsm"
    )
    monkeypatch.setattr(core.pd, "read_excel", _fake_read_excel)

    lu = core._load_master_speed_lookup_from_master_workbook()
    key = (
        core._normalize_process_name_for_rule_match("SEC"),
        core._normalize_equipment_match_key("SEC機　湖南"),
    )
    assert lu.get(key) == 27.0


def test_normalize_speed_label_strips_ideographic_space():
    assert core._normalize_speed_sheet_row_label("基本　速度") == "基本速度"
    assert core._normalize_speed_sheet_row_label("実稼働 比率") == "実稼働比率"
