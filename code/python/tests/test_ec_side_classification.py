# -*- coding: utf-8 -*-
"""EC面区分（両面EC/片面EC）判定のテスト。"""

from __future__ import annotations

from planning_core.core.ec_side_classification import (
    EC_SIDE_CLASS_DOUBLE,
    EC_SIDE_CLASS_SINGLE,
    EC_SIDE_CLASS_UNKNOWN,
    classify_ec_side,
    process_content_has_ec,
)


def test_process_content_has_ec_token():
    assert process_content_has_ec("EC,ゴミ除去")
    assert process_content_has_ec("EC（片面）")
    assert process_content_has_ec("Ec")
    assert not process_content_has_ec("スリット,SEC")
    assert not process_content_has_ec("")


def test_w8_16_juchu_ec_lowercase_process_content():
    """W8-16: 受注 加工内容が 'Ec'（小文字 c）でも EC 工程として判定する。"""
    assert process_content_has_ec("Ec")
    assert classify_ec_side("Ec", "両面") == EC_SIDE_CLASS_DOUBLE


def test_classify_ec_side_double_when_ec_men_blank_and_original_ref():
    assert (
        classify_ec_side("EC,ゴミ除去", "", original_ref_found=True)
        == EC_SIDE_CLASS_DOUBLE
    )
    assert (
        classify_ec_side("EC,ゴミ除去", None, original_ref_found=True)
        == EC_SIDE_CLASS_DOUBLE
    )
    assert (
        classify_ec_side("EC,ゴミ除去", "-", original_ref_found=True)
        == EC_SIDE_CLASS_DOUBLE
    )


def test_classify_ec_side_unknown_when_ec_men_blank_without_original():
    assert (
        classify_ec_side("EC,ゴミ除去", "", original_ref_found=False)
        == EC_SIDE_CLASS_UNKNOWN
    )
    assert (
        classify_ec_side("EC,ゴミ除去", "-", juchu_row_found=True, original_ref_found=False)
        == EC_SIDE_CLASS_UNKNOWN
    )


def test_classify_ec_side_single_when_ec_men_present():
    assert classify_ec_side("EC,ゴミ除去", "H") == EC_SIDE_CLASS_SINGLE
    assert classify_ec_side("EC", "Ｈ面") == EC_SIDE_CLASS_SINGLE


def test_classify_ec_side_double_when_juchu_ec_men_ryomen():
    assert classify_ec_side("EC", "両面") == EC_SIDE_CLASS_DOUBLE
    assert classify_ec_side("EC,スリット", "両面") == EC_SIDE_CLASS_DOUBLE


def test_w8_18_juchu_ryomen_is_double_sided():
    from planning_core.core.request_form_juchu_lookup import (
        lookup_juchu_ec_row,
        resolve_ec_men_for_side_classification,
    )

    juchu = {"W8-18": {"加工内容": "EC", "EC面": "両面"}}
    original = {"W8-18": {"加工内容": "EC", "EC面": ""}}
    row = lookup_juchu_ec_row(juchu, "W8-18", original)
    ec_men = resolve_ec_men_for_side_classification("W8-18", row, original)
    assert ec_men == "両面"
    assert (
        classify_ec_side(
            "EC",
            ec_men,
            juchu_row_found=True,
            original_ref_found=True,
        )
        == EC_SIDE_CLASS_DOUBLE
    )


def test_classify_ec_side_empty_without_ec_process():
    assert classify_ec_side("スリット,SEC", "") == ""
    assert classify_ec_side("スリット,SEC", "両面") == ""


def test_classify_ec_side_unknown_when_juchu_missing():
    assert (
        classify_ec_side("EC,ゴミ除去", "", juchu_row_found=False) == EC_SIDE_CLASS_UNKNOWN
    )
    assert (
        classify_ec_side("EC,ゴミ除去", "H", juchu_row_found=False) == EC_SIDE_CLASS_UNKNOWN
    )


def test_parent_irai_no_lookup_key():
    from planning_core.core.request_form_juchu_lookup import (
        lookup_juchu_ec_row,
        parent_irai_no_lookup_key,
    )

    assert parent_irai_no_lookup_key("W7-22-1") == "W7-22"
    assert parent_irai_no_lookup_key("CS-3-1") == "CS-3"
    assert parent_irai_no_lookup_key("W7-22") == ""
    lookup = {
        "W7-22": {"加工内容": "熟成,EC", "EC面": "H"},
    }
    row = lookup_juchu_ec_row(lookup, "W7-22-1", {"W7-22": {"加工内容": "熟成,EC", "EC面": "H"}})
    assert row.get("EC面") == "H"
    assert (
        classify_ec_side(
            row.get("加工内容"),
            row.get("EC面"),
            original_ref_found=True,
        )
        == "片面EC"
    )


def test_pn08_05_unknown_without_original_when_juchu_ec_men_blank():
    from planning_core.core.request_form_juchu_lookup import lookup_juchu_ec_row

    juchu = {
        "PN08-05": {"加工内容": "EC", "EC面": ""},
    }
    row = lookup_juchu_ec_row(juchu, "PN08-05", {})
    assert classify_ec_side(
        row.get("加工内容") or "EC",
        row.get("EC面", ""),
        juchu_row_found=bool(row),
        original_ref_found=False,
    ) == EC_SIDE_CLASS_UNKNOWN


def test_ec_dispatch_pass_count():
    from planning_core.core.ec_side_classification import (
        EC_DISPATCH_PASS_COUNT_DOUBLE,
        EC_SIDE_CLASS_DOUBLE,
        EC_SIDE_CLASS_SINGLE,
        EC_SIDE_CLASS_UNKNOWN,
        ec_dispatch_pass_count,
    )

    assert ec_dispatch_pass_count(EC_SIDE_CLASS_DOUBLE, "EC") == EC_DISPATCH_PASS_COUNT_DOUBLE
    assert ec_dispatch_pass_count(EC_SIDE_CLASS_DOUBLE, "SEC") == 1
    assert ec_dispatch_pass_count(EC_SIDE_CLASS_DOUBLE, "SEC機　湖南") == 1
    assert ec_dispatch_pass_count(EC_SIDE_CLASS_SINGLE, "EC") == 1
    assert ec_dispatch_pass_count(EC_SIDE_CLASS_UNKNOWN, "EC") == 1
    assert ec_dispatch_pass_count("", "EC") == 1
    assert ec_dispatch_pass_count(EC_SIDE_CLASS_DOUBLE) == 1


def test_lookup_juchu_ec_row_enriches_ec_men_from_original_when_juchu_blank():
    from planning_core.core.request_form_juchu_lookup import lookup_juchu_ec_row

    juchu = {
        "W7-22": {"加工内容": "EC,スリット", "EC面": ""},
    }
    original = {
        "W7-22": {"加工内容": "検反,EC", "EC面": "H"},
    }
    row = lookup_juchu_ec_row(juchu, "W7-22-1", original)
    assert row.get("EC面") == "H"
    assert (
        classify_ec_side(
            row.get("加工内容"),
            row.get("EC面"),
            original_ref_found=True,
        )
        == EC_SIDE_CLASS_SINGLE
    )


def test_w8_18_original_blank_ec_men_overrides_juchu_h():
    """受注 N 列がＨ面でも、原本 AJ 列が空なら両面EC（H面外巻き≠片面EC）。"""
    from planning_core.core.request_form_juchu_lookup import (
        lookup_juchu_ec_row,
        resolve_ec_men_for_side_classification,
    )

    juchu = {"W8-18": {"加工内容": "EC", "EC面": "Ｈ面"}}
    original = {"W8-18": {"加工内容": "EC", "EC面": ""}}
    row = lookup_juchu_ec_row(juchu, "W8-18", original)
    ec_men = resolve_ec_men_for_side_classification("W8-18", row, original)
    assert ec_men == ""
    assert (
        classify_ec_side(
            "EC",
            ec_men,
            juchu_row_found=True,
            original_ref_found=True,
        )
        == EC_SIDE_CLASS_DOUBLE
    )


def test_plan_input_column_order_includes_ec_side_class():
    from planning_core import _core as pc

    order = pc.plan_input_sheet_column_order()
    idx_kako = order.index(pc.TASK_COL_PROCESS_CONTENT)
    idx_ec = order.index(pc.PLAN_COL_EC_SIDE_CLASS)
    assert idx_ec == idx_kako + 1
