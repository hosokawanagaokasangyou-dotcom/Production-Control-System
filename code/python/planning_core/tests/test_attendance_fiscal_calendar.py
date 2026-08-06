# -*- coding: utf-8 -*-
"""Fiscal year company calendar payload."""

from __future__ import annotations

from planning_core.core.attendance_store import (
    build_company_calendar_payload_fiscal,
    empty_store,
    fiscal_year_date_range,
)


def test_fiscal_year_date_range_april_march():
    start, end = fiscal_year_date_range(2024, 4, 1)
    assert start.isoformat() == "2024-04-01"
    assert end.isoformat() == "2025-03-31"


def test_build_company_calendar_payload_fiscal_filters():
    store = empty_store(2024)
    store["company_calendar"]["days"]["2024-03-31"] = {"kind": "public_holiday"}
    store["company_calendar"]["days"]["2024-04-01"] = {"kind": "working_day"}
    store["company_calendar"]["days"]["2025-03-31"] = {"kind": "public_holiday"}
    store["company_calendar"]["days"]["2025-04-01"] = {"kind": "public_holiday"}
    payload = build_company_calendar_payload_fiscal(store, 2024, 4, 1)
    keys = set(payload["days"].keys())
    assert "2024-03-31" not in keys
    assert "2024-04-01" in keys
    assert "2025-03-31" in keys
    assert "2025-04-01" not in keys


def test_enrich_national_holidays_on_weekdays():
    from planning_core.core.attendance_store import enrich_company_calendar_days_with_national_holidays

    start, end = fiscal_year_date_range(2026, 4, 1)
    enriched = enrich_company_calendar_days_with_national_holidays({}, start, end)
  # 昭和の日
    assert "2026-04-29" in enriched
    assert enriched["2026-04-29"]["source"] == "national_holiday"
    assert "昭和" in enriched["2026-04-29"]["label"]
