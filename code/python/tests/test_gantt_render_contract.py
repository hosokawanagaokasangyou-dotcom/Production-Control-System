"""gantt_render_contract のエンコード・契約・JSON 入出力（planning_core が import できる環境のみ）。"""

from __future__ import annotations

import json
import os
import tempfile

import pytest

pytest.importorskip("pandas")


def test_encode_decode_datetime_like_roundtrip():
    from datetime import date, datetime, time

    from planning_core.gantt_render_contract import decode_value, encode_value

    v = {
        "d": date(2026, 5, 4),
        "dt": datetime(2026, 5, 4, 10, 0, 0),
        "pair": (time(8, 0), time(17, 0)),
    }
    dec = decode_value(encode_value(v))
    assert dec["d"] == v["d"]
    assert dec["dt"] == v["dt"]
    assert dec["pair"] == v["pair"]


def test_make_contract_unpack_roundtrip():
    from datetime import date, datetime, time

    from planning_core.gantt_render_contract import (
        GANTT_CONTRACT_SCHEMA_VERSION,
        make_gantt_render_contract,
        unpack_gantt_contract,
    )

    c = make_gantt_render_contract(
        timeline_events=[],
        equipment_list=["M1", "M2"],
        sorted_dates=[date(2026, 5, 4)],
        attendance_data={},
        data_extract_dt_str="2026/05/04 00:00:00",
        base_now_dt=datetime(2026, 5, 4, 8, 0, 0),
        regular_shift_times=(time(8, 0), time(17, 0)),
        plan_rows=True,
        kind="equipment_gantt",
    )
    assert c["schema_version"] == GANTT_CONTRACT_SCHEMA_VERSION
    c2 = json.loads(json.dumps(c, ensure_ascii=False))
    u = unpack_gantt_contract(c2)
    assert u["equipment_list"] == ["M1", "M2"]
    assert len(u["sorted_dates"]) == 1
    assert u["regular_shift_times"] == (time(8, 0), time(17, 0))


def test_write_and_load_gantt_contract_json():
    from datetime import date, datetime, time

    from planning_core.gantt_render_contract import (
        load_gantt_contract_json,
        make_gantt_render_contract,
        write_gantt_contract_json,
    )

    c = make_gantt_render_contract(
        timeline_events=[],
        equipment_list=["A"],
        sorted_dates=[date(2026, 5, 4)],
        attendance_data={},
        data_extract_dt_str=None,
        base_now_dt=datetime(2026, 5, 4, 12, 0, 0),
        regular_shift_times=None,
        plan_rows=True,
        kind="equipment_gantt",
    )
    td = tempfile.gettempdir()
    plan_fake = os.path.join(td, "pm_ai_gantt_contract_unit_test_plan.xlsx")
    out_path, strat = write_gantt_contract_json(plan_fake, "equipment", c)
    assert out_path and os.path.isfile(out_path), (out_path, strat)
    try:
        loaded = load_gantt_contract_json(out_path)
        assert loaded.get("kind") == "equipment_gantt"
        assert loaded.get("kwargs_packed") is not None
    finally:
        try:
            os.remove(out_path)
        except OSError:
            pass
