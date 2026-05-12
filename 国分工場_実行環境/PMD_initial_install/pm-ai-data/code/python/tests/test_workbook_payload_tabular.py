"""表シート JSON 正本 ↔ DataFrame のラウンドトリップ（planning_core が import できる環境のみ）。"""

from __future__ import annotations

import pytest


pytest.importorskip("pandas")


def test_build_payload_roundtrip_matches_shape():
    try:
        import pandas as pd

        from planning_core.workbook_payload import (
            build_workbook_payload_from_dataframes,
            sheet_payload_to_dataframe,
        )
    except Exception as exc:
        pytest.skip(f"planning_core / pandas 環境: {exc}")

    df = pd.DataFrame(
        {
            "日時帯": ["08:00-08:10", "08:10-08:20"],
            "設備A": [None, "タスク"],
        }
    )
    payload = build_workbook_payload_from_dataframes(
        {"結果_試験": df},
        source_xlsx_basename="dummy.xlsx",
        metadata_extra={"schema": "test"},
    )
    assert payload.get("schema") == "test"
    back = sheet_payload_to_dataframe(payload["sheets"]["結果_試験"])
    assert list(back.columns) == list(df.columns)
    assert len(back) == len(df)
