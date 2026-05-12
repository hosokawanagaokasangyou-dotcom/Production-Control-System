"""equipment_gantt_json_enrich: 結果_タスク一覧から設備ガント JSON セルへ担当をジョインする。"""

from __future__ import annotations

import pytest

pytest.importorskip("pandas")


def test_enrich_appends_hint_after_bracket_task_id():
    from planning_core.equipment_gantt_json_enrich import (
        enrich_workbook_payload_equipment_gantt_members,
    )

    payload = {
        "format_version": 2,
        "source_xlsx": "t.xlsx",
        "sheets": {
            "結果_タスク一覧": {
                "columns": [
                    "タスクID",
                    "履歴1",
                ],
                "row_count": 1,
                "rows": [
                    {
                        "タスクID": "A-1",
                        "履歴1": "・【2026/05/07】：1R/10m 主:山田,佐藤",
                    }
                ],
            },
            "結果_設備ガント": {
                "columns": [
                    "日付",
                    "dummy",
                    "8:00",
                    "8:10",
                ],
                "row_count": 1,
                "rows": [
                    {
                        "日付": "2026/05/07",
                        "dummy": "設備1",
                        "8:00": "[A-1] 加工",
                        "8:10": "休憩",
                    }
                ],
            },
        },
    }
    out = enrich_workbook_payload_equipment_gantt_members(payload)
    row0 = out["sheets"]["結果_設備ガント"]["rows"][0]
    assert "（" in row0["8:00"] and "）" in row0["8:00"]
    assert "山田" in row0["8:00"] or "佐藤" in row0["8:00"]
    assert row0["8:10"] == "休憩"
    assert out.get("equipment_gantt_member_enrich") is True


def test_enrich_skips_when_already_parenthesized():
    from planning_core.equipment_gantt_json_enrich import (
        enrich_workbook_payload_equipment_gantt_members,
    )

    orig_cell = "[A-1]（既存）"
    payload = {
        "format_version": 2,
        "sheets": {
            "結果_タスク一覧": {
                "columns": ["タスクID", "履歴1"],
                "rows": [{"タスクID": "A-1", "履歴1": "主:山田"}],
            },
            "結果_設備ガント": {
                "columns": ["日付", "8:00"],
                "rows": [{"日付": "d", "8:00": orig_cell}],
            },
        },
    }
    out = enrich_workbook_payload_equipment_gantt_members(payload)
    assert out["sheets"]["結果_設備ガント"]["rows"][0]["8:00"] == orig_cell
