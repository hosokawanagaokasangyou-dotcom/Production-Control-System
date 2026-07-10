# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import date

import pandas as pd

from planning_core.core.roll_pipeline import (
    _resolve_in_progress_aladdin_today_shortfall_m,
    append_in_progress_next_day_dialog_rows_to_dispatch_table,
)


def test_resolve_shortfall_from_remaining_minus_next_day_fallback():
    plan_row = pd.Series(
        {
            "依頼NO": "C7-4",
            "工程名": "SEC",
            "機械名": "SEC機　湖南",
            "換算数量": 10000,
            "実加工数": 4400,
            "未加工": 5600,
            "配台使用残数量": 5600,
        }
    )
    key = "C7-4\x1eSEC\x1eSEC機　湖南"
    sf = _resolve_in_progress_aladdin_today_shortfall_m(key, 2000.0, plan_row, {})
    assert sf == 3600.0


def test_append_in_progress_adds_shortfall_on_calendar_today(monkeypatch, tmp_path):
    json_path = tmp_path / "next_day.json"
    json_path.write_text(
        """
{
  "version": 1,
  "entries": [
    {
      "task_id": "C7-4",
      "process": "SEC",
      "machine_name": "SEC機　湖南",
      "next_day_dispatch_m": 2000.0,
      "aladdin_today_shortfall_m": 3600.0
    }
  ]
}
""".strip(),
        encoding="utf-8",
    )
    monkeypatch.setenv("PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON", str(json_path))

    tasks_df = pd.DataFrame(
        [
            {
                "依頼NO": "C7-4",
                "工程名": "SEC",
                "機械名": "SEC機　湖南",
                "換算数量": 10000,
                "実加工数": 4400,
                "未加工": 5600,
                "配台使用残数量": 5600,
            }
        ]
    )
    out = append_in_progress_next_day_dialog_rows_to_dispatch_table(
        pd.DataFrame(),
        tasks_df,
        None,
        run_date=date(2026, 7, 10),
        working_days=[date(2026, 7, 10), date(2026, 7, 13)],
        calendar_today=date(2026, 7, 10),
    )
    assert len(out) == 2
    by_day = {
        pd.to_datetime(r["配台日"]).date(): float(r["当日配台数量"])
        for r in out.to_dict(orient="records")
    }
    assert by_day[date(2026, 7, 10)] == 3600.0
    assert by_day[date(2026, 7, 13)] == 2000.0
