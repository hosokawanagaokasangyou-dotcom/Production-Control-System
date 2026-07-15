# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import date, datetime

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


def test_append_in_progress_does_not_add_calendar_today_shortfall_when_skip_today(
    monkeypatch, tmp_path
):
    json_path = tmp_path / "next_day-skip-today.json"
    json_path.write_text(
        """
{"version":1,"entries":[{"task_id":"C7-4","process":"SEC","machine_name":"SEC機　湖南",
"next_day_dispatch_m":2000.0,"aladdin_today_shortfall_m":3600.0}]}
""".strip(),
        encoding="utf-8",
    )
    monkeypatch.setenv("PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON", str(json_path))
    monkeypatch.setenv("PM_AI_STAGE2_SKIP_TODAY_DISPATCH", "1")
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
        run_date=date(2026, 7, 13),
        working_days=[date(2026, 7, 13), date(2026, 7, 14)],
        calendar_today=date(2026, 7, 10),
    )

    assert len(out) == 1
    assert pd.to_datetime(out.iloc[0]["配台日"]).date() == date(2026, 7, 13)
    assert float(out.iloc[0]["当日配台数量"]) == 5600.0


def test_skip_today_shortfall_tops_up_existing_target_day_row(monkeypatch, tmp_path):
    json_path = tmp_path / "next-day-existing-row.json"
    json_path.write_text(
        """
{"version":1,"entries":[{"task_id":"C7-4","process":"SEC","machine_name":"SEC機　湖南",
"next_day_dispatch_m":2000.0,"aladdin_today_shortfall_m":3600.0}]}
""".strip(),
        encoding="utf-8",
    )
    monkeypatch.setenv("PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON", str(json_path))
    monkeypatch.setenv("PM_AI_STAGE2_SKIP_TODAY_DISPATCH", "1")
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
    existing = pd.DataFrame(
        [
            {
                "依頼NO": "C7-4",
                "工程名": "SEC",
                "機械名": "SEC機　湖南",
                "配台日": date(2026, 7, 13),
                "当日配台数量": 2000.0,
            }
        ]
    )

    out = append_in_progress_next_day_dialog_rows_to_dispatch_table(
        existing,
        tasks_df,
        None,
        run_date=date(2026, 7, 13),
        working_days=[date(2026, 7, 13), date(2026, 7, 14)],
        calendar_today=date(2026, 7, 10),
    )

    assert len(out) == 1
    assert float(out.iloc[0]["当日配台数量"]) == 5600.0


def test_append_in_progress_skips_next_day_when_table_already_has_remaining(
    monkeypatch, tmp_path
):
    """W7-7 型: タイムライン行が別日に残量ぶん載っているとき翌日追補で二重計上しない。"""
    json_path = tmp_path / "next_day.json"
    json_path.write_text(
        """
{
  "version": 1,
  "entries": [
    {
      "task_id": "W7-7",
      "process": "検査",
      "machine_name": "熱融着機　湖南",
      "next_day_dispatch_m": 1500.0
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
                "依頼NO": "W7-7",
                "工程名": "検査",
                "機械名": "熱融着機　湖南",
                "換算数量": 3600,
                "実加工数": 2100,
                "未加工": 1500,
                "配台使用残数量": 1500,
            }
        ]
    )
    df_dispatch = pd.DataFrame(
        [
            {
                "依頼NO": "W7-7",
                "工程名": "検査",
                "機械名": "熱融着機　湖南",
                "換算数量": 3600,
                "実加工数": 2100,
                "配台日": date(2026, 7, 14),
                "当日配台数量": 1500.0,
                "加工開始日時": "2026/07/14 10:25",
            }
        ]
    )
    out = append_in_progress_next_day_dialog_rows_to_dispatch_table(
        df_dispatch,
        tasks_df,
        None,
        run_date=date(2026, 7, 12),
        working_days=[date(2026, 7, 12), date(2026, 7, 13), date(2026, 7, 14)],
    )
    assert len(out) == 1
    assert float(out.iloc[0]["当日配台数量"]) == 1500.0
    assert pd.to_datetime(out.iloc[0]["配台日"]).date() == date(2026, 7, 14)


def _run_shortfall_case(monkeypatch, tmp_path, *, existing_today_m=0.0, timeline_today_m=0.0):
    json_path = tmp_path / f"shortfall-{existing_today_m}-{timeline_today_m}.json"
    json_path.write_text(
        """
{"version":1,"entries":[{"task_id":"C7-4","process":"SEC","machine_name":"SEC機　湖南",
"next_day_dispatch_m":2000.0,"aladdin_today_shortfall_m":3600.0}]}
""".strip(),
        encoding="utf-8",
    )
    monkeypatch.setenv("PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON", str(json_path))
    tasks_df = pd.DataFrame(
        [{"依頼NO":"C7-4","工程名":"SEC","機械名":"SEC機　湖南","換算数量":10000,
          "実加工数":4400,"未加工":5600,"配台使用残数量":5600}]
    )
    rows = []
    if existing_today_m > 0:
        rows.append({"依頼NO":"C7-4","工程名":"SEC","機械名":"SEC機　湖南",
                     "配台日":date(2026, 7, 10),"当日配台数量":existing_today_m})
    timeline = []
    sorted_tasks = []
    if timeline_today_m > 0:
        timeline = [{"task_id":"C7-4","machine":"line-sec","date":date(2026, 7, 10),
                     "start_dt":datetime(2026, 7, 10, 8, 0),"units_done":timeline_today_m}]
        sorted_tasks = [{"task_id":"C7-4","equipment_line_key":"line-sec",
                         "machine":"SEC","machine_name":"SEC機　湖南"}]
    return append_in_progress_next_day_dialog_rows_to_dispatch_table(
        pd.DataFrame(rows), tasks_df, None, run_date=date(2026, 7, 10),
        working_days=[date(2026, 7, 10), date(2026, 7, 13)],
        calendar_today=date(2026, 7, 10), timeline_events=timeline,
        sorted_tasks_for_result=sorted_tasks,
    )


def _today_total(out):
    return sum(
        float(r["当日配台数量"])
        for r in out.to_dict(orient="records")
        if pd.to_datetime(r["配台日"]).date() == date(2026, 7, 10)
    )


def test_shortfall_adds_only_uncovered_after_existing_today_row(monkeypatch, tmp_path):
    out = _run_shortfall_case(monkeypatch, tmp_path, existing_today_m=1000.0)
    assert _today_total(out) == 3600.0


def test_shortfall_adds_only_uncovered_after_partial_timeline(monkeypatch, tmp_path):
    out = _run_shortfall_case(monkeypatch, tmp_path, timeline_today_m=1000.0)
    assert _today_total(out) == 2600.0


def test_shortfall_does_not_double_count_table_and_timeline_coverage(monkeypatch, tmp_path):
    out = _run_shortfall_case(
        monkeypatch, tmp_path, existing_today_m=1000.0, timeline_today_m=3600.0
    )
    assert _today_total(out) == 1000.0
