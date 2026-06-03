# -*- coding: utf-8 -*-
"""配台試行: 編集 JSON の暦日行がタイムライン meta 上書きで潰れないこと。"""
import os

import pandas as pd

from planning_core import _core as pc


def test_overlay_preserves_editor_dispatch_dates():
    os.environ["PM_AI_INTERACTIVE_DISPATCH_TRIAL"] = "1"
    try:
        cols = [
            "依頼NO",
            "工程名",
            "機械名",
            "換算数量",
            "当日配台数量",
            "配台日",
            "加工開始日時",
            "加工終了日時",
            "メンバー名",
        ]
        json_rows = [
            {
                "依頼NO": "V6-2",
                "工程名": "分割",
                "機械名": "スリット機1　湖南",
                "換算数量": 10000,
                "当日配台数量": 4000,
                "配台日": "2026-06-11",
            },
            {
                "依頼NO": "V6-2",
                "工程名": "分割",
                "機械名": "スリット機1　湖南",
                "換算数量": 10000,
                "当日配台数量": 6000,
                "配台日": "2026-06-12",
            },
        ]
        df_out = pc._dataframe_from_interactive_dispatch_json_rows(
            json_rows, cols, fallback_columns_from=None
        )
        # タイムライン側は加工開始が 06/12 のみ（配台日 06/11 とずれる）
        df_sim = df_out.copy()
        for i in range(len(df_sim)):
            df_sim.at[df_sim.index[i], "加工開始日時"] = "2026/06/12 08:55"
            df_sim.at[df_sim.index[i], "加工終了日時"] = "2026/06/12 17:00"
            df_sim.at[df_sim.index[i], "メンバー名"] = "OP1"
        merged = pc._overlay_timeline_meta_onto_interactive_dispatch_df(df_out, df_sim)
        dates = sorted(
            {
                str(pc._norm_ymd(merged.iloc[i].get("配台日")))
                for i in range(len(merged))
            }
        )
        assert dates == ["2026/06/11", "2026/06/12"]
        assert all(float(merged.iloc[i].get("換算数量") or 0) == 10000.0 for i in range(len(merged)))
    finally:
        os.environ.pop("PM_AI_INTERACTIVE_DISPATCH_TRIAL", None)


def test_use_editor_rows_keeps_two_calendar_rows():
    os.environ["PM_AI_INTERACTIVE_DISPATCH_TRIAL"] = "1"
    try:
        cols = [
            "依頼NO",
            "工程名",
            "機械名",
            "換算数量",
            "当日配台数量",
            "配台日",
            "加工開始日時",
            "加工終了日時",
            pc.INTERACTIVE_DISPATCH_ACTUAL_QTY_COL,
        ]
        json_rows = [
            {
                "依頼NO": "V6-2",
                "工程名": "分割",
                "機械名": "スリット機1　湖南",
                "換算数量": 10000,
                "当日配台数量": 4000,
                "配台日": "2026-06-11",
            },
            {
                "依頼NO": "V6-2",
                "工程名": "分割",
                "機械名": "スリット機1　湖南",
                "換算数量": 10000,
                "当日配台数量": 6000,
                "配台日": "2026-06-12",
            },
        ]
        recs = []
        for r in json_rows:
            rec = {c: "" for c in cols}
            rec.update(r)
            rec["加工開始日時"] = "2026/06/12 08:55"
            rec["加工終了日時"] = "2026/06/12 17:00"
            rec[pc.INTERACTIVE_DISPATCH_ACTUAL_QTY_COL] = 0.0
            recs.append(rec)
        df_sim = pd.DataFrame(recs)
        out = pc._interactive_dispatch_trial_use_editor_rows_for_result_table(
            df_sim,
            json_rows,
            cols,
            interactive_dispatch_targets=None,
            timeline_events=None,
            task_queue=None,
            working_days=None,
        )
        v62 = out[out["依頼NO"].astype(str).str.contains("V6-2")]
        assert len(v62) == 2
        dates = sorted(str(pc._norm_ymd(r.get("配台日"))) for _, r in v62.iterrows())
        assert dates == ["2026/06/11", "2026/06/12"]
        conv = {float(r.get("換算数量") or 0) for _, r in v62.iterrows()}
        assert conv == {10000.0}
        plan_sum = sum(float(r.get("当日配台数量") or 0) for _, r in v62.iterrows())
        assert abs(plan_sum - 10000.0) < 1e-6
        assert abs(float(v62.iloc[0].get("当日配台数量") or 0) - 4000.0) < 1e-6
        assert abs(float(v62.iloc[1].get("当日配台数量") or 0) - 6000.0) < 1e-6
    finally:
        os.environ.pop("PM_AI_INTERACTIVE_DISPATCH_TRIAL", None)
