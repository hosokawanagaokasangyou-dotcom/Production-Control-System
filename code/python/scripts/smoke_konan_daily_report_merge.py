# -*- coding: utf-8 -*-
"""Live smoke test for konan daily report merge (optional, needs network share)."""
from planning_core.konan_daily_report import (
    aggregate_daily_report_latest,
    merge_daily_report_into_plan_df,
    read_konan_daily_report_csv,
    resolve_daily_report_csv_path,
)
import pandas as pd

p = resolve_daily_report_csv_path()
print("path:", p)
dr = aggregate_daily_report_latest(read_konan_daily_report_csv(p))
print("aggregated keys:", len(dr))
sample = dr[dr["依頼NO"].astype(str).str.startswith("Y6-19")]
print("Y6-19 rows:", len(sample))
if not sample.empty:
    print(sample[["依頼NO", "工程名", "機械名", "加工実績累計", "完了区分"]].head().to_string())

if not sample.empty:
    row = sample.iloc[0]
    plan = pd.DataFrame(
        [
            {
                "依頼NO": row["依頼NO"],
                "工程名": row["工程名"],
                "機械名": row["機械名"],
                "換算数量": row["換算数量"],
                "未加工": row["換算数量"],
                "実加工数": "0",
                "加工完了区分": "",
            }
        ]
    )
    merged = merge_daily_report_into_plan_df(plan)
    print("merged:\n", merged[["依頼NO", "実加工数", "未加工", "加工完了区分"]].to_string())
