# -*- coding: utf-8 -*-
"""E6-4/EC のみ配台対象にした plan_input を output に書き出す（デバッグ用）。"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

import pandas as pd


def main() -> int:
    repo = Path(__file__).resolve().parents[1]
    src = (
        repo
        / "pm-ai-package-release"
        / "PMD_initial_install"
        / "pm-ai-data"
        / "output"
        / "plan_input_tasks.xlsx"
    )
    if not src.is_file():
        print(f"source missing: {src}", file=sys.stderr)
        return 1
    dst = repo / "output" / "plan_input_e64_only_debug.xlsx"
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    df = pd.read_excel(dst, sheet_name="タスク一覧")
    df.columns = df.columns.str.strip()
    ex_cols = [c for c in df.columns if "配台不要" in str(c)]
    if not ex_cols:
        print("配台不要 column missing", file=sys.stderr)
        return 1
    ex = ex_cols[0]
    for i, row in df.iterrows():
        tid = str(row.get("依頼NO", "")).strip()
        proc = str(row.get("工程名", "")).strip()
        if tid == "E6-4" and proc == "EC":
            df.at[i, ex] = ""
        else:
            df.at[i, ex] = "yes"
    with pd.ExcelWriter(dst, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="タスク一覧", index=False)
    yes = df[ex].astype(str).str.strip().str.lower().isin(
        ["yes", "y", "1", "true", "はい", "配台不要", "on"]
    )
    print(f"written {dst} active={len(df) - int(yes.sum())}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
