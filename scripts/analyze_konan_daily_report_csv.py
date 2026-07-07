# -*- coding: utf-8 -*-
"""湖南工場 加工日報発行問合せ CSV の構造・加工完了列を解析する。"""
from __future__ import annotations

import csv
import glob
import os
import re
from collections import Counter
from datetime import datetime

DIR = r"\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\生産管理システム\管理システム\●DATA\加工日報"


def read_csv(path: str):
    with open(path, encoding="utf-8-sig", newline="") as fp:
        rows = list(csv.reader(fp))
    return rows[:3], rows[3], rows[4:]


def main() -> None:
    files = sorted(
        glob.glob(os.path.join(DIR, "加工日報発行問合せ_*.csv")),
        key=os.path.getmtime,
        reverse=True,
    )
    print("=== ファイル一覧 (最新10) ===")
    for f in files[:10]:
        print(
            f"{os.path.basename(f)}\t{os.path.getsize(f):,}\t"
            f"{datetime.fromtimestamp(os.path.getmtime(f))}"
        )

    targets = {
        "latest": files[0],
        "largest_recent": max(files[:5], key=os.path.getsize),
    }
    single = next((f for f in files if "20260623_193059" in f), None)
    if single:
        targets["single_day"] = single

    for label, path in targets.items():
        meta, header, data = read_csv(path)
        print(f"\n=== {label}: {os.path.basename(path)} ===")
        for i, m in enumerate(meta):
            print(f"  meta[{i}]: {m[0] if m else ''}")
        print(f"  列数: {len(header)}, データ行: {len(data)}")

        idx = {c: i for i, c in enumerate(header)}
        key_cols = [
            "依頼NO", "工程名", "機械名", "加工日付", "換算数量", "実加工数",
            "加工日加工予定数", "実製品出来高", "完了区分", "加工完了日",
            "注文単位加工完了区分", "注文単位加工完了日", "加工実績累計", "実製品出来高累計",
        ]
        comp_cols = [c for c in key_cols if c in idx]

        print("  全列名:")
        for i, c in enumerate(header):
            print(f"    [{i:3d}] {c}")

        print("  完了関連列の値分布:")
        for col in [c for c in header if "完了" in c]:
            vals = Counter(
                (r[idx[col]].strip() if idx[col] < len(r) else "")
                for r in data
            )
            print(f"    {col}: {dict(vals.most_common(10))}")

        key_latest: dict[tuple[str, str, str], dict[str, str]] = {}
        for r in data:
            if idx["依頼NO"] >= len(r):
                continue
            irai = r[idx["依頼NO"]].strip()
            proc = r[idx["工程名"]].strip() if idx["工程名"] < len(r) else ""
            mach = r[idx["機械名"]].strip() if idx["機械名"] < len(r) else ""
            day = r[idx["加工日付"]].strip() if idx["加工日付"] < len(r) else ""
            k = (irai, proc, mach)
            if k not in key_latest or day > key_latest[k]["_day"]:
                rec = {
                    c: (r[idx[c]].strip() if idx[c] < len(r) else "")
                    for c in comp_cols
                    if c in idx
                }
                rec["_day"] = day
                key_latest[k] = rec

        order_done = sum(
            1
            for v in key_latest.values()
            if v.get("注文単位加工完了区分", "").endswith("完了")
        )
        order_mikan = sum(
            1 for v in key_latest.values() if "未完" in v.get("注文単位加工完了区分", "")
        )
        print(f"  ユニーク(依頼NO,工程,機械): {len(key_latest)}")
        print(f"  注文単位 完了={order_done}, 未完={order_mikan}")

        print("  サンプル (最新加工日順 8件):")
        samples = sorted(key_latest.values(), key=lambda x: x["_day"], reverse=True)[:8]
        for s in samples:
            print(
                f"    {s.get('_day')} {s.get('依頼NO')} {s.get('工程名')} {s.get('機械名')} | "
                f"注文={s.get('注文単位加工完了区分')} {s.get('注文単位加工完了日')} | "
                f"完了区分={s.get('完了区分')} 工程完了日={s.get('加工完了日')} | "
                f"累計={s.get('加工実績累計')}/{s.get('換算数量')} "
                f"日次実加工={s.get('実加工数')} 予定={s.get('加工日加工予定数')}"
            )

    # 未完サンプル詳細 (largest file)
    path = targets["largest_recent"]
    _, header, data = read_csv(path)
    idx = {c: i for i, c in enumerate(header)}
    print("\n=== 未完タスク詳細 (largest_recent) ===")
    seen = set()
    for r in sorted(data, key=lambda x: x[idx["加工日付"]] if idx["加工日付"] < len(x) else "", reverse=True):
        flag = r[idx["注文単位加工完了区分"]].strip() if idx["注文単位加工完了区分"] < len(r) else ""
        if "未完" not in flag:
            continue
        irai = r[idx["依頼NO"]].strip()
        proc = r[idx["工程名"]].strip()
        if (irai, proc) in seen:
            continue
        seen.add((irai, proc))
        cum = r[idx["加工実績累計"]].strip() if idx["加工実績累計"] < len(r) else ""
        conv = r[idx["換算数量"]].strip() if idx["換算数量"] < len(r) else ""
        day = r[idx["加工日付"]].strip() if idx["加工日付"] < len(r) else ""
        print(
            f"  {day} {irai} {proc} | 注文={flag} | "
            f"累計={cum}/{conv} | 完了区分={r[idx['完了区分']].strip() if idx['完了区分']<len(r) else ''}"
        )
        if len(seen) >= 10:
            break


if __name__ == "__main__":
    main()
