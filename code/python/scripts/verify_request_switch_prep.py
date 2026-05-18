# -*- coding: utf-8 -*-
"""
依頼切替・休憩再開の準備時間 — 実行確認（master 読込 + 配台ロジックのスモーク）。

マスタは環境変数 ``PM_AI_MASTER_WORKBOOK``（絶対パス・必須）のみ。段階2配台と同じ。

用法:
  cd code/python
  PM_AI_MASTER_WORKBOOK=/path/to/国分master.xlsm PYTHONPATH=. python3.14 scripts/verify_request_switch_prep.py
"""
from __future__ import annotations

import os
import sys
from datetime import date, datetime

_py_here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _py_here not in sys.path:
    sys.path.insert(0, _py_here)

import planning_core._core as core

# 実行確認のマスタ正本（国分工場）
VERIFY_MASTER_BASENAME = "国分master.xlsm"


def resolve_verify_master_path() -> str:
    """確認ロジック用マスタパス（``PM_AI_MASTER_WORKBOOK`` 必須。段階2と同じ解決）。"""
    return core._master_workbook_path_resolved()


def main() -> int:
    master = resolve_verify_master_path()
    print("=== 依頼切替準備 実行確認 ===")
    print(f"master: {master}")
    if not os.path.isfile(master):
        print(f"ERROR: マスタブックが見つかりません: {master}")
        return 2

    sp, sm, rp, rm = core.load_request_switch_prep_settings(master)
    core._STAGE2_REQUEST_SWITCH_PREP_BY_PROC_MACHINE = sp
    core._STAGE2_REQUEST_SWITCH_PREP_BY_MACHINE = sm
    core._STAGE2_BREAK_RESUME_PREP_BY_PROC_MACHINE = rp
    core._STAGE2_BREAK_RESUME_PREP_BY_MACHINE = rm

    n_switch = len({k for k in sp if isinstance(k, tuple)}) + len(sm)
    n_resume = len({k for k in rp if isinstance(k, tuple)}) + len(rm)
    print(f"読込: 依頼切替準備 {n_switch} 件 / 休憩再開準備 {n_resume} 件")
    if n_switch == 0:
        print(
            "WARN: 準備時間が 0 のみ、または列見出し不一致。"
            " シート「設定_依頼切替前後時間」の 準備時間_分 を確認してください。"
        )

    samples = list(sp.items())[:5]
    for (proc, mn), mins in samples:
        print(f"  例: ({proc}, {mn}) => {mins} 分")
    for (proc, mn), mins in list(rp.items())[:3]:
        print(f"  再開例: ({proc}, {mn}) => {mins} 分")

    proc, mn = "スライス", "スライス機1"
    if (proc, mn) not in sp:
        proc, mn = next(iter(sp.keys()), (proc, mn)) if sp else (proc, mn)
    prep = core._lookup_request_switch_prep_minutes(proc, mn)
    resume = core._lookup_break_resume_prep_minutes(proc, mn)
    print(f"\nルックアップ: {proc!r} + {mn!r} => 準備 {prep} 分 / 再開 {resume} 分")

    d = date.today()
    mh_switch = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
    }
    t0 = datetime.combine(d, datetime.strptime("12:50", "%H:%M").time())
    ts_sw, segs_sw = core._roll_prep_segments_for_assign(
        team_start=t0,
        team_breaks=[],
        machine_handoff=mh_switch,
        machine_occ_key="occ1",
        current_date=d,
        task_id="B002",
        machine_proc=proc,
        machine_name=mn,
        eq_line=f"{proc}+{mn}",
        abolish_limits=False,
    )
    print(f"\n[1] 依頼切替（A001→B002）:")
    print(f"  加工開始: {t0} -> {ts_sw} (+{(ts_sw - t0).total_seconds() / 60:.0f} 分)")
    for s in segs_sw:
        print(
            f"  セグメント: {s.get('event_kind')} "
            f"[{s.get('start_dt')}, {s.get('end_dt')})"
        )

    break_end = datetime.combine(d, datetime.strptime("13:00", "%H:%M").time())
    mh_same = {
        "last_tid": {"occ1": "A001"},
        "last_machining_date": {"occ1": d},
        "machining_today_occ": {"occ1"},
    }
    ts_rs, segs_rs = core._roll_prep_segments_for_assign(
        team_start=break_end,
        team_breaks=[(datetime.combine(d, datetime.strptime("12:00", "%H:%M").time()), break_end)],
        machine_handoff=mh_same,
        machine_occ_key="occ1",
        current_date=d,
        task_id="A001",
        machine_proc=proc,
        machine_name=mn,
        eq_line=f"{proc}+{mn}",
        abolish_limits=False,
    )
    print(f"\n[2] 同一依頼・休憩明け再開:")
    print(f"  加工開始: {break_end} -> {ts_rs} (+{(ts_rs - break_end).total_seconds() / 60:.0f} 分)")
    for s in segs_rs:
        print(
            f"  セグメント: {s.get('event_kind')} "
            f"[{s.get('start_dt')}, {s.get('end_dt')})"
        )

    ok_switch = (
        prep > 0
        and segs_sw
        and segs_sw[0].get("event_kind") == core.TIMELINE_EVENT_REQUEST_SWITCH_PREP
    )
    ok_resume = (
        resume > 0
        and segs_rs
        and segs_rs[0].get("event_kind") == core.TIMELINE_EVENT_BREAK_RESUME_PREP
    )
    if ok_switch and (resume == 0 or ok_resume):
        print("\nOK: 国分master 読込と準備セグメント生成を確認しました。")
        return 0
    if n_switch == 0:
        print("\nSKIP: マスタに有効な準備分がありません。")
        return 0
    print("\nFAIL: 期待どおりのセグメントが生成されませんでした。")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
