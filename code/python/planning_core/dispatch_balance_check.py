# -*- coding: utf-8 -*-
"""段階3配台照合（Java Stage3DispatchQtyBalanceCheck 相当）を JSON 行から判定。"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any

_EPS = 1e-3
_WIDE_IDENTITY = ("工程名", "機械名", "加工内容", "依頼NO", "換算数量", "実加工数", "計画合計")


def _nz(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _parse_float(v: Any) -> float:
    try:
        if v is None or (isinstance(v, str) and not v.strip()):
            return 0.0
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def roll_aligned_dispatch_m(raw_remaining_m: float, roll_unit_m: float) -> float:
    if raw_remaining_m <= _EPS:
        return 0.0
    if roll_unit_m <= _EPS:
        return raw_remaining_m
    n_rolls = int(math.ceil(raw_remaining_m / roll_unit_m - 1e-12))
    return roll_unit_m * n_rolls


def format_check(
    qty_converted: float,
    actual_processed: float,
    stage3_actual_total: float,
    *,
    roll_unit_m: float = 0.0,
    has_actual_col: bool = True,
) -> str:
    if not has_actual_col or stage3_actual_total <= _EPS:
        return ""
    raw_rem = max(0.0, qty_converted - actual_processed)
    expected = roll_aligned_dispatch_m(raw_rem, roll_unit_m)
    if abs(stage3_actual_total - expected) <= _EPS:
        if roll_unit_m > _EPS and expected > raw_rem + _EPS:
            return f"{_fmt_qty(raw_rem)} ({_fmt_qty(expected)}m)"
        return "OK"
    return f"NG (期待{_fmt_qty(expected)}／配台{_fmt_qty(stage3_actual_total)})"


def _fmt_qty(v: float) -> str:
    if abs(v - round(v)) <= _EPS:
        return str(int(round(v)))
    return f"{v:.4g}"


def _profile_key(row: dict[str, Any]) -> tuple[str, ...]:
    return tuple(_nz(row.get(h)) for h in _WIDE_IDENTITY)


def _roll_unit_for_profile(profile: dict[str, Any]) -> float:
    conv = _parse_float(profile.get("換算数量"))
    rolls = _parse_float(profile.get("原反数"))
    if rolls > _EPS and conv > _EPS:
        return conv / rolls
    pt = _parse_float(profile.get("計画合計"))
    return pt if pt > _EPS else (conv if conv > _EPS else 0.0)


@dataclass
class TaskBalanceResult:
    task_id: str
    process: str
    machine: str
    qty_converted: float
    actual_processed: float
    plan_total: float
    actual_dispatch_total: float
    expected: float
    check: str
    rows: list[dict[str, Any]]

    @property
    def ok(self) -> bool:
        return self.check == "OK" or (self.check and not self.check.startswith("NG"))


def check_task_balance(
    rows: list[dict[str, Any]],
    task_id: str,
    *,
    process: str | None = None,
    has_actual_col: bool = True,
) -> TaskBalanceResult | None:
    matched = [
        r
        for r in rows
        if _nz(r.get("依頼NO")) == task_id
        and (process is None or _nz(r.get("工程名")) == process)
    ]
    if not matched:
        return None
    profiles: dict[tuple[str, ...], dict[str, Any]] = {}
    for r in matched:
        pk = _profile_key(r)
        profiles.setdefault(pk, r)
    # 1 プロファイル想定（Y5-24 SEC 等）
    profile = next(iter(profiles.values()))
    pk = _profile_key(profile)
    task_rows = [r for r in matched if _profile_key(r) == pk]
    qty_conv = _parse_float(profile.get("換算数量"))
    actual_done = _parse_float(profile.get("実加工数"))
    plan_sum = sum(_parse_float(r.get("当日配台数量")) for r in task_rows)
    actual_sum = sum(_parse_float(r.get("実配台数量")) for r in task_rows)
    roll_u = _roll_unit_for_profile(profile)
    check = format_check(
        qty_conv,
        actual_done,
        actual_sum,
        roll_unit_m=roll_u,
        has_actual_col=has_actual_col,
    )
    raw_rem = max(0.0, qty_conv - actual_done)
    expected = roll_aligned_dispatch_m(raw_rem, roll_u)
    detail_rows = []
    for r in task_rows:
        detail_rows.append(
            {
                "dispatch_date": _nz(r.get("配台日")),
                "plan_m": _parse_float(r.get("当日配台数量")),
                "actual_m": _parse_float(r.get("実配台数量")),
                "start_dt": _nz(r.get("加工開始日時")),
            }
        )
    return TaskBalanceResult(
        task_id=task_id,
        process=_nz(profile.get("工程名")),
        machine=_nz(profile.get("機械名")),
        qty_converted=qty_conv,
        actual_processed=actual_done,
        plan_total=plan_sum,
        actual_dispatch_total=actual_sum,
        expected=expected,
        check=check,
        rows=detail_rows,
    )
