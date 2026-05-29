# -*- coding: utf-8 -*-
"""アラジン計画と配台表の乖離指標。"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from .dispatch_aladdin_align_json import (
    _normalize_date,
    _nz,
    _parse_float,
    _profile_key,
    _roll_unit_m_for_profile,
    load_aladdin_lookup,
)

_ALADDIN_DATE = re.compile(r"^\d{4}/\d{2}/\d{2}$")
_EPS = 1e-9
_WIDE_IDENTITY = ("工程名", "機械名", "加工内容", "依頼NO", "換算数量", "実加工数", "計画合計")


def _dispatch_qty_by_profile_date(rows: list[dict[str, Any]]) -> dict[tuple[str, ...], dict[str, float]]:
    out: dict[tuple[str, ...], dict[str, float]] = {}
    for row in rows:
        pk = _profile_key(row)
        d = _normalize_date(_nz(row.get("配台日")))
        if not d:
            continue
        q = _parse_float(row.get("当日配台数量"))
        bucket = out.setdefault(pk, {})
        bucket[d] = bucket.get(d, 0.0) + max(0.0, q)
    return out


def _l1_profile_deviation(
    dispatch_by_date: dict[str, float],
    aladdin_by_date: dict[str, float],
) -> float:
    dates = sorted(set(dispatch_by_date) | set(aladdin_by_date))
    return sum(
        abs(dispatch_by_date.get(d, 0.0) - aladdin_by_date.get(d, 0.0)) for d in dates
    )


def compute_metrics(
    dispatch_payload: dict[str, Any],
    aladdin_path: Path,
) -> dict[str, Any]:
    rows: list[dict[str, Any]] = list(dispatch_payload.get("rows") or [])
    lookup = load_aladdin_lookup(aladdin_path)
    dispatch_map = _dispatch_qty_by_profile_date(rows)

    row_metrics: list[dict[str, Any]] = []
    total_l1 = 0.0
    machine_day: dict[tuple[str, str], dict[str, float]] = {}

    profiles: dict[tuple[str, ...], dict[str, Any]] = {}
    for row in rows:
        pk = _profile_key(row)
        if pk not in profiles:
            profiles[pk] = {h: row.get(h) for h in _WIDE_IDENTITY}

    for pk, qty_by_date in dispatch_map.items():
        profile = profiles.get(pk, {})
        proc, mk, tid = pk[0], pk[1], pk[3]
        aladdin_key = (mk, tid, proc)
        aladdin_by_date = lookup.get(aladdin_key, {})
        l1 = _l1_profile_deviation(qty_by_date, aladdin_by_date)
        total_l1 += l1
        unit_m = _roll_unit_m_for_profile(profile)
        row_metrics.append(
            {
                "process": proc,
                "machine": mk,
                "task_id": tid,
                "l1_deviation_m": round(l1, 4),
                "roll_unit_m": unit_m,
            }
        )
        for d, q in qty_by_date.items():
            a = aladdin_by_date.get(d, 0.0)
            key = (mk, d)
            bucket = machine_day.setdefault(key, {"dispatch_m": 0.0, "aladdin_m": 0.0})
            bucket["dispatch_m"] += q
            bucket["aladdin_m"] += a

    machine_day_metrics = []
    for (mk, d), vals in sorted(machine_day.items()):
        disp = vals["dispatch_m"]
        ala = vals["aladdin_m"]
        fit_pct = (min(disp, ala) / ala * 100.0) if ala > _EPS else (100.0 if disp <= _EPS else 0.0)
        machine_day_metrics.append(
            {
                "machine": mk,
                "date": d,
                "dispatch_m": round(disp, 4),
                "aladdin_m": round(ala, 4),
                "l1_deviation_m": round(abs(disp - ala), 4),
                "fit_pct": round(fit_pct, 2),
            }
        )

    n_rows = len(row_metrics)
    mean_l1 = total_l1 / n_rows if n_rows else 0.0
    return {
        "summary": {
            "profile_count": n_rows,
            "total_l1_deviation_m": round(total_l1, 4),
            "mean_l1_deviation_m": round(mean_l1, 4),
        },
        "rows": row_metrics,
        "machine_by_date": machine_day_metrics,
    }


def write_metrics_file(path: Path, metrics: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(metrics, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
