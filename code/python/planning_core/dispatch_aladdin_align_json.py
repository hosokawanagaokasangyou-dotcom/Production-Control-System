# -*- coding: utf-8 -*-
"""結果_配台表.json の当日配台数量をアラジン計画に沿って再配分（Java DispatchAladdinPlanAligner 相当）。"""
from __future__ import annotations

import json
import math
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

_ALADDIN_DATE = re.compile(r"^\d{4}/\d{2}/\d{2}$")
_EPS = 1e-9
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


def _normalize_date(s: str) -> str:
    s = _nz(s)
    if not s:
        return ""
    for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return s


def _profile_key(row: dict[str, Any]) -> tuple[str, ...]:
    return tuple(_nz(row.get(h)) for h in _WIDE_IDENTITY)


def _allocate_rolls_by_weight(total_rolls: int, weights: list[float]) -> list[int]:
    n = len(weights)
    rolls = [0] * n
    if total_rolls < 1 or n == 0:
        return rolls
    wsum = sum(weights)
    if wsum <= _EPS:
        return rolls
    remainders: list[float] = []
    assigned = 0
    for w in weights:
        if w <= _EPS:
            remainders.append(-1.0)
            continue
        exact = total_rolls * w / wsum
        r = int(math.floor(exact + _EPS))
        rolls[len(remainders)] = r
        remainders.append(exact - r)
        assigned += r
    while assigned < total_rolls:
        best = -1
        best_rem = -1.0
        for i, rem in enumerate(remainders):
            if weights[i] <= _EPS or rem < 0:
                continue
            if rem > best_rem + _EPS:
                best_rem = rem
                best = i
        if best < 0:
            break
        rolls[best] += 1
        remainders[best] = -1.0
        assigned += 1
    return rolls


def _align_row(current: list[float], aladdin: list[float], unit_m: float, uses_conv: bool) -> list[float]:
    n = len(current)
    if n == 0 or n != len(aladdin) or unit_m <= _EPS:
        return current[:]
    total = sum(max(0.0, v) for v in current)
    if total <= _EPS:
        return current[:]
    total_rolls = int(round(total / unit_m))
    if total_rolls < 1 or abs(total_rolls * unit_m - total) > _EPS:
        return current[:]
    weights: list[float] = []
    for a in aladdin:
        a = max(0.0, a)
        weights.append(1.0 if (uses_conv and a > _EPS) else a)
    if sum(weights) <= _EPS:
        return current[:]
    rolls = _allocate_rolls_by_weight(total_rolls, weights)
    return [rolls[i] * unit_m for i in range(n)]


def load_aladdin_lookup(path: Path) -> dict[tuple[str, str, str], dict[str, float]]:
    """(machine, tid, process) -> {yyyy-MM-dd: qty}."""
    if not path.is_file():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    headers = data.get("columns") or data.get("headers") or []
    rows = data.get("rows") or []
    if not headers:
        return {}
    mk_i = next((i for i, h in enumerate(headers) if h == "機械名"), -1)
    tid_i = next((i for i, h in enumerate(headers) if h == "依頼NO"), -1)
    proc_i = next((i for i, h in enumerate(headers) if h == "工程名"), -1)
    if mk_i < 0 or tid_i < 0:
        return {}
    date_cols: list[tuple[int, str]] = []
    for i, h in enumerate(headers):
        hs = _nz(h)
        if _ALADDIN_DATE.match(hs):
            date_cols.append((i, _normalize_date(hs)))
    out: dict[tuple[str, str, str], dict[str, float]] = {}
    for row in rows:
        if isinstance(row, dict):
            mk = _nz(row.get("機械名"))
            tid = _nz(row.get("依頼NO"))
            proc = _nz(row.get("工程名"))
            vals = {headers[i]: row.get(headers[i]) for i in range(len(headers))}
        else:
            mk = _nz(row[mk_i] if mk_i < len(row) else "")
            tid = _nz(row[tid_i] if tid_i < len(row) else "")
            proc = _nz(row[proc_i] if proc_i >= 0 and proc_i < len(row) else "")
            vals = {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers))}
        if not mk or not tid:
            continue
        key = (mk, tid, proc)
        bucket = out.setdefault(key, {})
        for i, d_iso in date_cols:
            if not d_iso:
                continue
            q = _parse_float(vals.get(headers[i]))
            if q > _EPS:
                bucket[d_iso] = bucket.get(d_iso, 0.0) + q
    return out


def _roll_unit_m_for_profile(profile: dict[str, Any]) -> float:
    conv = _parse_float(profile.get("換算数量"))
    rolls_raw = profile.get("原反数")
    rolls = _parse_float(rolls_raw)
    if rolls > _EPS and conv > _EPS:
        u = conv / rolls
        if u > _EPS:
            return u
    plan_total = _parse_float(profile.get("計画合計"))
    if plan_total > _EPS:
        return plan_total
    return conv if conv > _EPS else 0.0


def align_dispatch_json_from_aladdin(
    payload: dict[str, Any],
    aladdin_path: Path,
    *,
    align_from_day: date | None = None,
) -> tuple[dict[str, Any], int]:
    """
    ワイド行ごとに当日配台数量を再配分。align_from_day 以降の暦日のみ（None=全日）。
    Returns (new_payload, changed_row_count).
    """
    rows: list[dict[str, Any]] = list(payload.get("rows") or [])
    if not rows:
        return payload, 0
    lookup = load_aladdin_lookup(aladdin_path)
    profiles: dict[tuple[str, ...], dict[str, Any]] = {}
    for row in rows:
        pk = _profile_key(row)
        if pk not in profiles:
            profiles[pk] = {h: row.get(h) for h in _WIDE_IDENTITY}

    # 日付軸
    dates_set: set[str] = set()
    for row in rows:
        d = _normalize_date(_nz(row.get("配台日")))
        if d:
            dates_set.add(d)
    for bucket in lookup.values():
        dates_set.update(bucket.keys())
    axis = sorted(dates_set)
    if not axis:
        return payload, 0
    align_from_idx = 0
    if align_from_day is not None:
        iso = align_from_day.isoformat()
        for i, d in enumerate(axis):
            if d >= iso:
                align_from_idx = i
                break
        else:
            align_from_idx = len(axis)

    changed_profiles = 0
    new_qty_by_profile_date: dict[tuple[str, ...], dict[str, float]] = {}

    for pk, profile in profiles.items():
        mk, tid, proc = pk[1], pk[3], pk[0]
        aladdin_key = (mk, tid, proc)
        aladdin_by_date = lookup.get(aladdin_key, {})
        current = [
            sum(
                _parse_float(r.get("当日配台数量"))
                for r in rows
                if _profile_key(r) == pk and _normalize_date(_nz(r.get("配台日"))) == d
            )
            for d in axis
        ]
        aladdin = [aladdin_by_date.get(d, 0.0) for d in axis]
        unit_m = _roll_unit_m_for_profile(profile)
        uses_conv = any(0 < v < unit_m - _EPS for v in aladdin if unit_m > _EPS)
        prefix = current[:align_from_idx]
        suffix_cur = current[align_from_idx:]
        suffix_ala = aladdin[align_from_idx:]
        aligned_suffix = _align_row(suffix_cur, suffix_ala, unit_m, uses_conv)
        target = prefix + aligned_suffix
        if all(abs(a - b) <= _EPS for a, b in zip(current, target)):
            continue
        changed_profiles += 1
        new_qty_by_profile_date[pk] = {axis[i]: target[i] for i in range(len(axis))}

    if not new_qty_by_profile_date:
        return payload, 0

    # 行を再構築: 既存行を更新し、新規日付行を追加
    out_rows: list[dict[str, Any]] = []
    seen_pd: set[tuple[tuple[str, ...], str]] = set()
    template_by_profile: dict[tuple[str, ...], dict[str, Any]] = {}
    for row in rows:
        pk = _profile_key(row)
        template_by_profile.setdefault(pk, dict(row))
        d = _normalize_date(_nz(row.get("配台日")))
        qty_map = new_qty_by_profile_date.get(pk)
        if qty_map is None or d not in qty_map:
            out_rows.append(row)
            seen_pd.add((pk, d))
            continue
        new_q = qty_map[d]
        if new_q <= _EPS:
            continue
        nr = dict(row)
        nr["当日配台数量"] = new_q
        if "実配台数量" in nr:
            nr["実配台数量"] = 0.0
        out_rows.append(nr)
        seen_pd.add((pk, d))

    for pk, qty_map in new_qty_by_profile_date.items():
        tpl = template_by_profile.get(pk)
        if not tpl:
            continue
        for d, q in qty_map.items():
            if q <= _EPS or (pk, d) in seen_pd:
                continue
            nr = dict(tpl)
            nr["配台日"] = d.replace("-", "/")
            nr["当日配台数量"] = q
            nr["実配台数量"] = 0.0
            nr["加工開始日時"] = ""
            nr["加工終了日時"] = ""
            nr["メンバー名"] = ""
            out_rows.append(nr)

    new_payload = dict(payload)
    new_payload["rows"] = out_rows
    new_payload["row_count"] = len(out_rows)
    return new_payload, changed_profiles
