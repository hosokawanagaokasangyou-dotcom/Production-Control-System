# -*- coding: utf-8 -*-
"""加工実績明細から (工程名, 機械名) 別の速度分布を蓄積する。"""
from __future__ import annotations

import json
import logging
import math
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd

from .dispatch_learning_dedup import (
    is_observation_seen,
    observation_id_from_fields,
    register_observation_ids,
    speed_source_fingerprint,
)
from .dispatch_workspace import resolve_actual_detail_workbook_path

SPEED_JSON_NAME = "process_machine_speed.json"
OBSERVATIONS_JSONL = "observations.jsonl"
ENV_LEARNED_SPEED_HISTOGRAM_BIN_WIDTH = "PM_AI_LEARNED_SPEED_HISTOGRAM_BIN_WIDTH"
ENV_LEARNED_SPEED_PERCENTILE = "PM_AI_LEARNED_SPEED_PERCENTILE"
ENV_LEARNED_SPEED_MIN_SAMPLES = "PM_AI_LEARNED_SPEED_MIN_SAMPLES"
DEFAULT_BIN_WIDTH = 1.0
_MIN_DURATION_MIN = 1.0
_EPS = 1e-9


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def speed_distributions_dir(archive_root: str | Path) -> Path:
    return Path(archive_root) / "speed-distributions"


def speed_json_path(archive_root: str | Path) -> Path:
    return speed_distributions_dir(archive_root) / SPEED_JSON_NAME


def observations_jsonl_path(archive_root: str | Path) -> Path:
    return speed_distributions_dir(archive_root) / OBSERVATIONS_JSONL


def _env_float(name: str, default: float) -> float:
    raw = (os.environ.get(name) or "").strip()
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


def _env_int(name: str, default: int) -> int:
    raw = (os.environ.get(name) or "").strip()
    if not raw:
        return default
    try:
        return int(float(raw))
    except ValueError:
        return default


def speed_key(process: str, machine: str) -> str:
    return f"{(process or '').strip()}|{(machine or '').strip()}"


def load_speed_store(archive_root: str | Path) -> dict[str, Any]:
    p = speed_json_path(archive_root)
    if not p.is_file():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def save_speed_store(archive_root: str | Path, data: dict[str, Any]) -> None:
    p = speed_json_path(archive_root)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp, p)


def percentile(values: list[float], p: float) -> float:
    if not values:
        return 0.0
    xs = sorted(values)
    if len(xs) == 1:
        return xs[0]
    rank = (p / 100.0) * (len(xs) - 1)
    lo = int(math.floor(rank))
    hi = int(math.ceil(rank))
    if lo == hi:
        return xs[lo]
    frac = rank - lo
    return xs[lo] * (1.0 - frac) + xs[hi] * frac


def histogram_from_observations(
    observations: list[float],
    *,
    bin_width: float,
) -> dict[str, Any]:
    if not observations:
        return {
            "unit": "m_per_min",
            "bin_width": bin_width,
            "bin_start": 0.0,
            "counts": [],
            "bin_edges": [],
        }
    bw = max(bin_width, _EPS)
    lo = math.floor(min(observations) / bw) * bw
    hi = math.ceil(max(observations) / bw) * bw
    if hi <= lo:
        hi = lo + bw
    n_bins = max(1, int(round((hi - lo) / bw)))
    edges = [lo + i * bw for i in range(n_bins + 1)]
    counts = [0] * n_bins
    for v in observations:
        if v < lo:
            idx = 0
        elif v >= edges[-1]:
            idx = n_bins - 1
        else:
            idx = min(n_bins - 1, int((v - lo) / bw))
        counts[idx] += 1
    return {
        "unit": "m_per_min",
        "bin_width": bw,
        "bin_start": lo,
        "counts": counts,
        "bin_edges": edges,
    }


def merge_observation_into_histogram(
    entry: dict[str, Any],
    speed_m_per_min: float,
    *,
    bin_width: float,
) -> None:
    hist = entry.setdefault(
        "histogram",
        histogram_from_observations([], bin_width=bin_width),
    )
    bw = float(hist.get("bin_width") or bin_width or DEFAULT_BIN_WIDTH)
    counts = list(hist.get("counts") or [])
    edges = list(hist.get("bin_edges") or [])
    if not counts or not edges:
        hist.update(histogram_from_observations([speed_m_per_min], bin_width=bw))
        return
    if speed_m_per_min < edges[0]:
        counts[0] += 1
    elif speed_m_per_min >= edges[-1]:
        counts[-1] += 1
    else:
        idx = min(len(counts) - 1, int((speed_m_per_min - edges[0]) / bw))
        counts[idx] += 1
    hist["counts"] = counts


def _recompute_summary(entry: dict[str, Any], speeds: list[float], percentile_p: float, min_samples: int) -> None:
    entry["n"] = len(speeds)
    if not speeds:
        return
    mean = sum(speeds) / len(speeds)
    var = sum((x - mean) ** 2 for x in speeds) / len(speeds)
    entry["mean_m_per_min"] = round(mean, 4)
    entry["std_m_per_min"] = round(math.sqrt(var), 4)
    entry["p25"] = round(percentile(speeds, 25), 4)
    entry["p50"] = round(percentile(speeds, 50), 4)
    entry["p75"] = round(percentile(speeds, 75), 4)
    p_apply = percentile(speeds, percentile_p)
    entry["applied_speed_m_per_min"] = round(p_apply, 4) if len(speeds) >= min_samples else None
    entry["last_observation_at"] = _utc_now_iso()


def _outlier_bounds(speeds: list[float]) -> tuple[float, float] | None:
    if len(speeds) < 20:
        return None
    return percentile(speeds, 1), percentile(speeds, 99)


_LOG = logging.getLogger(__name__)

_DURATION_FALLBACK_COLS = (
    "稼働時間分換算",
    "加工時間分換算",
    "稼働時間分",
    "加工時間分",
)


def _row_actual_qty_m(row) -> float:
    from ._core import ACT_COL_ACTUAL_QTY, ACT_COL_CONVERTED_QTY, parse_float_safe

    for col in (ACT_COL_ACTUAL_QTY, ACT_COL_CONVERTED_QTY, "換算数量"):
        qty = parse_float_safe(row.get(col), None)
        if qty is not None and qty > _EPS:
            return float(qty)
    return 0.0


def _row_duration_minutes(row) -> float | None:
    from ._core import _actual_row_time_bounds, parse_float_safe

    start, end = _actual_row_time_bounds(row)
    if start and end and start < end:
        minutes = (end - start).total_seconds() / 60.0
        if minutes >= _MIN_DURATION_MIN:
            return minutes
    for col in _DURATION_FALLBACK_COLS:
        minutes = parse_float_safe(row.get(col), 0.0)
        if minutes >= _MIN_DURATION_MIN:
            return float(minutes)
    return None


def extract_observations_from_detail_df(df: pd.DataFrame) -> list[dict[str, Any]]:
    if df is None or df.empty:
        return []
    from ._core import (
        ACT_COL_PROCESS,
        ACT_COL_TASK_ID,
        TASK_COL_MACHINE_NAME,
        _actual_row_time_bounds,
    )

    out: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        proc = str(row.get(ACT_COL_PROCESS) or "").strip()
        machine = str(row.get(TASK_COL_MACHINE_NAME) or row.get("機械名") or "").strip()
        tid = str(row.get(ACT_COL_TASK_ID) or "").strip()
        qty = _row_actual_qty_m(row)
        if qty <= _EPS or not proc:
            continue
        minutes = _row_duration_minutes(row)
        if minutes is None:
            continue
        speed = qty / minutes
        start, end = _actual_row_time_bounds(row)
        if start and end:
            start_iso = start.isoformat()
            end_iso = end.isoformat()
        else:
            row_tag = str(row.get("行NO") or row.get("加工実績NO") or "").strip()
            day_key = str(row.get("加工日") or row.get("日付") or "")
            start_iso = f"{day_key}|{row_tag}|{minutes:.4f}"
            end_iso = start_iso
        obs_id = observation_id_from_fields(tid, proc, machine, start_iso, end_iso, qty)
        out.append(
            {
                "observation_id": obs_id,
                "process": proc,
                "machine": machine,
                "task_id": tid,
                "speed_m_per_min": round(speed, 6),
                "actual_qty_m": qty,
                "duration_min": round(minutes, 4),
                "start_iso": start_iso,
                "end_iso": end_iso,
            }
        )
    return out


def update_speed_distribution(
    archive_root: str | Path,
    *,
    task_input_workbook: str = "",
    force_full: bool = False,
) -> dict[str, Any]:
    """実績明細から新規観測のみマージ。戻り値はサマリ。"""
    from .dispatch_learning_dedup import load_registry, save_registry

    archive_root = Path(archive_root)
    wb = resolve_actual_detail_workbook_path(task_input_workbook)
    if not wb or not os.path.isfile(wb):
        return {"added": 0, "skipped_dup": 0, "reason": "no_workbook"}
    fp = speed_source_fingerprint(wb)
    reg = load_registry(archive_root)
    speed_files = reg.setdefault("speed_source_files", {})
    if not force_full and fp in speed_files:
        prior_store = load_speed_store(archive_root)
        prior_seen = load_seen_observation_ids(archive_root)
        if prior_store or prior_seen:
            return {"added": 0, "skipped_dup": 0, "reason": "source_unchanged"}

    from ._core import load_machining_actual_detail_df

    df = load_machining_actual_detail_df()
    observations = extract_observations_from_detail_df(df)
    if not observations and df is not None and not df.empty:
        _LOG.warning(
            "速度分布: 実績明細 %s 行を読んだが有効観測 0 件（列不一致・時間0の可能性）。",
            len(df),
        )
    store = load_speed_store(archive_root)
    speeds_by_key: dict[str, list[float]] = {}
    for key, entry in store.items():
        if isinstance(entry, dict) and entry.get("n"):
            hist = entry.get("histogram") or {}
            bw = float(hist.get("bin_width") or DEFAULT_BIN_WIDTH)
            n = int(entry.get("n") or 0)
            p50 = float(entry.get("p50") or entry.get("mean_m_per_min") or 0.0)
            if n > 0 and p50 > 0:
                speeds_by_key[key] = [p50] * n

    bin_width = _env_float(ENV_LEARNED_SPEED_HISTOGRAM_BIN_WIDTH, DEFAULT_BIN_WIDTH)
    percentile_p = _env_float(ENV_LEARNED_SPEED_PERCENTILE, 50.0)
    min_samples = _env_int(ENV_LEARNED_SPEED_MIN_SAMPLES, 5)

    added = 0
    skipped_dup = 0
    new_ids: list[str] = []
    obs_path = observations_jsonl_path(archive_root)
    obs_path.parent.mkdir(parents=True, exist_ok=True)

    for obs in observations:
        obs_id = obs["observation_id"]
        if is_observation_seen(archive_root, obs_id):
            skipped_dup += 1
            continue
        key = speed_key(obs["process"], obs["machine"])
        speeds = speeds_by_key.setdefault(key, [])
        bounds = _outlier_bounds(speeds)
        spd = float(obs["speed_m_per_min"])
        if bounds and (spd < bounds[0] or spd > bounds[1]):
            skipped_dup += 1
            continue
        speeds.append(spd)
        entry = store.setdefault(
            key,
            {
                "process": obs["process"],
                "machine": obs["machine"],
            },
        )
        entry["process"] = obs["process"]
        entry["machine"] = obs["machine"]
        merge_observation_into_histogram(entry, spd, bin_width=bin_width)
        _recompute_summary(entry, speeds, percentile_p, min_samples)
        with obs_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(obs, ensure_ascii=False) + "\n")
        new_ids.append(obs_id)
        added += 1

    if new_ids:
        register_observation_ids(archive_root, new_ids)
    save_speed_store(archive_root, store)
    if observations:
        speed_files[fp] = {"fingerprint": fp, "updated_at": _utc_now_iso()}
        save_registry(archive_root, reg)
    return {
        "added": added,
        "skipped_dup": skipped_dup,
        "source": wb,
        "observation_candidates": len(observations),
        "detail_rows": len(df) if df is not None else 0,
    }


def write_ml_readiness(archive_root: str | Path) -> None:
    root = Path(archive_root)
    index_path = root / "index.jsonl"
    job_count = 0
    if index_path.is_file():
        job_count = sum(1 for line in index_path.read_text(encoding="utf-8").splitlines() if line.strip())
    speed_store = load_speed_store(root)
    applicable = sum(
        1
        for v in speed_store.values()
        if isinstance(v, dict) and v.get("applied_speed_m_per_min") is not None
    )
    payload = {
        "updated_at": _utc_now_iso(),
        "ml_mode_active": (os.environ.get("PM_AI_STAGE2_5_ML_MODE") or "off").strip() or "off",
        "archive_job_count": job_count,
        "strong_teacher_count": 0,
        "speed_key_count": len(speed_store),
        "speed_applicable_key_count": applicable,
        "layers": {
            "mvp": {"eligible": True, "enabled": True, "blockers": []},
            "ml0": {
                "eligible": job_count >= 10,
                "enabled": False,
                "blockers": [] if job_count >= 10 else [f"archive_job_count {job_count}/10"],
            },
            "ml1": {
                "eligible": job_count >= 20,
                "enabled": False,
                "blockers": [] if job_count >= 20 else [f"archive_job_count {job_count}/20"],
            },
            "ml2": {
                "eligible": job_count >= 50,
                "enabled": False,
                "blockers": [] if job_count >= 50 else [f"archive_job_count {job_count}/50"],
            },
            "ml3": {
                "eligible": False,
                "enabled": False,
                "blockers": ["strong_teacher_count 0/10"],
            },
        },
        "holdout": {"aladdin_l1_improvement_pct": None, "shortage_regression": False},
    }
    out = root / "ml_readiness.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
