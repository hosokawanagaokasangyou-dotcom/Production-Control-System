# -*- coding: utf-8 -*-
"""段階2.5 学習推論のみモード: 既存アーカイブを参照し新規蓄積は行わない。"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from .actual_speed_distribution import load_speed_store

ENV_STAGE25_LEARNING_MODE = "PM_AI_STAGE2_5_LEARNING_MODE"
MODE_ACCUMULATE = "accumulate"
MODE_INFERENCE_ONLY = "inference_only"


def resolve_learning_mode() -> str:
    raw = (os.environ.get(ENV_STAGE25_LEARNING_MODE) or MODE_ACCUMULATE).strip().lower()
    if raw in (MODE_INFERENCE_ONLY, "inference", "infer", "read_only", "readonly"):
        return MODE_INFERENCE_ONLY
    return MODE_ACCUMULATE


def is_inference_only_mode() -> bool:
    return resolve_learning_mode() == MODE_INFERENCE_ONLY


def _count_archive_jobs(archive_root: Path) -> int:
    index_path = archive_root / "index.jsonl"
    if not index_path.is_file():
        return 0
    return sum(1 for line in index_path.read_text(encoding="utf-8").splitlines() if line.strip())


def _count_applicable_speed_keys(archive_root: Path) -> int:
    store = load_speed_store(archive_root)
    return sum(
        1
        for v in store.values()
        if isinstance(v, dict) and v.get("applied_speed_m_per_min") is not None
    )


def summarize_archive_for_inference(archive_root: str | Path) -> dict[str, Any]:
    root = Path(archive_root)
    job_count = _count_archive_jobs(root)
    speed_keys = _count_applicable_speed_keys(root)
    return {
        "archive_job_count": job_count,
        "speed_applicable_key_count": speed_keys,
        "ready": job_count > 0 or speed_keys > 0,
    }


def validate_archive_for_inference(archive_root: str | Path) -> dict[str, Any]:
    summary = summarize_archive_for_inference(archive_root)
    if not summary["ready"]:
        raise FileNotFoundError(
            "学習推論のみモード: 参照可能な学習データがありません。"
            f" archive={archive_root}（job={summary['archive_job_count']},"
            f" 速度キー={summary['speed_applicable_key_count']}）"
        )
    return summary


def load_low_l1_profile_triples(
    archive_root: str | Path,
    *,
    max_jobs: int = 20,
    l1_threshold_m: float = 1.0,
) -> set[tuple[str, str, str]]:
    """
    過去アーカイブでアラジン L1 乖離が小さい (工程名, 機械名, 依頼NO) を返す。
    整列時にアラジン重みを強めるプロファイルのヒントに使う。
    """
    root = Path(archive_root)
    index_path = root / "index.jsonl"
    if not index_path.is_file():
        return set()
    lines = [ln for ln in index_path.read_text(encoding="utf-8").splitlines() if ln.strip()]
    out: set[tuple[str, str, str]] = set()
    for line in reversed(lines[-max_jobs:]):
        try:
            entry = json.loads(line)
        except json.JSONDecodeError:
            continue
        folder = str(entry.get("folder") or "").strip()
        if not folder:
            continue
        metrics_path = root / folder / "aladdin_metrics.json"
        if not metrics_path.is_file():
            continue
        try:
            metrics = json.loads(metrics_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        for row in metrics.get("rows") or []:
            if not isinstance(row, dict):
                continue
            l1_raw = row.get("l1_deviation_m")
            if l1_raw is None:
                l1 = 999.0
            else:
                try:
                    l1 = float(l1_raw)
                except (TypeError, ValueError):
                    l1 = 999.0
            if l1 > l1_threshold_m:
                continue
            proc = str(row.get("process") or "").strip()
            machine = str(row.get("machine") or "").strip()
            tid = str(row.get("task_id") or "").strip()
            if proc and machine and tid:
                out.add((proc, machine, tid))
    return out
