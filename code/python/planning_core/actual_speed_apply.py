# -*- coding: utf-8 -*-
"""学習推定速度を配台計画 DataFrame に適用する。"""
from __future__ import annotations

import logging
import os

import pandas as pd

from .actual_speed_distribution import load_speed_store, speed_json_path
from .dispatch_workspace import resolve_dispatch_learning_archive_root

ENV_LEARNED_SPEED_ENABLED = "PM_AI_LEARNED_SPEED_ENABLED"
ENV_LEARNED_SPEED_MIN_SAMPLES = "PM_AI_LEARNED_SPEED_MIN_SAMPLES"


def _truthy_env(name: str, default: bool = True) -> bool:
    raw = (os.environ.get(name) or "").strip().lower()
    if not raw:
        return default
    return raw not in ("0", "false", "no", "off", "none")


def _env_int(name: str, default: int) -> int:
    raw = (os.environ.get(name) or "").strip()
    if not raw:
        return default
    try:
        return int(float(raw))
    except ValueError:
        return default


def apply_learned_speed_to_plan_df(
    df: pd.DataFrame,
    *,
    log_prefix: str,
) -> None:
    """master 速度適用後に呼ぶ。列「加工速度」が正の行のみ更新（手入力済みは上書きしない）。"""
    if df is None or df.empty:
        return
    if not _truthy_env(ENV_LEARNED_SPEED_ENABLED, False):
        return
    from ._core import TASK_COL_MACHINE, TASK_COL_MACHINE_NAME, TASK_COL_SPEED

    if TASK_COL_SPEED not in df.columns:
        return
    archive_root = resolve_dispatch_learning_archive_root()
    store_path = speed_json_path(archive_root)
    if not store_path.is_file():
        return
    store = load_speed_store(archive_root)
    if not store:
        return
    min_samples = _env_int(ENV_LEARNED_SPEED_MIN_SAMPLES, 5)
    n_hit = 0
    for i, row in df.iterrows():
        cur = row.get(TASK_COL_SPEED)
        try:
            if cur is not None and float(cur) > 0:
                continue
        except (TypeError, ValueError):
            pass
        proc = str(row.get(TASK_COL_MACHINE) or "").strip()
        machine = str(row.get(TASK_COL_MACHINE_NAME) or "").strip()
        key = f"{proc}|{machine}"
        entry = store.get(key)
        if not isinstance(entry, dict):
            continue
        if int(entry.get("n") or 0) < min_samples:
            continue
        spd = entry.get("applied_speed_m_per_min")
        if spd is None:
            continue
        try:
            spd_f = float(spd)
        except (TypeError, ValueError):
            continue
        if spd_f <= 0:
            continue
        df.at[i, TASK_COL_SPEED] = spd_f
        n_hit += 1
    if n_hit:
        logging.info(
            "%s: 学習推定速度を %s 行に適用（store=%s）。",
            log_prefix,
            n_hit,
            store_path,
        )


# 段階1/配台読込から呼ぶ別名（計画の命名に合わせる）
_apply_learned_speed_to_plan_df = apply_learned_speed_to_plan_df
