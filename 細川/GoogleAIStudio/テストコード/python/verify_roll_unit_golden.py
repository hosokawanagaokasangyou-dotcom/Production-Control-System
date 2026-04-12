#!/usr/bin/env python3
"""
製品名,ロール単位の長さ.txt と infer_unit_m_from_product_name の整合を確認する。

使い方（テストコードの python フォルダで）:
  py -3 verify_roll_unit_golden.py

期待値はシート上の「ロール単位長さ」と同様、次のいずれかと一致すれば合格とみなす:
  - 推定 raw 値そのもの
  - または raw を ROLL_UNIT_LENGTH_CEIL_STEP_M（100m）刻みに切り上げた値
  （例: 40→100、95→100、125→200。250 は raw=250 のまま表に出ている行もある）

寸法が無い品番（例: FEL3002…）は換算数量フォールバックが必要なため本スクリプトではスキップする。
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
GOLDEN = ROOT.parent / "製品名,ロール単位の長さ.txt"


def _ceil_step(v: float, step: float) -> float:
    if v <= 0 or step <= 0:
        return v
    return float(math.ceil(v / step) * step)


def main() -> int:
    if not GOLDEN.is_file():
        print(f"ゴールデンファイルが見つかりません: {GOLDEN}", file=sys.stderr)
        return 2

    import logging

    # planning_core 読込時のログを抑止（検証には不要）
    try:
        _prev_disable = logging.root.manager.disable
    except AttributeError:
        _prev_disable = logging.NOTSET
    logging.disable(logging.WARNING)
    sys.path.insert(0, str(ROOT))
    try:
        from planning_core._core import (  # noqa: E402
            ROLL_UNIT_LENGTH_CEIL_STEP_M,
            infer_unit_m_from_product_name,
        )
    finally:
        logging.disable(_prev_disable)

    lines = GOLDEN.read_text(encoding="utf-8").splitlines()
    step = float(ROLL_UNIT_LENGTH_CEIL_STEP_M)
    ok = 0
    bad: list[tuple[str, int, float, float]] = []
    skipped_no_dim = 0

    for line in lines[1:]:
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        if "," not in s:
            continue
        name, _, exp_s = s.rpartition(",")
        name = name.strip()
        exp_s = exp_s.strip()
        if not name or not exp_s:
            continue
        try:
            expected = int(float(exp_s))
        except ValueError:
            continue

        raw = infer_unit_m_from_product_name(name, fallback_unit=0.0)
        try:
            raw_f = float(raw)
        except (TypeError, ValueError):
            raw_f = 0.0

        if raw_f <= 0:
            skipped_no_dim += 1
            print(
                f"[スキップ] 寸法から推定不可（換算数量フォールバック前提）: "
                f"expected={expected} name={name!r}"
            )
            continue

        ceiled = _ceil_step(raw_f, step)
        if expected == int(round(raw_f)) or expected == int(round(ceiled)):
            ok += 1
        else:
            bad.append((name, expected, raw_f, ceiled))

    print(
        f"ゴールデン: {GOLDEN.name}\n"
        f"  一致: {ok} 件\n"
        f"  寸法なしスキップ: {skipped_no_dim} 件\n"
        f"  不一致: {len(bad)} 件"
    )
    for name, expected, raw_f, ceiled in bad:
        print(
            f"  NG expected={expected} raw={raw_f:g} ceiled100={ceiled:g} | {name!r}"
        )
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
