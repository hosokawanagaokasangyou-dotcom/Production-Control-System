# -*- coding: utf-8 -*-
"""
配台シミュレーション中の「設備・人の占有区間」ミラー。

タイムライン（timeline_events）と並行して半開区間 [start_dt, end_dt) を保持し、
割付候補が既存占有と重なるかを検査する。ブロックテーブル／TEMP 時間割と同じ
壁時計ベースの占有モデルで、avail_dt / machine_avail_dt との整合を強める。

Excel セルを逐次読み書きする方式は行わない（I/O と COM 負荷のため）。

設備側のキーは **機械名ベースの占有キー**（イベントの machine_occupancy_key、
無い場合は machine 列から推定）とする。
"""
from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from datetime import datetime


def _mirror_normalize_occupancy_key(val) -> str:
    if val is None or val == "":
        return ""
    t = unicodedata.normalize("NFKC", str(val))
    t = t.replace("\u00a0", " ").replace("\u3000", " ")
    t = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _mirror_occupancy_key_from_event(ev: dict) -> str:
    k = str(ev.get("machine_occupancy_key") or "").strip()
    if k:
        return _mirror_normalize_occupancy_key(k)
    mach = str(ev.get("machine") or "").strip()
    if "+" in mach:
        return _mirror_normalize_occupancy_key(mach.split("+", 1)[1])
    return _mirror_normalize_occupancy_key(mach)


class DispatchIntervalMirror:
    """設備占有キー・メンバー名ごとの占有区間リスト（重なり検査用）。"""

    __slots__ = ("_eq", "_mem")

    def __init__(self) -> None:
        self._eq: dict[str, list[tuple[datetime, datetime]]] = defaultdict(list)
        self._mem: dict[str, list[tuple[datetime, datetime]]] = defaultdict(list)

    def clear(self) -> None:
        self._eq.clear()
        self._mem.clear()

    @staticmethod
    def _overlaps(
        st1: datetime, ed1: datetime, st2: datetime, ed2: datetime
    ) -> bool:
        return st1 < ed2 and st2 < ed1

    def would_block_equipment(
        self, machine_occupancy_key: str, st: datetime, ed: datetime
    ) -> bool:
        eq = _mirror_normalize_occupancy_key(machine_occupancy_key)
        if not eq or not isinstance(st, datetime) or not isinstance(ed, datetime):
            return False
        if ed <= st:
            return False
        for a, b in self._eq.get(eq, ()):
            if self._overlaps(st, ed, a, b):
                return True
        return False

    def would_block_member(self, name: str, st: datetime, ed: datetime) -> bool:
        n = (name or "").strip()
        if not n or not isinstance(st, datetime) or not isinstance(ed, datetime):
            return False
        if ed <= st:
            return False
        for a, b in self._mem.get(n, ()):
            if self._overlaps(st, ed, a, b):
                return True
        return False

    def would_block_roll(
        self,
        machine_occupancy_key: str,
        team: tuple,
        st: datetime,
        ed: datetime,
    ) -> bool:
        if self.would_block_equipment(machine_occupancy_key, st, ed):
            return True
        for m in team:
            if self.would_block_member(str(m), st, ed):
                return True
        return False

    def register_from_event(self, ev: dict) -> None:
        st = ev.get("start_dt")
        ed = ev.get("end_dt")
        if not isinstance(st, datetime) or not isinstance(ed, datetime):
            return
        if ed <= st:
            return
        occ = _mirror_occupancy_key_from_event(ev)
        if occ:
            self._eq[occ].append((st, ed))
        op = str(ev.get("op") or "").strip()
        if op:
            self._mem[op].append((st, ed))
        for s in str(ev.get("sub") or "").split(","):
            s = s.strip()
            if s:
                self._mem[s].append((st, ed))

    def rebuild_from_timeline(self, timeline_events: list) -> None:
        self.clear()
        for ev in timeline_events:
            self.register_from_event(ev)
