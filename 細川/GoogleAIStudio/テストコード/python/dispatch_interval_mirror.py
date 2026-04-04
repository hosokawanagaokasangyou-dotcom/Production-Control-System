# -*- coding: utf-8 -*-
"""
配台シミュレーション中の「設備・人の占有区間」ミラー。

タイムライン（timeline_events）と並行して半開区間 [start_dt, end_dt) を保持し、
割付候補が既存占有と重なるかを検査する。ブロックテーブル／TEMP 時間割と同じ
壁時計ベースの占有モデルで、avail_dt / machine_avail_dt との整合を強める。

Excel セルを逐次読み書きする方式は行わない（I/O と COM 負荷のため）。
"""
from __future__ import annotations

from collections import defaultdict
from datetime import datetime


class DispatchIntervalMirror:
    """設備キー・メンバー名ごとの占有区間リスト（重なり検査用）。"""

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
        self, eq_line: str, st: datetime, ed: datetime
    ) -> bool:
        eq = (eq_line or "").strip()
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
        eq_line: str,
        team: tuple,
        st: datetime,
        ed: datetime,
    ) -> bool:
        if self.would_block_equipment(eq_line, st, ed):
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
        eq = str(ev.get("machine") or "").strip()
        if eq:
            self._eq[eq].append((st, ed))
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
