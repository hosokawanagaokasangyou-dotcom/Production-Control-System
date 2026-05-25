# -*- coding: utf-8 -*-
"""B-6.1 / B-4.1: 前工程完走後に SEC ゲートが残ロールを塞がないこと。"""

from planning_core import _core as pc


def _y64_tasks(*, conn_rem: float, sec_rem: float, conn_done: float, sec_done: float):
  conn = {
      "task_id": "Y6-4",
      "machine": "接続",
      "machine_name": "熱融着機　湖南",
      "remaining_units": conn_rem,
      "initial_remaining_units": 8,
      "process_content_tokens": ["接続", "SEC"],
  }
  conn["remaining_units"] = conn_rem
  # simulate done via initial - remaining
  sec = {
      "task_id": "Y6-4",
      "machine": "SEC",
      "machine_name": "SEC機　湖南",
      "remaining_units": sec_rem,
      "initial_remaining_units": 8,
      "process_content_tokens": ["接続", "SEC"],
  }
  tq = [conn, sec]
  # Adjust to emulate partial dispatch counts used by gap helpers
  conn["initial_remaining_units"] = conn_done + conn_rem
  sec["initial_remaining_units"] = sec_done + sec_rem
  return tq, sec


def test_b61_does_not_block_sec_when_connection_complete():
    tq, sec = _y64_tasks(conn_rem=0, sec_rem=3, conn_done=8, sec_done=5)
    assert not pc._b61_sec_blocked_by_connection_min_rolls(sec, tq)


def test_b61_blocks_sec_when_connection_in_progress_and_gap_below_threshold():
    tq, sec = _y64_tasks(conn_rem=3, sec_rem=8, conn_done=2, sec_done=0)
    assert pc._b61_sec_blocked_by_connection_min_rolls(sec, tq)


def test_b61_allows_sec_when_connection_in_progress_and_gap_at_threshold():
    tq, sec = _y64_tasks(conn_rem=3, sec_rem=8, conn_done=5, sec_done=0)
    assert not pc._b61_sec_blocked_by_connection_min_rolls(sec, tq)
