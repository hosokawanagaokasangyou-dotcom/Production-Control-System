# -*- coding: utf-8 -*-
"""load_attendance_and_analyze: no master.xlsm legacy fallback."""

from __future__ import annotations

import pytest

from planning_core.core.attendance_paths import ENV_ATTENDANCE_JSON


def test_load_attendance_and_analyze_raises_without_json(tmp_path, monkeypatch):
  from planning_core.core.master_data import load_attendance_and_analyze

  jp = tmp_path / "attendance-data.json"
  monkeypatch.setenv(ENV_ATTENDANCE_JSON, str(jp))
  assert not jp.is_file()

  with pytest.raises(RuntimeError, match="attendance-data.json.*フォールバック"):
    load_attendance_and_analyze(["A"])
