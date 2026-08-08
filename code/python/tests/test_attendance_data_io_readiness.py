# -*- coding: utf-8 -*-
"""attendance_data_io readiness はマスタブック未設定でも名簿ベースで動くこと。"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from planning_core.core.attendance_paths import ENV_ATTENDANCE_JSON
from planning_core.core.attendance_store import (
    apply_company_calendar_to_members,
    empty_store,
    save_attendance_store,
)


def test_readiness_cli_without_master_workbook(tmp_path, monkeypatch):
    monkeypatch.delenv("PM_AI_MASTER_WORKBOOK", raising=False)
    att = tmp_path / "attendance-data.json"
    monkeypatch.setenv(ENV_ATTENDANCE_JSON, str(att))
    store = empty_store(2026)
    store["meta"]["company_calendar_revision"] = 1
    store["member_roster"] = [{"name": "A", "primary_role": ""}]
    store["company_calendar"]["days"]["2026-08-06"] = {"kind": "public", "label": "休"}
    members = ["A"]
    apply_company_calendar_to_members(store, members, 2026, 8)
    save_attendance_store(store, att)

    script = Path(__file__).resolve().parents[1] / "attendance_data_io.py"
    proc = subprocess.run(
        [sys.executable, str(script), "readiness", "2026", "8"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        cwd=str(script.parent),
        env={
            **{k: v for k, v in __import__("os").environ.items() if k != "PM_AI_MASTER_WORKBOOK"},
            ENV_ATTENDANCE_JSON: str(att),
            "PM_AI_SKIP_ERROR_PAUSE": "1",
        },
        check=False,
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr
    lines = [ln.strip() for ln in proc.stdout.splitlines() if ln.strip().startswith("{")]
    payload = json.loads(lines[-1])
    assert payload["ok"] is True
    assert payload["json_path"]
    assert payload["member_cells_in_month"] == 31
