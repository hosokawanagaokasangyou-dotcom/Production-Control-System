# -*- coding: utf-8 -*-
"""結果_配台表 JSON: 0 行でも空 JSON を出力する。"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from planning_core.core.roll_pipeline import _write_dispatch_table_standalone_json


def test_write_dispatch_table_standalone_json_writes_empty_payload(tmp_path: Path):
    out = _write_dispatch_table_standalone_json(pd.DataFrame(), str(tmp_path))
    assert out is not None
    path = Path(out)
    assert path.is_file()
    payload = json.loads(path.read_text(encoding="utf-8"))
    assert payload["row_count"] == 0
    assert payload["rows"] == []
    assert "依頼NO" in payload["columns"]
    assert "当日配台数量" in payload["columns"]
