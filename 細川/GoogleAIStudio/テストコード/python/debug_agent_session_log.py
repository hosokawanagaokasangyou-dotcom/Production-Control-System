# -*- coding: utf-8 -*-
"""Debug session NDJSON append (Cursor debug mode). Do not log secrets."""
from __future__ import annotations

import json
import os
import time

_SESSION = "8b603e"


def _log_path() -> str:
    # .../細川/GoogleAIStudio/テストコード/python/this_file.py -> repo root (----AI-------1)
    here = os.path.dirname(os.path.abspath(__file__))
    root = os.path.abspath(os.path.join(here, "..", "..", "..", ".."))
    return os.path.join(root, "debug-8b603e.log")


def append(
    hypothesis_id: str,
    location: str,
    message: str,
    data: dict | None = None,
    run_id: str = "pre",
) -> None:
    p = _log_path()
    rec: dict = {
        "sessionId": _SESSION,
        "runId": run_id,
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "timestamp": int(time.time() * 1000),
    }
    if data:
        rec["data"] = data
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            f.flush()
    except OSError:
        pass
