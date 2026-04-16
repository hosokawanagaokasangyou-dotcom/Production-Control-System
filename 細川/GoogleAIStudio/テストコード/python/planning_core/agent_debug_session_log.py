# -*- coding: utf-8 -*-
"""デバッグ用 NDJSON（セッション）。本番では未使用可。"""
from __future__ import annotations

import json
import os
import time


def _log_path() -> str:
    here = os.path.abspath(os.path.dirname(__file__))
    # planning_core -> python -> テストコード -> GoogleAIStudio -> 細川 -> リポジトリ直下
    root = os.path.abspath(os.path.join(here, "..", "..", "..", "..", ".."))
    return os.path.join(root, "debug-938d25.log")


def agent_ndjson_log(
    *,
    hypothesis_id: str,
    location: str,
    message: str,
    data: dict | None = None,
    run_id: str = "pre",
) -> None:
    # #region agent log
    payload = {
        "sessionId": "938d25",
        "runId": run_id,
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data or {},
        "timestamp": int(time.time() * 1000),
    }
    try:
        with open(_log_path(), "a", encoding="utf-8", newline="\n") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
