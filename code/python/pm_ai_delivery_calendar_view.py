# -*- coding: utf-8 -*-
"""Prints one-line JSON for delivery calendar view (JavaFX child process)."""
from __future__ import annotations

import json
import logging
import os
import sys

# python -P / PYTHONSAFEPATH ではスクリプト所在ディレクトリが sys.path に入らない。
# 同梱 embed 実行時は JAVA 側の PYTHONPATH が効かない／限定されることがあるため、段階1/2 と同様に明示する。
_py_here = os.path.dirname(os.path.abspath(__file__))
if _py_here:
    sys.path.insert(0, _py_here)

# INFO on stderr is merged into stdout by Java's ProcessBuilder.redirectErrorStream(true); keep quiet.
logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(message)s")

_DEBUG_LOG = "/mnt/c/工程管理AIプロジェクト_JAVA/.cursor/debug-delcal01.log"


def _agent_debug_ndjson(payload: dict) -> None:
    # #region agent log
    try:
        import time as _t

        payload = dict(payload)
        payload.setdefault("sessionId", "delcal01")
        payload.setdefault("timestamp", int(_t.time() * 1000))
        line = json.dumps(payload, ensure_ascii=False) + "\n"
        with open(_DEBUG_LOG, "a", encoding="utf-8") as _f:
            _f.write(line)
    except Exception:
        pass

    # #endregion


def main() -> int:
    try:
        # #region agent log
        _agent_debug_ndjson(
            {
                "hypothesisId": "H-src",
                "location": "pm_ai_delivery_calendar_view.py:pre-import",
                "message": "import_start",
                "data": {"sys_path0": sys.path[0] if sys.path else ""},
            }
        )
        # #endregion
        import planning_core.delivery_calendar_payload as _dcpmod

        # #region agent log
        _agent_debug_ndjson(
            {
                "hypothesisId": "H-src",
                "location": "pm_ai_delivery_calendar_view.py:post-import",
                "message": "delivery_calendar_payload_module_file",
                "data": {"__file__": getattr(_dcpmod, "__file__", "")},
            }
        )
        # #endregion
        from planning_core.delivery_calendar_payload import build_delivery_calendar_payload

        out = build_delivery_calendar_payload()
        # #region agent log
        _agent_debug_ndjson(
            {
                "hypothesisId": "H-ok",
                "location": "pm_ai_delivery_calendar_view.py:after-build",
                "message": "payload_ok",
                "data": {"ok": bool(out.get("ok")), "exit_if_ok": 0 if out.get("ok") else 2},
            }
        )
        # #endregion
        print(json.dumps(out, ensure_ascii=False), flush=True)
        return 0 if out.get("ok") else 2
    except Exception as e:
        logging.exception("pm_ai_delivery_calendar_view")
        # #region agent log
        _agent_debug_ndjson(
            {
                "hypothesisId": "H-err",
                "location": "pm_ai_delivery_calendar_view.py:except",
                "message": type(e).__name__,
                "data": {"error": str(e)[:2000]},
            }
        )
        # #endregion
        print(
            json.dumps({"ok": False, "error": str(e), "meta": {}}, ensure_ascii=False),
            flush=True,
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
