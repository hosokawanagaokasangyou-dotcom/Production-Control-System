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


def main() -> int:
    try:
        from planning_core.delivery_calendar_payload import build_delivery_calendar_payload

        out = build_delivery_calendar_payload()
        print(json.dumps(out, ensure_ascii=False), flush=True)
        return 0 if out.get("ok") else 2
    except Exception as e:
        logging.exception("pm_ai_delivery_calendar_view")
        print(
            json.dumps({"ok": False, "error": str(e), "meta": {}}, ensure_ascii=False),
            flush=True,
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
