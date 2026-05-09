# -*- coding: utf-8 -*-
"""Prints one-line JSON for delivery calendar view (JavaFX child process)."""
from __future__ import annotations

import json
import logging
import sys

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")


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
