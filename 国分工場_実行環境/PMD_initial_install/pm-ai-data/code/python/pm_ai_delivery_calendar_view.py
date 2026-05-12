# -*- coding: utf-8 -*-
"""Prints one-line JSON for delivery calendar view (JavaFX child process)."""

from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path

# python -P / PYTHONSAFEPATH ではスクリプト所在ディレクトリが sys.path に入らない。
# 同梱 embed 実行時は JAVA 側の PYTHONPATH が効かない／限定されることがあるため、段階1/2 と同様に明示する。
_py_here = os.path.dirname(os.path.abspath(__file__))
if _py_here:
    sys.path.insert(0, _py_here)

# INFO on stderr is merged into stdout by Java's ProcessBuilder.redirectErrorStream(true); keep quiet.
logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(message)s")


def main() -> int:
    try:
        _py_dir = Path(__file__).resolve().parent
        _dcp_source = _py_dir / "planning_core" / "delivery_calendar_payload.py"
        try:
            _dcp_source.read_bytes().decode("utf-8")
        except OSError as oe:
            msg = f"delivery_calendar_payload.py を読めません ({_dcp_source}): {oe}"
            print(
                json.dumps(
                    {
                        "ok": False,
                        "error": msg,
                        "meta": {"badSourcePath": str(_dcp_source)},
                    },
                    ensure_ascii=False,
                ),
                flush=True,
            )
            return 1
        except UnicodeDecodeError as ude:
            msg = (
                "delivery_calendar_payload.py が UTF-8 として不正です（別フォルダの古いコピーを実行している可能性）。"
                f" path={_dcp_source} ; {ude}"
            )
            print(
                json.dumps(
                    {
                        "ok": False,
                        "error": msg,
                        "meta": {"badSourcePath": str(_dcp_source)},
                    },
                    ensure_ascii=False,
                ),
                flush=True,
            )
            return 1
        from planning_core.delivery_calendar_payload import build_delivery_calendar_payload

        out = build_delivery_calendar_payload()
        print(json.dumps(out, ensure_ascii=False), flush=True)
        return 0 if out.get("ok") else 2
    except Exception as e:
        logging.exception("pm_ai_delivery_calendar_view")
        meta_out: dict = {}
        try:
            _pd = Path(__file__).resolve().parent
            meta_out["badSourcePath"] = str(_pd / "planning_core" / "delivery_calendar_payload.py")
        except Exception:
            pass
        try:
            import importlib.util as _iu

            _spec = _iu.find_spec("planning_core.delivery_calendar_payload")
            if _spec is not None:
                _org = getattr(_spec, "origin", None)
                if _org:
                    meta_out["importOrigin"] = str(_org)
                _loc = getattr(_spec, "submodule_search_locations", None)
                if _loc:
                    meta_out["packageSearchLocations"] = [str(x) for x in _loc]
        except Exception:
            pass
        print(
            json.dumps({"ok": False, "error": str(e), "meta": meta_out}, ensure_ascii=False),
            flush=True,
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
