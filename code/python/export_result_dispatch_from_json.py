# -*- coding: utf-8 -*-
"""Write dispatch-table xlsx next to the input JSON via planning_core (same layout as stage2)."""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
os.chdir(SCRIPT_DIR)


def main() -> int:
    if len(sys.argv) < 2:
        print(
            "usage: export_result_dispatch_from_json.py <path-to-dispatch-result.json>",
            file=sys.stderr,
        )
        return 2
    path = Path(sys.argv[1]).resolve()
    if not path.is_file():
        print(f"not a file: {path}", file=sys.stderr)
        return 1
    try:
        raw = path.read_text(encoding="utf-8")
        payload = json.loads(raw)
    except Exception as e:
        print(f"json read failed: {e}", file=sys.stderr)
        return 1
    cols = payload.get("columns") or []
    rows = payload.get("rows") or []
    if not cols:
        print("missing columns", file=sys.stderr)
        return 1
    try:
        import pandas as pd

        df = pd.DataFrame(rows)
        for c in cols:
            if c not in df.columns:
                df[c] = None
        df = df[cols]
    except Exception as e:
        print(f"pandas build failed: {e}", file=sys.stderr)
        return 1
    try:
        from planning_core._core import _write_dispatch_table_standalone_xlsx

        target_dir = str(path.parent)
        out = _write_dispatch_table_standalone_xlsx(df, target_dir)
        if out:
            print(out, flush=True)
            return 0
        print("xlsx export returned None", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"export failed: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
