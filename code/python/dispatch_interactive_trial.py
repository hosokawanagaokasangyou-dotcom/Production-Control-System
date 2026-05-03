# -*- coding: utf-8 -*-
"""Interactive dispatch trial stub: shortage JSON + export_result_dispatch_from_json."""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
os.chdir(SCRIPT_DIR)


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: dispatch_interactive_trial.py <path-to-result-dispatch.json>", file=sys.stderr)
        return 2
    path = Path(sys.argv[1]).resolve()
    if not path.is_file():
        print(f"not a file: {path}", file=sys.stderr)
        return 1
    try:
        raw = path.read_text(encoding="utf-8")
        json.loads(raw)
    except Exception as e:
        print(f"json read failed: {e}", file=sys.stderr)
        return 1

    shortage_path = path.with_name("dispatch_trial_shortages.json")
    shortages = {
        "format_version": 1,
        "source_json": str(path),
        "note": "stub: extend with planning_core two-phase allocation",
        "op_shortage": [],
        "as_shortage": [],
    }
    try:
        shortage_path.write_text(
            json.dumps(shortages, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
    except Exception as e:
        print(f"shortage json write failed: {e}", file=sys.stderr)
        return 1

    export_script = SCRIPT_DIR / "export_result_dispatch_from_json.py"
    if export_script.is_file():
        py = sys.executable or "python3"
        try:
            subprocess.run(
                [py, str(export_script), str(path)],
                cwd=str(SCRIPT_DIR),
                check=True,
                timeout=600,
            )
        except Exception as e:
            print(f"xlsx export warning: {e}", file=sys.stderr)

    print(str(shortage_path), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
