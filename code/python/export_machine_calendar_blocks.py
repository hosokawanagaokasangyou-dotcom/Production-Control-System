# -*- coding: utf-8 -*-
"""Emit JSON: equipment_key -> list of ISO dates that have at least one machine-calendar block (occupied slot)."""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
os.chdir(SCRIPT_DIR)

# region agent log
_DEBUG_LOG = Path(__file__).resolve().parent.parent.parent / ".cursor" / "debug-ecf65d.log"


def _agent_dbg(hypothesis_id: str, message: str, data: dict) -> None:
    import json
    import time

    try:
        line = (
            json.dumps(
                {
                    "sessionId": "ecf65d",
                    "hypothesisId": hypothesis_id,
                    "location": "export_machine_calendar_blocks.py",
                    "message": message,
                    "data": data,
                    "timestamp": int(time.time() * 1000),
                },
                ensure_ascii=False,
            )
            + "\n"
        )
        _DEBUG_LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(_DEBUG_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


# endregion


def main() -> int:
    if len(sys.argv) < 2:
        print("{}", flush=True)
        return 0
    master = os.path.abspath(sys.argv[1])
    if not os.path.isfile(master):
        print(json.dumps({"error": "file_not_found", "path": master}, ensure_ascii=False), flush=True)
        return 0
    os.environ["PM_AI_MASTER_WORKBOOK"] = master
    try:
        import workbook_env_bootstrap as wbe

        wbe.apply_from_task_input_workbook()
    except Exception:
        pass
    # region agent log
    _pc_init = SCRIPT_DIR / "planning_core" / "__init__.py"
    _agent_dbg(
        "H3",
        "pre-import state",
        {
            "cwd": os.getcwd(),
            "scriptDir": str(SCRIPT_DIR),
            "master": master,
            "planningCoreInitExists": _pc_init.is_file(),
            "sysPathHead": sys.path[:16],
        },
    )
    # endregion
    try:
        from planning_core._core import (
            load_machine_calendar_occupancy_blocks,
            load_skills_and_needs,
        )

        tup = load_skills_and_needs()
        equipment_list = tup[2]
        occ = load_machine_calendar_occupancy_blocks(master, equipment_list)
    except Exception as e:
        # region agent log
        _agent_dbg(
            "H2",
            "planning_core load failed",
            {"excType": type(e).__name__, "exc": str(e)},
        )
        # endregion
        print(json.dumps({"error": str(e), "blocks": {}}, ensure_ascii=False), flush=True)
        return 0
    blocks: dict[str, list[str]] = {}
    for d, eqm in occ.items():
        ds = d.isoformat()
        for eq, ivs in eqm.items():
            if ivs:
                ek = str(eq).strip()
                blocks.setdefault(ek, []).append(ds)
    for k in list(blocks.keys()):
        blocks[k] = sorted(set(blocks[k]))
    # region agent log
    _agent_dbg(
        "H5",
        "export ok",
        {"blockEquipmentKeys": len(blocks), "totalDates": sum(len(v) for v in blocks.values())},
    )
    # endregion
    print(json.dumps({"blocks": blocks}, ensure_ascii=False), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
