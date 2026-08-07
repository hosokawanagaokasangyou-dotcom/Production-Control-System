# -*- coding: utf-8 -*-
"""CLI for machine calendar JSON store."""
from __future__ import annotations

import json
import os
import sys

_py_here = os.path.dirname(os.path.abspath(__file__))
if _py_here:
    sys.path.insert(0, _py_here)
os.chdir(_py_here)

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def _emit(payload: dict, exit_code: int = 0) -> None:
    json.dump(payload, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    sys.exit(exit_code)


def _load_json_from_argv(start_index: int) -> dict:
    if len(sys.argv) <= start_index:
        raise ValueError("JSON ペイロードがありません")
    if sys.argv[start_index] == "--patch-file" and len(sys.argv) > start_index + 1:
        from pathlib import Path

        text = Path(sys.argv[start_index + 1]).read_text(encoding="utf-8")
    elif sys.argv[start_index].startswith("@") and len(sys.argv[start_index]) > 1:
        from pathlib import Path

        text = Path(sys.argv[start_index][1:]).read_text(encoding="utf-8")
    else:
        text = sys.argv[start_index]
    payload = json.loads(text)
    if not isinstance(payload, dict):
        raise ValueError("JSON はオブジェクトである必要があります")
    return payload


def main() -> int:
    try:
        action = (sys.argv[1] if len(sys.argv) > 1 else "status").strip().lower()
        from datetime import date
        from pathlib import Path

        from planning_core.core.machine_calendar_paths import machine_calendar_data_json_path
        from planning_core.core.machine_calendar_store import (
            apply_machine_calendar_patch,
            build_editor_payload,
            import_from_master_workbook,
            load_machine_calendar_store,
            save_machine_calendar_store,
            store_has_machine_calendar_data,
            validate_store_for_dispatch,
        )
        from planning_core.core.master_data import (
            load_skills_and_needs,
            _master_workbook_path_resolved,
        )

        store = load_machine_calendar_store()
        skills_pack = load_skills_and_needs()
        equipment_list = skills_pack[0]
        master_path = _master_workbook_path_resolved()

        if action == "status":
            jp = machine_calendar_data_json_path()
            _emit(
                {
                    "ok": True,
                    "json_path": str(jp),
                    "json_exists": jp.is_file(),
                    "has_data": store_has_machine_calendar_data(store),
                    "stage2_ready": validate_store_for_dispatch(store),
                    "master_workbook_path": str(master_path),
                    "master_workbook_exists": Path(master_path).is_file(),
                    "meta": store.get("meta", {}),
                    "column_count": len(store.get("columns") or []),
                    "occupancy_slot_count": len(store.get("occupancy") or {}),
                }
            )
            return 0

        if action == "import_from_master":
            result = import_from_master_workbook(store, master_path, equipment_list)
            path = save_machine_calendar_store(
                store,
                history_kind="import_from_master",
                history_label="master 機械カレンダー取込",
            )
            _emit({"ok": True, **result, "json_path": str(path)})
            return 0

        if action == "day_grid":
            day_s = sys.argv[2] if len(sys.argv) > 2 else date.today().isoformat()
            day = date.fromisoformat(day_s)
            _emit(build_editor_payload(store, day, equipment_list))
            return 0

        if action == "merge":
            patch = _load_json_from_argv(2)
            result = apply_machine_calendar_patch(store, patch)
            path = save_machine_calendar_store(
                store,
                history_kind="merge",
                history_label="機械カレンダー保存",
            )
            _emit({"ok": True, **result, "json_path": str(path)})
            return 0

        if action == "save":
            if len(sys.argv) > 2:
                payload = _load_json_from_argv(2)
                store.clear()
                store.update(payload)
            path = save_machine_calendar_store(
                store, history_kind="save_full", history_label="全体保存"
            )
            _emit({"ok": True, "json_path": str(path)})
            return 0

        _emit({"ok": False, "error": f"未知の action: {action}"}, exit_code=1)
        return 1
    except Exception as e:
        _emit({"ok": False, "error": str(e)}, exit_code=1)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
