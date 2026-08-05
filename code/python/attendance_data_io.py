# -*- coding: utf-8 -*-
"""CLI for attendance store: load / save / holidays / master export / view xlsx."""
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
    """Load JSON from --patch-file, @path, or inline argv (inline is legacy)."""
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
        from planning_core.core.attendance_store import (
            apply_company_calendar_to_members,
            apply_national_holidays_to_company_calendar,
            build_company_calendar_payload,
            build_editor_payload,
            export_attendance_to_master_new_sheets,
            generate_attendance_view_xlsx,
            load_attendance_store,
            save_attendance_store,
        )
        from planning_core.core.master_data import load_skills_and_needs, _master_workbook_path_resolved

        store = load_attendance_store()
        members = load_skills_and_needs()[1]

        if action == "status":
            from planning_core.core.attendance_paths import (
                attendance_data_json_path,
                attendance_view_xlsx_path,
            )

            jp = attendance_data_json_path()
            _emit(
                {
                    "ok": True,
                    "json_path": str(jp),
                    "json_exists": jp.is_file(),
                    "view_xlsx_path": str(attendance_view_xlsx_path()),
                    "meta": store.get("meta", {}),
                }
            )
            return 0

        if action == "company_calendar":
            year = int(sys.argv[2]) if len(sys.argv) > 2 else __import__("datetime").date.today().year
            _emit(build_company_calendar_payload(store, year))
            return 0

        if action == "member_grid":
            year = int(sys.argv[2]) if len(sys.argv) > 2 else __import__("datetime").date.today().year
            month = int(sys.argv[3]) if len(sys.argv) > 3 else __import__("datetime").date.today().month
            _emit(build_editor_payload(store, members, year, month))
            return 0

        if action == "fetch_holidays":
            year = int(sys.argv[2]) if len(sys.argv) > 2 else __import__("datetime").date.today().year
            include_weekends = "--weekends" in sys.argv
            overwrite = "--overwrite" in sys.argv
            result = apply_national_holidays_to_company_calendar(
                store,
                year,
                overwrite=overwrite,
                include_weekends=include_weekends,
                force_online=True,
            )
            save_attendance_store(store)
            generate_attendance_view_xlsx(store)
            _emit({"ok": True, **result})
            return 0

        if action == "sync_members":
            year = int(sys.argv[2]) if len(sys.argv) > 2 else __import__("datetime").date.today().year
            month = int(sys.argv[3]) if len(sys.argv) > 3 else __import__("datetime").date.today().month
            only_unedited = "--all" not in sys.argv
            result = apply_company_calendar_to_members(
                store, members, year, month, only_unedited=only_unedited
            )
            save_attendance_store(store)
            generate_attendance_view_xlsx(store)
            _emit({"ok": True, **result})
            return 0

        if action == "save":
            if len(sys.argv) > 2:
                payload = _load_json_from_argv(2)
                store.clear()
                store.update(payload)
            path = save_attendance_store(store)
            view = generate_attendance_view_xlsx(store)
            _emit({"ok": True, "json_path": str(path), "view_xlsx": str(view)})
            return 0

        if action == "merge_company_calendar":
            patch = _load_json_from_argv(2)
            cc = store.setdefault("company_calendar", {})
            if "year" in patch:
                cc["year"] = patch["year"]
            days = cc.setdefault("days", {})
            for k, v in (patch.get("days") or {}).items():
                days[k] = v
            meta = store.setdefault("meta", {})
            meta["company_calendar_revision"] = int(meta.get("company_calendar_revision") or 0) + 1
            path = save_attendance_store(store)
            view = generate_attendance_view_xlsx(store)
            _emit({"ok": True, "json_path": str(path), "view_xlsx": str(view)})
            return 0

        if action == "export_master":
            master = _master_workbook_path_resolved()
            result = export_attendance_to_master_new_sheets(store, master)
            save_attendance_store(store)
            _emit(result)
            return 0

        if action == "set_company_day":
            # argv: set_company_day YYYY-MM-DD kind [label]
            d_key = sys.argv[2]
            kind = sys.argv[3]
            label = sys.argv[4] if len(sys.argv) > 4 else ""
            days = store.setdefault("company_calendar", {}).setdefault("days", {})
            days[d_key] = {"kind": kind, "label": label, "manual_edit": True}
            store["meta"]["company_calendar_revision"] = int(
                store["meta"].get("company_calendar_revision") or 0
            ) + 1
            save_attendance_store(store)
            _emit({"ok": True, "date": d_key, "kind": kind})
            return 0

        _emit({"ok": False, "error": f"unknown action: {action}"}, 1)
        return 1
    except Exception as e:
        import traceback

        _emit(
            {
                "ok": False,
                "error": str(e),
                "traceback": traceback.format_exc()[:3000],
            },
            1,
        )
        return 1


if __name__ == "__main__":
    try:
        import workbook_env_bootstrap as _wbe_exit

        sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
    except ImportError:
        sys.exit(main())
