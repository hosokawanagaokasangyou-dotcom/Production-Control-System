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
            apply_holidays_to_fiscal_year,
            apply_member_attendance_patch,
            apply_national_holidays_to_company_calendar,
            build_company_calendar_payload,
            build_company_calendar_payload_fiscal,
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
            from planning_core.core.attendance_readiness import build_attendance_readiness

            jp = attendance_data_json_path()
            readiness = build_attendance_readiness(store, members)
            _emit(
                {
                    "ok": True,
                    "json_path": str(jp),
                    "json_exists": jp.is_file(),
                    "view_xlsx_path": str(attendance_view_xlsx_path()),
                    "meta": store.get("meta", {}),
                    **{k: v for k, v in readiness.items() if k not in ("ok", "format_version")},
                }
            )
            return 0

        if action == "readiness":
            from planning_core.core.attendance_readiness import build_attendance_readiness

            year = int(sys.argv[2]) if len(sys.argv) > 2 else __import__("datetime").date.today().year
            month = int(sys.argv[3]) if len(sys.argv) > 3 else __import__("datetime").date.today().month
            _emit(build_attendance_readiness(store, members, year, month))
            return 0

        if action == "company_calendar":
            fiscal_year = int(sys.argv[2]) if len(sys.argv) > 2 else __import__("datetime").date.today().year
            start_month = int(sys.argv[3]) if len(sys.argv) > 3 else 4
            start_day = int(sys.argv[4]) if len(sys.argv) > 4 else 1
            _emit(
                build_company_calendar_payload_fiscal(
                    store, fiscal_year, start_month, start_day
                )
            )
            return 0

        if action == "fetch_holidays_fiscal":
            fiscal_year = int(sys.argv[2]) if len(sys.argv) > 2 else __import__("datetime").date.today().year
            start_month = int(sys.argv[3]) if len(sys.argv) > 3 else 4
            start_day = int(sys.argv[4]) if len(sys.argv) > 4 else 1
            include_weekends = "--weekends" in sys.argv
            overwrite = "--overwrite" in sys.argv
            result = apply_holidays_to_fiscal_year(
                store,
                fiscal_year,
                start_month=start_month,
                start_day=start_day,
                overwrite=overwrite,
                include_weekends=include_weekends,
                force_online=True,
            )
            save_attendance_store(
                store, history_kind="fetch_holidays_fiscal", history_label="祝日取得"
            )
            generate_attendance_view_xlsx(store)
            _emit({"ok": True, **result})
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
            save_attendance_store(store, history_kind="fetch_holidays", history_label="祝日取得")
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
            save_attendance_store(
                store, history_kind="sync_members", history_label="会社カレンダー同期"
            )
            generate_attendance_view_xlsx(store)
            _emit({"ok": True, **result})
            return 0

        if action == "save":
            if len(sys.argv) > 2:
                payload = _load_json_from_argv(2)
                store.clear()
                store.update(payload)
            path = save_attendance_store(
                store, history_kind="save_full", history_label="全体保存"
            )
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
            if "fiscal_start_month" in patch:
                meta["fiscal_start_month"] = patch["fiscal_start_month"]
            if "fiscal_start_day" in patch:
                meta["fiscal_start_day"] = patch["fiscal_start_day"]
            meta["company_calendar_revision"] = int(meta.get("company_calendar_revision") or 0) + 1
            path = save_attendance_store(
                store,
                history_kind="merge_company_calendar",
                history_label="会社カレンダー保存",
            )
            view = generate_attendance_view_xlsx(store)
            _emit(
                {
                    "ok": True,
                    "json_path": str(path),
                    "view_xlsx": str(view),
                    "revision": meta.get("company_calendar_revision"),
                }
            )
            return 0

        if action == "merge_member_attendance":
            patch = _load_json_from_argv(2)
            result = apply_member_attendance_patch(store, patch)
            path = save_attendance_store(
                store,
                history_kind="merge_member_attendance",
                history_label="メンバー勤怠保存",
            )
            view = generate_attendance_view_xlsx(store)
            _emit({"ok": True, **result, "json_path": str(path), "view_xlsx": str(view)})
            return 0

        if action == "export_master":
            master = _master_workbook_path_resolved()
            result = export_attendance_to_master_new_sheets(store, master)
            save_attendance_store(
                store,
                history_kind="export_master",
                history_label="master出力",
            )
            _emit(result)
            return 0

        if action == "history_list":
            from planning_core.core.attendance_history_store import list_attendance_history

            data = list_attendance_history()
            _emit({"ok": True, **data})
            return 0

        if action == "history_restore":
            entry_id = sys.argv[2] if len(sys.argv) > 2 else ""
            if not entry_id.strip():
                raise ValueError("復元する世代 id が必要です")
            from planning_core.core.attendance_history_store import restore_attendance_snapshot

            jp = restore_attendance_snapshot(entry_id.strip())
            store.clear()
            store.update(load_attendance_store(jp))
            view = generate_attendance_view_xlsx(store)
            _emit(
                {
                    "ok": True,
                    "restored_id": entry_id.strip(),
                    "json_path": str(jp),
                    "view_xlsx": str(view),
                }
            )
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
            save_attendance_store(
                store,
                history_kind="set_company_day",
                history_label="会社カレンダー1日編集",
            )
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
