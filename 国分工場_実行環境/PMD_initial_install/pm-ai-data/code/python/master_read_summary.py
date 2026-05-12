# -*- coding: utf-8 -*-
"""CLI: print master workbook read summary as one JSON line (for JavaFX tab)."""
import json
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def _emit_minimal_probe_json(
    warnings: list[str], exit_code: int, extra: dict | None = None
) -> None:
    """Stdout に常に1行 JSON を出す（JavaFX は stderr を stdout にマージして読むため）。"""
    payload: dict = {
        "ok": False,
        "warnings": warnings,
        "resolved_path": "",
        "file_exists": False,
        "cwd": os.getcwd(),
        "master_workbook_file_env": "",
        "pm_ai_master_workbook_env": (os.environ.get("PM_AI_MASTER_WORKBOOK") or "").strip(),
        "master_use_speed_sheet_env": (os.environ.get("MASTER_USE_SPEED_SHEET") or "1").strip(),
        "speed": {
            "enabled": False,
            "sheet_name": "",
            "first_data_col_1based": 0,
            "lookup_entry_count": 0,
        },
        "main_sheet": {
            "resolved_name": None,
            "factory_operating": {"a12": None, "b12": None, "effective": False},
            "regular_shift": {"a15": None, "b15": None, "effective": False},
        },
        "sheet_checks": [],
        "attendance": {
            "skills_member_count": 0,
            "attendance_sheets_matched": 0,
            "note": "probe error (see warnings)",
        },
        "openpyxl_skip": False,
        "all_sheet_names": [],
    }
    if extra:
        payload.update(extra)
    json.dump(payload, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    sys.exit(exit_code)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("-h", "--help"):
        print("usage: master_read_summary.py", file=sys.stderr)
        sys.exit(0)

    try:
        from planning_core.master_read_ui import main as _main
    except SystemExit as e:
        code = e.code if isinstance(e.code, int) else 1
        _emit_minimal_probe_json(
            [
                "planning_core の import 中に終了しました（多くの場合 Python 3.14+ が必要、"
                "またはブートストラップ失敗）。stderr の [planning_core] 行と環境タブの PM_AI_PYTHON を確認してください。"
            ],
            code,
            {"probe_error": "system_exit_on_import", "system_exit_code": code},
        )
    except Exception as e:
        import traceback

        tb = traceback.format_exc()
        detail = (tb if len(tb) < 3000 else tb[:3000] + "\n...(trunc)")
        _emit_minimal_probe_json(
            [f"import 失敗: {e}"],
            1,
            {"probe_error": "import_exception", "traceback": detail},
        )

    try:
        _main()
    except SystemExit:
        raise
    except Exception as e:
        import traceback

        tb = traceback.format_exc()
        detail = (tb if len(tb) < 3000 else tb[:3000] + "\n...(trunc)")
        _emit_minimal_probe_json(
            [f"master_read 実行中の例外: {e}"],
            1,
            {"probe_error": "main_exception", "traceback": detail},
        )
