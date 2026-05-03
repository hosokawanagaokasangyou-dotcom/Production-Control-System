# -*- coding: utf-8 -*-
"""Master workbook read summary as JSON for JavaFX (same resolution as planning_core)."""

from __future__ import annotations

import json
import os
from typing import Any

_CALENDAR_IN_NAME = "\u30ab\u30ec\u30f3\u30c0\u30fc"  # "?J?????_?["


def _fmt_time(t) -> str | None:
    if t is None:
        return None
    return t.strftime("%H:%M")


def _skills_member_names_quick(path: str) -> list[str]:
    import pandas as pd

    try:
        raw = pd.read_excel(path, sheet_name="skills", header=None)
    except Exception:
        return []
    members: list[str] = []
    if raw.shape[0] < 3 or raw.shape[1] < 2:
        return members
    ne = 0
    for c in range(1, raw.shape[1]):
        p, m = raw.iat[0, c], raw.iat[1, c]
        if pd.isna(p) or pd.isna(m):
            continue
        ps, ms = str(p).strip(), str(m).strip()
        if ps and ms and ps.lower() != "nan" and ms.lower() != "nan":
            ne += 1
    if ne == 0:
        try:
            sdf = pd.read_excel(path, sheet_name="skills")
            sdf.columns = sdf.columns.str.strip()
            cols = [str(x).strip() for x in sdf.columns if not str(x).startswith("Unnamed")]
            mc = None
            for col in cols:
                if col in (
                    "\u30e1\u30f3\u30d0\u30fc",
                    "\u62c5\u5f53\u8005",
                    "\u4e26\u3073",
                    "\u4f5c\u696d\u8005",
                ):
                    mc = col
                    break
            if mc is None and cols:
                mc = cols[0]
            if mc:
                for _, row in sdf.iterrows():
                    mn = str(row.get(mc, "")).strip()
                    if mn and mn.lower() != "nan":
                        members.append(mn)
        except Exception:
            pass
        return members
    for r in range(2, raw.shape[0]):
        mn = raw.iat[r, 0]
        if mn is None or (isinstance(mn, float) and pd.isna(mn)):
            continue
        mname = str(mn).strip()
        if mname and mname.lower() not in ("nan", "none", "null"):
            members.append(mname)
    return members


def _count_attendance_sheets(sheet_names: list[str], members: set[str]) -> int:
    skip = {"skills", "need", "tasks"}
    n = 0
    for sn in sheet_names:
        if _CALENDAR_IN_NAME in sn:
            continue
        if sn.lower() in skip:
            continue
        if sn.strip() in members:
            n += 1
    return n


def build_master_read_summary_dict() -> dict[str, Any]:
    import planning_core._core as c

    warnings: list[str] = []
    path = c._master_workbook_path_resolved()
    cwd = os.getcwd()
    pm_override = (os.environ.get("PM_AI_MASTER_WORKBOOK") or "").strip()
    master_file_env = c.master_workbook_filename()
    use_speed_raw = (os.environ.get("MASTER_USE_SPEED_SHEET") or "1").strip()

    file_exists = bool(path and os.path.isfile(path))
    if not file_exists:
        warnings.append(
            "Master file missing at path resolved by planning_core."
        )

    sheet_names: list[str] = []
    openpyxl_skip = False
    if path:
        openpyxl_skip = bool(c._workbook_should_skip_openpyxl_io(path))
        if openpyxl_skip:
            warnings.append(
                "Workbook marked incompatible with openpyxl I/O (e.g. certain sheet names)."
            )
        if file_exists:
            try:
                import pandas as pd

                with pd.ExcelFile(path) as xls:
                    sheet_names = list(xls.sheet_names)
            except Exception as e:
                warnings.append(f"Could not list sheets: {e}")
                ox = c._ooxml_workbook_sheet_names(path)
                if ox:
                    sheet_names = list(ox)

    main_sheet: str | None = None
    a12 = b12 = a15 = b15 = None
    factory_effective = False
    regular_effective = False
    if file_exists and not openpyxl_skip:
        try:
            st, et = c._read_master_main_factory_operating_times(path)
            if st is not None and et is not None:
                a12, b12 = st, et
                factory_effective = True
        except Exception as e:
            warnings.append(f"Factory hours A12/B12: {e}")
        try:
            st2, et2 = c._read_master_main_regular_shift_times(path)
            if st2 is not None and et2 is not None:
                a15, b15 = st2, et2
                regular_effective = True
        except Exception as e:
            warnings.append(f"Regular shift A15/B15: {e}")
        try:
            main_sheet = c._pick_master_main_sheet_name(sheet_names)
            if main_sheet is None:
                warnings.append("Could not resolve main settings sheet name.")
        except Exception as e:
            warnings.append(f"Main sheet name: {e}")

    speed_enabled = bool(c._master_speed_sheet_apply_enabled())
    speed_sheet = c.MASTER_SHEET_SPEED
    speed_first_col = int(c._master_speed_first_excel_col_1based())
    speed_count = 0
    if file_exists:
        try:
            lu = c._load_master_speed_lookup_from_master_workbook()
            speed_count = len(lu)
        except Exception as e:
            warnings.append(f"speed lookup count: {e}")

    members = _skills_member_names_quick(path) if file_exists else []
    member_set = set(members)
    attend_count = _count_attendance_sheets(sheet_names, member_set)

    sk_machine = c.SHEET_MACHINE_CALENDAR
    sk_combo = c.MASTER_SHEET_TEAM_COMBINATIONS
    sk_startup = c.SHEET_MACHINE_DAILY_STARTUP
    key_sheets = [
        ("skills", "skills"),
        ("need", "need"),
        ("machine_calendar", sk_machine),
        ("team_combinations", sk_combo),
        ("speed", speed_sheet),
        ("machine_daily_startup", sk_startup),
    ]
    sheet_rows: list[dict[str, Any]] = []
    for label_key, sn in key_sheets:
        present = sn in sheet_names
        note = "" if present else "missing"
        sheet_rows.append({"key": label_key, "sheet_name": sn, "present": present, "note": note})

    ok = file_exists and ("skills" in sheet_names) and ("need" in sheet_names)
    if not members and "skills" in sheet_names:
        warnings.append("No member names parsed from skills sheet (check format).")

    return {
        "ok": ok,
        "warnings": warnings,
        "resolved_path": path or "",
        "file_exists": file_exists,
        "cwd": cwd,
        "master_workbook_file_env": master_file_env,
        "pm_ai_master_workbook_env": pm_override,
        "master_use_speed_sheet_env": use_speed_raw,
        "speed": {
            "enabled": speed_enabled,
            "sheet_name": speed_sheet,
            "first_data_col_1based": speed_first_col,
            "lookup_entry_count": speed_count,
        },
        "main_sheet": {
            "resolved_name": main_sheet,
            "factory_operating": {
                "a12": _fmt_time(a12),
                "b12": _fmt_time(b12),
                "effective": factory_effective,
            },
            "regular_shift": {
                "a15": _fmt_time(a15),
                "b15": _fmt_time(b15),
                "effective": regular_effective,
            },
        },
        "sheet_checks": sheet_rows,
        "attendance": {
            "skills_member_count": len(members),
            "attendance_sheets_matched": attend_count,
            "note": "Sheets whose name matches a skills member (attendance candidates).",
        },
        "openpyxl_skip": openpyxl_skip,
        "all_sheet_names": sheet_names,
    }


def main() -> None:
    import sys

    data = build_master_read_summary_dict()
    json.dump(data, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    sys.exit(0 if data.get("ok") else 1)


if __name__ == "__main__":
    main()
