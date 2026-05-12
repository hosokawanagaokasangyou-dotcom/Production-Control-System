#!/usr/bin/env python3
"""???????????? .bas ? CP932 ??????????? Python ??????????"""
from __future__ import annotations

from pathlib import Path

# tools/restructure/this.py -> repo root is parents[2]
ROOT = Path(__file__).resolve().parents[2]
# ASCII + escapes only (avoid editor encoding issues)
VBA_DIR = ROOT / "code" / "\u53c2\u7167\u7528" / "\u30e2\u30b8\u30e5\u30fc\u30eb\u5206\u5272\u30d0\u30fc\u30b8\u30e7\u30f3"


def write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.encode("cp932"))


def main() -> None:
    if not VBA_DIR.is_dir():
        raise SystemExit(f"missing: {VBA_DIR}")

    # --- ??????.bas ---
    p = VBA_DIR / "\u30d5\u30a9\u30f3\u30c8\u7ba1\u7406.bas"
    t = p.read_bytes().decode("cp932")
    for old, new in (
        ("python\\apply_", "\u53c2\u7167\u7528\\python\\apply_"),
        ("python\\dedupe_", "\u53c2\u7167\u7528\\python\\dedupe_"),
    ):
        t = t.replace(old, new)
    write_cp932(p, t)
    print("updated:", p.relative_to(ROOT))

    # --- ?????????.bas ---
    p = VBA_DIR / "\u8a08\u753b\u5b9f\u7e3e\u6bd4\u8f03\u30ac\u30f3\u30c8.bas"
    t = p.read_bytes().decode("cp932")
    t = t.replace(
        "python\\plan_compare_gantt_from_snapshot.py",
        "\u53c2\u7167\u7528\\python\\plan_compare_gantt_from_snapshot.py",
    )
    write_cp932(p, t)
    print("updated:", p.relative_to(ROOT))

    # --- ??????.bas ---
    p = VBA_DIR / "\u696d\u52d9\u30ed\u30b8\u30c3\u30af.bas"
    t = p.read_bytes().decode("cp932")
    t = t.replace(
        "python\\plan_refresh_actual_detail_gantt.py",
        "\u53c2\u7167\u7528\\python\\plan_refresh_actual_detail_gantt.py",
    )
    write_cp932(p, t)
    print("updated:", p.relative_to(ROOT))

    # --- Gemini??.bas ---
    p = VBA_DIR / "Gemini\u9023\u643a.bas"
    t = p.read_bytes().decode("cp932")
    t = t.replace("\\python\\encrypt_gemini_credentials.py", "\\\u53c2\u7167\u7528\\python\\encrypt_gemini_credentials.py")
    t = t.replace("python\\encrypt_gemini_credentials.py", "\u53c2\u7167\u7528\\python\\encrypt_gemini_credentials.py")
    write_cp932(p, t)
    print("updated:", p.relative_to(ROOT))

    # --- ????????.bas ---
    p = VBA_DIR / "\u74b0\u5883\u30bb\u30c3\u30c8\u30a2\u30c3\u30d7.bas"
    t = p.read_bytes().decode("cp932")
    t = t.replace("\\python\\setup_environment.py", "\\\u53c2\u7167\u7528\\python\\setup_environment.py")
    t = t.replace('"python\\setup_environment.py"', '"\u53c2\u7167\u7528\\python\\setup_environment.py"')
    write_cp932(p, t)
    print("updated:", p.relative_to(ROOT))

    # --- ??????.bas ---
    p = VBA_DIR / "\u6bb5\u968e\u5b9f\u884c\u5236\u5fa1.bas"
    t = p.read_bytes().decode("cp932")
    old = "os.path.join(os.path.dirname(str(wb.fullname)), 'python', 'xlwings_console_runner.py')"
    new = (
        "os.path.join(os.path.dirname(str(wb.fullname)), "
        "'\u53c2\u7167\u7528', 'python', 'xlwings_console_runner.py')"
    )
    if old not in t:
        raise SystemExit("\u6bb5\u968e\u5b9f\u884c\u5236\u5fa1.bas: xlwings_console_runner join string not found")
    t = t.replace(old, new)
    write_cp932(p, t)
    print("updated:", p.relative_to(ROOT))

    print("done")


if __name__ == "__main__":
    main()
