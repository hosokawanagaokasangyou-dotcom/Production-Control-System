import os
import sys

from oletools.olevba import VBA_Parser


def main() -> int:
    if sys.version_info < (3, 14):
        v = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
        print(
            "Python 3.14 以上が必要です（現在 "
            + v
            + "）。例: py -3.14 tools/extract_vba_search.py",
            file=sys.stderr,
        )
        return 2
    if len(sys.argv) < 2:
        print("usage: extract_vba_search.py <xlsm/xlsb/xls> [keyword ...]")
        return 2
    path = os.path.abspath(sys.argv[1])
    keywords = sys.argv[2:] or [
        "結果_設備ガント_実績明細",
        "actual_detail_gantt",
        "gantt_refresh",
        "ACTUAL_DETAIL",
        "実績明細",
        "設備ガント",
        "データ抽出",
        "D2",
        "Range(\"D2\")",
        "Cells(2, 4)",
    ]
    print("file:", path)
    vp = VBA_Parser(path)
    try:
        if not vp.detect_vba_macros():
            print("no vba macros detected")
            return 0
        mods = []
        for (_fn, stream_path, vba_filename, vba_code) in vp.extract_macros():
            mods.append((vba_filename or "", stream_path or "", vba_code or ""))
        print("modules:", len(mods))
        for k in keywords:
            hits = []
            for vba_filename, stream_path, code in mods:
                if k in code:
                    hits.append((vba_filename, stream_path))
            print(f"{k}: {len(hits)}")
            for vba_filename, stream_path in hits[:20]:
                print("  -", vba_filename, "|", stream_path)
        print("\n--- context (first 20 matches each) ---")
        for k in keywords:
            printed = 0
            for vba_filename, stream_path, code in mods:
                if k not in code:
                    continue
                lines = code.splitlines()
                for i, ln in enumerate(lines):
                    if k in ln:
                        lo = max(0, i - 3)
                        hi = min(len(lines), i + 4)
                        print(f"\n[{k}] {vba_filename} | {stream_path} | line {i+1}")
                        for j in range(lo, hi):
                            print(f"{j+1:5d}: {lines[j]}")
                        printed += 1
                        if printed >= 20:
                            break
                if printed >= 20:
                    break
        return 0
    finally:
        vp.close()


if __name__ == "__main__":
    raise SystemExit(main())

