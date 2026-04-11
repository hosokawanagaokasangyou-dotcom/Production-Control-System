# -*- coding: utf-8 -*-
"""Split 生産管理_AI配台テスト_xlsm_VBA.bas into README-listed modules (CP932 output)."""
from __future__ import annotations

import re
import sys
from pathlib import Path

ENC = "cp932"
ROOT = Path(__file__).resolve().parents[1]  # .../テストコード
VBA_DIR = ROOT / "VBA"
MONO = VBA_DIR / "生産管理_AI配台テスト_xlsm_VBA.bas"
OUT_DIR = VBA_DIR / "モジュール分割バージョン"

# Import order / output filenames (README)
MODULE_ORDER = [
    "共通定義.bas",
    "段階実行制御.bas",
    "スプラッシュ表示.bas",
    "サウンド制御.bas",
    "Gemini連携.bas",
    "環境セットアップ.bas",
    "業務ロジック.bas",
    "文字列入出力共通.bas",
    "ファイル探索.bas",
    "フォント管理.bas",
    "起動ショートカット.bas",
]

# name→module 解決は「狭いモジュール優先」。README 順だと 業務ロジック が先に名前を奪うため別順を使う
MAP_ORDER = [
    "共通定義.bas",
    "段階実行制御.bas",
    "スプラッシュ表示.bas",
    "サウンド制御.bas",
    "Gemini連携.bas",
    "環境セットアップ.bas",
    "文字列入出力共通.bas",
    "ファイル探索.bas",
    "フォント管理.bas",
    "起動ショートカット.bas",
    "業務ロジック.bas",
]

PROC_START = re.compile(
    r"^(?:(Public|Private)\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)\s*[\(:]",
    re.MULTILINE,
)


def extract_procedure_names_from_module(src: str) -> list[str]:
    names: list[str] = []
    for m in PROC_START.finditer(src):
        names.append(m.group(3))
    return names


def build_name_to_file() -> dict[str, str]:
    """First module in MAP_ORDER wins (specialized modules before 業務ロジック)."""
    name_to_file: dict[str, str] = {}
    for fn in MAP_ORDER:
        p = OUT_DIR / fn
        if not p.exists():
            continue
        old = p.read_text(encoding=ENC, errors="replace")
        for n in extract_procedure_names_from_module(old):
            if n not in name_to_file:
                name_to_file[n] = fn
            elif name_to_file[n] != fn:
                print(f"WARN: duplicate proc {n!r} kept in {name_to_file[n]} (skip {fn})", file=sys.stderr)
    return name_to_file


def promote_private_procedures_to_public(src: str) -> str:
    """分割モジュール間参照のため、手続きだけ Private→Public（Declare/Type/Const/m_ は対象外）。"""
    return re.sub(
        r"^Private\s+((?:Static\s+)?)(Function|Sub)\s+",
        r"Public \1\2 ",
        src,
        flags=re.MULTILINE,
    )


def find_procedure_spans(text: str) -> list[tuple[int, int, str]]:
    """(start_char, end_char, name) non-overlapping spans in monolithic text."""
    matches = list(PROC_START.finditer(text))
    spans: list[tuple[int, int, str]] = []
    for i, m in enumerate(matches):
        start = m.start()
        name = m.group(3)
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        spans.append((start, end, name))
    return spans


def main() -> int:
    if not MONO.exists():
        print("Missing monolithic:", MONO, file=sys.stderr)
        return 1

    name_to_file = build_name_to_file()
    text = MONO.read_text(encoding=ENC, errors="strict")

    header_end = text.find("Private Function ParseStage12CmdHideWindowBool")
    if header_end < 0:
        print("Could not find ParseStage12CmdHideWindowBool", file=sys.stderr)
        return 1

    header = text[:header_end]
    body = text[header_end:]

    spans = find_procedure_spans(body)
    unmapped: list[str] = []

    chunks: dict[str, list[str]] = {fn: [] for fn in MODULE_ORDER}
    chunks["共通定義.bas"] = [header]

    for start, end, name in spans:
        fn = name_to_file.get(name)
        if fn is None:
            unmapped.append(name)
            fn = "業務ロジック.bas"
        chunk = body[start:end]
        chunks[fn].append(chunk)

    if unmapped:
        uniq = sorted(set(unmapped))
        print("Unmapped procedures (assigned to 業務ロジック.bas):", ", ".join(uniq), file=sys.stderr)

    for fn in MODULE_ORDER:
        parts = chunks[fn]
        if not parts:
            print("WARN: empty module", fn, file=sys.stderr)
        raw = "".join(parts)
        if fn != "共通定義.bas":
            raw = "Option Explicit\r\n\r\n" + raw
        else:
            if not raw.lstrip().startswith("Option"):
                raw = "Option Explicit\r\n\r\n" + raw
        raw = promote_private_procedures_to_public(raw)
        out_path = OUT_DIR / fn
        out_bytes = raw.replace("\n", "\r\n").replace("\r\r\n", "\r\n")
        out_path.write_bytes(out_bytes.encode(ENC, errors="strict"))

    tw = VBA_DIR / "生産管理_AI配台テスト_ThisWorkbook_VBA.txt"
    if tw.exists():
        tw_out = OUT_DIR / "生産管理_AI配台テスト_ThisWorkbook_VBA.txt"
        try:
            t = tw.read_text(encoding="utf-8", errors="strict")
        except UnicodeDecodeError:
            t = tw.read_text(encoding=ENC, errors="strict")
        tw_out.write_bytes(t.replace("\n", "\r\n").encode(ENC, errors="strict"))

    print("Wrote", len(MODULE_ORDER), "modules to", OUT_DIR)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
