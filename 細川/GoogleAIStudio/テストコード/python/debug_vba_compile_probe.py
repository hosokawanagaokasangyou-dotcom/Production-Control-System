from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path


LOG_PATH = Path("debug-91edfc.log")
SESSION_ID = "91edfc"
RUN_ID = f"probe-{int(time.time())}"


def emit(hypothesis_id: str, location: str, message: str, data: dict) -> None:
    # region agent log
    payload = {
        "sessionId": SESSION_ID,
        "runId": RUN_ID,
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data,
        "timestamp": int(time.time() * 1000),
    }
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    # endregion


def main() -> int:
    target = (
        Path(sys.argv[1])
        if len(sys.argv) > 1
        else Path("細川/GoogleAIStudio/テストコード/生産管理_AI配台テスト_xlsm_VBA.txt")
    )
    text = target.read_text(encoding="utf-8")
    lines = text.splitlines()

    emit("H0", "debug_vba_compile_probe.py:35", "probe_start", {"target": str(target), "lineCount": len(lines)})

    # H1: 先頭条件コンパイル行が破損している
    first_non_empty = next((ln for ln in lines if ln.strip()), "")
    emit("H1", "debug_vba_compile_probe.py:40", "first_non_empty_line", {"value": first_non_empty[:120]})

    # H2: #If/#Else/#End If の対応が壊れている
    depth = 0
    unmatched = []
    for i, ln in enumerate(lines, 1):
        t = ln.strip()
        if re.match(r"^#If\b", t):
            depth += 1
        elif re.match(r"^#Else\b", t):
            if depth <= 0:
                unmatched.append({"line": i, "kind": "#Else"})
        elif re.match(r"^#End If\b", t):
            if depth <= 0:
                unmatched.append({"line": i, "kind": "#End If"})
            else:
                depth -= 1
    emit("H2", "debug_vba_compile_probe.py:57", "conditional_balance", {"depth": depth, "unmatched": unmatched[:10]})

    # H3: 先頭 API 宣言ブロックで # が抜けた条件コンパイル行
    missing_hash = []
    for i, ln in enumerate(lines[:120], 1):
        t = ln.strip()
        if t in {"If VBA7 Then", "If Win64 Then", "Else", "End If"}:
            missing_hash.append({"line": i, "text": t})
    emit("H3", "debug_vba_compile_probe.py:65", "missing_hash_candidates", {"count": len(missing_hash), "examples": missing_hash[:10]})

    # H4: 不正文字（全角空白・スマートクォート）
    fw_space = [i for i, ln in enumerate(lines, 1) if "\u3000#" in ln or ln.startswith("\u3000")]
    smart_quote = [i for i, ln in enumerate(lines, 1) if "“" in ln or "”" in ln or "’" in ln]
    emit(
        "H4",
        "debug_vba_compile_probe.py:71",
        "unicode_noise_check",
        {"fullWidthSpaceLines": fw_space[:10], "smartQuoteLines": smart_quote[:10]},
    )

    # H5: API Declare 重複件数（異常増殖の確認）
    decl_count = len(re.findall(r"^\s*(Private|Public)\s+Declare\b", text, flags=re.MULTILINE))
    emit("H5", "debug_vba_compile_probe.py:81", "declare_count", {"count": decl_count})

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
