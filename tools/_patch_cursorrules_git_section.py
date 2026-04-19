# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
p = ROOT / ".cursorrules"
t = p.read_text(encoding="utf-8")
old = (
    "`git push`**）。\n\n"
    "要点（Generate 用）: "
)
new = (
    "`git push`**）。\n\n"
    "**作業完了時（変更後）**: **`.cursor/rules/git-commit-push-after-code-changes.mdc`** "
    "— そのターンで変更・追加した版管理対象は **漏れなく `git add` → `commit` → `push`**。"
    "\u4f9d\u983c\u5916\u306e\u5dee\u5206\u304c\u6df7\u5728\u3059\u308b\u5834\u5408\u306f\u30b3\u30df\u30c3\u30c8\u3092\u5206\u3051\u3001\u5fdc\u7b54\u3067 "
    "`git status` \u306e\u8981\u7d04\u3092\u5831\u544a\u3059\u308b\u3002\n\n"
    "要点（Generate 用）: "
)
if old not in t:
    raise SystemExit("needle not found")
p.write_text(t.replace(old, new, 1), encoding="utf-8")
print("patched", p)
