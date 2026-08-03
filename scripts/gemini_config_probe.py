#!/usr/bin/env python3
"""モデルごとに生成設定の受け入れ可否を切り分ける（400 INVALID_ARGUMENT の原因特定用）。

    py -3.14 -X utf8 -u scripts/gemini_config_probe.py gemini-3.5-flash-lite
"""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "code" / "python"))
os.environ.setdefault("PM_AI_REPO_ROOT", str(REPO_ROOT))
os.environ.setdefault(
    "GEMINI_CREDENTIALS_JSON", str(REPO_ROOT / "gemini_credentials.encrypted.json")
)

from planning_core import _core as pc  # noqa: E402

PROMPT = 'JSON で {"entries": []} だけを返してください。'


def variants():
    schema = pc.ATTENDANCE_REMARK_AI_RESPONSE_SCHEMA
    yield "設定なし", None
    yield "思考0のみ", pc._gemini_generate_content_config(thinking_budget=0)
    yield "スキーマのみ", pc._gemini_generate_content_config(
        response_schema=schema, thinking_budget=-1
    )
    yield "思考0＋スキーマ", pc._gemini_generate_content_config(response_schema=schema)


def main() -> int:
    logging.basicConfig(level=logging.WARNING)
    model = sys.argv[1] if len(sys.argv) > 1 else "gemini-3.5-flash-lite"
    if not pc.API_KEY:
        print("Gemini API キーを解決できませんでした。")
        return 1
    client = pc._gemini_client(pc.API_KEY)
    for label, config in variants():
        try:
            if config is None:
                client.models.generate_content(model=model, contents=PROMPT)
            else:
                client.models.generate_content(
                    model=model, contents=PROMPT, config=config
                )
            print(f"{model}  {label:<16} OK")
        except Exception as ex:  # noqa: BLE001 - 切り分けが目的
            print(f"{model}  {label:<16} NG  {str(ex)[:200]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
