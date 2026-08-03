#!/usr/bin/env python3
"""Gemini 応答時間の実測（旧方式 / 差分出力 / 差分出力＋バッチ並列）。

実 API を呼ぶ。認証は planning_core と同じ経路で解決する
（GEMINI_CREDENTIALS_JSON、未設定ならリポジトリ直下の gemini_credentials.encrypted.json）。

マスタブックが無い環境でも勤怠備考 AI と同じ形の入力を合成して測れるようにしてある。

    py -3.14 -X utf8 -u scripts/gemini_latency_benchmark.py --rows 60
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from datetime import date, timedelta
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "code" / "python"))
os.environ.setdefault("PM_AI_REPO_ROOT", str(REPO_ROOT))
os.environ.setdefault(
    "GEMINI_CREDENTIALS_JSON", str(REPO_ROOT / "gemini_credentials.encrypted.json")
)

from planning_core import _core as pc  # noqa: E402

MEMBERS = [
    "冨田　裕子",
    "図司　智子",
    "宮島　　剛",
    "森下　　誠",
    "森岡　真由美",
    "竹内　正美",
    "菅沼　めぐみ",
]

REMARKS = [
    "通常勤務",
    "午前中は事務所で作業",
    "午後は会議",
    "有給休暇",
    "午前半休",
    "午後半休",
    "月次点検のみ",
    "9時出勤",
    "16時退勤",
    "教育のため現場不可",
    "体調不良で欠勤",
    "-",
]


def build_lines(rows: int) -> list[str]:
    """勤怠備考 AI に渡すのと同じ形（YYYY-MM-DD_メンバー名 の備考: 本文）の合成入力。"""
    base = date(2026, 4, 1)
    out: list[str] = []
    for i in range(rows):
        d = base + timedelta(days=i // len(MEMBERS))
        m = MEMBERS[i % len(MEMBERS)]
        out.append(f"{d.isoformat()}_{m} の備考: {REMARKS[i % len(REMARKS)]}")
    return out


def old_style_prompt(lines) -> str:
    """変更前のプロンプト（全行・全項目を書き戻させる契約）を再現したもの。"""
    joined = "\n".join(str(x) for x in lines)
    return f"""
以下は勤怠の備考・メンバーの備考です。出退勤時刻の変更や中抜き、休日の判定を行い JSON 形式で出力してください。
マークダウン記法(``` 等)は一切含めず、純粋な JSON 文字列のみを返してください。

【JSON の出力形式（キー型を厳密に守ること）】
{{
  "YYYY-MM-DD_メンバー名": {{
    "出勤時刻": "HH:MM",
    "退勤時刻": "HH:MM",
    "中抜き開始": "HH:MM",
    "中抜き終了": "HH:MM",
    "作業効率": 1.0,
    "is_holiday": false,
    "配台不参加": false
  }}
}}
・キー名は上記の日本語キーをそのまま使う（英語キーに置き換えない）
・【特記事項リスト】の全行について、対応するオブジェクトを必ず出力する
・出勤時刻 / 退勤時刻: 当該行の備考から推測。不明や変更なしなら null
・中抜き開始 / 中抜き終了: 一時的な離脱（中抜け・事務所・会議など）があるときはその開始・終了。ない場合は null
・is_holiday: 終日休暇・欠勤など勤務自体がないと判断できる場合のみ true
・配台不参加: 勤務はあるが加工ラインへの配台（OP/AS の割当）に載せてはいけないときは true
・作業効率: 0.0〜1.0 の数値

【特記事項リスト】
{joined}
"""


def usage_of(res) -> dict[str, int]:
    um = getattr(res, "usage_metadata", None)

    def g(name: str) -> int:
        v = getattr(um, name, None) if um is not None else None
        try:
            return int(v) if v is not None else 0
        except (TypeError, ValueError):
            return 0

    return {
        "prompt": g("prompt_token_count"),
        "output": g("candidates_token_count"),
        "thoughts": g("thoughts_token_count"),
        "total": g("total_token_count"),
    }


def report(label: str, sec: float, usage: dict[str, int], entries: int, model: str) -> None:
    print(
        f"{label:<28} {sec:7.2f}s  "
        f"入力 {usage['prompt']:>6}  出力 {usage['output']:>6}  思考 {usage['thoughts']:>6}  "
        f"取得 {entries:>3} 件  {model}"
    )


def run_old(client, lines) -> None:
    """旧方式: 全行・全項目、思考はモデル既定、単一リクエスト。"""
    config = pc._gemini_generate_content_config(thinking_budget=-1)
    t0 = time.perf_counter()
    res, model = pc._gemini_generate_content_with_retry(
        client, contents=old_style_prompt(lines), log_label="旧方式", config=config
    )
    sec = time.perf_counter() - t0
    parsed = pc._gemini_parse_json_object(pc._gemini_result_text(res)) or {}
    report("旧方式（全行・思考あり）", sec, usage_of(res), len(parsed), model)


def run_diff_single(client, lines) -> None:
    """差分のみ出力＋スキーマ＋思考無効、単一リクエスト。"""
    config = pc._gemini_generate_content_config(
        response_schema=pc.ATTENDANCE_REMARK_AI_RESPONSE_SCHEMA
    )
    t0 = time.perf_counter()
    res, model = pc._gemini_generate_content_with_retry(
        client,
        contents=pc._attendance_remark_ai_prompt(lines),
        log_label="差分・単発",
        config=config,
    )
    sec = time.perf_counter() - t0
    parsed = pc._attendance_ai_entries_to_map(
        pc._gemini_parse_json_object(pc._gemini_result_text(res))
    )
    report("差分のみ（単一リクエスト）", sec, usage_of(res), len(parsed), model)


def run_diff_batched(client, lines, batch_size: int, workers: int) -> None:
    """差分のみ出力＋バッチ分割並列（本番と同じ経路）。"""
    before = json.loads(json.dumps(pc._gemini_usage_session))
    t0 = time.perf_counter()
    merged, models, failed = pc._gemini_generate_json_map_in_batches(
        client,
        items=lines,
        build_prompt=pc._attendance_remark_ai_prompt,
        log_label="差分・バッチ",
        batch_size=batch_size,
        max_workers=workers,
        response_schema=pc.ATTENDANCE_REMARK_AI_RESPONSE_SCHEMA,
        parse_map=pc._attendance_ai_entries_to_map,
    )
    sec = time.perf_counter() - t0
    usage = {"prompt": 0, "output": 0, "thoughts": 0, "total": 0}
    for mid, cur in pc._gemini_usage_session.items():
        old = before.get(mid, {})
        usage["prompt"] += cur["prompt"] - int(old.get("prompt") or 0)
        usage["output"] += cur["candidates"] - int(old.get("candidates") or 0)
        usage["thoughts"] += cur["thoughts"] - int(old.get("thoughts") or 0)
        usage["total"] += cur["total"] - int(old.get("total") or 0)
    label = f"差分＋バッチ{batch_size}行×並列{workers}"
    report(label, sec, usage, len(merged), ",".join(sorted(set(models))) or "—")
    if failed:
        print(f"  ※ 失敗バッチ {failed} 件")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", type=int, default=60, help="合成する備考行数")
    ap.add_argument("--batch-size", type=int, default=20, help="バッチ 1 本あたりの行数")
    ap.add_argument("--workers", type=int, default=4, help="並列数")
    ap.add_argument("--skip-old", action="store_true", help="旧方式の計測を省く")
    ap.add_argument("--pause", type=float, default=5.0, help="計測間の待機秒（RPM 対策）")
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s"
    )
    if not pc.API_KEY:
        print("Gemini API キーを解決できませんでした（GEMINI_CREDENTIALS_JSON を確認）。")
        return 1

    lines = build_lines(args.rows)
    client = pc._gemini_client(pc.API_KEY)
    print(f"\n入力 {len(lines)} 行 / タイムアウト {pc._gemini_request_timeout_sec():.0f}s\n")

    if not args.skip_old:
        run_old(client, lines)
        time.sleep(args.pause)
    run_diff_single(client, lines)
    time.sleep(args.pause)
    run_diff_batched(client, lines, args.batch_size, args.workers)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
