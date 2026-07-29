# -*- coding: utf-8 -*-
"""CLI: 設備ガント担当割当の検証・契約/計画ブック同期（1 行 JSON を stdout）。"""
from __future__ import annotations

import argparse
import json
import os
import sys
import traceback

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


def _load_updates(path: str | None) -> dict:
    if not path or path == "-":
        raw = sys.stdin.read()
    else:
        with open(path, encoding="utf-8-sig") as f:
            raw = f.read()
    if not raw.strip():
        return {}
    return json.loads(raw)


def main() -> int:
    parser = argparse.ArgumentParser(description="設備ガント担当割当の同期")
    parser.add_argument("--contract", required=True, help="設備ガント契約 JSON パス")
    parser.add_argument(
        "--updates-json",
        required=True,
        help="更新 JSON パス（- で stdin）",
    )
    parser.add_argument("--plan-xlsx", default="", help="計画ブック .xlsx（任意）")
    parser.add_argument(
        "--expected-timeline-hash",
        default="",
        help="楽観ロック用タイムラインハッシュ",
    )
    parser.add_argument("--confirm-token", default="", help="警告確認トークン")
    parser.add_argument(
        "--force-warnings",
        action="store_true",
        help="警告を確認済みとして強制適用",
    )
    parser.add_argument("--dry-run", action="store_true", help="検証のみ")
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="書込みせず検証のみ（--dry-run と同義）",
    )
    args = parser.parse_args()

    try:
        from planning_core.gantt_assignment_sync import (
            parse_updates_map,
            sync_assignment_to_artifacts,
        )

        updates = parse_updates_map(_load_updates(args.updates_json))
        plan_xlsx = args.plan_xlsx.strip() or None
        dry_run = bool(args.dry_run or args.validate_only)
        result = sync_assignment_to_artifacts(
            contract_path=args.contract,
            updates=updates,
            plan_xlsx_path=plan_xlsx,
            expected_timeline_hash=args.expected_timeline_hash.strip() or None,
            confirm_token=args.confirm_token.strip() or None,
            force_warnings=bool(args.force_warnings),
            dry_run=dry_run,
        )
        payload = result.to_dict()
        exit_code = 0 if result.ok else 1
        _emit(payload, exit_code)
        return exit_code
    except Exception as e:
        _emit(
            {
                "format_version": 1,
                "ok": False,
                "status": "error",
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
