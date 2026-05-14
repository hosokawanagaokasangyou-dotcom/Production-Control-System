# -*- coding: utf-8 -*-
"""
インタラクティブ配台試行（段階3）: 結果_配台表.json を入力とし、
``planning_core.stage2_identical_dispatch_runner`` 経由で段階2と同一条件の配台を実行する。

オーケストレーションの正本は runner。本スクリプトは argv・ログ・xlsx エクスポートのみ担当する。
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
# python -P / PYTHONSAFEPATH ではスクリプト所在ディレクトリが sys.path に入らない。
sys.path.insert(0, str(SCRIPT_DIR))
os.chdir(str(SCRIPT_DIR))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def main() -> int:
    if len(sys.argv) < 2:
        print(
            "usage: dispatch_interactive_trial.py <path-to-result-dispatch.json>",
            file=sys.stderr,
        )
        return 2
    path = Path(sys.argv[1]).resolve()
    if not path.is_file():
        print(f"not a file: {path}", file=sys.stderr)
        return 1
    print("[dispatch trial] 入力JSONを読み込み中…", flush=True)

    from planning_core.stage2_identical_dispatch_runner import (
        run_interactive_dispatch_trial_from_result_dispatch_json,
    )

    s3_phase = None
    s3_two = None
    if len(sys.argv) >= 3:
        s3_phase = sys.argv[2].strip().lower()
        if s3_phase not in ("equipment", "people", "both"):
            print(
                "usage: dispatch_interactive_trial.py <path-to-result-dispatch.json> "
                "[equipment|people|both]",
                file=sys.stderr,
            )
            return 2
        s3_two = True
    code, shortage_path = run_interactive_dispatch_trial_from_result_dispatch_json(
        path, stage3_phase=s3_phase, stage3_two_phase=s3_two
    )
    if code != 0:
        return code
    if shortage_path is None:
        return 1

    export_script = SCRIPT_DIR / "export_result_dispatch_from_json.py"
    if export_script.is_file():
        print("[dispatch trial] 結果Excel(xlsx)をエクスポート中…", flush=True)
        py = sys.executable or "python3"
        try:
            subprocess.run(
                [py, str(export_script), str(path)],
                cwd=str(SCRIPT_DIR),
                check=True,
                timeout=600,
            )
            print("[dispatch trial] xlsx エクスポート完了。", flush=True)
        except Exception as e:
            print(f"xlsx export warning: {e}", file=sys.stderr)

    print(str(shortage_path), flush=True)
    return 0


if __name__ == "__main__":
    try:
        import workbook_env_bootstrap as _wbe_exit

        sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
    except ImportError:
        sys.exit(main())
