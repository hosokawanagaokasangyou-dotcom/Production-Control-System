# -*- coding: utf-8 -*-
"""段階2.5(AI) 前景エントリ（アラジン整列のみ）。"""
import json
import os
import sys

_py_here = os.path.dirname(os.path.abspath(__file__))
if _py_here:
    sys.path.insert(0, _py_here)

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass


def main():
    from pathlib import Path

    from planning_core.stage2_5_ai_runner import (
        ENV_JOB_ID,
        ENV_STAGE2_RAW,
        resolve_default_dispatch_json,
        run_stage2_5_foreground,
    )

    dispatch_json = resolve_default_dispatch_json()
    if len(sys.argv) >= 2 and sys.argv[1].strip():
        dispatch_json = Path(sys.argv[1]).resolve()
    job_id = (os.environ.get(ENV_JOB_ID) or "").strip()
    if not job_id:
        print("PM_AI_STAGE2_5_JOB_ID が未設定です。", file=sys.stderr)
        return 2
    raw_env = (os.environ.get(ENV_STAGE2_RAW) or "").strip()
    raw_path = Path(raw_env).resolve() if raw_env else None
    try:
        result = run_stage2_5_foreground(dispatch_json, job_id=job_id, stage2_raw=raw_path)
        print(json.dumps(result, ensure_ascii=False), flush=True)
        return 0
    except FileNotFoundError as e:
        print(str(e), file=sys.stderr)
        return 2
    except Exception as e:
        print(str(e), file=sys.stderr)
        return 1


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
