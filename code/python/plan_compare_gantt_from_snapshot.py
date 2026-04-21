# -*- coding: utf-8 -*-
"""
過去スナップショット（pdf\\<stamp>\\結果_タスク一覧.csv）と現在マスタの実績明細を比較した
「結果_設備ガント_計画実績比較」を output に出力する。

引数: スナップショットフォルダの絶対パス（未指定時は環境変数 COMPARE_GANTT_SNAPSHOT_DIR）。
マクロから TASK_INPUT_WORKBOOK を渡す想定（workbook_env_bootstrap）。
"""
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass

import planning_core as pc


def main():
    snap = (sys.argv[1] if len(sys.argv) > 1 else "").strip()
    if not snap:
        snap = (os.environ.get(pc.ENV_COMPARE_GANTT_SNAPSHOT_DIR) or "").strip()
    if not snap:
        print(
            "使い方: py plan_compare_gantt_from_snapshot.py <スナップショットフォルダ> "
            "または環境変数 COMPARE_GANTT_SNAPSHOT_DIR を設定してください。",
            file=sys.stderr,
        )
        sys.exit(2)
    try:
        out = pc.write_plan_actual_compare_gantt_from_snapshot_dir(snap)
        print(out)
    except pc.PlanningValidationError as e:
        msg = str(e).strip() or "計画実績比較ガントを中断しました。"
        print(msg, file=sys.stderr)
        sys.exit(3)
    except SystemExit:
        raise
    except Exception:
        import logging

        logging.exception("plan_compare_gantt_from_snapshot")
        sys.exit(1)


if __name__ == "__main__":
    main()
