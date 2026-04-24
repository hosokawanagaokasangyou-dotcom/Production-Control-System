# -*- coding: utf-8 -*-
"""
過去スナップショット（pdf\\<stamp>\\結果_タスク一覧.csv）と現在マスタの実績明細を比較した
「結果_設備ガント_計画実績比較」を output に出力する。

引数: スナップショット（日時）フォルダの絶対パス。未指定時は環境変数 COMPARE_GANTT_SNAPSHOT_DIR。
さらに未設定で TASK_INPUT_WORKBOOK があれば、マクロ実行ブックと同一フォルダの pdf 配下から
結果_タスク一覧.csv がある最新の履歴フォルダを選ぶ（VBA の一覧1と同じ並びの先頭相当）。
workbook_env_bootstrap で TASK_INPUT_WORKBOOK を反映する想定。
マクロ「計画実績比較ガント_リストから生成実行」経由では COMPARE_GANTT_PLAN_TASKS_SHEET で
アラジン参照シートを「加工計画DATA_実績比較用」に切替（未設定時は加工計画DATA。段階1は従来どおり）。
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


def _resolve_snapshot_dir() -> str:
    """第1引数 > COMPARE_GANTT_SNAPSHOT_DIR > マクロブック同階層 pdf 配下の最新履歴フォルダ。"""
    snap = (sys.argv[1] if len(sys.argv) > 1 else "").strip()
    if snap:
        return snap
    snap = (os.environ.get(pc.ENV_COMPARE_GANTT_SNAPSHOT_DIR) or "").strip()
    if snap:
        return snap
    task_wb = (os.environ.get("TASK_INPUT_WORKBOOK") or "").strip()
    if not task_wb:
        return ""
    pdf_root = os.path.join(
        os.path.dirname(os.path.abspath(task_wb)), "pdf"
    )
    csv_leaf = f"{pc.RESULT_TASK_SHEET_NAME}.csv"
    candidates: list[str] = []
    try:
        if not os.path.isdir(pdf_root):
            return ""
        for name in os.listdir(pdf_root):
            p = os.path.join(pdf_root, name)
            if os.path.isdir(p) and os.path.isfile(os.path.join(p, csv_leaf)):
                candidates.append(p)
    except OSError:
        return ""
    if not candidates:
        return ""
    # スタンプ名（フォルダ名）の降順で最新＝VBA RefreshCompareGanttSnapshotList と同様
    return max(candidates, key=lambda x: os.path.basename(x))


def main():
    snap = _resolve_snapshot_dir()
    if not snap:
        print(
            "使い方: py plan_compare_gantt_from_snapshot.py <スナップショットフォルダ> "
            "または環境変数 COMPARE_GANTT_SNAPSHOT_DIR を設定してください。"
            " TASK_INPUT_WORKBOOK があれば、マクロブックと同一フォルダの pdf 配下から最新を選べます。",
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
