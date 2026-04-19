# -*- coding: utf-8 -*-
"""
VBA / cmd から起動: 試行順パターンごとに段階2（配台シミュレーション）を実行する。

- 「配台計画_タスク入力」を xlwings で読み、P1～P4（既定）の
  試行順を DataFrame に反映して ``_generate_plan_impl`` を回す。
- 成果物はマクロブックではなく ``output/dispatch_pattern_stage2/<実行時刻>/<パターンID>/`` に保存
  （production_plan_multi_day_*.xlsx / member_schedule_*.xlsx）。
  バッチ実行時は結果ブックに「結果_設備ガント_実績明細」（加工実績明細ベースの設備ガント）を含めません。
- マクロブックにシート ``配台試行順_パターン別段階2``（環境変数 DISPATCH_PATTERN_STAGE2_SUMMARY_SHEET で上書き可）を
  作成し、上記ブックへの HYPERLINK・簡易スコア・参考スコア・各パターンの処理時間(秒)・合計秒(C2/D2)・採用用 B3・バッチルート B2 を書く。
- シミュレーション件数は ``DISPATCH_PATTERN_STAGE2_MAX_PATTERNS``（既定 20）で上限。バッチフォルダに ``pattern_jobs_meta.json`` を保存。

環境変数は DISPATCH_TRIAL_PATTERN_LIST_* / DISPATCH_PATTERN_STAGE2_* 等（設定_環境変数 シート可）。

環境変数 TASK_INPUT_WORKBOOK、Excel で本ブックを開いたまま実行すること。
"""
import os
import sys

if os.name == "nt" and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import logging

os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    import workbook_env_bootstrap as _wbe

    _wbe.apply_from_task_input_workbook()
except Exception:
    pass

import planning_core as pc  # noqa: E402


def main() -> int:
    logging.info("apply_dispatch_trial_pattern_stage2_batch: 開始")
    ok = pc.refresh_dispatch_trial_pattern_stage2_batch_only()
    logging.info("apply_dispatch_trial_pattern_stage2_batch: 終了 ok=%s", ok)
    return 0 if ok else 1


if __name__ == "__main__":
    import workbook_env_bootstrap as _wbe_exit

    sys.exit(_wbe_exit.run_cli_with_optional_pause_on_error(main))
