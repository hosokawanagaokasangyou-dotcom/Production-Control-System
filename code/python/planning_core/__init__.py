"""
planning_core は 工場向け「配台計画」シミュレーションの中核（パッケージ）

【このパッケージの役割】
    VBA / マクロ実行ブックから渡されたパスを読み、加工計画データ・勤怠・
    スキル need を統合して、設備・担当者への割付結果（Excel）を生成する。

【配台まわりの設定の読み方】
    マクロブックにシート ``設定_環境変数``（A列=変数名・B列=値）があるとき、各エントリ
    （``task_extract_stage1`` / ``plan_simulation_stage2`` / ``xlwings_console_runner`` 等）が
    ``import planning_core`` より前に ``workbook_env_bootstrap`` で内容を ``os.environ`` に反映する。
    運用上は **本シートにキーを集約**し、OS のユーザー環境変数に触れなくてよい（VBA が必ず設定する
    ``TASK_INPUT_WORKBOOK`` だけは従来どおりプロセス環境で渡す）。
    環境構築 ``setup_environment`` は pip で依存が入ったあと同シートを読み、``PM_AI_CMD_PAUSE_ON_ERROR`` / ``PM_AI_SETUP_COMPLETION_MSGBOX`` 等も反映する（前者は終了時 pause の成功・失敗共通の無効化にも使う）。
    ブックに ``配台_配台不要工程``
    があり openpyxl でブックを開けない場合はシートを読めない（そのときは OS 環境または既定値）。

【外部から直接使う入口（他 .py から import される想定）】
    - ``TASKS_INPUT_WORKBOOK`` … マクロブックのパス（環境変数 TASK_INPUT_WORKBOOK。上記シートでも上書き可）
    - ``run_stage1_extract()`` … 段階1。`task_extract_stage1.py` から呼ぶ
    - ``generate_plan()`` … 段階2。`plan_simulation_stage2.py` から呼ぶ
    - ``refresh_equipment_gantt_actual_detail_only()`` … 実績明細の設備ガントシートのみ再生成。
      既存 ``actual_detail_gantt_refresh.xlsx`` の「データ抽出」が今回と同一なら更新をスキップする。
      `plan_refresh_actual_detail_gantt.py` / マクロ「実績設備ガント_のみ更新_実行」から呼ぶ
    - ``apply_result_task_column_layout_only()`` … 取り込み済み「結果_タスク一覧」の列順・非表示を
      「列設定_結果_タスク一覧」に合わせる（`apply_result_task_column_layout.py` / VBA ボタン用・xlwings）
    - ``refresh_plan_input_dispatch_trial_order_only()`` … 「配台計画_タスク入力」の配台試行順番を
      再付与（`apply_plan_input_dispatch_trial_order.py` / xlwings）。設定シートの行同期・分割行の自動配台不要は
      段階2の ``load_planning_tasks_df`` と同様。「設定_配台不要工程」の C/E による配台不要の**再適用**は
      いずれも行わない（当該ルールの DataFrame 反映は段階1のみ）。段階2・試行順更新では設定シートの
      D 列→E 列（ロジック式）の **Gemini 補完は行わない**（段階1のみ）。
    - ``sort_plan_input_dispatch_trial_order_by_float_keys_only()`` … シート上の「配台試行順番」を
      小数キーとして昇順に並べ替え 1..n（`apply_plan_input_dispatch_trial_order_sort_by_float_keys.py`）。
      マスタ・上書きミューテーションは使わない。
    - ``refresh_dispatch_trial_pattern_list_sheet_only()`` … P1～P5 の決定論パターンを既定で一覧化
      （P3 は機械グループの納期−原反合計が短い順＋途中依頼NO優先、P4 は納期−原反日数の短い順＋途中依頼NO優先、
      P5 は P2 プローブ後に納期遅れ依頼のみ原反投入を1日前にしてから P2 を再適用）。
      （`apply_dispatch_trial_pattern_list_sheet.py` / xlwings）。値は ``DISPATCH_TRIAL_PATTERN_*`` または「設定_環境変数」シート（雛形 TSV 参照）。
    - ``refresh_dispatch_trial_pattern_stage2_batch_only()`` … 各パターンで段階2を実行し、
      ``output/dispatch_pattern_stage2/`` 配下に別ブックを保存。マクロブックに ``配台試行順_パターン別段階2`` で
      リンク・簡易スコア・参考スコア・採用用 B3 を書く（件数上限 ``DISPATCH_PATTERN_STAGE2_MAX_PATTERNS``、既定 20）。
      （`apply_dispatch_trial_pattern_stage2_batch.py` / xlwings ``run_dispatch_trial_pattern_stage2_batch_for_xlwings``）。
    - ``refresh_dispatch_pattern_stage2_selection_to_plan_only()`` … 上記サマリの採用パターンIDに従い
      ``pattern_jobs_meta.json`` を読み「配台計画_タスク入力」の配台試行順番を書き戻す
      （`apply_dispatch_pattern_stage2_selection.py` / xlwings ``run_dispatch_pattern_stage2_selection_for_xlwings``）。

【処理の流れ（ざっくり）】
    1. 段階1: 「加工計画DATA」→ 中間 `output/plan_input_tasks.xlsx`、
       マクロブック内「設定_配台不要工程」の行同期・D列→E列(Gemini)・保存（既定は xlwings。段階2では D→E AI なし）
    2. 段階2: master の skills / 勤怠 / 配台計画シートを読み、日付ループで割付、
       `output/` 直下に `production_plan_multi_day_*.xlsx` と `member_schedule_*.xlsx` を出力（実行前に同名パターンを削除・短いリトライ。
       ファイル名はデータ抽出時刻＋実行時刻サフィックスでペア一致・実行ごとに一意。Excel 占有時は削除に失敗しうるが新規名で保存は継続しやすい）

【ソース上の構成】
    - ``planning_core.bootstrap`` … カレントディレクトリ・ルートロガー・output/log 等
    - ``planning_core._core`` … 列定数・配台ロジック・段階1/2 本体（単一ファイルで極大。
      AI/局所編集用の行帯目安は同梱の ``_core_FILE_MAP.txt`` を参照）

【命名】
    - 先頭 ``_`` … モジュール内専用ヘルパ（外部から呼ばない想定）
    - ``PLAN_*`` / ``TASK_*`` … Excel 見出しと一致させる定数

【依存】
    pandas, openpyxl, xlwings（Excel 起動中の保存・設定シート同期）, google.genai, cryptography（暗号化認証 JSON 利用時）
    API キーはマクロブック「設定」B1 の JSON パスからのみ（環境変数 GEMINI_API_KEY は使わない）。
    暗号化 JSON の復号は ``planning_core`` 内の定数のみ使用（値はソースにあり、社内共有資料やログには書かないこと）。ログにキーは出さない。

【初回環境】 テストコード直下で ``py -3.14 python/setup_environment.py`` またはマクロ「環境構築」（requirements.txt 一括）
"""

# ``bootstrap`` が先に chdir するため、その前に同階層の兄弟モジュール（dispatch_interval_mirror 等）を
# 解決できるよう、本パッケージの親ディレクトリ（通常は ``python/``）を sys.path に含める。
import os
import sys

_pkg_dir = os.path.dirname(os.path.abspath(__file__))
_py_dir = os.path.dirname(_pkg_dir)
if _py_dir not in sys.path:
    sys.path.insert(0, _py_dir)

# Side effects: cwd, logging, output/log ディレクトリ
from . import bootstrap  # noqa: F401

from . import _core

# 旧 ``planning_core.py`` 単一モジュール互換:
# ``from planning_core import *`` はアンダースコア名を載せないため、
# ``pc._xlwings_*`` 等も使えるよう _core のトップレベル名をすべて公開する。
for _name in dir(_core):
    if _name.startswith("__"):
        continue
    globals()[_name] = getattr(_core, _name)

# bootstrap 側の例外・パス（_core に同名が無い想定だが、明示で上書き）
for _name in (
    "PlanningValidationError",
    "stage2_blocking_message_path",
    "_write_stage2_blocking_message",
    "_clear_stage2_blocking_message_file",
):
    globals()[_name] = getattr(bootstrap, _name)
