ゴールデンセット（Java 段階2一本化・Python 正本との厳密同一検証用）

■ ディレクトリ規則
  golden/<case_id>/
    task_queue_expected.json   … Stage2TaskQueueBuilder が出す依頼NO列の期待（移植進捗用）
    （任意）expected/          … Python 段階2を承認済みとして保存する計画 primary JSON 等。容量が大きい場合はリポジトリ外のパスを README に記すのみでも可

■ 収録例（task_queue_expected.json）
  case_minimal … 1 依頼・工程あり

■ result_task_list_sheet_expected.json（case_minimal）
  計画ブック JSON の「結果_タスク一覧」サブツリー。列順は Python _core.default_result_task_sheet_column_order(0) と一致。
  PassThrough は配台未実行のためセル値は主に空＋計画入力から転写した列のみ（依頼NO→タスクID 等）。Python 正本と全文一致させるときは同じ master/plan で段階2を実行して差し替える。
  case_two_tasks … 2 依頼の出現順
  case_skip_empty_process … 工程名が空の行はスキップ（Python TASK_COL_MACHINE 空ゲートに相当）

■ case_id の付け方
  英小文字・数字・アンダースコア。例: case_minimal, case_konan_202605

■ 正本出力の取り込み（手順）
  1. 対象 case の PM_AI_PLAN_INPUT_PATH・PM_AI_MASTER_WORKBOOK を固定する
  2. PM_AI_STAGE2_ENGINE=python（または委譲）で段階2を実行し、計画*.json・人員*.json・*設.json を承認スナップショットとして保管
  3. Java 同一検証（JavaFX ボタンまたは将来の CI）で突き合わせ、一致した版のみ golden に登録する
