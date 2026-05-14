ゴールデンセット（段階2成果物の手動・JUnit 照合用）

■ ディレクトリ規則
  golden/<case_id>/
    （任意）expected/ … Python 段階2を承認済みとして保存する計画 primary JSON 等。容量が大きい場合はリポジトリ外のパスを README に記すのみでも可

■ result_task_list_sheet_expected.json（case_minimal）
  計画ブック JSON の「結果_タスク一覧」サブツリーの例。列順は Python _core.default_result_task_sheet_column_order(0) と一致。
  正本と全文一致させるときは同じ master/plan で Python 段階2を実行して差し替える。

■ case_id の付け方
  英小文字・数字・アンダースコア。例: case_minimal, case_konan_202605

■ 正本出力の取り込み（手順）
  1. 対象 case の PM_AI_PLAN_INPUT_PATH・PM_AI_MASTER_WORKBOOK を固定する
  2. Python 子で段階2を実行し、計画*.json・人員*.json・*設.json を承認スナップショットとして保管
  3. 必要に応じて上記 JUnit ユーティリティや手動 diff で突き合わせ、承認した版のみ golden に登録する
