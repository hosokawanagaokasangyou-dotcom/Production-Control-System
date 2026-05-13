段階2 Java エンジン（PM_AI_STAGE2_ENGINE=java）と Python 子プロセスの回帰比較用メモ

1. JavaFX「配台計画_タスク入力」タブの「Java/Python 同一検証」ボタン: 同一 child 用 env で Python 段階2のあと Java 段階2を順実行し、(a) 配台計画タブの表と PM_AI_PLAN_INPUT_PATH の実体が一致するか、(b) 新しく出力された計画 primary JSON・人員 JSON（PM_AI_MEMBER_SCHEDULE_JSON 有効時）、(c) 計画／人員 xlsx（PM_AI_STAGE2_WRITE_EXCEL 有効時、セル内容の意味論比較）を突き合わせる。検証中のみ PM_AI_STAGE2_ENGINE を python / java に上書きする。
2. 手動: 同一の PM_AI_* 環境で Python 段階2を一度実行し、出力フォルダの 計画*.json を基準として保存する。その後 PM_AI_STAGE2_ENGINE=java で同じ入力を実行し、成果物を比較する。
3. 本リポジトリの JUnit では `Stage2JavaEngineSmokeTest` が最小マスタ＋CSV で入出力の足場を検証する（Python 基準ファイルの同梱は任意）。`Stage2ProductionPlanJsonParityTest` が JSON ツリー比較、`Stage2PlanInputUiParityTest` が UI／ファイル表の照合、`Stage2WorkbookSemanticParityTest` が xlsx 意味論比較、`Stage2OutputNamingParityTest` が「更新後の計画JSON」検出を検証する。
4. Python 正本（_generate_plan_impl）と同一の表・ガントを出すには、(a) 従来どおり `PM_AI_STAGE2_ENGINE=python`、または (b) `PM_AI_STAGE2_ENGINE=java` に加え `PM_AI_STAGE2_JAVA_DELEGATE_PYTHON_DISPATCH=1` で Java 起動経路から Python 子へ委譲する。リポジトリに置いた段階1の `output/plan_input_tasks.xlsx` 等で CLI 検証する場合は `scripts/stage2_python_dispatch_verify_repo.sh` を参照（planning_core が要求する Python バージョンに合わせる）。

5. ゴールデンセット（Java 一本化移行）: `golden/README_GOLDEN.txt` と `golden/<case_id>/` を参照。JUnit `Stage2TaskQueueGoldenTest` が `case_minimal` の task_queue を検証する。承認済み Python 成果物の同梱は任意（容量・機微情報に注意）。

6. 本番既定を `PM_AI_STAGE2_ENGINE=java` のみへ切り替える条件: チーム合意の golden 全件で「Java/Python 同一検証」がパスした後。切替後もリリース前に同一検証を実施する（EnvVarDocs の PM_AI_STAGE2_ENGINE 説明も参照）。

7. CI でゴールデン系テストを追加実行する場合: 環境変数 `PM_AI_STAGE2_GOLDEN_CI=1` を付けて `scripts/stage2_golden_parity_ci.sh` を実行（詳細はスクリプト先頭コメント）。
