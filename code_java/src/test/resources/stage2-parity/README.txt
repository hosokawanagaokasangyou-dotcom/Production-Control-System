段階2 Java エンジン（PM_AI_STAGE2_ENGINE=java）と Python 子プロセスの回帰比較用メモ

1. JavaFX「配台計画_タスク入力」タブの「Java/Python 同一検証」ボタン: 同一 child 用 env で Python 段階2のあと Java 段階2を順実行し、新しく出力された計画 primary JSON をツリー比較（Jackson JsonNode.equals）する。検証中のみ PM_AI_STAGE2_ENGINE を python / java に上書きする。
2. 手動: 同一の PM_AI_* 環境で Python 段階2を一度実行し、出力フォルダの 計画*.json を基準として保存する。その後 PM_AI_STAGE2_ENGINE=java で同じ入力を実行し、成果物を比較する。
3. 本リポジトリの JUnit では `Stage2JavaEngineSmokeTest` が最小マスタ＋CSV で入出力の足場を検証する（Python 基準ファイルの同梱は任意）。`Stage2ProductionPlanJsonParityTest` が JSON 比較、`Stage2OutputNamingParityTest` が「更新後の計画JSON」検出を検証する。
