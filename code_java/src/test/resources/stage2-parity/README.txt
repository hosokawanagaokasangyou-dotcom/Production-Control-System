段階2は JavaFX から常に Python 子プロセス（plan_simulation_stage2.py）のみ実行する。

1. 成果物の JSON／xlsx を任意の基準と突き合わせるときは、次の JUnit ユーティリティを参照する。
   - Stage2ProductionPlanJsonParityTest … 計画 primary JSON のツリー比較
   - Stage2PlanInputUiParityTest … 表データとディスク上ファイルの照合
   - Stage2WorkbookSemanticParityTest … xlsx の意味論比較
   - jp.co.pm.ai.desktop.io.Stage2OutputNamingParityTest … 「更新後の計画 JSON」検出など

2. ゴールデン JSON（結果_タスク一覧 等）の置き場と命名は golden/README_GOLDEN.txt を参照。

3. CLI で段階2正本を検証する例は scripts/stage2_python_dispatch_verify_repo.sh を参照（planning_core が要求する Python バージョンに合わせる）。
