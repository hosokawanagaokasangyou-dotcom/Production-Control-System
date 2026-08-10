# メインタブ整理（学習速度・ワークスペース履歴削除）とマスタ読込サマリ改善

**日付:** 2026-08-10  
**状態:** 設計承認済み（3体レビュー反映・実装前）  
**レビュー反映:** 2026-08-10（削除境界・JSON 契約・実装順・枯らしチェックを固定）

## 背景

メインシェル「その他」グループのうち、日常運用に不要なタブと、現行ロジックと乖離した診断タブがある。

| タブ | 決定 |
|------|------|
| 学習速度データ | **全部削除**（UI＋適用＋分布更新） |
| 配台ワークスペース履歴 | **UI＋スナップショットストア削除** |
| マスタ読込サマリ | **残す／案 A**（Python を現行経路に追従＋Java バインド更新） |

## 成功基準

1. メインシェルに `learnedSpeedData` / `planWorkspaceHistory` タブが無い。
2. 段階1／配台読込が `PM_AI_LEARNED_SPEED_*` 無しで動作し、`actual_speed_apply` 参照が本番経路に無い。
3. 工場切替ワークスペースとキャッシュ履歴が従来どおり動く。
4. マスタ読込サマリ更新で、下表の必須 JSON キーが埋まり、空欄だらけにならない。
5. `version.txt` は commit 時 hooks で +0.01。

## 実装順（ゲート）

同一 PR 可。コミットはこの順。

1. **A 学習速度削除** → Java コンパイル＋段階1関連テスト緑＋grep で適用経路ゼロ  
2. **B ワークスペース履歴削除** → コンパイル＋工場 WS／CACHE_HISTORY 手動確認  
3. **C マスタサマリ案 A** → Python 契約テスト緑＋サマリ画面で主要セクションが埋まる  

A/B 完了前に C の `EnvVarDocs` 大規模編集をしない（競合回避）。

---

## 要件（確定）

### A. 学習速度 — 全部削除（適用＋分布）

会話確定: タブ＋実行チェック＋env＋Python 適用／**分布まで**削除。

#### 削除対象（パス単位）

| 層 | 対象 |
|----|------|
| UI | `LearnedSpeedDataTab.fxml` / `LearnedSpeedDataTabController` / `MainShellTabId.LEARNED_SPEED_DATA` |
| 配線 | `MainShell.fxml` / `MainShellController`（bind・タブ選択・`refreshLearnedSpeedDataQuietly`・適用ログ・`overlayMainRunSkipGeminiApiEnv` の LEARNED_SPEED 行） / `MainShellTabLayoutDefaults` / `MainShellInnerTabCatalog`（LEARNED_SPEED case） |
| 他 Controller | `DispatchInteractiveTabController` の `refreshLearnedSpeedDataQuietly` 呼出 |
| 実行・ログ | `MainRunTab.fxml` の「加工速度」ブロック全体（見出し・チェック・説明）/ `MainRunTabController` の snapshot／restore |
| Session | `DesktopSessionState.mainRunApplyLearnedSpeedFromActuals` / `DesktopSessionStateStore` 入出力。旧 session-state に残っても **optional 無視**（書込しない） |
| 既定 JSON | `init_setting/session_defaults.json`・`session_defaults_konan.json`・`session_defaults_kokubu.json` から `learnedSpeedData` と `mainRunApplyLearnedSpeedFromActuals` を除去 |
| Java ストア／env | `DispatchMlReadinessStore` / `AppPaths`・`EnvVarDocs`・`code/設定_環境変数_雛形.tsv` の `PM_AI_LEARNED_SPEED_*`。タブ専用なら `resolveDispatchLearningArchiveRoot`（Java）も削除 |
| Python 適用 | `stage1.py` / `plan_input.py` の apply 呼出、`planning_core/actual_speed_apply.py` |
| Python 分布 | `planning_core/actual_speed_distribution.py` / `update_actual_speed_distribution.py` / `tests/test_actual_speed_histogram.py` |
| アーカイブジョブ修正 | `dispatch_learning_archive.py` から `update_speed_distribution` / `write_ml_readiness` 呼出を除去 |

#### 残す（明示）

| 対象 | 理由 |
|------|------|
| master `speed` シート適用（`_apply_master_speed_sheet_*` / `MASTER_USE_SPEED_SHEET`） | マスタ由来速度。学習速度と別系統 |
| `dispatch_learning_archive.py` の run 退避・`aladdin_deviation_metrics` | 速度分布以外の学習アーカイブ本体 |
| `dispatch_run_archiver` / `dispatch_workspace.resolve_dispatch_learning_archive_root` / `PM_AI_DISPATCH_LEARNING_ARCHIVE_SUBDIR` | 上記アーカイブ用 |
| `PM_AI_LEARNING_ARCHIVE_ENABLED` | 定義のみ／死旗なら AppPaths・EnvVarDocs・TSV から掃除可（アーカイブ CLI が参照していなければ削除） |

#### 非対象（学習速度まわり）

- `code/要件定義/*.md` の LEARNED_SPEED 記述更新（別依頼・既知の陳腐化）
- ディスク上の学習アーカイブ物理削除
- `_core.py.bak` 掃除
- `ui_ref_env_defaults.json` に LEARNED_SPEED キーは無い（誤掃除禁止）

### B. 配台ワークスペース履歴 — UI＋スナップショットストア削除

#### 削除対象

| 層 | 対象 |
|----|------|
| UI | `PlanWorkspaceHistoryTab.fxml` / `PlanWorkspaceHistoryTabController` / `PLAN_WORKSPACE_HISTORY` |
| 配線 | `MainShell.fxml` / `MainShellController`（`snapshotDispatchDocumentForPlanWorkspace` / `restorePlanWorkspaceSnapshot` / 工場リセット内 `PlanWorkspaceSnapshotStore.deleteAllSilently`） / LayoutDefaults / session_defaults 3 ファイルの `planWorkspaceHistory` |
| ストア | `PlanWorkspaceSnapshotStore` / `PlanWorkspaceSessionFragment` |
| 付随 | `TableColumnOrderPersistence.capture/mergePlanWorkspaceColumnOrderPartial` / `DispatchInteractiveTabController.copyDispatchDocumentForSnapshot`（PlanWorkspace 経由のみならセット削除） |

#### 残す（別機能・触らない）

- キャッシュ履歴（`CACHE_HISTORY` / `WorkspaceCacheArchiveStore` / `restoreWorkspaceCacheArchive`）
- 工場切替（`FactorySiteWorkspace*` / `FactorySiteWorkspaceSnapshot`）
- 結果_配台表の通常 I/O

### C. マスタ読込サマリ — 案 A

**目的**  
本番と同じ解決・読込経路で「読めたか」を確認する。Python が薄いプローブ・Java だけ肥大、という乖離を解消する。

#### Python→Java JSON 契約（キー改名禁止）

正本は `MasterReadSummaryTabController` の既存読取。Python が埋める。ルートに別名キーを新設しない。

| パス | 必須 | 内容 |
|------|------|------|
| `ok` / `warnings` / `resolved_path` / `file_exists` / `cwd` / `openpyxl_skip` | はい | 総合・互換マーカー |
| `pm_ai_master_workbook_env` / `master_use_speed_sheet_env` / `pm_ai_exclude_rules_json_env` / `team_assign_*_env` | はい | env 表示（現行どおりルート） |
| `speed.*`（＋可能なら `lookup_sample`） | はい | マスタ speed シート |
| `main_sheet` / `sheet_checks` / `all_sheet_names` | はい | シート診断。`sheet_checks` の machine_calendar は**シート有無のみ** |
| `skills_need` | はい | 本番 `load_skills_and_needs()` 要約。**常にオブジェクトを返す**。失敗時: `loaded:false`＋`validation_error` / `parse_error` / `skip_reason`＋`warnings` |
| `team_combinations` | はい | 組合せ表ローダ要約（現行 Java フィールド） |
| `app_config_sheet` / `planning_constants` / `machine_daily_startup` | はい | 本番ローダがあれば呼ぶ。重い場合は present＋件数に落とす（対象を実装時に明記） |
| `exclude_rules_sheet` | はい | **キー名固定**。JSON 正本時: `source:"json"`, `present`, `pm_ai_exclude_rules_json_set/path`, `rules_count`, `rules_sample[]`, `stage1_effective_source_note`。Excel シートはフォールバック時のみ |
| `machine_calendar_sheet` | 任意 | レガシーシート寸法用。JSON readiness の正本にしない |
| `attendance` | はい | 既存＋Java が読むキーを追加: `stage2_ready`, `canonical_json`（または flatten: path/exists/company/member/machine ready/`issues[]`）, `matched_sheet_names`, `skills_member_count`, `attendance_sheets_matched` |

機械カレンダー JSON readiness は **`attendance.canonical_json.machine_calendar_*`（既存 `build_attendance_readiness`）を正本**とする。ルート `machine_calendar` は新設しない。

#### Java / FXML（必須改修・ラベルだけでは不可）

- `applyExcludeRulesSection` / FXML「配台不要工程シート」文言を JSON 正本表示に合わせて改修
- 勤怠セクションで `stage2_ready` / readiness 要約を表示（現状未バインド）
- 機械カレンダーは `attendance.canonical_json` の machine 系＋必要ならシート寸法
- アラジンマスタは対象外

#### 読取専用境界（禁止）

サマリ経路で次を呼んではならない:

- exclude JSON マージ書込・`run_exclude_rules_sheet_maintenance`
- 勤怠／機械カレンダーの save

配台不要は読取のみ（例: `_get_exclude_rules_from_json_env` 相当）。

#### 実行条件

- env は段階1/2 と同じ（`childEnvForPython`＋workbook bootstrap）
- `openpyxl_skip=true` 時: メインシート openpyxl 読込は抑止。skills/need は pandas 経路で可能な範囲。不可なら `skills_need.loaded=false`＋理由
- `build_attendance_readiness` には **members を渡す**（内部の二重 `load_skills_and_needs` を避ける）
- 想定: 読取専用・書込なし。重い I/O は UI 更新のたびに子プロセス全ロード（許容。タイムアウトは既存 Python 起動枠に従う）

#### ドキュメント

- `EnvVarDocs` / `ui_ref_env_defaults.json` の「マスタ読込サマリで確認可」を実内容に合わせて更新
- マニュアル md/html は非対象

---

## MainShell 枯らしチェック（必須）

学習速度:

- [ ] `overlayMainRunSkipGeminiApiEnv` の `PM_AI_LEARNED_SPEED_ENABLED` 行のみ除去（Gemini 行は残す）
- [ ] `refreshLearnedSpeedDataQuietly` 定義・呼出ゼロ
- [ ] `DesktopSessionStateStore` の `mainRunApplyLearnedSpeedFromActuals` 削除
- [ ] session_defaults 3 ファイル掃除
- [ ] `dispatch_learning_archive.py` から速度分布／ml_readiness 除去後も run 退避が動く

ワークスペース履歴:

- [ ] restore / snapshot / `deleteAllSilently`（PlanWorkspace）ゼロ
- [ ] 隣接する `WorkspaceCacheArchiveStore` を誤削除していない

---

## エラー方針

| 状況 | 動作 |
|------|------|
| マスタファイル無し | `ok=false`、`warnings` に理由。セクションは空オブジェクトまたは欠損フィールド明示 |
| `load_skills_and_needs` で `PlanningValidationError` 等 | プロセス全体 abort しない。`skills_need` オブジェクトは必ず返す（`loaded:false`＋エラー欄）。他セクションは継続 |
| 配台不要 JSON 未設定／無し | `exclude_rules_sheet.present=false`、パス解決結果を明示 |
| 勤怠／機械カレンダー readiness 例外 | `warnings`、`stage2_ready=false` |
| `openpyxl_skip` | warnings に含め、抑止した読込を `skip_reason` 等で示す |
| `ok` / exit | skills 失敗でも他が埋まれば exit 0 可。`ok` は「skills+need シートが読め本番相当」など総合判定を実装時に1行で固定 |

---

## テスト

**Python**

- `master_read_ui`: 必須キー存在、exclude JSON 有／無、`skills_need` 部分失敗で warnings＋オブジェクト返却、`attendance.stage2_ready`
- 学習速度削除後 grep: `actual_speed_apply` / `PM_AI_LEARNED_SPEED_` / `update_speed_distribution` が本番経路に残らない
- `dispatch_learning_archive` が速度分布無しでも完走（可能な範囲の単体または手動）

**Java**

- `mvnw` コンパイル
- LayoutDefaults / TabId / session_defaults 3 ファイルに削除キー無し
- 旧 session JSON（削除フィールド・削除タブキー付き）が例外なく読める
- `FactorySiteWorkspaceStore` / `WorkspaceCacheArchiveStore` 参照が残る
- （可能なら）MasterReadSummary 契約: 上記キーで主要セクションが空にならない

**手動**

- 両タブ非表示
- サマリ更新で skills／exclude JSON／勤怠 readiness が埋まる
- 段階1が学習速度 UI 無しで動作
- 工場切替後のサマリ再実行
- キャッシュ履歴の退避／復元が生きている

---

## 非対象（明示）

- APIモデルベンチマーク／実行時間分析／キャッシュ履歴タブの削除
- アラジンマスタの本タブ統合
- マニュアル HTML/MD・要件定義 md の学習速度記述更新（別依頼）
- ディスク上の `plan-workspace-snapshots/`・学習アーカイブ物理削除
- `_core.py.bak` 掃除

## 完了時

版管理対象を commit（`version.txt` は hooks で +0.01）。push はリポジトリ運用に従う。
