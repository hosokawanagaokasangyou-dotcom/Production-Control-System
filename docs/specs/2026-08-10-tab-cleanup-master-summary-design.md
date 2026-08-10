# メインタブ整理（学習速度・ワークスペース履歴削除）とマスタ読込サマリ改善

**日付:** 2026-08-10  
**状態:** 設計承認済み（実装前）

## 背景

メインシェル「その他」グループのうち、日常運用に不要なタブと、現行ロジックと乖離した診断タブがある。

| タブ | 決定 |
|------|------|
| 学習速度データ | **全部削除**（UI＋適用ロジック） |
| 配台ワークスペース履歴 | **ロジックごと削除**（UI＋スナップショット） |
| マスタ読込サマリ | **残す／案 A**（Python を現行経路に追従） |

## 要件（確定）

### A. 学習速度 — 全部削除

**削除対象**

- UI: `LearnedSpeedDataTab.fxml` / `LearnedSpeedDataTabController` / `MainShellTabId.LEARNED_SPEED_DATA`
- 配線: `MainShell.fxml` / `MainShellController` / `MainShellTabLayoutDefaults` / `MainShellInnerTabCatalog`
- session 既定: `init_setting/session_defaults*.json` の当該キー
- 実行・ログ: 「実績由来学習速度を適用」チェック（`MainRunTab.fxml` / `MainRunTabController`）と `DesktopSessionState.mainRunApplyLearnedSpeedFromActuals`
- Java: `DispatchMlReadinessStore`、`AppPaths` / `EnvVarDocs` / 環境変数雛形 TSV の `PM_AI_LEARNED_SPEED_*`
- Python: `actual_speed_apply` の呼び出し（`stage1` / `plan_input`）、適用・分布モジュールと専用テスト（参照が無くなるもの）

**残す**

- master ブックの `speed` シート読込（マスタ由来速度。学習速度とは別）
- 学習アーカイブ path は他参照が無ければ削除、残参照があれば path 解決のみ残す

### B. 配台ワークスペース履歴 — ロジックごと削除

**削除対象**

- UI: `PlanWorkspaceHistoryTab.fxml` / `PlanWorkspaceHistoryTabController` / `PLAN_WORKSPACE_HISTORY`
- 配線・session 既定（学習速度と同様）
- `PlanWorkspaceSnapshotStore` / `PlanWorkspaceSessionFragment`
- `MainShellController.restorePlanWorkspaceSnapshot` および工場リセット内の `PlanWorkspaceSnapshotStore.deleteAllSilently`
- スナップショット専用の列順 partial API（`TableColumnOrderPersistence` の該当メソッド）

**残す（別機能）**

- キャッシュ履歴タブ（`CACHE_HISTORY` / `WorkspaceCacheArchiveStore`）
- 工場切替ワークスペース（`FactorySiteWorkspace*`）
- 結果_配台表の通常 I/O

### C. マスタ読込サマリ — 案 A

**目的**  
本番と同じ解決・読込経路で「読めたか」を確認できる画面にする。Java だけ肥大・Python が薄いプローブのまま、という乖離を解消する。

**Python（`planning_core/master_read_ui.py`）**

| キー／領域 | 内容 |
|------------|------|
| `skills_need` | 本番 `load_skills_and_needs()`（失敗は `warnings`、部分結果可） |
| `exclude_rules`（または既存 Java が読むキーに合わせる） | `PM_AI_EXCLUDE_RULES_JSON` の解決・存在・ルール件数・サンプル。Excel「設定_配台不要工程」を正本としない |
| `attendance` | 既存 `canonical_json` / `stage2_ready` を維持し、Java が表示可能な要約を明確化 |
| `machine_calendar` | JSON 正本（機械カレンダー）の readiness。シート有無だけに依存しない |
| `team_combinations` 等 | 本番ローダがあるものは呼ぶ。重い／無いものは present＋件数に落とす |

既存の `resolved_path` / `sheet_checks` / `speed`（マスタ sheet）/ `main_sheet` は維持。

**Java / FXML**

- 既存セクションを新 JSON で埋める
- 配台不要は JSON 正本表示に差し替え（Excel シート前提の文言・欄を改める）
- 勤怠・機械カレンダーは readiness（`stage2_ready` 等）を表示
- アラジンマスタは対象外のまま

**ドキュメント**

- `EnvVarDocs` / `ui_ref_env_defaults.json` の「マスタ読込サマリで確認可」説明を、実際の確認内容に合わせて更新
- `*.md` / `*.html` マニュアルは本仕様の範囲外（別依頼時のみ）

## 方式・構成

1. **削除はタブ単位で完結させる**  
   enum → LayoutDefaults → FXML → Controller → session_defaults → 参照呼び出しの順で枯らす。コンパイルが通るまでを 1 単位とする。
2. **学習速度の Python 削除は適用入口から**  
   `stage1` / `plan_input` の import・呼出を先に外し、孤立モジュールとテストを削除する。
3. **マスタサマリは契約先行**  
   Python が Java の既存キー契約に合わせて出力を増やす。キー名は `MasterReadSummaryTabController` の読取箇所に合わせ、必要なら Controller 側のラベル／セクション見出しのみ更新する。
4. **セッション互換**  
   既存 `session-state` に削除済みタブキーが残っても、sanitize／欠落マージで無視される想定。既定 JSON は掃除必須。

## エラー方針

| 状況 | 動作 |
|------|------|
| マスタファイル無し | `ok=false`、`warnings` に理由、可能な限り空セクション |
| `load_skills_and_needs` 失敗 | 全体を落とさず `warnings`＋当該セクション欠損 |
| 配台不要 JSON 未設定／無し | パス解決結果と `present=false` を明示（段階1と同様の解釈をコメントで揃える） |
| 勤怠／機械カレンダー readiness 例外 | `warnings`、当該ブロックは空または `stage2_ready=false` |

## テスト

**Python**

- `master_read_ui` のキー存在（skills_need / exclude / attendance readiness / machine_calendar）
- 配台不要 JSON がある／無いケース
- 学習速度削除後: `actual_speed_apply` 参照が残っていないこと（grep／既存 stage1 テストが緑）

**Java**

- `mvnw` コンパイル／関連単体（タブ ID・LayoutDefaults・session 既定に削除キーが無いこと）
- 学習速度チェック・ワークスペース restore の参照がコンパイル上残っていないこと

**手動**

- メインシェルに両タブが無い
- マスタ読込サマリ更新で skills／exclude JSON／勤怠 readiness が埋まる
- 段階1が学習速度チェック無しで従来どおり動く

## 非対象（明示）

- APIモデルベンチマーク／実行時間分析／キャッシュ履歴タブの削除
- アラジンマスタを本タブへ統合すること
- マニュアル HTML/MD の同期（依頼時のみ）
- ディスク上の既存 `~/.pm-ai-desktop/plan-workspace-snapshots/` や学習アーカイブの物理削除（コード参照削除後、残骸は運用放置可）
