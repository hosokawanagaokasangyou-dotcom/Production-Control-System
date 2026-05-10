---
name: ポータブル初回／バージョンアップのUI・配台不要既定
overview: 初回インストール用・バージョンアップ用の両ZIPに、タブ／ガント等のセッション既定JSON・テーブル列順・配台不要ルールJSONを同梱し、バージョンアップ同期後は ~/.pm-ai-desktop のセッションと列順ストアを正本に合わせて上書きする方針で実装済み。2026-05-10 にコード実装と差分が出ていた API 名・呼び出し順を最新に追従。残リスクと運用検討事項を整理する。
todos:
  - id: review-merge-scope
    content: bundled_session_ui_defaults.json の全キーをセッションにマージ上書きしているため、ユーザーがカスタマイズした設備ガント／バッジ等もアップグレードで失う。正本に含めたいキーだけに絞るか、別フラグでオプトアウトするかを検討する
    status: pending
  - id: review-table-column-overwrite
    content: table-column-order.json を全置換している。列幅カスタムを温存したい要望があれば、マージ戦略またはバージョンキー付きの部分上書きを検討する
    status: pending
  - id: review-open-tabs-ux
    content: 同期直後は既に開いている TableView の列が見た目で変わらない場合がある。必要ならユーザー向けREADMEに「タブの開き直し／再起動」を明記する
    status: pending
  - id: align-legacy-plan-doc
    content: 旧プラン fast_package_除外の再設計_d181c2f0.plan.md は既に削除済み。後継として 2026-05-10 に fast_package_zip_artifacts_exclude.plan.md を起票し、ZIP 巨大化（hprof / 入れ子 .venv 等）と Step 8 並列・リトライ等のパッケージ高速化系をそちらに集約した
    status: completed
  - id: sync-plan-with-code
    content: 2026-05-10 の実装に追従し、Java 側の同期成功後フローを「InitSettingPersistence.applyPortableUpgradeOverwriteFromPmAiData → applyPortableUpgradeBundledPolicyToSessionStore → overwriteTableColumnOrderStoreAfterPortableUpgrade(ui) → resetEnvRowsToDefaults → applyBundledPortableDefaultsIfPresent → applyDesktopSession(load(),false) → applyRepoFolderPathNormalization → applyUncNetworkSourceDirDefaults → DesktopSessionStateStore.save」に更新
    status: completed
isProject: false
---

# ポータブル：タブ設定・配台不要ルールのバンドル同梱とアップグレード上書き（実装済みの整理）

## 背景・目的

- **初回インストーラー**に、メインシェルタブ順・エイリアス・タブオーガナイザー周り、設備ガント／バッジ既定などを **`bundled_session_ui_defaults.json`** として同梱する。
- **配台不要ルール**を **`code/exclude_rules.json`** として同梱し、環境タブの `PM_AI_EXCLUDE_RULES_JSON` ブートストラップ（`MainShellController.maybeFillEmptyBootstrap`）が実ファイルを指せるようにする。
- **バージョンアップ用ZIP**にも同じファイル群を含め、正本から `pm-ai-data` へ同期したあと、**ユーザー別の `~/.pm-ai-desktop` 側の該当設定を上書き**して製品既定に追従させる。

## 実装サマリ（コード参照）

### パッケージ（PowerShell）

- **ファイル**: リポジトリ直下 [`fast_package_app.ps1`](../../fast_package_app.ps1) の `Copy-BundleToDist`
- **条件**: `$BundleKind -eq 'InitialInstall' -or $BundleKind -eq 'VersionUpgrade'`
- **処理**:
  1. `code_java/src/main/resources/jp/co/pm/ai/desktop/config/bundled_session_ui_defaults.json` → `pm-ai-data/config/`
  2. 同パスの `bundled_table_column_order.json` → `pm-ai-data/config/`
  3. `pm-ai-data/code/json/stage1_exclude_rules.json` があれば → `pm-ai-data/code/exclude_rules.json`、無ければ `bundled_exclude_rules.json`（ASCII の空 `rules`）をフォールバック
- **ドキュメント**: `pm-ai-data/README_PORTABLE.txt` に Initial / Upgrade 双方の説明と「同期後に Java がホーム配下を上書きする」旨を記載

### ランタイム（Java）- バージョンアップ同期成功後

- **ファイル**: [`MainShellController.java`](../../code_java/src/main/java/jp/co/pm/ai/desktop/MainShellController.java)（`maybePortableBundleSelfUpdate` 内 `Task#setOnSucceeded`、2026-05-10 時点で 3444-3473 付近）
- **現行順序（実コード追従・引数省略）**:
  1. `InitSettingPersistence.applyPortableUpgradeOverwriteFromPmAiData(localData, ui)`
  2. `DesktopSessionStateStore.applyPortableUpgradeBundledPolicyToSessionStore(ui)`
  3. `TableColumnOrderPersistence.overwriteTableColumnOrderStoreAfterPortableUpgrade(ui)`
     - 内部で旧 API `overwriteTableColumnOrderStoreFromBundledAfterPortableUpgrade()` を呼んだあと、UI 側の追補処理を行う
  4. `resetEnvRowsToDefaults()`
  5. `applyBundledPortableDefaultsIfPresent()`
  6. `applyDesktopSession(DesktopSessionStateStore.load(), false)`
  7. `applyRepoFolderPathNormalization()`
  8. `applyUncNetworkSourceDirDefaults()`
  9. `DesktopSessionStateStore.save(collectDesktopSession())`
- **失敗時**: 上記 1〜3 のいずれかが `IOException` を投げた場合、`appendLog("[startup] バージョンアップ後のバンドル既定（タブ／列順／配台不要 JSON パス）の上書きに失敗: " + ex.getMessage())` を出してフローを継続する（同期そのものは成功扱い）。

### セッション上書きの中身

- **ファイル**: [`DesktopSessionStateStore.java`](../../code_java/src/main/java/jp/co/pm/ai/desktop/config/DesktopSessionStateStore.java)
- **内容**:
  - `readBundledSessionUiDefaultsNode()` と同じ解決順（`pm-ai-data/config/` → クラスパス）でバンドル JSON を読み、**ルートオブジェクトの各キーを** `session-state.json` に **上書きセット**
  - バンドルに `mainShellTabOrder` がある場合は **`mainShellTabLayout` をセッションから削除**（入れ子レイアウトが残ると新しいタブ順が効かないため）
  - `pm-ai-data/code/exclude_rules.json` が存在する場合、`excludeRulesPath` と `uiEnvRows` 内の `PM_AI_EXCLUDE_RULES_JSON` をその絶対パスに更新

### テーブル列順の上書き

- **ファイル**: [`TableColumnOrderPersistence.java`](../../code_java/src/main/java/jp/co/pm/ai/desktop/ui/TableColumnOrderPersistence.java)
- **公開 API**:
  - `overwriteTableColumnOrderStoreAfterPortableUpgrade(Map<String,String> ui)` … 現行フローの正面入口（1265 付近）
  - `overwriteTableColumnOrderStoreFromBundledAfterPortableUpgrade()` … 引数なしの旧 API（1252 付近）。上の `Map` 版から内部的に呼ばれる
- **内容**: `readBundledTableColumnOrderRoot()` の結果で `~/.pm-ai-desktop/table-column-order.json` を **全置換**

### 既知の互換対応（UTF-8）

- `bundled_session_ui_defaults.json` 内の一部ラベルは **Unicode エスケープ**で ASCII 化済み（環境によって JSON パースが失敗しないようにするため）。テスト: `BundledDesktopUiDefaultsResourceTest`

## 検討事項（このプランの TODO と対応関係）

| 論点 | 現状 | 検討の方向 |
|------|------|------------|
| 上書き範囲 | バンドル JSON に含まれるキーはすべてセッションへ反映（タブ以外のガント／バッジ既定も含む） | 製品として「アップグレードで必ず揃える」なら現状でよい。ユーザー保持を優先するなら **キーのホワイトリスト化**や **設定でスキップ** |
| 列順ファイル | ファイル全体上書き | 温存要件が出たら **差分マージ**または **スコープ別**（タブID単位）のマージ |
| UI の即時反映 | 既存タブの TableView は再構築されないことがある | README 追記、または同期後に特定タブへ `notify`／再バインド（コストと効果の見極め） |
| 旧プランとの整合 | `fast_package_除外の再設計_*.plan.md` は **削除済み**。後継として `fast_package_zip_artifacts_exclude.plan.md` を 2026-05-10 に新設し、ZIP 巨大化・hprof 混入・入れ子 `.venv` の除外漏れ・Step 8 並列化／open リトライ等を集約 | このプラン本体（UI 既定とアップグレード上書き）と **役割を分離して保守**する |

## 関連プラン

- [`fast_package_zip_artifacts_exclude.plan.md`](./fast_package_zip_artifacts_exclude.plan.md) … パッケージ ZIP 側の除外漏れ・性能・リトライ設計の正本。

## 関連コミット（参照用・履歴）

実装は複数コミットに分割されている。直近の意図を追う場合は次を参照する。

```bash
git log --oneline -- \
  fast_package_app.ps1 \
  code_java/package_workspace_copy.ps1 \
  code_java/src/main/java/jp/co/pm/ai/desktop/MainShellController.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/config/DesktopSessionStateStore.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/config/InitSettingPersistence.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/ui/TableColumnOrderPersistence.java
```

## 次のアクション（運用）

- リリースノートに「**バージョンアップ後はタブ順・配台不要JSON・列幅既定が製品同梱版に揃う**」ことを短く記載するか判断する。
- 社内で「ユーザー変更の温存」が要件に入るか確認し、入るなら TODO `review-merge-scope` / `review-table-column-overwrite` を起票して設計変更する。

## エンコーディング（プレビュー文字化け対策）

- 本ファイルは **UTF-8（改行 LF）** で保存する。`/mnt/c/` 上で Shift_JIS として保存すると Markdown プレビューで文字化けするため、エディタの「文字コード: UTF-8」を確認すること。
