---
name: ポータブル初回／バージョンアップのUI・配台不要既定
overview: "初回インストール用・バージョンアップ用の両ZIPに、タブ／ガント等のセッション既定JSON・テーブル列順・配台不要ルールJSONを同梱し、バージョンアップ同期後は ~/.pm-ai-desktop のセッションと列順ストアを正本に合わせて上書きする方針で実装済み。残リスクと運用検討事項を整理する。"
todos:
  - id: review-merge-scope
    content: "bundled_session_ui_defaults.json の全キーをセッションにマージ上書きしているため、ユーザーがカスタマイズした設備ガント／バッジ等もアップグレードで失う。正本に含めたいキーだけに絞るか、別フラグでオプトアウトするかを検討する"
    status: pending
  - id: review-table-column-overwrite
    content: "table-column-order.json を全置換している。列幅カスタムを温存したい要望があれば、マージ戦略またはバージョンキー付きの部分上書きを検討する"
    status: pending
  - id: review-open-tabs-ux
    content: "同期直後は既に開いている TableView の列が見た目で変わらない場合がある。必要ならユーザー向けREADMEに「タブの開き直し／再起動」を明記する"
    status: pending
  - id: align-legacy-plan-doc
    content: "既存プラン fast_package_除外の再設計_d181c2f0.plan.md の「未実装」記述と現状 fast_package_app.ps1（ルート配置・二系統バンドル等）の差分を整理し、プランをアーカイブまたは更新する"
    status: pending
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

- **ファイル**: [`MainShellController.java`](../../code_java/src/main/java/jp/co/pm/ai/desktop/MainShellController.java)（`maybePortableBundleSelfUpdate` 内 `Task#setOnSucceeded`）
- **順序**:
  1. `DesktopSessionStateStore.applyPortableUpgradeBundledPolicyToSessionStore()`
  2. `TableColumnOrderPersistence.overwriteTableColumnOrderStoreFromBundledAfterPortableUpgrade()`
  3. `applyBundledPortableDefaultsIfPresent()`
  4. `applyDesktopSession(DesktopSessionStateStore.load())`
  5. `DesktopSessionStateStore.save(collectDesktopSession())`

### セッション上書きの中身

- **ファイル**: [`DesktopSessionStateStore.java`](../../code_java/src/main/java/jp/co/pm/ai/desktop/config/DesktopSessionStateStore.java)
- **内容**:
  - `readBundledSessionUiDefaultsNode()` と同じ解決順（`pm-ai-data/config/` → クラスパス）でバンドル JSON を読み、**ルートオブジェクトの各キーを** `session-state.json` に **上書きセット**
  - バンドルに `mainShellTabOrder` がある場合は **`mainShellTabLayout` をセッションから削除**（入れ子レイアウトが残ると新しいタブ順が効かないため）
  - `pm-ai-data/code/exclude_rules.json` が存在する場合、`excludeRulesPath` と `uiEnvRows` 内の `PM_AI_EXCLUDE_RULES_JSON` をその絶対パスに更新

### テーブル列順の上書き

- **ファイル**: [`TableColumnOrderPersistence.java`](../../code_java/src/main/java/jp/co/pm/ai/desktop/ui/TableColumnOrderPersistence.java)
- **内容**: `readBundledTableColumnOrderRoot()` の結果で `~/.pm-ai-desktop/table-column-order.json` を **全置換**

### 既知の互換対応（UTF-8）

- `bundled_session_ui_defaults.json` 内の一部ラベルは **Unicode エスケープ**で ASCII 化済み（環境によって JSON パースが失敗しないようにするため）。テスト: `BundledDesktopUiDefaultsResourceTest`

## 検討事項（このプランの TODO と対応関係）

| 論点 | 現状 | 検討の方向 |
|------|------|------------|
| 上書き範囲 | バンドル JSON に含まれるキーはすべてセッションへ反映（タブ以外のガント／バッジ既定も含む） | 製品として「アップグレードで必ず揃える」なら現状でよい。ユーザー保持を優先するなら **キーのホワイトリスト化**や **設定でスキップ** |
| 列順ファイル | ファイル全体上書き | 温存要件が出たら **差分マージ**または **スコープ別**（タブID単位）のマージ |
| UI の即時反映 | 既存タブの TableView は再構築されないことがある | README 追記、または同期後に特定タブへ `notify`／再バインド（コストと効果の見極め） |
| 旧プランとの整合 | `fast_package_除外の再設計_*.plan.md` は一部「未実装」前提の記述が残る可能性 | ドキュメントの **更新またはクローズ** |

## 関連コミット（参照用・履歴）

実装は複数コミットに分割されている可能性がある。直近の意図を追う場合は `git log --oneline -- fast_package_app.ps1 DesktopSessionStateStore.java MainShellController.java TableColumnOrderPersistence.java` を参照。

## 次のアクション（運用）

- リリースノートに「**バージョンアップ後はタブ順・配台不要JSON・列幅既定が製品同梱版に揃う**」ことを短く記載するか判断する。
- 社内で「ユーザー変更の温存」が要件に入るか確認し、入るなら TODO `review-merge-scope` / `review-table-column-overwrite` を起票して設計変更する。

## エンコーディング（プレビュー文字化け対策）

- 本ファイルは **UTF-8（改行 LF）** で保存する。`/mnt/c/` 上で Shift_JIS として保存すると Markdown プレビューで文字化けするため、エディタの「文字コード: UTF-8」を確認すること。
