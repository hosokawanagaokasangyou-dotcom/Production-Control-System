# fast_package_app.ps1 リリース成果物の世代管理

**日付:** 2026-09-04  
**状態:** 統合先へ移行（正本は `2026-09-04-version-management-downgrade-design.md`）

## 背景

`fast_package_app.ps1` の Step 8 は、`pm-ai-package-release`（または `-PackageReleaseDir` 先）の同名成果物を削除してから再生成する。前バージョンが残りず、ロールバック用の手元コピーが失われる。

## 要件（確定）

| 項目 | 内容 |
|------|------|
| 方式 | 案 A: 最新はルート固定名のまま。旧成果物を `previous\v{version}\` へ退避 |
| 退避対象 | `version.txt`, `PMD_initial_install.zip`, `PMD_version_upgrade.zip`, `build-manifest.json`（存在時） |
| 最新（配布用） | ルートの固定名を従来どおり上書き生成。配布パスは変更しない |
| 退避タイミング | Step 8 で同名削除する直前（ZIP 生成前） |
| バージョン衝突 | 同名 `v{version}` が既にある場合は `v{version}_{yyyyMMdd-HHmmss}` |
| 保持数 | `-KeepGenerations`（既定 5）。環境変数 `PM_AI_PACKAGE_KEEP_GENERATIONS` |
| 保持数 0 | 剪定なし（無制限） |
| 退避スキップ | `-SkipArchive`（従来どおり削除のみ） |
| 剪定順 | `previous\` 直下フォルダを更新日時の新しい順で N 個残し、超過分を削除 |

## 方式

1. Step 8 開始時、ルートに退避対象のいずれかが存在すれば退避処理を実行する（`-SkipArchive` 時はスキップ）。
2. 旧バージョン文字列は既存の `version.txt` 先頭行から取得（無い／読めない場合は `unknown` + タイムスタンプ）。
3. 退避先ディレクトリを作成し、存在する対象ファイルを **移動**（Move-Item）する。
4. 退避後、従来どおりルートに新 `version.txt` / ZIP / `build-manifest.json` を生成する。
5. `-KeepGenerations` が 1 以上なら、`previous\` 直下の世代フォルダを剪定する。

## パラメータ・環境変数

| 名前 | 既定 | 説明 |
|------|------|------|
| `-KeepGenerations` | `5` | 残す世代数。`0` = 無制限 |
| `PM_AI_PACKAGE_KEEP_GENERATIONS` | （未設定時はパラメータ既定） | 同上の環境変数上書き |
| `-SkipArchive` | off | 退避せず従来の削除のみ |

優先順位: 明示パラメータ > 環境変数 > 既定。

## 構成

- 変更: リポジトリ直下 `fast_package_app.ps1`（正本）
- `code_java/fast_package_app.ps1` は転送ラッパーのため変更不要
- README 文言（`Copy-BundleToDist` 内 Step 8 説明）を世代退避に合わせて更新

## エラー方針

| 状況 | 動作 |
|------|------|
| 退避先への移動失敗（ロック等） | Stop（`$ErrorActionPreference = 'Stop'`）。ZIP 上書き前に失敗させる |
| `previous\` が無い | 作成して続行 |
| 退避対象が 1 つも無い（初回） | 何もしない |
| 剪定時の削除失敗 | Warning を出し続行（最新生成は成功させる） |

## テスト

**手動**

1. 既存成果物がある状態でパッケージ実行 → `previous\v{旧}\` に ZIP 等が移り、ルートに新成果物
2. 6 回以上繰り返す（既定 Keep=5）→ `previous\` が最大 5 世代
3. `-KeepGenerations 0` → 剪定されない
4. `-SkipArchive` → 従来どおりルート同名を削除して再生成、`previous\` に増えない

## 非対象（明示）

- 配布クライアント側の自動ロールバック UI
- Git への ZIP コミット
- `rpa_luncher_release` / `fast_package_rdp_launcher.ps1` の世代管理
- `Cash_PMD` キャッシュの世代管理
