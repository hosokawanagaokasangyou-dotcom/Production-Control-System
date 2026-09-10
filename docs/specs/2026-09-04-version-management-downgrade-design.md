# リリース世代退避 ＋ バージョン管理タブ（ダウングレード対応）

**日付:** 2026-09-04  
**状態:** 設計承認済み（実装前）  
**関連:** 初期案 `2026-09-04-package-release-generations-design.md` を本仕様に統合・拡張

## 背景

1. `fast_package_app.ps1` Step 8 は同名成果物を削除して再生成するため、前バージョンが残らない。
2. 起動時ポータブル自動更新（`PortableBundleSelfUpdater.shouldUpdate`）は **正本がローカルより新しいときだけ** 動く。ダウングレード導線がない。

## 要件（確定）

### パッケージ側（世代退避）

| 項目 | 内容 |
|------|------|
| 方式 | 最新はルート固定名。旧成果物を `previous\v{version}\` へ退避 |
| 退避対象 | `version.txt`, `PMD_initial_install.zip`, `PMD_version_upgrade.zip`, `build-manifest.json`（存在時） |
| タイミング | Step 8 の同名削除直前 |
| 衝突 | 同名 `v{version}` がある場合は `v{version}_{yyyyMMdd-HHmmss}` |
| 保持数 | `-KeepGenerations` 既定 5。環境変数 `PM_AI_PACKAGE_KEEP_GENERATIONS`。`0` = 無制限 |
| スキップ | `-SkipArchive` で従来どおり削除のみ |

### デスクトップ側（バージョン管理タブ）

| 項目 | 内容 |
|------|------|
| 主目的 | 正本リリース直下の最新＋`previous\v*` を一覧し、選択版を適用（アップ／ダウン共通） |
| 確認 | ダウングレード（ローカル版 ≥ 選択版）のときだけ追加確認。アップは確認1回 |
| 正本解決 | `PM_AI_PORTABLE_BUNDLE_SOURCE_DIR` が ZIP → 親フォルダをリリース直下。フォルダ → その直下＋`previous\` |
| 適用経路 | 既存ポータブル更新を「指定正本で強制実行」するよう拡張（アプローチ 1） |
| タブ配置 | 「環境設定」グループ。キー `versionManagement` |
| 非ポータブル | タブは表示。適用不可の説明のみ |

## 方式

### A. `fast_package_app.ps1`

1. Step 8 開始時、`-SkipArchive` でなければ既存ルート成果物を `previous\v{旧}\` へ **移動**。
2. 従来どおりルートに新成果物を生成。
3. `-KeepGenerations` ≥ 1 なら `previous\` 直下を更新日時の新しい順で N 個残し超過を削除。

### B. バージョン管理タブ

1. **新タブ**: `MainShellTabId.VERSION_MANAGEMENT`（`versionManagement`）。FXML＋Controller。`MainShellTabLayoutDefaults`（`DEFAULT_FLAT` 末尾＋`groupedLayout` の環境設定）を更新。
2. **一覧データ**: リリース直下を解決 → ルート最新（`version.txt`＋Upgrade ZIP）と `previous\v*` 各フォルダを行として表示（版番号・フォルダ名・更新日時・ZIP 有無・現在適用中か）。
3. **適用**: 選択行のディレクトリ（ルート最新ならリリース直下、世代なら `previous\v…`）を正本フォルダとし、その中の `PMD_version_upgrade.zip`（＋隣の `version.txt` / `build-manifest.json`）を既存フローに渡す。
4. **強制適用 API**: `PortableBundleSelfUpdateService`（または同等）に「版比較をスキップして指定正本で同期＋終了後 exe 適用」を追加。起動時自動更新の「新しいときだけ」は変更しない。
5. **確認ダイアログ**:
   - 選択版 > ローカル: 既存に近い1回確認
   - 選択版 ≤ ローカル: 二重確認（ダウングレード警告）の後に適用

### C. 正本パス解決（共通ヘルパ推奨）

```
resolveReleaseRoot(sourceDirOrZip):
  if path is .zip file -> parent directory
  else if path is directory -> that directory
```

一覧・適用の両方で同一ロジックを使う。

## 構成（主な変更ファイル）

| 役割 | ファイル |
|------|----------|
| パッケージ退避 | `fast_package_app.ps1` |
| タブ ID・既定レイアウト | `MainShellTabId.java`, `MainShellTabLayoutDefaults.java`, `MainShell.fxml`, `MainShellController.java` |
| タブ UI | 新規 `VersionManagementTab.fxml` / `VersionManagementTabController.java`（名前は実装時に既存命名に合わせる） |
| 強制適用 | `PortableBundleSelfUpdateService.java` / `PortableBundleSelfUpdater.java`（最小差分で拡張） |
| 正本解決ヘルパ | 既存 `AppPaths` 近傍または `PortableBundleSelfUpdater` 内 |
| テスト | FXML ロード系（既存 `*FxmlTest` パターン）、版比較／正本解決の単体テスト |

## エラー方針

| 状況 | 動作 |
|------|------|
| パッケージ: 退避移動失敗 | Stop（ZIP 上書き前に失敗） |
| パッケージ: 剪定削除失敗 | Warning、最新生成は続行 |
| タブ: 正本パス空／開けない | 一覧空＋案内メッセージ、適用ボタン無効 |
| タブ: 選択世代に Upgrade ZIP なし | 適用不可＋理由表示 |
| タブ: 適用中失敗 | 既存ポータブル更新と同様にログ＋ダイアログ |

## テスト

**パッケージ（手動）**

1. 既存成果物ありで実行 → `previous\v旧\` に退避、ルートに新版
2. Keep=5 で超過剪定
3. `-SkipArchive` で退避なし

**デスクトップ**

1. 正本 ZIP パス時、親の `previous\` が一覧に出る
2. 新しい版を適用 → 確認1回 → 再起動フロー
3. 古い版を適用 → 二重確認 → 再起動フロー
4. 非ポータブル起動 → 適用不可メッセージ

## 非対象（明示）

- RDP ランチャー（`fast_package_rdp_launcher.ps1` / `rpa_luncher_release`）の世代管理
- 任意 ZIP の手動パス指定 UI（以前の案 B）
- Git への ZIP コミット
- 起動時自動更新でのダウングレード（起動時は従来どおりアップのみ）
- `Cash_PMD` キャッシュの世代管理

## パラメータ・環境変数（パッケージ）

| 名前 | 既定 | 説明 |
|------|------|------|
| `-KeepGenerations` | `5` | 残す世代数。`0` = 無制限 |
| `PM_AI_PACKAGE_KEEP_GENERATIONS` | （未設定時パラメータ既定） | 同上 |
| `-SkipArchive` | off | 退避せず削除のみ |

優先: 明示パラメータ > 環境変数 > 既定。
