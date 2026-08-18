# 同一化チェック結果の操作者別保存と閲覧タブ

**日付:** 2026-08-18  
**状態:** 実装完了（2026-08-18）

## 背景

同一化チェックの比較対象（アラジン入力用配台計画 Excel と、その時点のアラジン加工計画）を後から見返せない。  
操作者ごとにしばらく残し、閲覧用メインタブで確認できるようにする。

## 要件（確定）

| 項目 | 内容 |
|------|------|
| 保存タイミング | 同一化チェック実行のたび（一致／差異とも）。比較用の表が取れたときのみセット保存 |
| 保存内容 | 比較に使った Excel のコピー ＋ チェック時に読んだ加工計画の JSON スナップショット ＋ meta |
| JSON の出所 | チェック時にディスクから読んだ加工計画表をその場で JSON 化（`shaped_aladdin_plan.json` のコピーではない） |
| 保存先 | `{工場共有 DATA}/同一化チェック履歴/{sanitize(操作者)}/{yyyyMMdd-HHmmss}/` |
| 工場共有 DATA | `AppPaths.summarySharedDataDir`（操作ログ・サマリ Excel と同じ親） |
| 操作者 | ログイン中セッション名。ディレクトリ名は `OperatorUserPaths.sanitizeOperatorDirName` |
| 保持 | 操作者あたりタイムスタンプフォルダ **最新 20 件**（古い順に削除。アラジン入力 Excel 世代と同数） |
| 閲覧 | 新メインタブ。操作者コンボで共有上の操作者を切替（既定は自分） |
| エラー時 | 比較エラーで表が取れないときはセット保存しない（既存の操作ログ `identity_check` 記録は従来どおり） |

## 非対象

- 同一化チェックの比較ロジック変更（シス計のみ比較・完了行除外は別件で済み）
- 加工計画ソース xlsx 自体のコピー
- 90日日次削除（件数上限 20 のみ）
- リモート／他工場への同期

## フォルダ構成

```
{summarySharedDataDir}/同一化チェック履歴/
  {操作者}/
    20260818-172233/
      meta.json
      配台計画.xlsx
      加工計画.json
    20260818-180101/
      ...
```

### meta.json（例）

```json
{
  "savedAt": "2026-08-18T17:22:33+09:00",
  "operator": "細川",
  "result": "mismatch",
  "badgeText": "差異 9件",
  "diffCount": 9,
  "excelSourcePath": "...",
  "planSourcePath": "...",
  "excelFileName": "配台計画.xlsx",
  "planJsonFileName": "加工計画.json"
}
```

`result` は `ok` / `mismatch` / `error`（error でセット保存する場合は将来拡張。当面は ok/mismatch のみ保存）。

### 加工計画.json

既存の shaped 表と同じ配列表形式（`JsonTableIo.saveArrayTable` 相当: headers + rows）。  
閲覧時は同一形式で読めること。

## 保存処理

1. `AladdinEntryDispatchPlanIdentityCheck.evaluate` が Excel・加工計画表の読込に成功したあと、比較結果とともにスナップショット書き出しを行う（または `ResultDispatchTableTabController.finishAladdinEntryIdentityCheck` から専用 Store を呼ぶ）。
2. 書き出し内容:
   - 比較に使った Excel パスを `配台計画.xlsx` へコピー
   - 読込済み `TabularSheet` を `加工計画.json` へ保存
   - `meta.json` を書く
3. 当該操作者ディレクトリで prune（フォルダ数が 20 超なら古いタイムスタンプフォルダをディレクトリごと削除）
4. 失敗してもチェック結果ダイアログ表示は阻害しない（ログに警告）

正本クラス案: `IdentityCheckHistoryStore`（`OperatorActionLogStore` / 世代 prune と同系統）。

## 閲覧タブ

| 項目 | 内容 |
|------|------|
| タブ ID | `identityCheckHistory`（`MainShellTabId` 新規） |
| 表示名 | 同一化チェック履歴 |
| 配置 | `MainShellTabLayoutDefaults` の「その他」グループ（操作ログの直後） |
| UI | 操作者コンボ（既定＝ログイン中）・履歴一覧（日時・結果・差異件数）・選択時の詳細・「Excelを開く」「JSONを表で表示」 |
| タブ整理 | `MainShellInnerTabCatalog` は子タブが無いなら更新不要 |

`MainShell.fxml` / `MainShellController` にタブ実体を追加する（既存メインタブ追加手順に従う）。

## 実装の正本

| 役割 | ファイル |
|------|----------|
| パス | `AppPaths`（履歴ルート解決） |
| 保存・prune・一覧 | `IdentityCheckHistoryStore` |
| チェック後フック | `AladdinEntryDispatchPlanIdentityCheck` および／または `ResultDispatchTableTabController` |
| タブ UI | `IdentityCheckHistoryTab.fxml` / `IdentityCheckHistoryTabController` |
| タブ登録 | `MainShellTabId` / `MainShellTabLayoutDefaults` / `MainShell.fxml` / `MainShellController` |

## テスト方針

- Store: 一時ディレクトリでセット書き込み・20件 prune・操作者ディレクトリ分離
- meta / JSON が読めること
- タブは手動確認（一覧・操作者切替・Excel オープン）

## 承認メモ

- 保存: 毎回 / 保持: 20件 / 閲覧: 自分既定＋操作者切替 / JSON: チェック時スナップショット  
- 方式: 共有 DATA 配下のセットフォルダ（方式 A）
