# 操作ログと段階2後の同一化未整合終了ゲート

**日付:** 2026-08-17  
**状態:** 設計確定（実装前）

## 背景

段階2のあと、配台計画 Excel とアラジン加工計画が揃わないままアプリを閉じると、入力漏れに気づきにくい。  
終了を止める警告と、配台まわりの重要操作を操作者別に残す監査が必要である。

既存の `remote_log`（段階終了時の実行ログアーカイブ、保持3日）とは別系統とする。

## 要件（確定）

| 項目 | 内容 |
|------|------|
| 記録対象 | 段階2完了、同一化チェック、Excel出力、終了警告（表示時点） |
| 保存先 | `{工場共有 DATA}/操作ログ/{操作者}/{yyyy-MM-dd}.ndjson` |
| 工場共有 DATA | `AppPaths.summarySharedDataDir`（サマリ Excel と同じ親） |
| 操作者 | ログイン中セッション名。ディレクトリ名は `OperatorUserPaths.sanitizeOperatorDirName` |
| 閲覧 | 新メインタブ「操作ログ」。操作者コンボで共有上の全操作者を切替（既定は自分） |
| 保持 | **90日**（ファイル最終更新が90日超の ndjson を削除） |
| 終了ゲート条件 | **この起動で段階2が正常完了したあと**のみ |
| 終了時比較 | ローカル最新 `アラジン入力用_配台計画.xlsx` とソース最新の加工計画を再比較 |
| 警告解除 | ダイアログに表示したランダム7桁を入力。キャンセル・✕・Esc では閉じられない |
| 再起動 | 段階2未完了の起動ではゲートしない |

## 非対象

- 実行・ログタブ全文の複製
- ボタン押下・タブ切替など UI 全操作の計装
- ログの暗号化
- 環境変数キーの新規追加（共有 DATA 既存解決を使う）
- 段階2.1 / 段階3 完了をゲート条件にすること
- 世代フォルダ側 Excel との終了時比較（終了時はローカル最新のみ）

## 保存フォーマット

NDJSON。1操作1行。UTF-8。秘密情報・個人情報を `detail` に入れない。

| フィールド | 型 | 意味 |
|------------|----|------|
| `ts` | string | ローカル時刻の ISO-8601 |
| `operator` | string | 操作者名（サニタイズ前の表示名） |
| `action` | string | `stage2_complete` / `identity_check` / `excel_export` / `close_warning` |
| `result` | string | `ok` / `mismatch` / `error` / `shown` |
| `detail` | string | 短い説明（差異件数、Excel ファイル名、エラー要約） |

`close_warning` の `result` は表示時点で常に `shown`。7桁の入力成功は記録しない（表示したことが監査対象）。

## アーキテクチャ

```
段階2 正常完了
  ├─ セッションフラグ stage2CompletedThisLaunch = true
  └─ OperatorActionLogStore.append(stage2_complete, ok)

同一化チェック完了
  └─ append(identity_check, ok|mismatch|error)

アラジン入力用 Excel 出力完了
  └─ append(excel_export, ok|error)

主窓 close / 終了確認の直前
  ├─ フラグなし → 通常の終了確認
  ├─ フラグあり → evaluate(ui, ローカル最新 xlsx)
  │    ├─ identical → 通常の終了確認
  │    └─ 非同一（ファイル無し・比較失敗含む）
  │         ├─ 7桁ダイアログ表示
  │         ├─ append(close_warning, shown)  ※表示時点
  │         └─ 数字一致後 → 通常の終了確認
  └─ ログ書き込み失敗は実行・ログに1行。終了は止めない
```

比較ロジックは既存 `AladdinEntryDispatchPlanIdentityCheck.evaluate(ui, path)` を使う。  
パスは `AppPaths.aladdinEntryDispatchPlanLocalXlsxPath`。

## 7桁ダイアログ

- 範囲: `1000000`〜`9999999`（7桁、先頭0なし）
- 同じダイアログに数字を大きく表示し、数字のみの入力欄を置く
- キャンセルボタンなし。`setOnCloseRequest` で consume。Esc 無効
- 不一致: 入力を空にして再入力。アプリは閉じない
- 一致: ダイアログを閉じ、既存の「終了確認」へ進む
- 表示した瞬間に操作ログへ1件（入力完了を待たない）

## 専用タブ

- `MainShellTabId` 新規（key 例: `operatorActionLog`、見出し「操作ログ」）
- `DEFAULT_FLAT_TAB_KEY_ORDER` の末尾（タブ整理の直前）に追加
- `groupedLayout()` の「その他」グループへ追加
- `MainShell.fxml` / `MainShellController` の `@FXML`・`mainShellTabFor` と整合
- UI: 操作者コンボ、再読みボタン、表（日時 / 操作 / 結果 / 詳細）
- 操作者一覧: `操作ログ` 直下のディレクトリ名。ログイン中操作者が一覧に無ければ先頭に足す
- 表示前に90日超ファイルを削除してから読む。日付ファイル名の新しい順、ファイル内は新しい行を上

## 失敗時

| 状況 | 動作 |
|------|------|
| 共有フォルダが無い・書けない | 実行・ログに1行。ゲートは通常どおり。タブは空＋理由 |
| ローカル Excel 無し | 非同一として7桁ダイアログ |
| 加工計画ソース無し・比較例外 | 非同一として7桁ダイアログ |
| 7桁不一致 | 再入力。終了しない |
| 段階2未完了 | ゲートなし |

## テスト

- 操作者別パス解決と日次ファイル名
- 90日超 ndjson の削除、90日以内は残す
- 段階2フラグなし → ゲートしない
- フラグあり＋比較不一致／Excel無し → 要7桁
- フラグあり＋比較一致 → ゲートしない
- 7桁は一致のみ通過
- ダイアログ表示で `close_warning` / `shown` が1件付く

## 構成（実装時）

- 新規: `OperatorActionLogStore`（追記・削除・一覧）
- 新規: `Stage2IdentityCloseGate`（フラグ＋要7桁判定）
- 新規: 7桁確認ダイアログ
- 新規: 操作ログタブ FXML / Controller
- 変更: `MainShellController`（段階2完了・終了フック）
- 変更: `ResultDispatchTableTabController`（同一化チェック・Excel出力の記録）
- 変更: `MainShellTabId` / `MainShellTabLayoutDefaults` / `MainShell.fxml`
