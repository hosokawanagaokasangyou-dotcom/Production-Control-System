# アラジン入力用 Excel — ローカル最新／世代を開くボタン

**日付:** 2026-07-14  
**状態:** 設計承認済み（実装前レビュー待ち）

## 背景

結果_配台表タブには共有側の「最新を開く」「世代を開く…」と「ローカルへ出力」があるが、**ローカル出力先のファイルを開く UI がない**。ローカル最新とローカル世代を開くボタンを追加する。

## 要件（確定）

| 項目 | 内容 |
|------|------|
| ボタン1 | 文言 **ローカル最新を開く**。`AppPaths.aladdinEntryDispatchPlanLocalXlsxPath` を読取専用で開く |
| ボタン2 | 文言 **ローカル世代を開く…**。世代ダイアログをローカルルート（`aladdinEntryDispatchPlanLocalDir`）で表示 |
| 配置 | `ローカルへ出力` の直後（共有の「最新／世代」の前） |
| ファイル無し | 警告ダイアログ（共有側「最新を開く」と同型。先に「ローカルへ出力」を促す） |
| 共有側 | 「最新を開く」「世代を開く…」の挙動・注意グローは変更しない |

## 非対象

- ローカル出力完了後のグロー誘導（共有側と同等の誘導を付けるかは本仕様に含めない）
- ボタン文言の国際化・ショートカットキー
- フォルダを開く専用ボタンの新規追加（世代ダイアログ内の「フォルダを開く」でローカルルートを開ける）

## アーキテクチャ

```
ResultDispatchTableTab.fxml
  └─ ローカルへ出力
  └─ ローカル最新を開く  → open local latest xlsx
  └─ ローカル世代を開く… → GenerationDialog(LOCAL root)
  └─ 最新を開く / 世代を開く…（既存・SHARED）
```

`DispatchAladdinEntryGenerationDialog` は現状 `AppPaths.aladdinEntryDispatchPlanDir`（共有）固定。  
オーバーロードでルート `Path`（または `Destination`）とタイトル接尾辞を受け取り、一覧走査先を切替える。既存 `show(owner, ui, defaultOperator)` は共有ルートを渡す互換ラッパとする。

## 変更ファイル（想定）

| ファイル | 変更 |
|----------|------|
| `ResultDispatchTableTab.fxml` | 2 ボタン追加 |
| `ResultDispatchTableTabController.java` | `@FXML` フィールド／ハンドラ追加 |
| `DispatchAladdinEntryGenerationDialog.java` | ルート指定可能な `show` |
| （任意）既存テスト or 小テスト | ダイアログルート解決の単体があれば Destination 切替をカバー。UI ハンドラは結合よりパス解決の再利用でよい |

## 成功基準

- 「ローカルへ出力」後に「ローカル最新を開く」で当該 xlsx が開く
- 「ローカル世代を開く…」でローカル側操作者フォルダ配下の世代が一覧される
- 共有側の最新／世代開くは従来どおり共有パスを参照する
