# 実装計画: 加工トレンド Excel 固定出力

## ファイル

| ファイル | 役割 |
|----------|------|
| `ProcessingTrendExcelExportStore.java` | 出力先解決・xlsx 全削除・最新検出 |
| `ProcessingTrendExcelExportStoreTest.java` | 上記の単体テスト |
| `ProcessingTrendTabController.java` | FileChooser 廃止・Store 利用・開くボタン |
| `ProcessingTrendTab.fxml` | `openExcelButton` 追加 |
| `ProcessingTrendTabFxmlTest.java` | 開くボタン存在確認 |
| `.gitignore` | `.pm-ai-cache/exports/` を除外 |

## タスク

1. Store の失敗テスト → 実装
2. FXML / Controller 配線
3. テスト実行・commit/push
