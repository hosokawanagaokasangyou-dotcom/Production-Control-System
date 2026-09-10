# 依頼書入力：TPI PDF 一覧フィルタ

## 目的

依頼書入力の「検索・絞り込み」ラジオに **TPI PDF** を追加し、`PM_AI_REQUEST_FORM_TPI_PDF_DIR` 由来／関連の行だけを一覧表示する。

## 対象行（方針 2）

次のいずれかを満たす `OrderRecord`：

1. `原本種別` = `TPI_PDF`（TPI フォルダ取込）
2. ステータス文字列に `TPI PDF` を含む（例: `既存登録 (相違あり・TPI PDF)`）

## UI

- ラベル: `TPI PDF`
- ツールチップ: TPI 依頼書フォルダ（`PM_AI_REQUEST_FORM_TPI_PDF_DIR`）由来・関連付け行のみ表示
- 配置: 既存 2×2 の下に 3 行目（単独）、または 2 行目に収まるなら右隣
- TPI 無効工場（国分など）ではラジオを非表示

## 実装要点

- `RecordListFilterMode.TPI_PDF`
- `recordIncludedInListFilter` に分岐追加
- `isTpiRelatedRecord(OrderRecord)` 静的ヘルパ（テスト可能）
- テキスト検索は既存どおりモード適用後に併用

## テスト

- `ReconciliationAppRecordFilterTest` に TPI モードの包含／除外を追加
