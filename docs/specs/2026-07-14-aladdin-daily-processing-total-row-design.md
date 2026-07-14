# アラジン入力用 Excel「日加工合計数」行

**日付:** 2026-07-14  
**状態:** 実装完了

## 背景

結果_配台表タブの次ボタンが生成するアラジン入力用配台計画 Excel に、見出し直下へ「日加工合計数」行を追加する。

- アラジン入力用Excel出力
- アラジン加工計画読込→Excel出力
- ローカルへ出力

3ボタンはいずれも `DispatchAladdinEntryWorkbookExporter` を通るため、Exporter の変更で一括反映する。

## 要件（確定）

| 項目 | 内容 |
|------|------|
| 合算単位 | **機械シートごと**。当該シートのデータ行について、日付列ごとに（現アラ計）・（シス計）をそれぞれ合算 |
| 行位置 | 各機械シートの **2 行目**（1 行目=見出し、3 行目以降=データ） |
| A 列 | 文言 `日加工合計数` |
| B〜J 列 | 空 |
| 日付列表示 | 既存と同じ 2 段文字列（`（現アラ計）…` / `（シス計）…`）。両方が誤差範囲内で 0 の日は空セル |
| 不一致色 | **付けない**（合計行は常に通常の日付セルスタイル） |
| フリーズ | 見出し＋合計の **2 行**（`createFreezePane(FIXED_COLUMN_COUNT, 2)`） |
| 印刷タイトル行 | 見出し＋合計の **2 行**（`setRepeatingRows(0, 1)`） |

## 非対象

- 日付セルの数値化や Excel `SUM` 数式化（現状の 2 段文字列セルを維持）
- Builder モデル（`MachineSheet` 等）への合計フィールド追加（本設計では Exporter 側で合算）
- 固定列（換算数量・配台合計など）の合算表示
- UI 文言・ボタン追加

## アーキテクチャ

```
ResultDispatchTableTabController
  └─ DispatchAladdinEntryWorkbookExporter.writeMachineSheet
        ├─ row0: 見出し（既存）
        ├─ row1: 日加工合計数（新規・シート内合算）
        └─ row2…: EntryRow（既存・行番号が +1）
```

合算は Exporter（または同パッケージの小さな静的ヘルパ）で行う。

```text
for each LocalDate d in dates:
  aladdinSum = Σ entry.cells().get(d).aladdinQty()   // null は 0
  systemSum  = Σ entry.cells().get(d).systemQty()
  cell = new EntryCell(aladdinSum, systemSum)
  描画は既存データ行と同じ（cellText / dateCellRichText）
  ただし styles.dateCellFor(d, mismatch=false) 固定
```

## 変更ファイル（想定）

| ファイル | 変更 |
|----------|------|
| `DispatchAladdinEntryWorkbookExporter.java` | 合計行挿入、データ行開始を 2 に、フリーズ／印刷タイトルを 2 行化。必要なら合算ヘルパを同クラス private static に |
| `DispatchAladdinEntryWorkbookExporterTest.java` | 合計行のラベル・日付合算・空セル・不一致色なし・フリーズ／RepeatingRows を検証 |

（任意）合算ロジックをテストしやすいよう package-private ヘルパを切り出す場合は、同テストまたは隣接テストクラスでカバーする。

## エラー処理・境界

- データ行 0 件の機械シート: 合計行は出してよい。日付セルはすべて空。
- 空ブック（データなしメッセージシート）: 現状どおり。合計行は追加しない。
- AutoFilter: 現状どおり見出し行に設定。合計行が見出し直後にあるため、フィルタ操作時に合計行も対象になり得る（Excel の制約。本件では許容）。

## テスト方針

1. 同一シートに複数 `EntryRow`、同一日に異なる現アラ計／シス計 → 日付セル文字列が合算結果と一致
2. 全日 0 の日 → 空文字
3. 合算結果が不一致でも日付セルスタイルは mismatch 用色を使わない
4. A1=固定見出し、A2=`日加工合計数`、最初のデータ行は 3 行目（0-based index 2）
5. FreezePane の分割行・`getRepeatingRows` が 0〜1 を指す

## 成功基準

- 上記 3 ボタンいずれで出力しても、各機械シートに「日加工合計数」行があり、シート内日次合算が正しい
- スクロール・印刷プレビューで見出し＋合計が固定／繰り返しされる
- 既存の日付 2 段表示・不一致色（データ行）・固定列の挙動を壊さない
