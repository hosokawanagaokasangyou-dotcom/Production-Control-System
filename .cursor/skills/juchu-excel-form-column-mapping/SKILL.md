---
name: juchu-excel-form-column-mapping
description: >-
  Maps 受注ファイル（Excel「受注ﾌｧｲﾙ」シート）見出し行と依頼書フォーム項目の対応を設計・修正する方法論。
  列定義ウィザード、JuchuHeaderAliasRegistry、見出し行変更、REDEFINE/別名/除外の運用。
  Use when implementing or fixing 受注シート列定義, header mismatch, 転記/吸出し, masterBase列,
  JuchuSheetColumnLayout, JuchuSheetHeaderRepairWizard, or Excel↔フォーム列対応。
---

# 受注 Excel ↔ 依頼書フォーム列対応

## 設計思想（必読）

| 原則 | 内容 |
|------|------|
| **列位置は固定** | 転記・読込は `JuchuSheetColumnLayout.Col` の **列 index（A, AP, AQ…）** が正本。Excel 上で見出しが別列に移っても、**物理列位置は変えない**。 |
| **見出しは検証用** | 行3（既定）の見出し文字列は「その列に何と書いてあるか」の検証。不一致は警告・ウィザードで解消。 |
| **Excel を書き換えない** | 受注 xlsm の見出しセルをアプリから直書きしない。期待定義の上書き・別名・除外で吸収。 |
| **ファイル別設定** | 別名・期待上書き・除外・見出し行は **受注ファイル絶対パス** 単位で `JuchuHeaderAliasRegistry` に保存。 |
| **工場別永続化** | ユーザーホーム `~/.pm-ai-desktop/request-form-juchu-header-aliases_<工場>.properties`。出荷既定は `init_setting/juchu_header_aliases_<工場>.json`。 |

## 正本ファイル

| 役割 | パス |
|------|------|
| 列 enum・検証・読込 | `code_java/.../reconciliation/JuchuSheetColumnLayout.java` |
| 別名・期待上書き・見出し行 | `code_java/.../reconciliation/JuchuHeaderAliasRegistry.java` |
| ウィザード UI | `code_java/.../reconciliation/JuchuSheetHeaderRepairWizard.java` |
| 転記・読込・警告 | `code_java/.../reconciliation/ReconciliationApp.java` |
| 起動・レジストリ注入 | `RequestFormInputTabController.java`, `MainShellController.java` |
| init_setting パス | `InitSettingPaths.juchuHeaderAliasesFileForFactory` |

## データモデル

### 既知列（`Col`）

- フォーム項目（例: 【製品】商品 masterBase = **AP 列**）と `primaryHeader` / 組込別名。
- `readDbValuesFromRow` は **列 index** から値を読む（見出し文字列に依存しない）。

### レジストリ種別（ファイル \| kind \| …）

| kind | 用途 |
|------|------|
| `expected` | 期待見出しの上書き（REDEFINE 結果） |
| `expectedPick` | REDEFINE 時に選んだ **`XX列: 見出し`** 表示ラベル（再表示用） |
| `alias` | 実際の Excel 見出しを許容する別名（複数可） |
| `exclude` | 転記/吸出しから除外する既知列 |
| `ignored` | 未知列（定義外 Excel 見出し）を検証対象外 |
| `headerRow` | 見出し行（**1-based**。既定 **3**） |

工場既定見出し行: JSON ルート `defaultHeaderRowOneBased` または properties `@default|headerRow`。

### 見出し行とデータ行

- **見出し行** … `registry.headerRowOneBasedFor(path)`（1-based）
- **先頭データ行** … 見出し行 + 1（0-based index）
- ウィザード・転記・`readMismatches` はすべてレジストリから行 index を解決すること。`HEADER_ROW_INDEX` 定数直参照は **テスト既定値** としてのみ。

## 不一致の解消（ウィザード）

起動: 転記前警告 `showTransferPrompt` / 手動 `showManage`（依頼書タブ）。

### 既知列 — 対応（FixAction）

| 操作 | 効果 |
|------|------|
| **期待定義をExcel見出しで再定義** (REDEFINE) | 採用見出しを `expected` に保存。**実際の列見出しと採用見出しが異なるときは実見出しを自動で `alias` 登録**（検証通過のため必須）。 |
| **実際の見出しを別名として許容** (ALIAS) | 当該列の Excel 実見出しのみ別名追加。 |
| **転記/吸出しから除外** (EXCLUDE) | その列は転記・読込・不一致一覧からスキップ。 |
| **対応しない** (SKIP) | 変更なし。 |

### 採用 Excel 見出し ComboBox

- 表示は **`XX列: 見出し`**（`ExcelHeaderPick.displayLabel()`）。
- 行モデルは **`selectedPickLabel`**（表示）と **`selectedExcelHeader`**（保存用文字列）を分離。
- 解決は **`resolvePick(comboValue, picks)`**: ① displayLabel 完全一致 → ② `XX列:` プレフィックスで列文字 → ③ 見出し文字のみ（同文複数列は列文字必須）。
- セル再利用対策: ComboBox は **`boundRow` + `syncingCombo`**。Apply 前に `knownTable.edit(-1, null)` と `commitKnownRowPickSelections`。

### 未知列

- **無視** / **既知列の別名として登録** / スキップ。

### 見出し行 UI

- Spinner（1–200、1始まり）+ **見出し行を反映** + Apply 時保存。
- 反映・Apply で `loadSheetContext(juchuFile, registry)` から見出し・候補一覧を再構築。

## 検証ロジック（`headerMatches`）

1. 列が **exclude** なら不一致収集から除外。
2. **`expected` override** あり: 実見出しが空 → OK。実見出し == override → OK。
3. 上記以外は **`Col.matchesHeader(actual, registry 別名)`**。

REDEFINE で「BU列: 商品(製品)」を選び AP 列実見出しが「タイプ」のとき:

- `expected` = `商品(製品)`
- 自動 `alias` = `タイプ`
- → AP 列は一致扱い。

## 実装・変更チェックリスト

```
- [ ] 見出し行を読む箇所は registry.headerRowIndexFor(path) 経由か
- [ ] データ行走査は firstDataRow = headerRowIndex + 1 か
- [ ] REDEFINE で actual ≠ 採用見出しのとき alias 自動登録が維持されているか
- [ ] resolvePick が列文字で一意解決できるか（同見出し複数列）
- [ ] ComboBox が boundRow パターンか（getIndex() 直参照禁止）
- [ ] 永続化: properties / init_setting JSON export に headerRow が含まれるか
- [ ] テスト: JuchuSheetColumnLayoutTest, JuchuHeaderAliasRegistryHeaderRowTest, ResolvePickTest
- [ ] Excel 見出しセルの直書きを追加していないか
```

## 典型的な業務シナリオ

### 見出しが別列に移った（AP=タイプ、BU=商品(製品)）

1. AP 行で **REDEFINE** + 採用 **BU列: 商品(製品)**。
2. Apply → expected=商品(製品)、alias=タイプ → 不一致解消。
3. 転記は引き続き **AP 列 index** から読書（列位置は不変）。

### 見出し行だけがファイルで異なる（例: 5 行目）

1. ウィザードで見出し行 **5** → **見出し行を反映**。
2. 列定義を再確認 → Apply。

### 列自体を転記したくない

- **EXCLUDE**（例: 原反「色」列が空で運用しない）。

## テスト実行

```bash
cd code_java && rtk ./mvnw test -Dtest=JuchuSheetColumnLayoutTest,JuchuHeaderAliasRegistryHeaderRowTest,JuchuSheetHeaderRepairWizardResolvePickTest
```

## 関連ルール

- `.cursor/rules/` … Git commit、UTF-8、agent debug（本スキルと重複記載しない）
- 依頼書・受注の VBA/Excel 側は `code/VBA/` を参照。コンパイル確認: `vba-compile-verify.mdc`

## 追加詳細

ファイル別キー形式・JSON 例・拡張手順は [reference.md](reference.md)。
