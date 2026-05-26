# 受注 Excel ↔ フォーム列対応 — リファレンス

## properties キー例（UTF-8）

```properties
# 工場既定見出し行（1-based）
@default|headerRow=3

# ファイル別見出し行
C\:\\data\\juchu.xlsm|headerRow=5

# 期待見出し上書き（AP = MASTER_BASE_SHOHIN_PRODUCT）
C\:\\data\\juchu.xlsm|expected|MASTER_BASE_SHOHIN_RAW=商品(原反)

# 別名（区切り \u0001）
C\:\\data\\juchu.xlsm|alias|MASTER_BASE_SHOHIN_PRODUCT=タイプ

# 転記除外
C\:\\data\\juchu.xlsm|exclude|IRO=1

# 未知列無視（列文字）
C\:\\data\\juchu.xlsm|ignored|BX=1
```

## init_setting JSON 例

```json
{
  "defaultHeaderRowOneBased": 3,
  "entries": {
    "C:\\data\\juchu.xlsm|headerRow": "5",
    "C:\\data\\juchu.xlsm|expected|MASTER_BASE_SHOHIN_PRODUCT": "商品(製品)",
    "C:\\data\\juchu.xlsm|alias|MASTER_BASE_SHOHIN_PRODUCT": "タイプ"
  }
}
```

## 主要 API 一覧

| API | 用途 |
|-----|------|
| `JuchuSheetColumnLayout.collectHeaderMismatches` | 不一致一覧 |
| `JuchuSheetColumnLayout.collectAllKnownColumns` | ウィザード既知列一覧 |
| `JuchuSheetColumnLayout.collectUnknownExcelColumns` | 定義外見出し |
| `JuchuSheetColumnLayout.readExcelHeaderPicks` | ComboBox 候補 |
| `JuchuSheetColumnLayout.headerMatches` | 1列の一致判定 |
| `JuchuSheetColumnLayout.resolveHeaderRowIndex` | 見出し行 0-based |
| `JuchuHeaderAliasRegistry.setExpectedOverride` | REDEFINE 保存 |
| `JuchuHeaderAliasRegistry.addAlias` | 別名追加 |
| `JuchuHeaderAliasRegistry.setHeaderRowOneBasedFor` | 見出し行 |
| `JuchuSheetHeaderRepairWizard.showTransferPrompt` | 転記前フロー |
| `JuchuSheetHeaderRepairWizard.resolvePick` | ComboBox → pick 解決 |

## 新規 `Col` 追加時

1. `JuchuSheetColumnLayout.Col` に列文字・primaryHeader・`formItemDescription` / `dbKey`。
2. 転記マッピング（`ReconciliationApp` の write/read 経路）に db キーを接続。
3. 既存ファイルでウィザードを開き不一致が出る場合は REDEFINE / ALIAS で解消。
4. 工場既定を更新する場合はグローバル設定タブから `init_setting` へ export。

## 既知の落とし穴

| 症状 | 原因 | 対処 |
|------|------|------|
| Apply 後も不一致 | REDEFINE のみで alias 未登録（実見出し≠採用見出し） | 自動 alias ロジックを維持する |
| 別列が選ばれる | `resolvePick` が見出し文字だけ先頭一致 | displayLabel / 列文字優先 |
| ComboBox 選択が別行に反映 | セル再利用 + `getIndex()` | boundRow パターン |
| 見出し行変更が効かない | `HEADER_ROW_INDEX` 直参照 | registry 経由に統一 |
| WSL で設定が見えない | 別ホーム / 工場 suffix | `FactorySite` と store パス確認 |

## グローバル設定連携

- `GlobalSettingsTabController` … 現在状態を init_setting へ
- `MainShellController.applyFactoryRequestFormGlobalSettings` … 工場切替時に import
- `RequestFormInputTabController.reloadJuchuHeaderAliasRegistry` … 依頼書タブでの reload
