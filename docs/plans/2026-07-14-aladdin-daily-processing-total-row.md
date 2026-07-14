# アラジン入力用 Excel「日加工合計数」行 Implementation Plan

> **For implementers:** Execute this plan task-by-task. Use the `test-driven-development` skill for each task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** アラジン入力用配台計画 Excel の各機械シート見出し直下に、シート内の日次（現アラ計）／（シス計）合算行「日加工合計数」を出し、見出し＋合計の 2 行をフリーズ／印刷タイトルにする。

**Architecture:** 合算は `DispatchAladdinEntryWorkbookExporter` 内の静的ヘルパで行い、`writeMachineSheet` が row1 に描画する（Builder モデルは変更しない）。3 ボタンは同一 Exporter のため追加 UI 変更は不要。日付セルは既存 `EntryCell.cellText()` と `dateCellRichText` を再利用し、合計行だけ `mismatch=false` 固定。

**Tech Stack:** Java 21、Apache POI（XSSF）、JUnit 5、Maven（`code_java/mvnw.cmd`）

**仕様正本:** `docs/specs/2026-07-14-aladdin-daily-processing-total-row-design.md`

---

## File structure

| ファイル | 責務 |
|----------|------|
| `code_java/.../io/DispatchAladdinEntryWorkbookExporter.java` | 合算ヘルパ、合計行描画、フリーズ／印刷タイトル 2 行化 |
| `code_java/.../io/DispatchAladdinEntryWorkbookExporterTest.java` | 合算・シート内容・フリーズ／RepeatingRows の検証 |

変更しない: `DispatchAladdinEntrySheetBuilder.java`、`ResultDispatchTableTabController.java`、FXML

---

### Task 1: 日次合算ヘルパ（単体）

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/io/DispatchAladdinEntryWorkbookExporter.java`
- Test: `code_java/src/test/java/jp/co/pm/ai/desktop/io/DispatchAladdinEntryWorkbookExporterTest.java`

- [ ] **Step 1: Write the failing test**

`DispatchAladdinEntryWorkbookExporterTest` に以下を追加する（ヘルパ未定義のためコンパイル失敗でよい）。

```java
@Test
void sumDateColumn_sumsAladdinAndSystemAcrossRows() {
    LocalDate d1 = LocalDate.of(2026, 7, 14);
    LocalDate d2 = LocalDate.of(2026, 7, 15);
    DispatchAladdinEntrySheetBuilder.EntryRow row1 =
            new DispatchAladdinEntrySheetBuilder.EntryRow(
                    "W1",
                    "",
                    "巻返し",
                    "",
                    "",
                    "",
                    0,
                    0,
                    0,
                    Map.of(
                            d1, new DispatchAladdinEntrySheetBuilder.EntryCell(100, 200),
                            d2, new DispatchAladdinEntrySheetBuilder.EntryCell(50, 0)),
                    d1,
                    2026);
    DispatchAladdinEntrySheetBuilder.EntryRow row2 =
            new DispatchAladdinEntrySheetBuilder.EntryRow(
                    "W2",
                    "",
                    "巻返し",
                    "",
                    "",
                    "",
                    0,
                    0,
                    0,
                    Map.of(d1, new DispatchAladdinEntrySheetBuilder.EntryCell(10, 20)),
                    d1,
                    2026);

    DispatchAladdinEntrySheetBuilder.EntryCell sum1 =
            DispatchAladdinEntryWorkbookExporter.sumDateColumn(List.of(row1, row2), d1);
    DispatchAladdinEntrySheetBuilder.EntryCell sum2 =
            DispatchAladdinEntryWorkbookExporter.sumDateColumn(List.of(row1, row2), d2);
    DispatchAladdinEntrySheetBuilder.EntryCell sumEmpty =
            DispatchAladdinEntryWorkbookExporter.sumDateColumn(List.of(row1, row2), LocalDate.of(2026, 7, 16));

    assertEquals(110d, sum1.aladdinQty(), 1e-9);
    assertEquals(220d, sum1.systemQty(), 1e-9);
    assertEquals(50d, sum2.aladdinQty(), 1e-9);
    assertEquals(0d, sum2.systemQty(), 1e-9);
    assertTrue(sumEmpty.isEmpty());
    assertEquals("", sumEmpty.cellText());
    assertEquals("（現アラ計）110\n（シス計）220", sum1.cellText());
}
```

既存テストファイル先頭の import に不足があれば追加する:

```java
import java.time.LocalDate;
```

- [ ] **Step 2: Run test to verify it fails**

Run（`code_java` ディレクトリで）:

```powershell
.\mvnw.cmd -q test "-Dtest=DispatchAladdinEntryWorkbookExporterTest#sumDateColumn_sumsAladdinAndSystemAcrossRows"
```

Expected: コンパイルエラー（`sumDateColumn` が未定義）またはテスト失敗。

- [ ] **Step 3: Write minimal implementation**

`DispatchAladdinEntryWorkbookExporter` に定数と package-visible ヘルパを追加する（クラス先頭付近の定数群の後ろ、`writeMachineSheet` の前が望ましい）。

```java
/** 機械シート 2 行目（見出し直下）のラベル。 */
static final String DAILY_PROCESSING_TOTAL_LABEL = "日加工合計数";

/**
 * シート内データ行について、指定日の（現アラ計）／（シス計）を合算する。
 * セルが無い／null の行は 0 として扱う。
 */
static DispatchAladdinEntrySheetBuilder.EntryCell sumDateColumn(
        List<DispatchAladdinEntrySheetBuilder.EntryRow> rows, LocalDate date) {
    double aladdin = 0;
    double system = 0;
    if (rows != null && date != null) {
        for (DispatchAladdinEntrySheetBuilder.EntryRow row : rows) {
            if (row == null || row.cells() == null) {
                continue;
            }
            DispatchAladdinEntrySheetBuilder.EntryCell cell = row.cells().get(date);
            if (cell == null) {
                continue;
            }
            aladdin += cell.aladdinQty();
            system += cell.systemQty();
        }
    }
    return new DispatchAladdinEntrySheetBuilder.EntryCell(aladdin, system);
}
```

- [ ] **Step 4: Run test to verify it passes**

```powershell
.\mvnw.cmd -q test "-Dtest=DispatchAladdinEntryWorkbookExporterTest#sumDateColumn_sumsAladdinAndSystemAcrossRows"
```

Expected: BUILD SUCCESS / tests pass.

- [ ] **Step 5: Commit**

```powershell
git add code_java/src/main/java/jp/co/pm/ai/desktop/io/DispatchAladdinEntryWorkbookExporter.java `
  code_java/src/test/java/jp/co/pm/ai/desktop/io/DispatchAladdinEntryWorkbookExporterTest.java
git commit -m "feat: アラジン入力Excelの日次合算ヘルパを追加"
```

---

### Task 2: 合計行のシート内容テスト（失敗させる）

**Files:**
- Test: `code_java/src/test/java/jp/co/pm/ai/desktop/io/DispatchAladdinEntryWorkbookExporterTest.java`

- [ ] **Step 1: Write the failing integration-style test**

同じテストクラスに追加する。`write(..., Destination.LOCAL)` で xlsx を書き、POI で開いて検証する。

```java
@Test
void write_insertsDailyProcessingTotalRowUnderHeader() throws IOException {
    Path repo = tempDir.resolve("repo");
    Files.createDirectories(repo.resolve("code"));
    Map<String, String> ui =
            Map.of(jp.co.pm.ai.desktop.config.AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());

    LocalDate dMatch = LocalDate.of(2026, 7, 14); // 火
    LocalDate dMismatch = LocalDate.of(2026, 7, 15); // 水
    LocalDate dEmpty = LocalDate.of(2026, 7, 16); // 木

    DispatchAladdinEntrySheetBuilder.EntryRow row1 =
            new DispatchAladdinEntrySheetBuilder.EntryRow(
                    "W7-4",
                    "C1",
                    "巻返し",
                    "",
                    "",
                    "",
                    1000,
                    0,
                    300,
                    Map.of(
                            dMatch, new DispatchAladdinEntrySheetBuilder.EntryCell(100, 100),
                            dMismatch, new DispatchAladdinEntrySheetBuilder.EntryCell(50, 200)),
                    dMatch,
                    2026);
    DispatchAladdinEntrySheetBuilder.EntryRow row2 =
            new DispatchAladdinEntrySheetBuilder.EntryRow(
                    "W7-5",
                    "C2",
                    "巻返し",
                    "",
                    "",
                    "",
                    2000,
                    0,
                    400,
                    Map.of(
                            dMatch, new DispatchAladdinEntrySheetBuilder.EntryCell(200, 200),
                            dMismatch, new DispatchAladdinEntrySheetBuilder.EntryCell(100, 100)),
                    dMatch,
                    2026);

    DispatchAladdinEntrySheetBuilder.EntryWorkbook model =
            new DispatchAladdinEntrySheetBuilder.EntryWorkbook(
                    List.of(dMatch, dMismatch, dEmpty),
                    List.of(
                            new DispatchAladdinEntrySheetBuilder.MachineSheet(
                                    "テスト機", List.of(row1, row2))));

    DispatchAladdinEntryWorkbookExporter.ExportResult result =
            DispatchAladdinEntryWorkbookExporter.write(
                    ui, model, DispatchAladdinEntryWorkbookExporter.Destination.LOCAL);

    try (XSSFWorkbook wb =
            new XSSFWorkbook(Files.newInputStream(result.latestPath()))) {
        org.apache.poi.ss.usermodel.Sheet sh = wb.getSheetAt(0);
        assertEquals("依頼NO", sh.getRow(0).getCell(0).getStringCellValue());
        assertEquals(
                DispatchAladdinEntryWorkbookExporter.DAILY_PROCESSING_TOTAL_LABEL,
                sh.getRow(1).getCell(0).getStringCellValue());
        assertEquals("", sh.getRow(1).getCell(1).getStringCellValue());
        assertEquals("W7-4", sh.getRow(2).getCell(0).getStringCellValue());

        // 日付列は FIXED_COLUMN_COUNT(=10) から
        assertEquals(
                "（現アラ計）300\n（シス計）300",
                sh.getRow(1).getCell(10).getStringCellValue());
        assertEquals(
                "（現アラ計）150\n（シス計）300",
                sh.getRow(1).getCell(11).getStringCellValue());
        assertEquals("", sh.getRow(1).getCell(12).getStringCellValue());

        // 合計行は不一致色なし: 黄系 FFF2CC でないこと
        org.apache.poi.xssf.usermodel.XSSFCellStyle totalStyle =
                (org.apache.poi.xssf.usermodel.XSSFCellStyle)
                        sh.getRow(1).getCell(11).getCellStyle();
        org.apache.poi.xssf.usermodel.XSSFColor fill = totalStyle.getFillForegroundXSSFColor();
        if (fill != null && fill.getRGB() != null) {
            byte[] rgb = fill.getRGB();
            assertFalse(
                    rgb[0] == (byte) 0xFF && rgb[1] == (byte) 0xF2 && rgb[2] == (byte) 0xCC,
                    "合計行に不一致色を付けてはならない");
        }

        // データ行の不一致セルは不一致色あり
        org.apache.poi.xssf.usermodel.XSSFCellStyle dataMismatchStyle =
                (org.apache.poi.xssf.usermodel.XSSFCellStyle)
                        sh.getRow(2).getCell(11).getCellStyle();
        org.apache.poi.xssf.usermodel.XSSFColor dataFill =
                dataMismatchStyle.getFillForegroundXSSFColor();
        assertNotNull(dataFill);
        byte[] dataRgb = dataFill.getRGB();
        assertEquals((byte) 0xFF, dataRgb[0]);
        assertEquals((byte) 0xF2, dataRgb[1]);
        assertEquals((byte) 0xCC, dataRgb[2]);

        assertEquals(2, sh.getPaneInformation().getHorizontalSplitTopRow());
        org.apache.poi.ss.util.CellRangeAddress repeating = sh.getRepeatingRows();
        assertNotNull(repeating);
        assertEquals(0, repeating.getFirstRow());
        assertEquals(1, repeating.getLastRow());
    }
}
```

必要 import:

```java
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
```

（既存でカバーされていれば重複追加しない。）

- [ ] **Step 2: Run test to verify it fails**

```powershell
.\mvnw.cmd -q test "-Dtest=DispatchAladdinEntryWorkbookExporterTest#write_insertsDailyProcessingTotalRowUnderHeader"
```

Expected: FAIL（現状 row1 が最初のデータ行 `W7-4`、フリーズ分割行が 1、RepeatingRows が 0 のみ）。

- [ ] **Step 3: Implement writeMachineSheet + applyPrintSetup**

`writeMachineSheet` を次の意図で変更する。

1. 見出し行（row 0）の直後、データ行ループの前に合計行を書く。
2. データ行開始を `r = 2` にする（現状 `r = 1`）。
3. `createFreezePane(FIXED_COLUMN_COUNT, 2)` にする。
4. `applyPrintSetup` の `setRepeatingRows` を `new CellRangeAddress(0, 1, -1, -1)` にする。

合計行描画の挿入イメージ（見出し作成直後）:

```java
Row totalRow = sh.createRow(1);
totalRow.setHeightInPoints(33f);
writeFixedCell(totalRow, 0, DAILY_PROCESSING_TOTAL_LABEL, styles.data());
for (int c = 1; c < FIXED_COLUMN_COUNT; c++) {
    writeFixedCell(totalRow, c, "", styles.data());
}
for (int i = 0; i < dates.size(); i++) {
    LocalDate d = dates.get(i);
    DispatchAladdinEntrySheetBuilder.EntryCell ec =
            sumDateColumn(machineSheet.rows(), d);
    Cell cell = totalRow.createCell(FIXED_COLUMN_COUNT + i);
    if (ec.isEmpty()) {
        cell.setCellValue("");
        cell.setCellStyle(styles.dateCellFor(d, false));
    } else {
        String cellText = ec.cellText();
        cell.setCellStyle(styles.dateCellFor(d, false)); // 不一致色なし
        XSSFRichTextString rich = styles.dateCellRichText(cellText);
        if (rich != null) {
            cell.setCellValue(rich);
        } else {
            cell.setCellValue(cellText);
        }
    }
}

int r = 2;
for (DispatchAladdinEntrySheetBuilder.EntryRow entry : machineSheet.rows()) {
    // 既存のデータ行ループ本体は変更しない（r++ 開始が 2 になるだけ）
    ...
}

sh.createFreezePane(FIXED_COLUMN_COUNT, 2);
```

`applyPrintSetup` 内:

```java
// 印刷タイトル: 1〜2 行目（見出し＋日加工合計数）・固定列（タイトル列）。
sh.setRepeatingRows(new CellRangeAddress(0, 1, -1, -1));
```

Javadoc コメント「印刷タイトルは 1 行目」があれば「1〜2 行目」に合わせて更新する。

- [ ] **Step 4: Run tests to verify they pass**

```powershell
.\mvnw.cmd -q test "-Dtest=DispatchAladdinEntryWorkbookExporterTest"
```

Expected: 全テスト PASS（Task 1 の合算テスト含む）。

- [ ] **Step 5: Commit**

```powershell
git add code_java/src/main/java/jp/co/pm/ai/desktop/io/DispatchAladdinEntryWorkbookExporter.java `
  code_java/src/test/java/jp/co/pm/ai/desktop/io/DispatchAladdinEntryWorkbookExporterTest.java
git commit -m "feat: アラジン入力Excelに日加工合計数行を追加"
```

---

### Task 3: 仕様ステータス更新と手動確認メモ

**Files:**
- Modify: `docs/specs/2026-07-14-aladdin-daily-processing-total-row-design.md`

- [ ] **Step 1: Update spec status**

先頭の状態行を次に変更する:

```markdown
**状態:** 実装完了
```

- [ ] **Step 2: Manual smoke（任意・アプリ起動中なら）**

結果_配台表タブで「アラジン入力用Excel出力」または「ローカルへ出力」を実行し、いずれかの機械シートで次を目視確認する。

1. 2 行目 A 列が `日加工合計数`
2. 日付列の合算がデータ行の合計と一致
3. 不一致の合計でもオレンジ／黄系背景が付かない
4. 下へスクロールしても見出し＋合計が残る

- [ ] **Step 3: Commit**

```powershell
git add docs/specs/2026-07-14-aladdin-daily-processing-total-row-design.md
git commit -m "docs: 日加工合計数行の仕様を実装完了に更新"
```

---

## Spec coverage checklist（自己レビュー済）

| 仕様要件 | 対応 Task |
|----------|-----------|
| 機械シート内の日次合算 | Task 1 + Task 2 |
| 2 行目ラベル・B〜J 空・2 段表示・空セル | Task 2 |
| 不一致色なし | Task 2 |
| フリーズ 2 行・印刷タイトル 2 行 | Task 2 |
| Builder / 数式 / UI 非変更 | 全 Task（非対象を触らない） |
| 空ブックは現状どおり | 変更なし（`writeEmptySheet` 非改修） |

## Placeholder scan

TBD / 「後で実装」なし。テストコード・実装スニペット・コマンドを各 Step に具体記述済み。
