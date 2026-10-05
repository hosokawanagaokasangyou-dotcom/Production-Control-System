package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;

class IdentityCheckHistoryWorkbookAnnotatorTest {

    private static final LocalDate D1 = LocalDate.of(2026, 7, 7);
    private static final LocalDate D2 = LocalDate.of(2026, 7, 8);

    @Test
    void rewriteAladdinLine_replacesStaleUpperLineAndFillsOnlyDiffs(@TempDir Path tempDir)
            throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Files.createDirectories(Path.of(ui.get(AppPaths.KEY_PM_AI_REPO_ROOT)).resolve("code"));
        Path xlsx = DispatchAladdinEntryWorkbookExporter.write(ui, staleWorkbook()).generationPath();
        Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                AladdinShapedPlanQtyLookup.buildLookup(
                        List.of("機械名", "依頼NO", "工程名", "2026/07/07", "2026/07/08"),
                        List.of(
                                List.of("M1", "T001", "工程A", "10", "4"),
                                List.of("M1", "T002", "工程A", "9", "")));

        int diffs = IdentityCheckHistoryWorkbookAnnotator.rewriteAladdinLine(xlsx, lookup, D1, s -> s);

        assertEquals(1, diffs);
        try (InputStream in = Files.newInputStream(xlsx);
                XSSFWorkbook wb = new XSSFWorkbook(in)) {
            Sheet sh = wb.getSheetAt(0);
            Map<Integer, LocalDate> dateCols =
                    AladdinEntryDispatchPlanWorkbookReader.dateColumns(sh.getRow(0), D1);
            int c1 = colOf(dateCols, D1);
            int c2 = colOf(dateCols, D2);
            Row t001 = rowOf(sh, "T001");
            Row t002 = rowOf(sh, "T002");
            Row total = rowOf(sh, DispatchAladdinEntryWorkbookExporter.DAILY_PROCESSING_TOTAL_LABEL);

            assertEquals("（現アラ計）10\n（シス計）10", text(t001.getCell(c1)));
            assertFalse(isDiffFill(t001.getCell(c1)), "一致セルは塗らない");
            assertEquals("（現アラ計）4\n（シス計）0", text(t001.getCell(c2)));
            assertFalse(isDiffFill(t001.getCell(c2)), "シス計なしは差異件数に数えないので塗らない");
            assertEquals("（現アラ計）9\n（シス計）7", text(t002.getCell(c1)));
            assertTrue(isDiffFill(t002.getCell(c1)));
            assertEquals("（現アラ計）19\n（シス計）17", text(total.getCell(c1)));
            assertEquals("（現アラ計）4\n（シス計）0", text(total.getCell(c2)));
        }
    }

    @Test
    void evaluate_historyExcelUsesCheckedPlanButSourceExcelUnchanged(@TempDir Path tempDir)
            throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Path sourceDir = AppPaths.resolveTaskInputSourceDir(ui);
        Files.createDirectories(sourceDir);
        TestAladdinPlanXlsx.writeGrid(
                sourceDir,
                "aladdin-plan.xlsx",
                new String[][] {
                    {"列1", "列2", "列3", "列4", "列5"},
                    {"上段1", "", "", "", ""},
                    {"上段2", "", "", "", ""},
                    {"上段3", "", "", "", ""},
                    {"機械名", "依頼NO", "工程名", "2026/07/07", "2026/07/08"},
                    {"", "", "", "", ""},
                    {"M1", "T001", "工程A", "10", "4"},
                    {"M1", "T002", "工程A", "9", ""}
                });
        Files.createDirectories(Path.of(ui.get(AppPaths.KEY_PM_AI_REPO_ROOT)).resolve("code"));
        Path xlsx = DispatchAladdinEntryWorkbookExporter.write(ui, staleWorkbook()).generationPath();
        byte[] sourceBefore = Files.readAllBytes(xlsx);

        AladdinEntryDispatchPlanIdentityCheck.Result result =
                AladdinEntryDispatchPlanIdentityCheck.evaluate(ui, xlsx);

        assertEquals(1, result.diffs().size(), result.dialogBody());
        assertArrayEquals(sourceBefore, Files.readAllBytes(xlsx), "比較元 Excel は書き換えない");
        Path hist =
                IdentityCheckHistoryStore.listNewestFirst(ui, "テスト太郎")
                        .getFirst()
                        .dir()
                        .resolve(IdentityCheckHistoryStore.EXCEL_FILE);
        try (InputStream in = Files.newInputStream(hist);
                XSSFWorkbook wb = new XSSFWorkbook(in)) {
            Sheet sh = wb.getSheetAt(0);
            int c1 =
                    colOf(AladdinEntryDispatchPlanWorkbookReader.dateColumns(sh.getRow(0), D1), D1);
            assertEquals("（現アラ計）10\n（シス計）10", text(rowOf(sh, "T001").getCell(c1)));
            assertTrue(isDiffFill(rowOf(sh, "T002").getCell(c1)));
        }
        AladdinEntryDispatchPlanIdentityCheck.Result again =
                AladdinEntryDispatchPlanIdentityCheck.evaluateSnapshot(
                        ui, hist.getParent(), false);
        assertEquals(1, again.diffs().size(), "書き換え後もシス計の再比較結果は変わらない");
    }

    /** 上段は古いアラジン計画（3 / 5 / 1）。シス計は T001=10(7/7)、T002=7(7/7)。 */
    private static DispatchAladdinEntrySheetBuilder.EntryWorkbook staleWorkbook() {
        return new DispatchAladdinEntrySheetBuilder.EntryWorkbook(
                List.of(D1, D2),
                List.of(
                        new DispatchAladdinEntrySheetBuilder.MachineSheet(
                                "M1",
                                List.of(
                                        entryRow(
                                                "T001",
                                                10,
                                                Map.of(
                                                        D1,
                                                        new DispatchAladdinEntrySheetBuilder.EntryCell(3, 10),
                                                        D2,
                                                        new DispatchAladdinEntrySheetBuilder.EntryCell(5, 0))),
                                        entryRow(
                                                "T002",
                                                7,
                                                Map.of(
                                                        D1,
                                                        new DispatchAladdinEntrySheetBuilder.EntryCell(1, 7)))))));
    }

    private static DispatchAladdinEntrySheetBuilder.EntryRow entryRow(
            String tid,
            double qty,
            Map<LocalDate, DispatchAladdinEntrySheetBuilder.EntryCell> cells) {
        return new DispatchAladdinEntrySheetBuilder.EntryRow(
                tid, "", "工程A", "", "", "", qty, 0, qty, cells, D1, 2026);
    }

    private static int colOf(Map<Integer, LocalDate> dateCols, LocalDate d) {
        return dateCols.entrySet().stream()
                .filter(e -> e.getValue().equals(d))
                .findFirst()
                .orElseThrow()
                .getKey();
    }

    private static Row rowOf(Sheet sh, String tid) {
        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row != null
                    && tid.equals(ExcelCellReadSupport.cellToDisplayString(row.getCell(0)).strip())) {
                return row;
            }
        }
        throw new AssertionError("row not found: " + tid);
    }

    private static String text(Cell cell) {
        return ExcelCellReadSupport.cellToDisplayString(cell);
    }

    private static boolean isDiffFill(Cell cell) {
        XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();
        if (style.getFillPattern() != FillPatternType.SOLID_FOREGROUND) {
            return false;
        }
        XSSFColor color = style.getFillForegroundXSSFColor();
        assertNotNull(color);
        return Arrays.equals(IdentityCheckHistoryWorkbookAnnotator.DIFF_FILL_RGB, color.getRGB());
    }

    private static Map<String, String> testUi(Path tempDir) {
        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, tempDir.resolve("repo").toString());
        ui.put(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, tempDir.resolve("task-input").toString());
        ui.put(AppPaths.KEY_PM_AI_OUTPUT_DIR, tempDir.resolve("output").toString());
        ui.put(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, tempDir.resolve("shared").toString());
        ui.put(AppPaths.KEY_PM_AI_OPERATOR_USER, "テスト太郎");
        return ui;
    }
}
