package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;
import jp.co.pm.ai.desktop.reconciliation.PostProcessingPlanMachineLookup;

class DispatchAladdinEntryWorkbookExporterTest {

    @TempDir
    Path tempDir;

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
                DispatchAladdinEntryWorkbookExporter.sumDateColumn(
                        List.of(row1, row2), LocalDate.of(2026, 7, 16));

        assertEquals(110d, sum1.aladdinQty(), 1e-9);
        assertEquals(220d, sum1.systemQty(), 1e-9);
        assertEquals(50d, sum2.aladdinQty(), 1e-9);
        assertEquals(0d, sum2.systemQty(), 1e-9);
        assertTrue(sumEmpty.isEmpty());
        assertEquals("", sumEmpty.cellText());
        assertEquals("（現アラ計）110\n（シス計）220", sum1.cellText());
    }

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
                DispatchAladdinEntryWorkbookExporter.write(ui, model);

        try (XSSFWorkbook wb =
                new XSSFWorkbook(Files.newInputStream(result.latestPath()))) {
            org.apache.poi.ss.usermodel.Sheet sh = wb.getSheetAt(0);
            assertEquals("依頼NO", sh.getRow(0).getCell(0).getStringCellValue());
            assertEquals(
                    DispatchAladdinEntryWorkbookExporter.DAILY_PROCESSING_TOTAL_LABEL,
                    sh.getRow(1).getCell(0).getStringCellValue());
            assertEquals("", sh.getRow(1).getCell(1).getStringCellValue());
            assertEquals("W7-4", sh.getRow(2).getCell(0).getStringCellValue());

            assertEquals(
                    "（現アラ計）300\n（シス計）300",
                    sh.getRow(1).getCell(10).getStringCellValue());
            assertEquals(
                    "（現アラ計）150\n（シス計）300",
                    sh.getRow(1).getCell(11).getStringCellValue());
            assertEquals("", sh.getRow(1).getCell(12).getStringCellValue());

            org.apache.poi.xssf.usermodel.XSSFCellStyle totalStyle =
                    (org.apache.poi.xssf.usermodel.XSSFCellStyle)
                            sh.getRow(1).getCell(11).getCellStyle();
            org.apache.poi.xssf.usermodel.XSSFColor fill = totalStyle.getFillForegroundXSSFColor();
            assertNotNull(fill);
            byte[] rgb = fill.getRGB();
            assertNotNull(rgb);
            assertEquals((byte) 0x00, rgb[0]);
            assertEquals((byte) 0x6B, rgb[1]);
            assertEquals((byte) 0x6B, rgb[2]);
            assertEquals(
                    org.apache.poi.ss.usermodel.IndexedColors.WHITE.getIndex(),
                    totalStyle.getFont().getColor());

            org.apache.poi.xssf.usermodel.XSSFCellStyle totalLabelStyle =
                    (org.apache.poi.xssf.usermodel.XSSFCellStyle)
                            sh.getRow(1).getCell(0).getCellStyle();
            org.apache.poi.xssf.usermodel.XSSFColor labelFill =
                    totalLabelStyle.getFillForegroundXSSFColor();
            assertNotNull(labelFill);
            byte[] labelRgb = labelFill.getRGB();
            assertEquals((byte) 0x00, labelRgb[0]);
            assertEquals((byte) 0x6B, labelRgb[1]);
            assertEquals((byte) 0x6B, labelRgb[2]);
            assertEquals(
                    org.apache.poi.ss.usermodel.IndexedColors.WHITE.getIndex(),
                    totalLabelStyle.getFont().getColor());

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

    @Test
    void sheetNameForMachine_prefixesMachineCodeWhenKnown() {
        PostProcessingPlanMachineLookup.Snapshot snap =
                new PostProcessingPlanMachineLookup.Snapshot(
                        Path.of("plan.csv"),
                        0L,
                        true,
                        true,
                        Map.of("2011", "スライス機1 湖南"),
                        Map.of("スライス機1 湖南", "2011"),
                        List.of("2011 スライス機1 湖南"));
        DispatchAladdinEntrySheetBuilder.MachineSheet ms =
                new DispatchAladdinEntrySheetBuilder.MachineSheet("スライス機1 湖南", List.of());

        assertEquals(
                "2011 スライス機1 湖南",
                DispatchAladdinEntryWorkbookExporter.sheetNameForMachine(ms, snap));
    }

    @Test
    void sheetNameForMachine_fallsBackToMachineNameOnlyWhenCodeUnknown() {
        DispatchAladdinEntrySheetBuilder.MachineSheet ms =
                new DispatchAladdinEntrySheetBuilder.MachineSheet("EC機 湖南", List.of());

        assertEquals(
                "EC機 湖南",
                DispatchAladdinEntryWorkbookExporter.sheetNameForMachine(
                        ms, PostProcessingPlanMachineLookup.Snapshot.empty()));
    }

    @Test
    void write_writesLatestToLocalDiskAndGenerationToSharedDir() throws IOException {
        Path repo = tempDir.resolve("repo");
        Files.createDirectories(repo.resolve("code"));
        Path sharedDataDir = tempDir.resolve("shared-data");
        Files.createDirectories(sharedDataDir);
        Map<String, String> ui =
                Map.of(
                        jp.co.pm.ai.desktop.config.AppPaths.KEY_PM_AI_REPO_ROOT,
                        repo.toString(),
                        jp.co.pm.ai.desktop.config.AppPaths
                                .KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        sharedDataDir
                                .resolve(
                                        jp.co.pm.ai.desktop.config.AppPaths
                                                .SUMMARY_AI_DISPATCH_XLSX)
                                .toString());
        DispatchAladdinEntrySheetBuilder.EntryWorkbook model =
                new DispatchAladdinEntrySheetBuilder.EntryWorkbook(List.of(), List.of());

        DispatchAladdinEntryWorkbookExporter.ExportResult result =
                DispatchAladdinEntryWorkbookExporter.write(ui, model);

        assertEquals(
                jp.co.pm.ai.desktop.config.AppPaths.aladdinEntryDispatchPlanLocalXlsxPath(ui),
                result.latestPath());
        assertTrue(Files.isRegularFile(result.latestPath()));
        assertTrue(
                result.generationPath()
                        .startsWith(
                                jp.co.pm.ai.desktop.config.AppPaths.aladdinEntryDispatchPlanDir(
                                        ui)));
        assertTrue(result.generationPath().startsWith(sharedDataDir));
        assertTrue(Files.isRegularFile(result.generationPath()));
    }

    @Test
    void dateCellRichTextAppliesPerLineFontSizes() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            String fontName = DispatchAladdinEntryWorkbookExporter.DEFAULT_WORKBOOK_FONT_FAMILY;
            Font aladdinLineFont = wb.createFont();
            aladdinLineFont.setFontName(fontName);
            aladdinLineFont.setFontHeightInPoints((short) 9);
            Font systemLineFont = wb.createFont();
            systemLineFont.setFontName(fontName);
            systemLineFont.setFontHeightInPoints((short) 12);

            String text =
                    DispatchAladdinEntrySheetBuilder.ALADDIN_LINE_PREFIX
                            + "4400\n"
                            + DispatchAladdinEntrySheetBuilder.SYSTEM_LINE_PREFIX
                            + "0";
            XSSFRichTextString rich =
                    DispatchAladdinEntryWorkbookExporter.buildDateCellRichText(
                            text, aladdinLineFont, systemLineFont);

            assertNotNull(rich);
            assertEquals(text, rich.getString());
            assertTrue(rich.hasFormatting());
            // XSSFRichTextString は改行境界でもランを分割するため 2 以上
            assertTrue(rich.numFormattingRuns() >= 2);
        }
    }

    @Test
    void pruneKeepsAtMostMaxGenerationsDeletingOldest() throws IOException {
        int total = DispatchAladdinEntryWorkbookExporter.MAX_GENERATIONS_PER_USER + 5;
        Instant base = Instant.parse("2026-07-01T00:00:00Z");
        for (int i = 0; i < total; i++) {
            Path f = tempDir.resolve(String.format("アラジン入力用_配台計画_%03d.xlsx", i));
            Files.writeString(f, "x");
            Files.setLastModifiedTime(f, FileTime.from(base.plusSeconds(i * 60L)));
        }

        DispatchAladdinEntryWorkbookExporter.pruneGenerations(tempDir);

        try (var stream = Files.list(tempDir)) {
            assertEquals(
                    DispatchAladdinEntryWorkbookExporter.MAX_GENERATIONS_PER_USER,
                    stream.filter(Files::isRegularFile).count());
        }
        // 最古の5件が削除され、新しい側が残る
        assertFalse(Files.exists(tempDir.resolve("アラジン入力用_配台計画_000.xlsx")));
        assertFalse(Files.exists(tempDir.resolve("アラジン入力用_配台計画_004.xlsx")));
        assertTrue(Files.exists(tempDir.resolve("アラジン入力用_配台計画_005.xlsx")));
        assertTrue(
                Files.exists(
                        tempDir.resolve(
                                String.format("アラジン入力用_配台計画_%03d.xlsx", total - 1))));
    }

    @Test
    void resolveMachineCodeFromName_normalizesWhitespace() throws Exception {
        Path csv = tempDir.resolve("plan-ws.csv");
        Files.writeString(
                csv,
                "機械,機械名,依頼NO\n" + "2011,スライス機1 湖南,R1\n");

        PostProcessingPlanMachineLookup.invalidate();
        PostProcessingPlanMachineLookup.Snapshot snap =
                PostProcessingPlanMachineLookup.snapshotFromFile(csv);

        assertEquals(
                "2011",
                PostProcessingPlanMachineLookup.resolveMachineCodeFromName(
                        snap, "スライス機1　湖南"));
    }

    @Test
    void pruneIgnoresNonXlsxFiles() throws IOException {
        for (int i = 0; i < DispatchAladdinEntryWorkbookExporter.MAX_GENERATIONS_PER_USER; i++) {
            Files.writeString(tempDir.resolve("gen_" + i + ".xlsx"), "x");
        }
        Path other = tempDir.resolve("readme.txt");
        Files.writeString(other, "keep");

        DispatchAladdinEntryWorkbookExporter.pruneGenerations(tempDir);

        assertTrue(Files.exists(other));
        try (var stream = Files.list(tempDir)) {
            assertEquals(
                    DispatchAladdinEntryWorkbookExporter.MAX_GENERATIONS_PER_USER + 1,
                    stream.filter(Files::isRegularFile).count());
        }
    }
}
