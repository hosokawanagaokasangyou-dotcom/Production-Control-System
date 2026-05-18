package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.SummaryAiDispatchExportPrefs;
import jp.co.pm.ai.desktop.ui.DeliveryCalendarMainCell;

class SummaryAiDispatchWorkbookExporterTest {

    @TempDir
    Path prefsDir;

    @BeforeEach
    void isolatePrefsStore() {
        System.setProperty(
                "pm.ai.test.summaryExportPrefs",
                prefsDir.resolve("summary-ai-dispatch-export.json").toString());
    }

    @AfterEach
    void clearPrefsStoreProperty() {
        System.clearProperty("pm.ai.test.summaryExportPrefs");
    }

    @Test
    void mainCompareFromUi_insertsSystemDispatchQtyTotalBeforeDateColumns() {
        PlanInputTabularIo.TabularSheet sheet =
                SummaryAiDispatchWorkbookExporter.mainCompareFromUi(
                        List.of("依頼NO", "2026年5月17日(土)", "2026年5月18日(日)"),
                        List.of(
                                List.of(
                                        new DeliveryCalendarMainCell.PlainText("Y5-1"),
                                        new DeliveryCalendarMainCell.TripleQty("", "", "40", ""),
                                        new DeliveryCalendarMainCell.TripleQty("", "", "60", ""))));
        assertEquals(
                List.of(
                        "依頼NO",
                        SummaryAiDispatchWorkbookExporter.COL_SYSTEM_DISPATCH_QTY_TOTAL,
                        "2026年5月17日(土)",
                        "2026年5月18日(日)"),
                sheet.headers());
        assertEquals("100", sheet.rows().get(0).get(1));
    }

    @Test
    void writeOverwrite_appliesFrozenColumnCount(@TempDir Path tmp) throws Exception {
        Path fakeRepo = tmp.resolve("repo");
        Files.createDirectories(fakeRepo.resolve("code"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());

        SummaryAiDispatchExportPrefs.save(
                new SummaryAiDispatchExportPrefs.ExportPrefs(
                        SummaryAiDispatchExportPrefs.DEFAULT_FONT_FAMILY,
                        SummaryAiDispatchExportPrefs.DEFAULT_FONT_SIZE_PT,
                        SummaryAiDispatchExportPrefs.DEFAULT_THEME,
                        Map.of(
                                SummaryAiDispatchExportPrefs.SheetKey.DISPATCH,
                                new SummaryAiDispatchExportPrefs.SheetPrefs(5, List.of()))));

        PlanInputTabularIo.TabularSheet dispatch =
                new PlanInputTabularIo.TabularSheet(
                        List.of("C1", "C2", "C3", "C4", "C5", "C6"),
                        List.of(List.of("1", "2", "3", "4", "5", "6")));
        Path out =
                SummaryAiDispatchWorkbookExporter.writeOverwrite(
                        ui, empty(), empty(), empty(), dispatch);
        try (var wb = WorkbookFactory.create(out.toFile())) {
            Sheet sh = wb.getSheet(SummaryAiDispatchWorkbookExporter.SHEET_DISPATCH);
            assertEquals(5, sh.getPaneInformation().getVerticalSplitLeftColumn());
            assertEquals(1, sh.getPaneInformation().getHorizontalSplitTopRow());
        }
    }

    @Test
    void buildMainCompareRichText_highlightsStage3QtyOnly() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Font base = wb.createFont();
            Font qty = wb.createFont();
            String text = "(アラ計画)900\n(段階3後)3300";
            XSSFRichTextString rich =
                    (XSSFRichTextString)
                            SummaryAiDispatchWorkbookExporter.buildMainCompareRichText(
                                    text, base, qty);
            assertTrue(rich != null);
            assertEquals(text, rich.getString());
            assertTrue(rich.getCTRst().toString().contains("highlight"));
            assertTrue(rich.getCTRst().toString().contains("green"));
        }
    }

    @Test
    void writeOverwrite_mainCompareStage3QtyUsesRichText(@TempDir Path tmp) throws Exception {
        Path fakeRepo = tmp.resolve("repo");
        Files.createDirectories(fakeRepo.resolve("code"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());

        PlanInputTabularIo.TabularSheet main =
                SummaryAiDispatchWorkbookExporter.mainCompareFromUi(
                        List.of("機械名", "2026年5月17日(土)"),
                        List.of(
                                List.of(
                                        new DeliveryCalendarMainCell.PlainText("M1"),
                                        new DeliveryCalendarMainCell.TripleQty(
                                                "10", "", "", "3300"))));
        Path out =
                SummaryAiDispatchWorkbookExporter.writeOverwrite(
                        ui, main, empty(), empty(), empty());

        try (var wb = WorkbookFactory.create(out.toFile())) {
            Sheet sh = wb.getSheet(SummaryAiDispatchWorkbookExporter.SHEET_MAIN_COMPARE);
            XSSFRichTextString rich =
                    (XSSFRichTextString) sh.getRow(1).getCell(2).getRichStringCellValue();
            assertTrue(rich.getString().contains("(段階3後)3300"));
            assertTrue(rich.getCTRst().toString().contains("highlight"));
        }
    }

    @Test
    void writeOverwrite_createsFourSheets(@TempDir Path tmp) throws Exception {
        Path fakeRepo = tmp.resolve("repo");
        Files.createDirectories(fakeRepo.resolve("code"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());

        PlanInputTabularIo.TabularSheet main =
                SummaryAiDispatchWorkbookExporter.mainCompareFromUi(
                        List.of("機械名", "2026年5月17日(土)", "2026年5月18日(日)"),
                        List.of(
                                List.of(
                                        new DeliveryCalendarMainCell.PlainText("M1"),
                                        new DeliveryCalendarMainCell.TripleQty("10", "", "100", ""),
                                        new DeliveryCalendarMainCell.TripleQty("10", "", "50", ""))));
        PlanInputTabularIo.TabularSheet dispatch =
                new PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO"), List.of(List.of("A001")));
        Path out = SummaryAiDispatchWorkbookExporter.writeOverwrite(ui, main, empty(), empty(), dispatch);

        assertTrue(Files.isRegularFile(out));
        assertEquals("サマリ_AI配台.xlsx", out.getFileName().toString());
        try (var wb = WorkbookFactory.create(out.toFile())) {
            assertEquals(4, wb.getNumberOfSheets());
            Sheet sh = wb.getSheet(SummaryAiDispatchWorkbookExporter.SHEET_MAIN_COMPARE);
            assertEquals("機械名", sh.getRow(0).getCell(0).getStringCellValue());
            assertEquals(
                    SummaryAiDispatchWorkbookExporter.COL_SYSTEM_DISPATCH_QTY_TOTAL,
                    sh.getRow(0).getCell(1).getStringCellValue());
            assertEquals("M1", sh.getRow(1).getCell(0).getStringCellValue());
            assertEquals("150", sh.getRow(1).getCell(1).getStringCellValue());
            assertTrue(sh.getRow(1).getCell(2).getStringCellValue().contains("(アラ計画)"));
            assertEquals(
                    "A001",
                    wb.getSheet(SummaryAiDispatchWorkbookExporter.SHEET_DISPATCH)
                            .getRow(1)
                            .getCell(0)
                            .getStringCellValue());
        }
    }

    private static PlanInputTabularIo.TabularSheet empty() {
        return new PlanInputTabularIo.TabularSheet(List.of(), List.of());
    }
}
