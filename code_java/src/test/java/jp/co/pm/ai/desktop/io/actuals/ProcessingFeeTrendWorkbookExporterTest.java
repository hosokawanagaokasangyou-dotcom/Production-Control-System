package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.io.DispatchAladdinEntryWorkbookExporter;
import jp.co.pm.ai.desktop.io.actuals.JuchuProcessingFeeRateLoader.FeeInfo;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.QuantityLine;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.Result;

class ProcessingFeeTrendWorkbookExporterTest {

    @TempDir Path tempDir;

    @Test
    void writeCreatesStyledDayRequestAndMismatchSheets() throws Exception {
        LocalDate d = LocalDate.of(2026, 9, 1);
        FeeInfo fee = new FeeInfo(null, 1_000.0, "最終", null, null, 100.0);
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                new QuantityLine(
                                        d, "R", 60, "最終", java.time.LocalDateTime.of(d, java.time.LocalTime.of(15, 0)))),
                        List.of(),
                        Map.of("R", fee),
                        d,
                        d.plusDays(2),
                        d);
        Path out = tempDir.resolve("fee.xlsx");
        ProcessingFeeTrendWorkbookExporter.write(r, out);
        assertTrue(Files.isRegularFile(out));
        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(out))) {
            assertEquals(3, wb.getNumberOfSheets());
            XSSFSheet days = wb.getSheet("日別");
            Sheet reqs = wb.getSheet("依頼NO別");
            Sheet mismatch = wb.getSheet("受注額と実績の相違");
            assertNotNull(days);
            assertNotNull(reqs);
            assertNotNull(mismatch);

            assertEquals("日付", days.getRow(0).getCell(0).getStringCellValue());
            assertEquals("曜日", days.getRow(0).getCell(1).getStringCellValue());
            assertEquals("実績", days.getRow(0).getCell(2).getStringCellValue());
            assertEquals("未了", days.getRow(0).getCell(3).getStringCellValue());
            assertEquals("実績累計", days.getRow(0).getCell(4).getStringCellValue());
            assertEquals("未了累計", days.getRow(0).getCell(5).getStringCellValue());
            assertEquals("見込累計", days.getRow(0).getCell(6).getStringCellValue());
            assertFalse(days.getRow(0).getCell(2).getStringCellValue().contains("円"));

            assertEquals("2026/09/01", days.getRow(1).getCell(0).getStringCellValue());
            assertEquals("火", days.getRow(1).getCell(1).getStringCellValue());
            assertTrue(days.getColumnWidth(0) >= 15 * 256, "日付列が切れない幅であること");
            assertTrue(days.getColumnWidth(2) >= 14 * 256, "円列がカンマ表示できる幅であること");
            assertEquals(1, days.getPaneInformation().getHorizontalSplitPosition());
            assertTrue(days.getCTWorksheet().isSetAutoFilter());

            String yenFmt =
                    wb.createDataFormat().getFormat(days.getRow(1).getCell(2).getCellStyle().getDataFormat());
            assertTrue(yenFmt.contains("#,##0"), "円はカンマ区切り: " + yenFmt);
            assertFalse(yenFmt.contains(".0"), "円は整数書式: " + yenFmt);

            String fontName = wb.getFontAt(days.getRow(0).getCell(0).getCellStyle().getFontIndex()).getFontName();
            assertEquals(DispatchAladdinEntryWorkbookExporter.DEFAULT_WORKBOOK_FONT_FAMILY, fontName);

            int last = days.getLastRowNum();
            assertEquals("合計", days.getRow(last).getCell(0).getStringCellValue());
            assertNotNull(days.getDrawingPatriarch());
            assertFalse(days.getDrawingPatriarch().getCharts().isEmpty());

            assertEquals("依頼NO", reqs.getRow(0).getCell(0).getStringCellValue());
            assertEquals("受注額", reqs.getRow(0).getCell(1).getStringCellValue());
            assertEquals("実績", reqs.getRow(0).getCell(5).getStringCellValue());
            assertEquals("未了", reqs.getRow(0).getCell(6).getStringCellValue());
            assertEquals("合計", reqs.getRow(1).getCell(0).getStringCellValue());

            assertEquals("依頼NO", mismatch.getRow(0).getCell(0).getStringCellValue());
            assertEquals("受注額", mismatch.getRow(0).getCell(1).getStringCellValue());
            assertEquals("差額", mismatch.getRow(0).getCell(3).getStringCellValue());
        }
    }
}
