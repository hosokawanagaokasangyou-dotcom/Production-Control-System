package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.io.actuals.JuchuProcessingFeeRateLoader.FeeInfo;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.QuantityLine;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.Result;

class ProcessingFeeTrendWorkbookExporterTest {

    @TempDir Path tempDir;

    @Test
    void writeCreatesDayAndRequestSheets() throws Exception {
        LocalDate d = LocalDate.of(2026, 9, 1);
        FeeInfo fee = new FeeInfo(null, 1_000.0, "最終", null, null, 100.0);
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(new QuantityLine(d, "R", 60, "最終")),
                        List.of(),
                        Map.of("R", fee),
                        d,
                        d.plusDays(2),
                        d);
        Path out = tempDir.resolve("fee.xlsx");
        ProcessingFeeTrendWorkbookExporter.write(r, out);
        assertTrue(Files.isRegularFile(out));
        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(out))) {
            assertEquals(2, wb.getNumberOfSheets());
            Sheet days = wb.getSheet("日別");
            Sheet reqs = wb.getSheet("依頼NO別");
            assertEquals("日付", days.getRow(0).getCell(0).getStringCellValue());
            assertEquals("依頼NO", reqs.getRow(0).getCell(0).getStringCellValue());
            assertTrue(days.getLastRowNum() >= 1);
            assertTrue(reqs.getLastRowNum() >= 1);
        }
    }
}
