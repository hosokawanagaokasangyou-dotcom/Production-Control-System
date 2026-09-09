package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.reconciliation.JuchuSheetColumnLayout;

class JuchuProcessingFeeRateLoaderTest {

    @TempDir Path tempDir;

    @Test
    void loadsIraiToAhRate_lastWins_skipsEmptyAh() throws Exception {
        Path xlsx = tempDir.resolve("juchu.xlsx");
        int irai = JuchuSheetColumnLayout.Col.IRAI_NO.columnIndex();
        int ah = JuchuSheetColumnLayout.Col.KAKOCHIN.columnIndex();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sh = wb.createSheet(JuchuProcessingFeeRateLoader.SHEET_NAME);
            Row h = sh.createRow(2);
            h.createCell(irai).setCellValue("依頼Ｎｏ");
            h.createCell(ah).setCellValue("加工賃");
            Row r0 = sh.createRow(3);
            r0.createCell(irai).setCellValue("R-1");
            r0.createCell(ah).setCellValue("100");
            Row r1 = sh.createRow(4);
            r1.createCell(irai).setCellValue("R-2");
            r1.createCell(ah).setCellValue("");
            Row r2 = sh.createRow(5);
            r2.createCell(irai).setCellValue("R-1");
            r2.createCell(ah).setCellValue("120");
            Row r3 = sh.createRow(6);
            r3.createCell(irai).setCellValue("R-3");
            r3.createCell(ah).setCellValue("1,250.5");
            try (var out = Files.newOutputStream(xlsx)) {
                wb.write(out);
            }
        }

        Map<String, Double> rates = JuchuProcessingFeeRateLoader.loadRates(xlsx);
        assertEquals(120.0, rates.get("R-1"));
        assertFalse(rates.containsKey("R-2"));
        assertEquals(1250.5, rates.get("R-3"));
        assertEquals(2, rates.size());
    }

    @Test
    void parseFeeRate_takesFirstLineOnly() {
        assertEquals(80.0, JuchuProcessingFeeRateLoader.parseFeeRate("80\n90"));
        assertTrue(JuchuProcessingFeeRateLoader.parseFeeRate("  ") == null);
        assertTrue(JuchuProcessingFeeRateLoader.parseFeeRate(null) == null);
    }
}
