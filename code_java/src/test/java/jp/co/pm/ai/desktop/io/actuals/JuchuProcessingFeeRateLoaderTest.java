package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.io.actuals.JuchuProcessingFeeRateLoader.FeeInfo;
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
    void loadFeeInfo_readsAoProcessContent_andAhLastLine() throws Exception {
        Path xlsx = tempDir.resolve("juchu-ao.xlsx");
        int irai = JuchuSheetColumnLayout.Col.IRAI_NO.columnIndex();
        int ah = JuchuSheetColumnLayout.Col.KAKOCHIN.columnIndex();
        int kako = JuchuSheetColumnLayout.Col.KAKO_NAIYO.columnIndex();
        int ao = JuchuProcessingFeeRateLoader.AO_COLUMN_INDEX;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sh = wb.createSheet(JuchuProcessingFeeRateLoader.SHEET_NAME);
            Row h = sh.createRow(2);
            h.createCell(irai).setCellValue("依頼Ｎｏ");
            h.createCell(ah).setCellValue("加工賃");
            h.createCell(ao).setCellValue("加工賃合計");
            h.createCell(kako).setCellValue("加工内容");
            Row r0 = sh.createRow(3);
            r0.createCell(irai).setCellValue("E9-2");
            r0.createCell(ah).setCellValue("80\n90");
            r0.createCell(ao).setCellValue(10_000);
            r0.createCell(kako).setCellValue("スリット,E9-2");
            try (var out = Files.newOutputStream(xlsx)) {
                wb.write(out);
            }
        }

        Map<String, FeeInfo> fees = JuchuProcessingFeeRateLoader.loadFeeInfo(xlsx);
        FeeInfo info = fees.get("E9-2");
        assertEquals(90.0, info.rateAhYenPerM());
        assertEquals(10_000.0, info.totalAoYen());
        assertEquals("スリット,E9-2", info.processContent());
        assertTrue(info.hasAo());
        assertTrue(info.hasAh());
    }

    @Test
    void parseFeeRate_takesFirstLineOnly() {
        assertEquals(80.0, JuchuProcessingFeeRateLoader.parseFeeRate("80\n90"));
        assertTrue(JuchuProcessingFeeRateLoader.parseFeeRate("  ") == null);
        assertTrue(JuchuProcessingFeeRateLoader.parseFeeRate(null) == null);
    }

    @Test
    void parseFeeRateLastLine_takesTrailingNumericLine() {
        assertEquals(90.0, JuchuProcessingFeeRateLoader.parseFeeRateLastLine("80\n90"));
        assertEquals(90.0, JuchuProcessingFeeRateLoader.parseFeeRateLastLine("80\n90\n"));
        assertNull(JuchuProcessingFeeRateLoader.parseFeeRateLastLine("  "));
    }

    @Test
    void computeAoFromAhAmProductSum_pairsLinesLikeExcel() {
        assertEquals(
                80.0 * 10 + 90.0 * 20,
                JuchuProcessingFeeRateLoader.computeAoFromAhAmProductSum("80\n90", "10\n20"),
                1e-9);
        assertNull(JuchuProcessingFeeRateLoader.computeAoFromAhAmProductSum("", ""));
        assertNull(JuchuProcessingFeeRateLoader.computeAoFromAhAmProductSum("0", "0"));
    }

    @Test
    void loadFeeInfo_computesAoFromAhAmWhenAoFormulaUnusable() throws Exception {
        Path xlsx = tempDir.resolve("juchu-ao-fallback.xlsx");
        int irai = JuchuSheetColumnLayout.Col.IRAI_NO.columnIndex();
        int ah = JuchuSheetColumnLayout.Col.KAKOCHIN.columnIndex();
        int am = JuchuProcessingFeeRateLoader.AM_COLUMN_INDEX;
        int ao = JuchuProcessingFeeRateLoader.AO_COLUMN_INDEX;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sh = wb.createSheet(JuchuProcessingFeeRateLoader.SHEET_NAME);
            Row h = sh.createRow(2);
            h.createCell(irai).setCellValue("依頼Ｎｏ");
            Row r0 = sh.createRow(3);
            r0.createCell(irai).setCellValue("W9-3");
            r0.createCell(ah).setCellValue("38");
            r0.createCell(am).setCellValue("4500");
            // TEXTSPLIT 数式は POI が評価できずキャッシュ 0 → AH×AM 再計算
            r0.createCell(ao).setCellFormula("SUM(IFERROR(VALUE(_xlfn.TEXTSPLIT(AH4,CHAR(10))),0))");
            try (var out = Files.newOutputStream(xlsx)) {
                wb.write(out);
            }
        }

        Map<String, FeeInfo> fees = JuchuProcessingFeeRateLoader.loadFeeInfo(xlsx);
        FeeInfo info = fees.get("W9-3");
        assertEquals(38.0, info.rateAhYenPerM());
        assertEquals(38.0 * 4500.0, info.totalAoYen(), 1e-6);
        assertTrue(info.hasAo());
    }
}
