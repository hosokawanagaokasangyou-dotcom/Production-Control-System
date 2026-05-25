package jp.co.pm.ai.desktop.reconciliation;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class JuchuSheetColumnLayoutTest {

    @Test
    void columnLetterToIndex_apAndAq() {
        assertEquals(41, JuchuSheetColumnLayout.columnLetterToIndex("AP"));
        assertEquals(42, JuchuSheetColumnLayout.columnLetterToIndex("AQ"));
        assertEquals(4, JuchuSheetColumnLayout.columnLetterToIndex("E"));
        assertEquals("AP", JuchuSheetColumnLayout.indexToColumnLetter(41));
    }

    @Test
    void matchesHeader_acceptsAliases() {
        JuchuSheetColumnLayout.Col irai = JuchuSheetColumnLayout.Col.IRAI_NO;
        assertTrue(irai.matchesHeader("依頼No"));
        assertTrue(irai.matchesHeader("依頼Ｎｏ"));

        JuchuSheetColumnLayout.Col ec = JuchuSheetColumnLayout.Col.EC_MEN;
        assertTrue(ec.matchesHeader("EC面"));
        assertTrue(ec.matchesHeader("ＥＣ面"));
    }

    @Test
    void validateHeaders_reportsMismatch() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(JuchuSheetColumnLayout.Col.NYURYOKU_BI.columnIndex()).setCellValue("投入日");

            List<String> warnings = JuchuSheetColumnLayout.validateHeaders(header);
            assertFalse(warnings.isEmpty());
            assertTrue(warnings.stream().anyMatch(w -> w.contains("E列") && w.contains("入力日")));
        }
    }

    @Test
    void buildAndParseSpecName() {
        String spec = JuchuSheetColumnLayout.buildSpecName("20010", "H600", "1180", "250");
        assertEquals("20010-H600-1180X250", spec);

        String[] parts = JuchuSheetColumnLayout.parseSpecName("20010-H600-1180X250");
        assertEquals("20010", parts[0]);
        assertEquals("H600", parts[1]);
        assertEquals("1180", parts[2]);
        assertEquals("250", parts[3]);
    }

    @Test
    void readDbValuesFromRow_usesLayoutColumns() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow row = sheet.createRow(3);
            row.createCell(JuchuSheetColumnLayout.Col.HINMEI.columnIndex()).setCellValue("6713");
            row.createCell(JuchuSheetColumnLayout.Col.SEIHIN.columnIndex())
                    .setCellValue("20010-H600-1180X250");
            row.createCell(JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT.columnIndex())
                    .setCellValue("A2K10H6B8250FW3");

            var vals = JuchuSheetColumnLayout.readDbValuesFromRow(row);
            assertEquals("6713", vals.get("品名"));
            assertEquals("20010-H600-1180X250", vals.get("製品"));
            assertEquals("A2K10H6B8250FW3", vals.get("masterBase商品(製品)"));
        }
    }
}
