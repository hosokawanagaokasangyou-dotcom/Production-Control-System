package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Method;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellFormulaType;

class ReconciliationAppJuchuFormulaTest {

    @Test
    void buildJuchuAoTextsSplitSumFormula_embedsRowNumber() {
        String formula = ReconciliationApp.buildJuchuAoTextsSplitSumFormula(390);

        assertEquals(
                "SUM(IFERROR(VALUE(_xlfn.TEXTSPLIT(受注ﾌｧｲﾙ!$AH390,CHAR(10))),0)"
                        + "*IFERROR(VALUE(_xlfn.TEXTSPLIT(受注ﾌｧｲﾙ!$AM390,CHAR(10))),0))",
                formula);
        assertTrue(formula.contains("受注ﾌｧｲﾙ!$AH390"));
        assertTrue(formula.contains("受注ﾌｧｲﾙ!$AM390"));
        assertTrue(formula.contains("CHAR(10)"));
        assertTrue(formula.startsWith("SUM("));
        assertTrue(!formula.contains("@TEXTSPLIT"));
    }

    /** {@code _xlfn.} 無しの TEXTSPLIT は Excel が未知名として扱い #NAME? → IFERROR で 0 になる。 */
    @Test
    void buildJuchuAoTextsSplitSumFormula_alwaysPrefixesTextsplitWithXlfn() {
        String formula = ReconciliationApp.buildJuchuAoTextsSplitSumFormula(445);

        assertEquals(2, countOccurrences(formula, "_xlfn.TEXTSPLIT("));
        assertEquals(2, countOccurrences(formula, "TEXTSPLIT("));
    }

    @Test
    void isJuchuAoFormulaMissingXlfnPrefix_detectsLegacyAndExcelRewrittenNames() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.setCellFormulaValidation(false);
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow row = sheet.createRow(444);
            XSSFCell legacy = row.createCell(0);
            legacy.setCellFormula(
                    "SUM(IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AH445,CHAR(10))),0)"
                            + "*IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AM445,CHAR(10))),0))");
            XSSFCell udf = row.createCell(1);
            udf.setCellFormula("SUM(IFERROR(VALUE(_xludf.TEXTSPLIT($AH445,CHAR(10))),0))");
            XSSFCell fixed = row.createCell(2);
            fixed.setCellFormula(ReconciliationApp.buildJuchuAoTextsSplitSumFormula(445));
            XSSFCell other = row.createCell(3);
            other.setCellFormula("AI445*AH445");
            XSSFCell number = row.createCell(4);
            number.setCellValue(38750);

            assertTrue(ReconciliationApp.isJuchuAoFormulaMissingXlfnPrefix(legacy));
            assertTrue(ReconciliationApp.isJuchuAoFormulaMissingXlfnPrefix(udf));
            assertTrue(!ReconciliationApp.isJuchuAoFormulaMissingXlfnPrefix(fixed));
            assertTrue(!ReconciliationApp.isJuchuAoFormulaMissingXlfnPrefix(other));
            assertTrue(!ReconciliationApp.isJuchuAoFormulaMissingXlfnPrefix(number));
            assertTrue(!ReconciliationApp.isJuchuAoFormulaMissingXlfnPrefix(null));
        }
    }

    @Test
    void repairJuchuAoFormulasMissingXlfnPrefix_rewritesOnlyLegacyRowsInRange() throws Exception {
        int ao = JuchuSheetColumnLayout.columnLetterToIndex("AO");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.setCellFormulaValidation(false);
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            // row 442: 旧版（接頭辞無し・通常数式）
            sheet.createRow(441)
                    .createCell(ao)
                    .setCellFormula(
                            "SUM(IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AH442,CHAR(10))),0)"
                                    + "*IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AM442,CHAR(10))),0))");
            // row 443: 既に正しい配列数式
            XSSFRow ok = sheet.createRow(442);
            ok.createCell(ao);
            sheet.setArrayFormula(
                    ReconciliationApp.buildJuchuAoTextsSplitSumFormula(443),
                    new CellRangeAddress(442, 442, ao, ao));
            // row 444: 手入力の数値（触らない）
            sheet.createRow(443).createCell(ao).setCellValue(900);
            // row 445: 旧版（接頭辞無し・配列数式）
            XSSFRow legacyArray = sheet.createRow(444);
            legacyArray.createCell(ao);
            sheet.setArrayFormula(
                    "SUM(IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AH445,CHAR(10))),0)"
                            + "*IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AM445,CHAR(10))),0))",
                    new CellRangeAddress(444, 444, ao, ao));
            // row 446: 範囲外（触らない）
            sheet.createRow(445)
                    .createCell(ao)
                    .setCellFormula("SUM(IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AH446,CHAR(10))),0))");

            int repaired =
                    ReconciliationApp.repairJuchuAoFormulasMissingXlfnPrefix(sheet, 3, 444);

            assertEquals(2, repaired);
            assertAoArrayFormula(sheet.getRow(441).getCell(ao), 442);
            assertAoArrayFormula(sheet.getRow(442).getCell(ao), 443);
            assertEquals(CellType.NUMERIC, sheet.getRow(443).getCell(ao).getCellType());
            assertEquals(900, sheet.getRow(443).getCell(ao).getNumericCellValue(), 0);
            assertAoArrayFormula(sheet.getRow(444).getCell(ao), 445);
            assertEquals(
                    "SUM(IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AH446,CHAR(10))),0))",
                    sheet.getRow(445).getCell(ao).getCellFormula());
        }
    }

    /** 保存後の XML に {@code _xlfn.TEXTSPLIT} が配列数式として残ることを確認する。 */
    @Test
    void writtenWorkbook_keepsXlfnPrefixedArrayFormulaAfterRoundTrip() throws Exception {
        int ao = JuchuSheetColumnLayout.columnLetterToIndex("AO");
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow row = sheet.createRow(444);
            invokeApplyDefault(row, Map.of(), 445);
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                wb.write(out);
                bytes = out.toByteArray();
            }
        }
        try (XSSFWorkbook reloaded = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFCell cell = reloaded.getSheet("受注ﾌｧｲﾙ").getRow(444).getCell(ao);
            assertAoArrayFormula(cell, 445);
            assertTrue(cell.getCTCell().getF().getStringValue().contains("_xlfn.TEXTSPLIT("));
        }
    }

    private static int countOccurrences(String text, String needle) {
        int count = 0;
        for (int i = text.indexOf(needle); i >= 0; i = text.indexOf(needle, i + needle.length())) {
            count++;
        }
        return count;
    }

    @Test
    void applyDefaultJuchuFormula_writesAoColumnAsArrayFormulaWhenEmpty() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(2);
            header.createCell(JuchuSheetColumnLayout.columnLetterToIndex("AO")).setCellValue("単価");
            XSSFRow row = sheet.createRow(389);

            invokeApplyDefault(row, Map.of("単価", JuchuSheetColumnLayout.columnLetterToIndex("AO")), 390);

            assertAoArrayFormula(row.getCell(JuchuSheetColumnLayout.columnLetterToIndex("AO")), 390);
        }
    }

    @Test
    void applyDefaultJuchuFormula_overwritesCopiedRegularFormulaWithArrayFormula() throws Exception {
        int ao = JuchuSheetColumnLayout.columnLetterToIndex("AO");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow src = sheet.createRow(427);
            src.createCell(ao)
                    .setCellFormula(ReconciliationApp.buildJuchuAoTextsSplitSumFormula(428));
            sheet.createRow(428);
            sheet.copyRows(
                    427,
                    427,
                    428,
                    new CellCopyPolicy.Builder()
                            .cellFormula(true)
                            .cellStyle(true)
                            .cellValue(false)
                            .mergedRegions(true)
                            .rowHeight(true)
                            .build());

            invokeApplyDefault(sheet.getRow(428), Map.of(), 429);

            assertAoArrayFormula(sheet.getRow(428).getCell(ao), 429);
        }
    }

    @Test
    void applyDefaultJuchuFormula_overwritesImplicitIntersectionFormula() throws Exception {
        int ao = JuchuSheetColumnLayout.columnLetterToIndex("AO");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow row = sheet.createRow(428);
            wb.setCellFormulaValidation(false);
            row.createCell(ao)
                    .setCellFormula(
                            "SUM(IFERROR(VALUE(@TEXTSPLIT(受注ﾌｧｲﾙ!$AH429,CHAR(10))),0)"
                                    + "*IFERROR(VALUE(@TEXTSPLIT(受注ﾌｧｲﾙ!$AM429,CHAR(10))),0))");

            invokeApplyDefault(row, Map.of(), 429);

            assertAoArrayFormula(row.getCell(ao), 429);
        }
    }

    private static void invokeApplyDefault(Row row, Map<String, Integer> colMap, int excelRow)
            throws Exception {
        Method method =
                ReconciliationApp.class.getDeclaredMethod(
                        "applyDefaultJuchuFormulasIfMissing",
                        Row.class,
                        Map.class,
                        int.class);
        method.setAccessible(true);
        method.invoke(null, row, colMap, excelRow);
    }

    private static void assertAoArrayFormula(XSSFCell cell, int excelRow) {
        assertEquals(CellType.FORMULA, cell.getCellType());
        String formula = cell.getCellFormula();
        assertEquals(ReconciliationApp.buildJuchuAoTextsSplitSumFormula(excelRow), formula);
        assertTrue(!formula.contains("@TEXTSPLIT"), formula);
        assertEquals(STCellFormulaType.ARRAY, cell.getCTCell().getF().getT());
    }
}
