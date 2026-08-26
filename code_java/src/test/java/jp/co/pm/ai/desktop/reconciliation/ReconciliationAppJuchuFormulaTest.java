package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.reflect.Method;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
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
                "SUM(IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AH390,CHAR(10))),0)"
                        + "*IFERROR(VALUE(TEXTSPLIT(受注ﾌｧｲﾙ!$AM390,CHAR(10))),0))",
                formula);
        assertTrue(formula.contains("受注ﾌｧｲﾙ!$AH390"));
        assertTrue(formula.contains("受注ﾌｧｲﾙ!$AM390"));
        assertTrue(formula.contains("TEXTSPLIT"));
        assertTrue(formula.contains("CHAR(10)"));
        assertTrue(formula.startsWith("SUM("));
        assertTrue(!formula.contains("@TEXTSPLIT"));
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
