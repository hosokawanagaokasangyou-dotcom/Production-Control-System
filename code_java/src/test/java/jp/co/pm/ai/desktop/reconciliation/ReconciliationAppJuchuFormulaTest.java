package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class ReconciliationAppJuchuFormulaTest {

    @Test
    void buildJuchuAoTextsSplitSumFormula_embedsRowNumber() {
        String formula = ReconciliationApp.buildJuchuAoTextsSplitSumFormula(390);

        assertTrue(formula.contains("$AH390"));
        assertTrue(formula.contains("$AM390"));
        assertTrue(formula.contains("TEXTSPLIT"));
        assertTrue(formula.contains("CHAR(10)"));
        assertTrue(formula.startsWith("SUM("));
    }

    @Test
    void applyDefaultJuchuFormula_writesAoColumnWhenEmpty() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(2);
            header.createCell(JuchuSheetColumnLayout.columnLetterToIndex("AO")).setCellValue("単価");
            XSSFRow row = sheet.createRow(389);

            var method =
                    ReconciliationApp.class.getDeclaredMethod(
                            "applyDefaultJuchuFormulasIfMissing",
                            org.apache.poi.ss.usermodel.Row.class,
                            java.util.Map.class,
                            int.class);
            method.setAccessible(true);
            method.invoke(
                    null,
                    row,
                    java.util.Map.of(
                            "単価", JuchuSheetColumnLayout.columnLetterToIndex("AO")),
                    390);

            assertEquals(
                    CellType.FORMULA,
                    row.getCell(JuchuSheetColumnLayout.columnLetterToIndex("AO")).getCellType());
            assertEquals(
                    ReconciliationApp.buildJuchuAoTextsSplitSumFormula(390),
                    row.getCell(JuchuSheetColumnLayout.columnLetterToIndex("AO"))
                            .getCellFormula());
        }
    }
}
