package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class ResultExcelSummaryBorderTest {

    @Test
    @DisplayName("見出しの下線は結合範囲の全セルに引く")
    void sectionUnderlineSpansMergedRange() throws Exception {
        try (XSSFWorkbook wb = ResultExcelExporter.buildWorkbook(stubResult(), null, null)) {
            XSSFSheet ws = wb.getSheet("サマリ");
            int rowIndex = findRow(ws, "総額");
            assertTrue(rowIndex >= 0, "総額 行が無い");
            Row row = ws.getRow(rowIndex);
            for (int c = 1; c <= 6; c++) {
                Cell cell = row.getCell(c);
                assertNotNull(cell, "見出し結合の欠落 col=" + c);
                assertEquals(BorderStyle.MEDIUM, cell.getCellStyle().getBorderBottom(), "bottom col=" + c);
            }
        }
    }

    @Test
    @DisplayName("KPIカードは3行の外枠がある")
    void kpiCardsHaveOuterBox() throws Exception {
        try (XSSFWorkbook wb = ResultExcelExporter.buildWorkbook(stubResult(), null, null)) {
            XSSFSheet ws = wb.getSheet("サマリ");
            for (int col = 1; col <= 6; col++) {
                CellStyle top = cellStyle(ws, 3, col);
                CellStyle mid = cellStyle(ws, 4, col);
                CellStyle bot = cellStyle(ws, 5, col);
                assertEquals(BorderStyle.THIN, top.getBorderTop(), "kpi top col=" + col);
                assertEquals(BorderStyle.THIN, bot.getBorderBottom(), "kpi bottom col=" + col);
                assertEquals(BorderStyle.THIN, top.getBorderLeft(), "kpi left col=" + col);
                assertEquals(BorderStyle.THIN, top.getBorderRight(), "kpi right col=" + col);
                assertEquals(BorderStyle.THIN, mid.getBorderLeft(), "kpi mid left col=" + col);
                assertEquals(BorderStyle.THIN, mid.getBorderRight(), "kpi mid right col=" + col);
            }
        }
    }

    @Test
    @DisplayName("明細行の罫線はラベルから金額列まで閉じる")
    void kvRowBoxSpansLabelToValue() throws Exception {
        try (XSSFWorkbook wb = ResultExcelExporter.buildWorkbook(stubResult(), null, null)) {
            XSSFSheet ws = wb.getSheet("サマリ");
            int rowIndex = findRowPrefix(ws, "① 東レ検収");
            assertTrue(rowIndex >= 0, "① 東レ検収 行が無い");
            Row row = ws.getRow(rowIndex);
            assertEquals(BorderStyle.THIN, row.getCell(1).getCellStyle().getBorderLeft(), "label left");
            assertEquals(BorderStyle.THIN, row.getCell(1).getCellStyle().getBorderTop(), "label top");
            assertEquals(BorderStyle.THIN, row.getCell(1).getCellStyle().getBorderBottom(), "label bottom");
            for (int c = 1; c <= 6; c++) {
                Cell cell = row.getCell(c);
                assertNotNull(cell, "明細結合の欠落 col=" + c);
                assertEquals(BorderStyle.THIN, cell.getCellStyle().getBorderTop(), "top col=" + c);
                assertEquals(BorderStyle.THIN, cell.getCellStyle().getBorderBottom(), "bottom col=" + c);
            }
            assertEquals(BorderStyle.THIN, row.getCell(6).getCellStyle().getBorderRight(), "value right");
        }
    }

    private static VerifyResult stubResult() {
        Map<String, Object> info = new HashMap<>();
        info.put("対象月ラベル", "2026年8月度");
        info.put("実行日時", "2026-09-17");
        info.put("実行者", "test");
        info.put("許容差", 0.5);
        info.put("入庫場所", "A010");
        info.put("②名称", "長岡明細");
        info.put("②金額列", "AA");
        info.put("①総額", 1L);
        info.put("②総額", 2L);
        info.put("③総額", 3L);
        return new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                new YearMonthKey(2026, 8),
                List.of(),
                List.of(),
                info,
                List.of(),
                null,
                null,
                null);
    }

    private static int findRow(XSSFSheet ws, String text) {
        for (Row row : ws) {
            Cell c = row.getCell(1);
            if (c != null && text.equals(c.getStringCellValue())) {
                return row.getRowNum();
            }
        }
        return -1;
    }

    private static int findRowPrefix(XSSFSheet ws, String prefix) {
        for (Row row : ws) {
            Cell c = row.getCell(1);
            if (c != null && c.getStringCellValue() != null && c.getStringCellValue().startsWith(prefix)) {
                return row.getRowNum();
            }
        }
        return -1;
    }

    private static CellStyle cellStyle(XSSFSheet ws, int r, int c) {
        Cell cell = ws.getRow(r).getCell(c);
        assertNotNull(cell, "cell r=" + r + " c=" + c);
        return cell.getCellStyle();
    }
}
