package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RequestFormSheetPreviewRendererTest {

    @Test
    void loadPreviewData_clipsToA1ThroughAn28(@TempDir Path tmp) throws Exception {
        File excel = tmp.resolve("preview-range.xlsx").toFile();
        try (Workbook wb = new XSSFWorkbook();
                FileOutputStream out = new FileOutputStream(excel)) {
            Sheet sheet = wb.createSheet("E5-1");
            sheet.createRow(0).createCell(0).setCellValue("左上");
            Row row29 = sheet.createRow(28);
            row29.createCell(0).setCellValue("範囲外行");
            row29.createCell(39).setCellValue("AN29");
            row29.createCell(40).setCellValue("AO29");
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
            wb.write(out);
        }

        RequestFormSheetPreviewRenderer.PreviewData data =
                RequestFormSheetPreviewRenderer.loadPreviewData(excel, "E5-1");

        assertEquals(28, data.rowCount());
        assertEquals(40, data.colCount());
        assertEquals(0, data.firstRow());
        assertEquals(0, data.firstCol());
        assertEquals("左上", data.texts()[0][0]);
        assertEquals("", data.texts()[27][0]);
        assertEquals("", data.texts()[27][39]);
        assertEquals(3, data.colSpans()[0][0]);
    }

    @Test
    void previewRangeSpec_isA1An28() {
        assertEquals("A1:AN28", RequestFormSheetPreviewRenderer.PREVIEW_RANGE_SPEC);
    }

    @Test
    void loadPreviewData_appliesFontSizeFromExcel(@TempDir Path tmp) throws Exception {
        File excel = tmp.resolve("font-size.xlsx").toFile();
        try (Workbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
                java.io.FileOutputStream out = new java.io.FileOutputStream(excel)) {
            Sheet sheet = wb.createSheet("E5-1");
            var cell = sheet.createRow(0).createCell(0);
            cell.setCellValue("大");
            Font font = wb.createFont();
            font.setFontHeightInPoints((short) 18);
            font.setColor(org.apache.poi.ss.usermodel.IndexedColors.RED.getIndex());
            font.setBold(true);
            font.setItalic(true);
            font.setStrikeout(true);
            var style = wb.createCellStyle();
            style.setFont(font);
            cell.setCellStyle(style);
            wb.write(out);
        }

        RequestFormSheetPreviewRenderer.PreviewData data =
                RequestFormSheetPreviewRenderer.loadPreviewData(excel, "E5-1");
        RequestFormPreviewCellStyle style = data.styles()[0][0];
        assertEquals(18.0 * 96.0 / 72.0, style.fontSizePx(), 0.01);
        assertTrue(style.bold());
        assertTrue(style.italic());
        assertTrue(style.strike());
    }

    @Test
    void loadPreviewData_numericCell_doesNotThrow(@TempDir Path tmp) throws Exception {
        File excel = tmp.resolve("numeric.xlsx").toFile();
        try (Workbook wb = new XSSFWorkbook();
                FileOutputStream out = new FileOutputStream(excel)) {
            Sheet sheet = wb.createSheet("E5-1");
            sheet.createRow(0).createCell(0).setCellValue(12345.67);
            sheet.createRow(1).createCell(0).setCellValue(20260524);
            wb.write(out);
        }

        RequestFormSheetPreviewRenderer.PreviewData data =
                RequestFormSheetPreviewRenderer.loadPreviewData(excel, "E5-1");

        assertEquals("12345.67", data.texts()[0][0]);
        assertEquals("20260524", data.texts()[1][0]);
    }

    @Test
    void loadPreviewData_readsMergedAnchorValue(@TempDir Path tmp) throws Exception {
        File excel = tmp.resolve("merge-anchor.xlsx").toFile();
        try (Workbook wb = new XSSFWorkbook();
                FileOutputStream out = new FileOutputStream(excel)) {
            Sheet sheet = wb.createSheet("E5-3");
            sheet.createRow(1).createCell(0).setCellValue("長岡産業(株) 湖南工場 殿");
            sheet.createRow(4).createCell(12).setCellValue("E5-3");
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 10));
            wb.write(out);
        }

        RequestFormSheetPreviewRenderer.PreviewData data =
                RequestFormSheetPreviewRenderer.loadPreviewData(excel, "E5-3");

        assertEquals("長岡産業(株) 湖南工場 殿", data.texts()[1][0]);
        assertEquals(11, data.colSpans()[1][0]);
        assertEquals("E5-3", data.texts()[4][12]);
    }
}
