package jp.co.pm.ai.desktop.reconciliation;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RequestFormOriginalWorkbookParserTest {

    @TempDir
    Path tempDir;

    @Test
    void parse_indexSheetOverridesSheetValuesAndRecordsConflicts() throws Exception {
        File file = tempDir.resolve("book.xlsm").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet index = wb.createSheet("目次");
            index.createRow(1).createCell(0).setCellValue("加工依頼NO");
            var dataRow = index.createRow(2);
            dataRow.createCell(0).setCellValue("T6-20");
            dataRow.createCell(8).setCellValue("6/9");
            dataRow.createCell(9).setCellValue("6/10");
            dataRow.createCell(10).setCellValue("6/22");
            dataRow.createCell(11).setCellValue("L備考");
            dataRow.createCell(13).setCellValue("185821Z");
            dataRow.createCell(14).setCellValue("O備考");

            XSSFSheet sheet = wb.createSheet("T6-20");
            sheet.createRow(4).createCell(17).setCellValue("T6-20");
            fillProductRow(sheet, 9, "6783", "15025", "JP17", "1360", "250");
            var deliveryRow = sheet.createRow(19);
            deliveryRow.createCell(8).setCellValue("6/15");
            deliveryRow.createCell(20).setCellValue("6/10");
            sheet.createRow(20).createCell(4).setCellValue("OLD-CONTRACT");
            fillRawRow(sheet, 22, "6780", "20020", "AP17", "1330", "300", "湖南", "5/15");

            try (FileOutputStream out = new FileOutputStream(file)) {
                wb.write(out);
            }
        }

        List<Map<String, String>> parsed = RequestFormOriginalWorkbookParser.parse(file);
        assertEquals(1, parsed.size());
        Map<String, String> raw = parsed.get(0);

        assertEquals("6/10", raw.get("投入日"));
        assertEquals("6/15", raw.get("希望納期"));
        assertEquals("6/22", raw.get("納期回答"));
        assertEquals("185821Z", raw.get("契約Ｎｏ"));
        assertEquals("6/9", raw.get(RequestFormOriginalIndexSheetMeta.KEY_RESPONSE_DATE));
        assertEquals("6/10", raw.get(RequestFormOriginalIndexSheetMeta.KEY_INPUT_DATE));
        assertEquals("6/22", raw.get(RequestFormOriginalIndexSheetMeta.KEY_DELIVERY_DATE));
        assertEquals("L備考", raw.get(RequestFormOriginalIndexSheetMeta.KEY_DELIVERY_REMARKS));
        assertEquals("185821Z", raw.get(RequestFormOriginalIndexSheetMeta.KEY_CONTRACT_NO));
        assertEquals("O備考", raw.get(RequestFormOriginalIndexSheetMeta.KEY_CONTRACT_REMARKS));
        assertEquals("true", raw.get(RequestFormOriginalIndexSheetMerger.META_INDEX_APPLIED));
        assertTrue(raw.containsKey(RequestFormOriginalIndexSheetMerger.META_INDEX_CONFLICTS));
        String conflicts = raw.get(RequestFormOriginalIndexSheetMerger.META_INDEX_CONFLICTS);
        assertTrue(conflicts.contains("投入日"));
        assertTrue(conflicts.contains("納期回答"));
        assertFalse(conflicts.contains("希望納期"));
        assertTrue(conflicts.contains("契約Ｎｏ"));
    }

    @Test
    void parse_withoutIndexSheet_keepsSheetValues() throws Exception {
        File file = tempDir.resolve("no-index.xlsm").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("E5-4");
            sheet.createRow(4).createCell(17).setCellValue("E5-4");
            fillProductRow(sheet, 9, "6783", "15025", "JP17", "1360", "250");
            sheet.createRow(19).createCell(8).setCellValue("2026-05-20");
            sheet.createRow(20).createCell(4).setCellValue("183784G");

            try (FileOutputStream out = new FileOutputStream(file)) {
                wb.write(out);
            }
        }

        List<Map<String, String>> parsed = RequestFormOriginalWorkbookParser.parse(file);
        assertEquals(1, parsed.size());
        Map<String, String> raw = parsed.get(0);
        assertEquals("2026-05-20", raw.get("希望納期"));
        assertEquals("183784G", raw.get("契約Ｎｏ"));
        assertFalse(raw.containsKey(RequestFormOriginalIndexSheetMerger.META_INDEX_APPLIED));
    }

    @Test
    void parse_indexMatchesSheetValues_noConflictMeta() throws Exception {
        File file = tempDir.resolve("match.xlsm").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet index = wb.createSheet("目次");
            index.createRow(1).createCell(0).setCellValue("加工依頼NO");
            var dataRow = index.createRow(2);
            dataRow.createCell(0).setCellValue("T6-21");
            dataRow.createCell(9).setCellValue("6/11");
            dataRow.createCell(10).setCellValue("6/22");
            dataRow.createCell(13).setCellValue("185821Z");

            XSSFSheet sheet = wb.createSheet("T6-21");
            sheet.createRow(4).createCell(17).setCellValue("T6-21");
            fillProductRow(sheet, 9, "6783", "15025", "JP17", "1360", "250");
            sheet.createRow(19).createCell(8).setCellValue("6/29");
            sheet.createRow(19).createCell(20).setCellValue("6/22");
            sheet.createRow(20).createCell(4).setCellValue("185821Z");
            fillRawRow(sheet, 22, "6780", "20020", "AP17", "1330", "300", "湖南", "6/11");

            try (FileOutputStream out = new FileOutputStream(file)) {
                wb.write(out);
            }
        }

        List<Map<String, String>> parsed = RequestFormOriginalWorkbookParser.parse(file);
        Map<String, String> raw = parsed.get(0);
        assertEquals("true", raw.get(RequestFormOriginalIndexSheetMerger.META_INDEX_APPLIED));
        assertFalse(raw.containsKey(RequestFormOriginalIndexSheetMerger.META_INDEX_CONFLICTS));
    }

    private static void fillProductRow(
            XSSFSheet sheet,
            int rowIndex,
            String hinmei,
            String part,
            String type,
            String width,
            String length) {
        var row = sheet.createRow(rowIndex);
        row.createCell(1).setCellValue(hinmei);
        row.createCell(5).setCellValue(part);
        row.createCell(10).setCellValue(type);
        row.createCell(15).setCellValue(width);
        row.createCell(20).setCellValue(length);
        row.createCell(30).setCellValue("250");
        row.createCell(23).setCellValue("F-A");
        row.createCell(26).setCellValue("色1");
        row.createCell(28).setCellValue("B");
    }

    private static void fillRawRow(
            XSSFSheet sheet,
            int rowIndex,
            String hinmei,
            String part,
            String type,
            String width,
            String length,
            String storage,
            String inputDate) {
        var row = sheet.createRow(rowIndex);
        row.createCell(7).setCellValue(hinmei);
        row.createCell(10).setCellValue(part);
        row.createCell(13).setCellValue(type);
        row.createCell(16).setCellValue(width);
        row.createCell(19).setCellValue(length);
        row.createCell(28).setCellValue("250");
        row.createCell(31).setCellValue(storage);
        if (inputDate != null) {
            row.createCell(RequestFormOriginalCellLayout.RawColumn.INPUT_DATE.columnIndex())
                    .setCellValue(inputDate);
        }
    }
}
