package jp.co.pm.ai.desktop.reconciliation;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RequestFormOriginalExtractorTest {

    @Test
    void translateYoto_maps3HbaToExport() {
        assertEquals("B（輸出）", RequestFormOriginalExtractor.translateYoto("3HBA"));
        assertEquals("W（自動車）", RequestFormOriginalExtractor.translateYoto("WA"));
    }

    @Test
    void normalizeEcSideForForm_mapsSingleLetter() {
        assertEquals("Ｈ面", RequestFormOriginalExtractor.normalizeEcSideForForm("H"));
        assertEquals("Ｑ面", RequestFormOriginalExtractor.normalizeEcSideForForm("Q"));
    }

    @Test
    void inferFeedLocation_detectsSlit() {
        assertEquals("ｽﾘｯﾄ", RequestFormOriginalExtractor.inferFeedLocation("EC, スリット, 輸出梱包"));
        assertEquals("EC", RequestFormOriginalExtractor.inferFeedLocation("EC"));
    }

    @Test
    void buildDbDefaultsFromRaw_onlyExtractableFields() {
        Map<String, String> raw =
                Map.ofEntries(
                        Map.entry("依頼Ｎｏ", "E5-4"),
                        Map.entry("品名", "6783"),
                        Map.entry("製品", "15025-JP17-1360X250"),
                        Map.entry("ＥＣ面", "H"),
                        Map.entry("用途", "3HBA"),
                        Map.entry("加工内容", "EC, スリット"),
                        Map.entry("特記事項1", "Q面外巻き"),
                        Map.entry("特記事項2", "タイプ変更"),
                        Map.entry("ユーザー", "共和興"));

        Map<String, String> db = RequestFormOriginalExtractor.buildDbDefaultsFromRaw(raw);
        assertEquals("6783", db.get("品名"));
        assertEquals("共和興", db.get("ユーザー"));
        assertEquals("Q面外巻き", db.get("特記事項1"));
        assertEquals("タイプ変更", db.get("特記事項2"));
        assertNull(db.get("用途"));
        assertNull(db.get("ＥＣ面"));
        assertNull(db.get("加工内容"));
        assertNull(db.get("入力区分"));
        assertNull(db.get("割数"));
    }

    @Test
    void buildRawMapFromSheet_multiProductRows() throws Exception {
        File file = new File("sample.xlsm");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("E5-4");
            sheet.createRow(4).createCell(17).setCellValue("E5-4");

            fillProductRow(sheet, 9, "6783", "15025", "JP17", "1360", "250");
            fillProductRow(sheet, 10, "6784", "15026", "JP18", "1370", "260");
            sheet.createRow(20).createCell(4).setCellValue("183784G");
            sheet.getRow(20).createCell(11).setCellValue("183785H");

            fillRawRow(sheet, 22, "6780", "15025", "AH1F", "1550", "250", "倉庫A", "2026-05-20");

            sheet.createRow(18).createCell(4).setCellValue("共和興");
            sheet.createRow(13).createCell(23).setCellValue("特記A");
            sheet.createRow(14).createCell(23).setCellValue("特記B");
            sheet.createRow(18).createCell(23).setCellValue("特記2");

            Map<String, String> raw = RequestFormOriginalExtractor.buildRawMapFromSheet(file, "E5-4", sheet);
            assertEquals("6783\n6784", raw.get("品名"));
            assertEquals("15025-JP17-1360X250\n15026-JP18-1370X260", raw.get("製品"));
            assertEquals("183784G\n183785H", raw.get("契約Ｎｏ"));
            assertEquals("6780", raw.get("原反品名"));
            assertEquals("15025-AH1F-1550X250", raw.get("原反"));
            assertEquals("2026-05-20", raw.get("投入日"));
            assertEquals("特記A 特記B", raw.get("特記事項1"));
            assertEquals("特記2", raw.get("特記事項2"));
        }
    }

    @Test
    void buildRawMapFromSheet_evaluatesFormulaKakochin() throws Exception {
        File file = new File("sample.xlsm");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("E6-2");
            sheet.createRow(4).createCell(17).setCellValue("E6-2");
            fillProductRow(sheet, 9, "6783", "15020", "NP17", "1300", "250");
            sheet.createRow(19).createCell(30).setCellFormula("18+18+13");

            Map<String, String> raw = RequestFormOriginalExtractor.buildRawMapFromSheet(file, "E6-2", sheet);
            assertEquals("49", raw.get("加工賃").replace(".0", "").strip());
        }
    }

    @Test
    void buildRawMapFromSheet_basicCells() throws Exception {
        File file = new File("sample.xlsm");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("E5-4");
            sheet.createRow(4).createCell(17).setCellValue("E5-4");
            fillProductRow(sheet, 9, "6783", "15025", "JP17", "1360", "250");
            sheet.createRow(18).createCell(4).setCellValue("共和興");
            sheet.createRow(19).createCell(8).setCellValue("2026-05-20");
            sheet.createRow(19).createCell(30).setCellValue("48");
            sheet.createRow(20).createCell(4).setCellValue("183784G");

            Map<String, String> raw = RequestFormOriginalExtractor.buildRawMapFromSheet(file, "E5-4", sheet);
            assertEquals("E5-4", raw.get("依頼Ｎｏ"));
            assertEquals("共和興", raw.get("ユーザー"));
            assertEquals("48", raw.get("加工賃"));
            assertEquals("183784G", raw.get("契約Ｎｏ"));
            assertTrue(raw.containsKey("加工内容"));
        }
    }

    @Test
    void buildRawMapFromSheet_threeProductContractsFromE21L21S21() throws Exception {
        File file = new File("sample.xlsm");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("E6-2");
            sheet.createRow(4).createCell(17).setCellValue("E6-2");
            fillProductRow(sheet, 9, "6783", "15020", "NP17", "1300", "250");
            fillProductRow(sheet, 10, "6784", "15021", "NP18", "1310", "260");
            fillProductRow(sheet, 11, "6785", "15022", "NP19", "1320", "270");
            var contractRow = sheet.createRow(20);
            contractRow.createCell(4).setCellValue("C-E21");
            contractRow.createCell(11).setCellValue("C-L21");
            contractRow.createCell(18).setCellValue("C-S21");

            Map<String, String> raw = RequestFormOriginalExtractor.buildRawMapFromSheet(file, "E6-2", sheet);
            assertEquals("C-E21\nC-L21\nC-S21", raw.get("契約Ｎｏ"));
        }
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
            String storage) {
        fillRawRow(sheet, rowIndex, hinmei, part, type, width, length, storage, null);
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
