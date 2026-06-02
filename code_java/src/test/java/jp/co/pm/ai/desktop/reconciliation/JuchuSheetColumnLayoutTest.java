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
        assertEquals(34, JuchuSheetColumnLayout.columnLetterToIndex("AI"));
        assertEquals(4, JuchuSheetColumnLayout.columnLetterToIndex("E"));
        assertEquals(24, JuchuSheetColumnLayout.columnLetterToIndex("Y"));
        assertEquals("AP", JuchuSheetColumnLayout.indexToColumnLetter(41));
        assertEquals("Y", JuchuSheetColumnLayout.Col.TONYU_BI.columnLetter());
        assertEquals("投入日", JuchuSheetColumnLayout.Col.TONYU_BI.primaryHeader());
    }

    @Test
    void matchesHeader_acceptsAliases() {
        JuchuSheetColumnLayout.Col irai = JuchuSheetColumnLayout.Col.IRAI_NO;
        assertTrue(irai.matchesHeader("依頼No"));
        assertTrue(irai.matchesHeader("依頼Ｎｏ"));

        JuchuSheetColumnLayout.Col ec = JuchuSheetColumnLayout.Col.EC_MEN;
        assertTrue(ec.matchesHeader("EC面"));
        assertTrue(ec.matchesHeader("ＥＣ面"));

        JuchuSheetColumnLayout.Col warisu = JuchuSheetColumnLayout.Col.WARISU;
        assertTrue(warisu.matchesHeader("加工回数（加工換算数に利用）"));
    }

    @Test
    void collectHeaderMismatches_expectedOverrideWithAliasAcceptsDifferentActualHeader() throws Exception {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:\\test\\juchu.xlsm";
        var col = JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT;
        registry.setExpectedOverride(path, col, "商品(製品)");
        registry.addAlias(path, col, "タイプ");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(col.columnIndex()).setCellValue("タイプ");

            var mismatches =
                    JuchuSheetColumnLayout.collectHeaderMismatches(header, registry, path);
            assertTrue(mismatches.stream().noneMatch(m -> m.column() == col));
            assertTrue(
                    JuchuSheetColumnLayout.headerMatches(
                            col, "タイプ", registry, path));
        }
    }

    @Test
    void collectHeaderMismatches_expectedOverrideWithoutAliasStillMismatchWhenActualDiffers()
            throws Exception {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:\\test\\juchu.xlsm";
        var col = JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT;
        registry.setExpectedOverride(path, col, "商品(製品)");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(col.columnIndex()).setCellValue("タイプ");

            var mismatches =
                    JuchuSheetColumnLayout.collectHeaderMismatches(header, registry, path);
            assertTrue(mismatches.stream().anyMatch(m -> m.column() == col));
        }
    }

    @Test
    void collectHeaderMismatches_expectedOverrideSuppressesWarning() throws Exception {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:\\test\\juchu.xlsm";
        registry.setExpectedOverride(path, JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT, "タイプ");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT.columnIndex())
                    .setCellValue("タイプ");

            var mismatches =
                    JuchuSheetColumnLayout.collectHeaderMismatches(header, registry, path);
            assertTrue(
                    mismatches.stream()
                            .noneMatch(
                                    m ->
                                            m.column()
                                                    == JuchuSheetColumnLayout.Col
                                                            .MASTER_BASE_SHOHIN_PRODUCT));
            assertEquals(
                    "タイプ",
                    registry.expectedHeaderFor(
                            path, JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT));
        }
    }

    @Test
    void collectHeaderMismatches_expectedOverrideAllowsEmptyHeader() throws Exception {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:\\test\\juchu.xlsm";
        registry.setExpectedOverride(path, JuchuSheetColumnLayout.Col.IRO, "色");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);

            var mismatches =
                    JuchuSheetColumnLayout.collectHeaderMismatches(header, registry, path);
            assertTrue(
                    mismatches.stream().noneMatch(m -> m.column() == JuchuSheetColumnLayout.Col.IRO));
        }
    }

    @Test
    void readExcelHeaderPicks_listsNonEmptyHeaders() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT.columnIndex())
                    .setCellValue("タイプ");

            var picks = JuchuSheetColumnLayout.readExcelHeaderPicks(header);
            assertEquals(1, picks.size());
            assertEquals("AP", picks.get(0).columnLetter());
            assertEquals("タイプ", picks.get(0).headerText());
        }
    }

    @Test
    void readExcelHeaderPicks_includesHeaderBeforeShortEmptyGap() throws Exception {
        int brIndex = JuchuSheetColumnLayout.columnLetterToIndex("BR");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(brIndex).setCellValue("BR見出し");
            header.createCell(brIndex + 1);

            var picks = JuchuSheetColumnLayout.readExcelHeaderPicks(header);
            assertEquals(1, picks.size());
            assertEquals("BR", picks.get(0).columnLetter());
            assertEquals("BR見出し", picks.get(0).headerText());
        }
    }

    @Test
    void readExcelHeaderPicks_stopsAfterTenConsecutiveEmptyCells() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(5).setCellValue("手前");
            for (int c = 6; c < 16; c++) {
                header.createCell(c);
            }
            header.createCell(20).setCellValue("以降除外");

            var picks = JuchuSheetColumnLayout.readExcelHeaderPicks(header);
            assertEquals(1, picks.size());
            assertEquals("手前", picks.get(0).headerText());
            assertEquals(
                    6, JuchuSheetColumnLayout.resolveHeaderPickScanExclusiveEnd(header));
        }
    }

    @Test
    void collectHeaderMismatches_registryAliasSuppressesWarning() throws Exception {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:\\test\\juchu.xlsm";
        registry.addAlias(path, JuchuSheetColumnLayout.Col.IRO, "色呼称");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(JuchuSheetColumnLayout.Col.IRO.columnIndex()).setCellValue("色呼称");

            var mismatches =
                    JuchuSheetColumnLayout.collectHeaderMismatches(header, registry, path);
            assertTrue(
                    mismatches.stream().noneMatch(m -> m.column() == JuchuSheetColumnLayout.Col.IRO));
        }
    }

    @Test
    void collectUnknownExcelColumns_listsOnlyNonKnownIndices() throws Exception {
        int afterAqIndex = JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_RAW.columnIndex() + 1;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_RAW.columnIndex())
                    .setCellValue("原反商品");
            header.createCell(afterAqIndex).setCellValue("追加列");

            var unknown =
                    JuchuSheetColumnLayout.collectUnknownExcelColumns(header, null, "C:\\test\\juchu.xlsm");
            assertEquals(1, unknown.size());
            assertEquals("追加列", unknown.get(0).headerText());
        }
    }

    @Test
    void collectAllKnownColumns_listsEveryLayoutColumn() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(JuchuSheetColumnLayout.Col.IRO.columnIndex()).setCellValue("色");

            var all =
                    JuchuSheetColumnLayout.collectAllKnownColumns(header, null, "C:\\test\\juchu.xlsm");
            assertEquals(JuchuSheetColumnLayout.Col.values().length, all.size());
            assertTrue(
                    all.stream().anyMatch(m -> m.column() == JuchuSheetColumnLayout.Col.IRO));
        }
    }

    @Test
    void formItemDescription_mapsToFormSections() {
        assertEquals(
                "【原反（材料）】色",
                JuchuSheetColumnLayout.Col.IRO.formItemDescription());
        assertEquals(
                "【製品（仕上がり）】契約Ｎｏ",
                JuchuSheetColumnLayout.Col.KEIYAKU_NO.formItemDescription());
        assertEquals(
                "【製品（仕上がり）】商品（masterBase）",
                JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT.formItemDescription());
    }

    @Test
    void summaryLine_includesFormItem() {
        var mismatch =
                new JuchuHeaderMismatch(
                        JuchuSheetColumnLayout.Col.IRO,
                        "色",
                        "",
                        true);
        assertTrue(mismatch.summaryLine().contains("【原反（材料）】色"));
        assertTrue(mismatch.summaryLine().contains("T列"));
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
    void computeRawRollCountFromQtyAndLength_floorsInteger() {
        assertEquals(1, JuchuSheetColumnLayout.computeRawRollCountFromQtyAndLength("250", "250").orElse(-1));
        assertEquals(0, JuchuSheetColumnLayout.computeRawRollCountFromQtyAndLength("249", "250").orElse(-1));
        assertEquals(2, JuchuSheetColumnLayout.computeRawRollCountFromQtyAndLength("500", "250").orElse(-1));
        assertEquals(27, JuchuSheetColumnLayout.computeRawRollCountFromQtyAndLength("8,100", "300").orElse(-1));
        assertEquals(27, JuchuSheetColumnLayout.computeRawRollCountFromQtyAndLength("8，100", "300").orElse(-1));
        assertTrue(JuchuSheetColumnLayout.computeRawRollCountFromQtyAndLength("", "250").isEmpty());
        assertTrue(JuchuSheetColumnLayout.computeRawRollCountFromQtyAndLength("250", "0").isEmpty());
        assertEquals(35, JuchuSheetColumnLayout.columnLetterToIndex("AJ"));
        assertEquals("AJ", JuchuSheetColumnLayout.Col.GENPAN_ROLL_SU.columnLetter());
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

    @Test
    void collectHeaderMismatches_skipsExcludedColumns() throws Exception {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:\\test\\juchu.xlsm";
        registry.setExcludedFromTransfer(path, JuchuSheetColumnLayout.Col.NYURYOKU_BI);

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow header = sheet.createRow(JuchuSheetColumnLayout.HEADER_ROW_INDEX);
            header.createCell(JuchuSheetColumnLayout.Col.NYURYOKU_BI.columnIndex()).setCellValue("投入日");

            var mismatches =
                    JuchuSheetColumnLayout.collectHeaderMismatches(header, registry, path);
            assertTrue(
                    mismatches.stream()
                            .noneMatch(m -> m.column() == JuchuSheetColumnLayout.Col.NYURYOKU_BI));
        }
    }

    @Test
    void readDbValuesFromRow_skipsExcludedColumns() throws Exception {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:\\test\\juchu.xlsm";
        registry.setExcludedFromTransfer(path, JuchuSheetColumnLayout.Col.HINMEI);

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow row = sheet.createRow(3);
            row.createCell(JuchuSheetColumnLayout.Col.HINMEI.columnIndex()).setCellValue("6713");
            row.createCell(JuchuSheetColumnLayout.Col.SEIHIN.columnIndex())
                    .setCellValue("20010-H600-1180X250");

            var vals = JuchuSheetColumnLayout.readDbValuesFromRow(row, registry, path);
            assertFalse(vals.containsKey("品名"));
            assertEquals("20010-H600-1180X250", vals.get("製品"));
        }
    }
}
