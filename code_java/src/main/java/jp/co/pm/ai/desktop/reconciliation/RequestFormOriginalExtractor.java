package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/** 加工依頼書原本シートから受注フォーム向けの値を抽出・正規化する。 */
final class RequestFormOriginalExtractor {

    private static final DataFormatter CELL_FORMATTER = new DataFormatter();

    private RequestFormOriginalExtractor() {}

    static Map<String, String> buildRawMapFromSheet(File file, String sName, Sheet rawSheet) {
        Map<String, String> rawMap = new LinkedHashMap<>();

        for (RequestFormOriginalCellLayout.BasicField field : RequestFormOriginalCellLayout.BasicField.values()) {
            RequestFormOriginalCellLayout.CellAddress cell = field.cell();
            rawMap.put(field.rawKey(), cellString(rawSheet, cell.rowIndex(), cell.columnIndex()));
        }

        List<String> hinmeis = new ArrayList<>();
        List<String> products = new ArrayList<>();
        List<String> qtys = new ArrayList<>();
        List<String> grades = new ArrayList<>();
        List<String> colors = new ArrayList<>();
        List<String> categories = new ArrayList<>();

        for (int rowIndex : RequestFormOriginalCellLayout.PRODUCT_ROW_INDICES) {
            if (!RequestFormOriginalCellLayout.isProductRowPopulated(thisCellReader(rawSheet), rowIndex)) {
                continue;
            }
            String hinmei = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.ProductColumn.HINMEI.columnIndex());
            String part = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.ProductColumn.PART_NO.columnIndex());
            String type = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.ProductColumn.TYPE.columnIndex());
            String width = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.ProductColumn.WIDTH.columnIndex());
            String length = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.ProductColumn.LENGTH.columnIndex());

            hinmeis.add(hinmei);
            products.add(JuchuSheetColumnLayout.buildSpecName(part, type, width, length));
            qtys.add(cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.ProductColumn.QTY.columnIndex()));
            grades.add(cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.ProductColumn.GRADE.columnIndex()));
            colors.add(cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.ProductColumn.COLOR.columnIndex()));
            categories.add(cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.ProductColumn.CATEGORY.columnIndex()));
        }

        if (!hinmeis.isEmpty()) {
            rawMap.put("品名", RequestFormOriginalCellLayout.joinNonBlankLines(hinmeis));
            rawMap.put("製品", RequestFormOriginalCellLayout.joinNonBlankLines(products));
            rawMap.put("数量1", RequestFormOriginalCellLayout.joinNonBlankLines(qtys));
            rawMap.put("梱-等1", RequestFormOriginalCellLayout.joinNonBlankLines(grades));
            rawMap.put("色1", RequestFormOriginalCellLayout.joinNonBlankLines(colors));
            rawMap.put("区分1", RequestFormOriginalCellLayout.joinNonBlankLines(categories));
        }

        int firstProductRow = RequestFormOriginalCellLayout.PRODUCT_ROW_INDICES[0];
        rawMap.put(
                "ＥＣ面",
                cellString(
                        rawSheet,
                        firstProductRow,
                        RequestFormOriginalCellLayout.ProductColumn.EC_SIDE.columnIndex()));
        rawMap.put(
                "ﾄﾘﾐﾝｸﾞ",
                cellString(
                        rawSheet,
                        firstProductRow,
                        RequestFormOriginalCellLayout.ProductColumn.TRIMMING.columnIndex()));

        List<String> rawHinmeis = new ArrayList<>();
        List<String> rawSpecs = new ArrayList<>();
        List<String> rawQtys = new ArrayList<>();
        List<String> rawGrades = new ArrayList<>();
        List<String> rawColors = new ArrayList<>();
        List<String> rawCategories = new ArrayList<>();
        List<String> storages = new ArrayList<>();

        for (int rowIndex : RequestFormOriginalCellLayout.RAW_ROW_INDICES) {
            if (!RequestFormOriginalCellLayout.isRawRowPopulated(thisCellReader(rawSheet), rowIndex)) {
                continue;
            }
            String hinmei = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.HINMEI.columnIndex());
            String part = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.PART_NO.columnIndex());
            String type = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.TYPE.columnIndex());
            String width = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.WIDTH.columnIndex());
            String length = cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.LENGTH.columnIndex());

            rawHinmeis.add(hinmei);
            rawSpecs.add(JuchuSheetColumnLayout.buildSpecName(part, type, width, length));
            rawQtys.add(cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.QTY.columnIndex()));
            rawGrades.add(cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.GRADE.columnIndex()));
            rawColors.add(cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.COLOR.columnIndex()));
            rawCategories.add(cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.CATEGORY.columnIndex()));
            storages.add(cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.STORAGE.columnIndex()));
        }

        if (!rawHinmeis.isEmpty()) {
            String joinedHinmei = RequestFormOriginalCellLayout.joinNonBlankLines(rawHinmeis);
            rawMap.put("原反品名", joinedHinmei);
            rawMap.put("品名1", joinedHinmei);
            rawMap.put("原反", RequestFormOriginalCellLayout.joinNonBlankLines(rawSpecs));
            rawMap.put("原反数量", RequestFormOriginalCellLayout.joinNonBlankLines(rawQtys));
            rawMap.put("原反梱-等", RequestFormOriginalCellLayout.joinNonBlankLines(rawGrades));
            rawMap.put("原反色", RequestFormOriginalCellLayout.joinNonBlankLines(rawColors));
            rawMap.put("原反区分", RequestFormOriginalCellLayout.joinNonBlankLines(rawCategories));
            rawMap.put("在庫場所", RequestFormOriginalCellLayout.joinNonBlankLines(storages));
        }

        rawMap.put("加工内容", readProcessingSteps(rawSheet));
        assignTokkiFromSheet(rawMap, rawSheet);

        String reqNo = rawMap.getOrDefault("依頼Ｎｏ", "").trim();
        if (reqNo.isEmpty()) {
            reqNo = sName;
            rawMap.put("依頼Ｎｏ", reqNo);
        }

        rawMap.put("原本ファイル名", file.getName());
        rawMap.put("原本シート名", sName);
        return rawMap;
    }

    /** 受注ファイル未登録行向けに、原本 raw からフォーム初期値 dbValues を組み立てる。 */
    static Map<String, String> buildDbDefaultsFromRaw(Map<String, String> raw) {
        Map<String, String> db = new LinkedHashMap<>();
        if (raw == null || raw.isEmpty()) {
            return db;
        }
        for (String key : RequestFormOriginalCellLayout.FORM_EXTRACT_RAW_KEYS) {
            putIfPresent(db, key, raw.get(key));
        }
        if (db.containsKey("原反品名") && !db.containsKey("品名1")) {
            db.put("品名1", db.get("原反品名"));
        }
        return db;
    }

    static String translateYoto(String useCode) {
        if (useCode == null || useCode.isBlank()) {
            return "";
        }
        String code = useCode.strip().toUpperCase();
        if (code.contains("WA") || code.equals("W")) {
            return "W（自動車）";
        }
        if (code.contains("BA") || code.equals("B")) {
            return "B（輸出）";
        }
        if (code.contains("YA") || code.equals("Y")) {
            return "Y（工材）";
        }
        if (code.contains("VA") || code.equals("V")) {
            return "V（TPI）";
        }
        if (code.contains("ZA") || code.equals("Z")) {
            return useCode.strip();
        }
        return useCode.strip();
    }

    static String normalizeEcSideForForm(String ec) {
        if (ec == null || ec.isBlank()) {
            return "";
        }
        String trimmed = ec.strip();
        String normalized =
                java.text.Normalizer.normalize(trimmed, java.text.Normalizer.Form.NFKC).toUpperCase();
        if (normalized.equals("H") || normalized.equals("Ｈ") || normalized.startsWith("H面")) {
            return "Ｈ面";
        }
        if (normalized.equals("Q") || normalized.equals("Ｑ") || normalized.startsWith("Q面")) {
            return "Ｑ面";
        }
        if (trimmed.contains("両面")) {
            return "両面";
        }
        if (trimmed.contains("スライス") || trimmed.contains("ｽﾗｲｽ")) {
            return "ｽﾗｲｽ面";
        }
        if (trimmed.contains("スキン") || trimmed.contains("ｽｷﾝ")) {
            return "ｽｷﾝ面";
        }
        return trimmed;
    }

    static String inferFeedLocation(String processingContent) {
        if (processingContent == null || processingContent.isBlank()) {
            return "";
        }
        String p = processingContent.replace("、", ",").toUpperCase();
        if (p.contains("スリット") || p.contains("SLIT") || p.contains("ｽﾘｯﾄ")) {
            return "ｽﾘｯﾄ";
        }
        if (p.contains("SEC")) {
            return "SEC";
        }
        if (p.contains("EC")) {
            return "EC";
        }
        if (p.contains("スライス") || p.contains("SLICE") || p.contains("ｽﾗｲｽ")) {
            return "ｽﾗｲｽ";
        }
        if (p.contains("エンボス") || p.contains("EMBOSS") || p.contains("ｴﾝﾎﾞｽ")) {
            return "ｴﾝﾎﾞｽ";
        }
        if (p.contains("検反")) {
            return "検反";
        }
        if (p.contains("融着")) {
            return "融着";
        }
        return "";
    }

    private static void assignTokkiFromSheet(Map<String, String> rawMap, Sheet sheet) {
        List<String> tokki1Parts = new ArrayList<>();
        int lastAnchorRow = -1;
        int lastAnchorCol = -1;
        for (int rowIndex : RequestFormOriginalCellLayout.TOKKI_1_ROW_INDICES) {
            int[] anchor =
                    mergeAnchor(
                            sheet, rowIndex, RequestFormOriginalCellLayout.TOKKI_COLUMN_INDEX);
            if (anchor[0] == lastAnchorRow && anchor[1] == lastAnchorCol) {
                continue;
            }
            lastAnchorRow = anchor[0];
            lastAnchorCol = anchor[1];
            String text = cellString(sheet, rowIndex, RequestFormOriginalCellLayout.TOKKI_COLUMN_INDEX);
            if (!text.isBlank()) {
                tokki1Parts.add(text.strip());
            }
        }
        String tokki1 = RequestFormOriginalCellLayout.joinNonBlankParts(tokki1Parts);
        if (!tokki1.isBlank()) {
            rawMap.put("特記事項1", tokki1);
        }

        String tokki2 =
                cellString(sheet, RequestFormOriginalCellLayout.TOKKI_2_ROW_INDEX, RequestFormOriginalCellLayout.TOKKI_COLUMN_INDEX);
        if (!tokki2.isBlank()) {
            rawMap.put("特記事項2", tokki2.strip());
        }
    }

    private static String readProcessingSteps(Sheet sheet) {
        List<String> steps = new ArrayList<>();
        for (int rowIndex : RequestFormOriginalCellLayout.PROCESS_STEP_ROW_INDICES) {
            String step = cellString(sheet, rowIndex, RequestFormOriginalCellLayout.PROCESS_STEP_COLUMN_INDEX).trim();
            if (!step.isEmpty()) {
                steps.add(step);
            }
        }
        return String.join(", ", steps);
    }

    private static java.util.function.BiFunction<Integer, Integer, String> thisCellReader(Sheet sheet) {
        return (row, col) -> cellString(sheet, row, col);
    }

    private static String cellString(Sheet sheet, int rowIndex, int colIndex) {
        if (sheet == null) {
            return "";
        }
        int[] anchor = mergeAnchor(sheet, rowIndex, colIndex);
        Row row = sheet.getRow(anchor[0]);
        if (row == null) {
            return "";
        }
        return getCellValueAsString(row.getCell(anchor[1]));
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        try {
            FormulaEvaluator evaluator =
                    cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            String formatted = CELL_FORMATTER.formatCellValue(cell, evaluator);
            return formatted != null ? formatted.trim() : "";
        } catch (RuntimeException ex) {
            String formatted = CELL_FORMATTER.formatCellValue(cell);
            return formatted != null ? formatted.trim() : "";
        }
    }

    private static void putIfPresent(Map<String, String> db, String key, String value) {
        if (value != null && !value.isBlank()) {
            db.put(key, value.trim());
        }
    }

    private static int[] mergeAnchor(Sheet sheet, int rowIndex, int colIndex) {
        if (sheet == null) {
            return new int[] {rowIndex, colIndex};
        }
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(rowIndex, colIndex)) {
                return new int[] {region.getFirstRow(), region.getFirstColumn()};
            }
        }
        return new int[] {rowIndex, colIndex};
    }
}
