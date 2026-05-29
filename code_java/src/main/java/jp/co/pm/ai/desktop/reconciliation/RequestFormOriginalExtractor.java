package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/** 加工依頼書原本シートから受注フォーム向けの値を抽出・正規化する。 */
final class RequestFormOriginalExtractor {

    private static final DataFormatter CELL_FORMATTER = new DataFormatter();

    /**
     * 依頼書原本の契約Ｎｏ改訂表記で使われる矢印（{@code →} / {@code ⇒} / {@code ->} 等）。
     * 長いパターンを先に並べ、最後に出現した矢印以降を改訂先契約Ｎｏとして採用する。
     */
    private static final java.util.regex.Pattern CONTRACT_REVISION_ARROW =
            java.util.regex.Pattern.compile(
                    "(?:"
                            + "==>|-->|=>|->"
                            + "|-\\s*>|=\\s*>"
                            + "|→|⇒|⇨|⇾|⟹|⟶|➔|➜|➝|➞|➡|➤|»|›"
                            + "|[＞>](?=\\s*[0-9A-Za-z])"
                            + ")");

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
        List<String> contracts = new ArrayList<>();

        int productSlot = 0;
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
            if (productSlot < RequestFormOriginalCellLayout.PRODUCT_CONTRACT_COLUMN_INDICES.length) {
                contracts.add(
                        resolveContractNoFromOriginalCell(
                                cellString(
                                        rawSheet,
                                        RequestFormOriginalCellLayout.PRODUCT_CONTRACT_ROW_INDEX,
                                        RequestFormOriginalCellLayout.PRODUCT_CONTRACT_COLUMN_INDICES[productSlot])));
            } else {
                contracts.add("");
            }
            productSlot++;
        }

        if (!hinmeis.isEmpty()) {
            rawMap.put("品名", RequestFormOriginalCellLayout.joinNonBlankLines(hinmeis));
            rawMap.put("製品", RequestFormOriginalCellLayout.joinNonBlankLines(products));
            rawMap.put("数量1", RequestFormOriginalCellLayout.joinNonBlankLines(qtys));
            rawMap.put("梱-等1", RequestFormOriginalCellLayout.joinNonBlankLines(grades));
            rawMap.put("色1", RequestFormOriginalCellLayout.joinNonBlankLines(colors));
            rawMap.put("区分1", RequestFormOriginalCellLayout.joinNonBlankLines(categories));
            rawMap.put(
                    "契約Ｎｏ",
                    String.join(
                            "\n",
                            contracts.stream().map(s -> s == null ? "" : s.strip()).toList()));
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
        List<String> inputDates = new ArrayList<>();

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
            inputDates.add(
                    cellString(rawSheet, rowIndex, RequestFormOriginalCellLayout.RawColumn.INPUT_DATE.columnIndex()));
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
            rawMap.put("投入日", RequestFormOriginalCellLayout.joinNonBlankLines(inputDates));
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

    /**
     * 依頼書原本の契約Ｎｏセル値を受注フォーム向けに解決する。
     * {@code A655440 → A22222} のように矢印で改訂先が書かれているときは、最後の矢印以降（{@code A22222}）を採用する。
     * 矢印は {@code →} / {@code ⇒} / {@code ➡} / {@code ->} / {@code =>} / {@code ＞} 等の表記ゆれに対応する。
     */
    static String resolveContractNoFromOriginalCell(String cellValue) {
        if (cellValue == null || cellValue.isBlank()) {
            return "";
        }
        String normalized =
                java.text.Normalizer.normalize(cellValue.strip(), java.text.Normalizer.Form.NFKC);
        java.util.regex.Matcher matcher = CONTRACT_REVISION_ARROW.matcher(normalized);
        int lastArrowEnd = -1;
        while (matcher.find()) {
            lastArrowEnd = matcher.end();
        }
        if (lastArrowEnd >= 0 && lastArrowEnd < normalized.length()) {
            return normalized.substring(lastArrowEnd).strip();
        }
        return normalized;
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

    /** 取り消し線セルの訂正値を下方向へ何行まで辿るかの上限（無限ループ防止）。 */
    private static final int MAX_STRIKE_CORRECTION_DEPTH = 10;

    private static String cellString(Sheet sheet, int rowIndex, int colIndex) {
        if (sheet == null) {
            return "";
        }
        int[] anchor = mergeAnchor(sheet, rowIndex, colIndex);
        return resolveCorrectedCellValue(sheet, anchor[0], anchor[1], 0);
    }

    /**
     * セル値を返す。取り消し線（strikeout）が付いたセルは、訂正値がすぐ下のセルに
     * 書かれていることが多いため、下のセルに非空の値があればそれを採用する。
     * 下のセルも取り消し線ならさらに下へ辿る（{@link #MAX_STRIKE_CORRECTION_DEPTH} 回まで）。
     * 下が空のときは取り消し線セルの値をそのまま返す。
     */
    private static String resolveCorrectedCellValue(
            Sheet sheet, int rowIndex, int colIndex, int depth) {
        Row row = sheet.getRow(rowIndex);
        Cell cell = row != null ? row.getCell(colIndex) : null;
        String value = getCellValueAsString(cell);
        if (depth >= MAX_STRIKE_CORRECTION_DEPTH || !cellHasStrikethrough(cell)) {
            return value;
        }
        int belowRowIndex = rowBelow(sheet, rowIndex, colIndex);
        Row belowRow = sheet.getRow(belowRowIndex);
        Cell belowCell = belowRow != null ? belowRow.getCell(colIndex) : null;
        String belowValue = getCellValueAsString(belowCell);
        if (belowValue.isBlank()) {
            return value;
        }
        if (cellHasStrikethrough(belowCell)) {
            return resolveCorrectedCellValue(sheet, belowRowIndex, colIndex, depth + 1);
        }
        return belowValue;
    }

    /** {@code (rowIndex, colIndex)} の直下行。縦結合セルのときは結合範囲の下の行。 */
    private static int rowBelow(Sheet sheet, int rowIndex, int colIndex) {
        if (sheet != null) {
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = sheet.getMergedRegion(i);
                if (region.isInRange(rowIndex, colIndex)) {
                    return region.getLastRow() + 1;
                }
            }
        }
        return rowIndex + 1;
    }

    /** セルのフォントに取り消し線が設定されているか。 */
    private static boolean cellHasStrikethrough(Cell cell) {
        if (cell == null) {
            return false;
        }
        try {
            CellStyle style = cell.getCellStyle();
            if (style == null) {
                return false;
            }
            Workbook wb = cell.getSheet().getWorkbook();
            Font font = wb.getFontAt(style.getFontIndexAsInt());
            return font != null && font.getStrikeout();
        } catch (RuntimeException ex) {
            return false;
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        try {
            FormulaEvaluator evaluator =
                    cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            String formatted =
                    RequestFormCellTextUtil.stripFormatLiteralQuotes(
                            cell, CELL_FORMATTER.formatCellValue(cell, evaluator));
            return formatted != null ? formatted.trim() : "";
        } catch (RuntimeException ex) {
            String formatted =
                    RequestFormCellTextUtil.stripFormatLiteralQuotes(
                            cell, CELL_FORMATTER.formatCellValue(cell));
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
