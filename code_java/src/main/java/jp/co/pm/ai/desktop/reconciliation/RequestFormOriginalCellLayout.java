package jp.co.pm.ai.desktop.reconciliation;

import java.util.List;

/**
 * 加工依頼書原本シートのフォーム項目↔セル座標の正本。
 * POI は 0-based（Excel 行10 → rowIndex 9）。
 */
public final class RequestFormOriginalCellLayout {

    /** 製品（仕上がり）ブロック: Excel 行 10–12。 */
    public static final int[] PRODUCT_ROW_INDICES = {9, 10, 11};

    /** 製品行ごとの契約Ｎｏ: Excel 行 21 の E / L / S（POI row 20）。 */
    public static final int PRODUCT_CONTRACT_ROW_INDEX = 20;

    public static final int[] PRODUCT_CONTRACT_COLUMN_INDICES = {
        columnLetterToIndex("E"), columnLetterToIndex("L"), columnLetterToIndex("S")
    };

    /** 原反（材料）ブロック: Excel 行 23–25。 */
    public static final int[] RAW_ROW_INDICES = {22, 23, 24};

    /** 特記事項1: Excel X14–X18（POI row 13–17, col 23）。 */
    public static final int[] TOKKI_1_ROW_INDICES = {13, 14, 15, 16, 17};

    /** 特記事項2: Excel X19（POI row 18, col 23）。 */
    public static final int TOKKI_2_ROW_INDEX = 18;

    public static final int TOKKI_COLUMN_INDEX = columnLetterToIndex("X");

    /** 照合用: 加工内容ステップ I13–I17。 */
    public static final int[] PROCESS_STEP_ROW_INDICES = {12, 13, 14, 15, 16};

    public static final int PROCESS_STEP_COLUMN_INDEX = columnLetterToIndex("I");

    public record CellAddress(int rowIndex, int columnIndex, String excelRef) {}

    /** 【受注データ手入力・修正フォーム】単一セル項目。 */
    public enum BasicField {
        IRAI_NO("依頼Ｎｏ", true, 4, "R", 5),
        USER("ユーザー", true, 18, "E", 19),
        /** 出荷希望（依頼シート I20:O20 結合セル）。 */
        KIBO_NOKI("希望納期", true, 19, "I", 20),
        /** 納期回答（依頼シート U20:Z20 結合セル）。目次「納期」と照合する正本。 */
        NOKI_KAITO("納期回答", true, 19, "U", 20),
        KAKOCHIN("加工賃", true, 19, "AE", 20),
        YOTO_COMPARE("用途", false, 17, "E", 18);

        private final String rawKey;
        private final boolean extractToForm;
        private final int rowIndex;
        private final int columnIndex;
        private final String excelRef;

        BasicField(String rawKey, boolean extractToForm, int rowIndex, String colLetter, int excelRow) {
            this.rawKey = rawKey;
            this.extractToForm = extractToForm;
            this.rowIndex = rowIndex;
            this.columnIndex = columnLetterToIndex(colLetter);
            this.excelRef = colLetter + excelRow;
        }

        public String rawKey() {
            return rawKey;
        }

        public boolean extractToForm() {
            return extractToForm;
        }

        public CellAddress cell() {
            return new CellAddress(rowIndex, columnIndex, excelRef);
        }
    }

    /** 【製品（仕上がり）情報】行内列（行 10/11/12 共通）。 */
    public enum ProductColumn {
        HINMEI("品名", true, "B"),
        PART_NO(null, true, "F"),
        TYPE(null, true, "K"),
        WIDTH(null, true, "P"),
        LENGTH(null, true, "U"),
        QTY("数量1", true, "AE"),
        GRADE("梱-等1", true, "X"),
        COLOR("色1", true, "AA"),
        CATEGORY("区分1", true, "AC"),
        EC_SIDE("ＥＣ面", false, "AJ"),
        TRIMMING("ﾄﾘﾐﾝｸﾞ", false, "AM");

        private final String rawKey;
        private final boolean extractToForm;
        private final int columnIndex;

        ProductColumn(String rawKey, boolean extractToForm, String colLetter) {
            this.rawKey = rawKey;
            this.extractToForm = extractToForm;
            this.columnIndex = columnLetterToIndex(colLetter);
        }

        public String rawKey() {
            return rawKey;
        }

        public boolean extractToForm() {
            return extractToForm;
        }

        public int columnIndex() {
            return columnIndex;
        }
    }

    /** 【原反（材料）情報】行内列（行 23/24/25 共通）。 */
    public enum RawColumn {
        HINMEI("原反品名", true, "H"),
        PART_NO(null, true, "K"),
        TYPE(null, true, "N"),
        WIDTH(null, true, "Q"),
        LENGTH(null, true, "T"),
        QTY("原反数量", true, "AC"),
        GRADE("原反梱-等", true, "V"),
        COLOR("原反色", true, "Y"),
        CATEGORY("原反区分", true, "AA"),
        STORAGE("在庫場所", true, "AF"),
        INPUT_DATE("投入日", true, "AM");

        private final String rawKey;
        private final boolean extractToForm;
        private final int columnIndex;

        RawColumn(String rawKey, boolean extractToForm, String colLetter) {
            this.rawKey = rawKey;
            this.extractToForm = extractToForm;
            this.columnIndex = columnLetterToIndex(colLetter);
        }

        public String rawKey() {
            return rawKey;
        }

        public boolean extractToForm() {
            return extractToForm;
        }

        public int columnIndex() {
            return columnIndex;
        }
    }

    /** フォームへ転記する rawMap キー（{@link #extractToForm} が true の項目）。 */
    public static final List<String> FORM_EXTRACT_RAW_KEYS =
            List.of(
                    "依頼Ｎｏ",
                    "ユーザー",
                    "希望納期",
                    "納期回答",
                    "加工賃",
                    "契約Ｎｏ",
                    "品名",
                    "製品",
                    "数量1",
                    "梱-等1",
                    "色1",
                    "区分1",
                    "原反品名",
                    "品名1",
                    "原反",
                    "原反数量",
                    "原反梱-等",
                    "原反色",
                    "原反区分",
                    "在庫場所",
                    "投入日",
                    "特記事項1",
                    "特記事項2");

    /** 依頼書原本で色セルが空のときの既定値（製品・原反共通）。 */
    public static final String DEFAULT_COLOR_WHEN_BLANK = "ナチュラル";

    private RequestFormOriginalCellLayout() {}

    public static int columnLetterToIndex(String letters) {
        return JuchuSheetColumnLayout.columnLetterToIndex(letters);
    }

    public static String indexToColumnLetter(int index) {
        return JuchuSheetColumnLayout.indexToColumnLetter(index);
    }

    public static String excelRef(int rowIndex, int columnIndex) {
        return indexToColumnLetter(columnIndex) + (rowIndex + 1);
    }

    public static boolean isProductRowPopulated(
            java.util.function.BiFunction<Integer, Integer, String> cellReader, int rowIndex) {
        String hinmei = cellReader.apply(rowIndex, ProductColumn.HINMEI.columnIndex());
        String part = cellReader.apply(rowIndex, ProductColumn.PART_NO.columnIndex());
        return !hinmei.isBlank() || !part.isBlank();
    }

    /**
     * 製品行スロット（Excel 10–12 行）に入力があるか。
     * 品名・品番に加え、原本で数量・長さだけが書かれた行（2・3 行目の追記）も対象とする。
     */
    public static boolean isProductRowSlotUsed(
            java.util.function.BiFunction<Integer, Integer, String> cellReader, int rowIndex) {
        if (isProductRowPopulated(cellReader, rowIndex)) {
            return true;
        }
        String qty = cellReader.apply(rowIndex, ProductColumn.QTY.columnIndex());
        String length = cellReader.apply(rowIndex, ProductColumn.LENGTH.columnIndex());
        return !qty.isBlank() || !length.isBlank();
    }

    public static boolean isRawRowPopulated(
            java.util.function.BiFunction<Integer, Integer, String> cellReader, int rowIndex) {
        String hinmei = cellReader.apply(rowIndex, RawColumn.HINMEI.columnIndex());
        String part = cellReader.apply(rowIndex, RawColumn.PART_NO.columnIndex());
        return !hinmei.isBlank() || !part.isBlank();
    }

    public static String joinNonBlankLines(List<String> lines) {
        return String.join("\n", lines.stream().filter(s -> s != null && !s.isBlank()).map(String::strip).toList());
    }

    public static String joinNonBlankParts(List<String> parts) {
        return String.join(" ", parts.stream().filter(s -> s != null && !s.isBlank()).map(String::strip).toList());
    }
}
