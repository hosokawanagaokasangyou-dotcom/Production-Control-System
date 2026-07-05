package jp.co.pm.ai.desktop.reconciliation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.text.Normalizer;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.OptionalInt;

/**
 * 受注ファイル「受注ﾌｧｲﾙ」シートの列位置と見出し名の正本。
 * 転記・読込は列 index を優先し、見出しは検証用（別名許容）。
 * 見出し行は {@link JuchuHeaderAliasRegistry} でファイル別に設定（既定 3 行目）。
 */
public final class JuchuSheetColumnLayout {

    /** 見出し行の既定（0-based、3 行目）。 */
    public static final int HEADER_ROW_INDEX = JuchuHeaderAliasRegistry.DEFAULT_HEADER_ROW_ONE_BASED - 1;

    /** 行3見出し走査: この列数まで（BR 付近の見出し取りこぼし防止）。 */
    public static final int HEADER_PICK_MAX_SCAN_COLUMNS = 512;

    /** 行3見出し走査: 連続空セルがこの数に達したら以降の列は採用候補に含めない。 */
    public static final int HEADER_PICK_EMPTY_RUN_STOP = 10;

    public enum Col {
        IRAI_NO("A", "依頼No", "依頼Ｎｏ", "依頼NO"),
        NYURYOKU_KBN("B", "入力区分"),
        KAKO_KBN("C", "加工区分"),
        NYURYOKU_TANTO("D", "入力担当", "入力者"),
        NYURYOKU_BI("E", "入力日"),
        UKETSUKE_NO("F", "受付Ｎｏ", "受付No", "受付NO"),
        HINMEI("G", "品名"),
        SEIHIN("H", "製品"),
        KON_TO_1("I", "梱-等1", "梱－等1"),
        IRO_1("J", "色1"),
        KUBUN_1("K", "区分1"),
        EDABAN("L", "枝番"),
        SURYO_1("M", "数量1"),
        EC_MEN("N", "EC面", "ＥＣ面"),
        TRIMMING("O", "トリミング", "ﾄﾘﾐﾝｸﾞ"),
        WARISU("P", "割数", "加工回数（加工換算数に利用）"),
        HINMEI_1("Q", "品名1"),
        GENPAN("R", "原反"),
        KON_TO("S", "梱-等", "梱－等"),
        IRO("T", "色"),
        KUBUN("U", "区分"),
        SURYO("V", "数量"),
        ZAIKO_BASHO("W", "在庫場所"),
        TONYU_BASHO("X", "投入場所"),
        TONYU_BI("Y", "投入日"),
        KAKO_NAIYO("Z", "加工内容"),
        TOKKI_1("AA", "特記事項1"),
        TOKKI_2("AB", "特記事項2"),
        TOKKI_3("AC", "特記事項3"),
        YOTO("AD", "用途"),
        USER("AE", "ユーザー"),
        KIBO_NOKI("AF", "希望納期"),
        CHOSEI_NOKI("AG", "調整納期"),
        KAKOCHIN("AH", "加工賃"),
        KEIYAKU_NO("AI", "契約Ｎｏ", "契約No", "契約NO"),
        GENPAN_ROLL_SU("AJ", "原反ロール数"),
        MASTER_BASE_SHOHIN_PRODUCT("AP", "masterBase商品(製品)"),
        MASTER_BASE_SHOHIN_RAW("AQ", "masterBase商品(原反)");

        private final String columnLetter;
        private final int columnIndex;
        private final String primaryHeader;
        private final List<String> aliases;

        Col(String columnLetter, String primaryHeader, String... aliases) {
            this.columnLetter = columnLetter;
            this.columnIndex = columnLetterToIndex(columnLetter);
            this.primaryHeader = primaryHeader;
            List<String> all = new ArrayList<>();
            all.add(primaryHeader);
            if (aliases != null) {
                all.addAll(Arrays.asList(aliases));
            }
            this.aliases = List.copyOf(all);
        }

        public String columnLetter() {
            return columnLetter;
        }

        public int columnIndex() {
            return columnIndex;
        }

        public String primaryHeader() {
            return primaryHeader;
        }

        public List<String> aliases() {
            return aliases;
        }

        /** {@code dbValues} 等で使う内部キー。 */
        public String dbKey() {
            return switch (this) {
                case HINMEI_1 -> "品名1";
                case KON_TO -> "原反梱-等";
                case IRO -> "原反色";
                case KUBUN -> "原反区分";
                case SURYO -> "原反数量";
                case EC_MEN -> "ＥＣ面";
                case TRIMMING -> "ﾄﾘﾐﾝｸﾞ";
                case MASTER_BASE_SHOHIN_PRODUCT -> "masterBase商品(製品)";
                case MASTER_BASE_SHOHIN_RAW -> "masterBase商品(原反)";
                default -> primaryHeader;
            };
        }

        /** 依頼書フォーム上のブロックと項目名（ウィザード表示用）。 */
        public String formItemDescription() {
            return switch (this) {
                case IRAI_NO -> "【受注転記】依頼Ｎｏ";
                case NYURYOKU_KBN -> "【作業指示】入力区分";
                case KAKO_KBN -> "【作業指示】加工区分";
                case NYURYOKU_TANTO -> "【作業指示】入力担当";
                case NYURYOKU_BI -> "【基本情報】入力日";
                case UKETSUKE_NO -> "【基本情報】受付Ｎｏ";
                case HINMEI -> "【製品（仕上がり）】品名";
                case SEIHIN -> "【製品（仕上がり）】製品名";
                case KON_TO_1 -> "【製品（仕上がり）】梱-等";
                case IRO_1 -> "【製品（仕上がり）】色";
                case KUBUN_1 -> "【製品（仕上がり）】区分";
                case EDABAN -> "【製品（仕上がり）】枝番";
                case SURYO_1 -> "【製品（仕上がり）】数量";
                case EC_MEN -> "【製品（仕上がり）】EC面";
                case TRIMMING -> "【製品（仕上がり）】トリミング";
                case WARISU -> "【原反（材料）】割数";
                case HINMEI_1 -> "【原反（材料）】原反名";
                case GENPAN -> "【原反（材料）】原反（仕様）";
                case KON_TO -> "【原反（材料）】梱-等";
                case IRO -> "【原反（材料）】色";
                case KUBUN -> "【原反（材料）】区分";
                case SURYO -> "【原反（材料）】数量";
                case ZAIKO_BASHO -> "【原反（材料）】在庫場所";
                case TONYU_BASHO -> "【原反（材料）】投入場所";
                case TONYU_BI -> "【原反（材料）】投入日";
                case KAKO_NAIYO -> "【基本情報】加工内容";
                case TOKKI_1 -> "【作業指示】特記事項1";
                case TOKKI_2 -> "【作業指示】特記事項2";
                case TOKKI_3 -> "【作業指示】特記事項3";
                case YOTO -> "【作業指示】用途";
                case USER -> "【基本情報】ユーザー";
                case KIBO_NOKI -> "【基本情報】希望納期";
                case CHOSEI_NOKI -> "【基本情報】調整納期";
                case KAKOCHIN -> "【基本情報】加工賃";
                case KEIYAKU_NO -> "【製品（仕上がり）】契約Ｎｏ";
                case GENPAN_ROLL_SU -> "【原反（材料）】原反ロール数";
                case MASTER_BASE_SHOHIN_PRODUCT -> "【製品（仕上がり）】商品（masterBase）";
                case MASTER_BASE_SHOHIN_RAW -> "【原反（材料）】商品（masterBase）";
            };
        }

        public boolean matchesHeader(String actual) {
            return matchesHeader(actual, List.of());
        }

        public boolean matchesHeader(String actual, List<String> extraAliases) {
            if (actual == null || actual.isBlank()) {
                return false;
            }
            String normActual = normalizeHeader(actual);
            for (String alias : aliases) {
                if (normalizeHeader(alias).equals(normActual)) {
                    return true;
                }
            }
            if (extraAliases != null) {
                for (String extra : extraAliases) {
                    if (extra != null
                            && !extra.isBlank()
                            && normalizeHeader(extra).equals(normActual)) {
                        return true;
                    }
                }
            }
            return false;
        }
    }

    private JuchuSheetColumnLayout() {}

    /** 転記・読込対象の列（定義順）。 */
    public static List<Col> transferColumns() {
        return List.of(Col.values());
    }

    public static int columnLetterToIndex(String letters) {
        Objects.requireNonNull(letters, "letters");
        String upper = letters.trim().toUpperCase(Locale.ROOT);
        if (upper.isEmpty()) {
            throw new IllegalArgumentException("empty column letters");
        }
        int index = 0;
        for (int i = 0; i < upper.length(); i++) {
            char ch = upper.charAt(i);
            if (ch < 'A' || ch > 'Z') {
                throw new IllegalArgumentException("invalid column letter: " + letters);
            }
            index = index * 26 + (ch - 'A' + 1);
        }
        return index - 1;
    }

    public static String indexToColumnLetter(int index) {
        if (index < 0) {
            throw new IllegalArgumentException("negative index: " + index);
        }
        int n = index + 1;
        StringBuilder sb = new StringBuilder();
        while (n > 0) {
            n--;
            sb.insert(0, (char) ('A' + (n % 26)));
            n /= 26;
        }
        return sb.toString();
    }

    public static String normalizeHeader(String header) {
        if (header == null) {
            return "";
        }
        String text = Normalizer.normalize(header.trim(), Normalizer.Form.NFKC);
        return text.replace(" ", "").replace("　", "");
    }

    public static Optional<Col> findByDbKey(String dbKey) {
        if (dbKey == null || dbKey.isBlank()) {
            return Optional.empty();
        }
        String norm = normalizeHeader(dbKey);
        for (Col col : Col.values()) {
            if (normalizeHeader(col.dbKey()).equals(norm)) {
                return Optional.of(col);
            }
        }
        if ("原反品名".equals(dbKey)) {
            return Optional.of(Col.HINMEI_1);
        }
        return Optional.empty();
    }

    /**
     * 行3見出しと定義列位置を照合。不一致は警告メッセージ文字列のリスト。
     */
    public static List<String> validateHeaders(Row headerRow) {
        return validateHeaders(headerRow, null, null);
    }

    /**
     * ファイル別別名レジストリを考慮した見出し検証。
     */
    public static List<String> validateHeaders(
            Row headerRow, JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        int headerRowOneBased = resolveHeaderRowOneBased(registry, juchuFileAbsolutePath);
        if (headerRow == null) {
            return List.of("受注ﾌｧｲﾙ: 見出し行（行" + headerRowOneBased + "）が存在しません。");
        }
        return collectHeaderMismatches(headerRow, registry, juchuFileAbsolutePath).stream()
                .map(JuchuHeaderMismatch::summaryLine)
                .toList();
    }

    public static int resolveHeaderRowOneBased(
            JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        if (registry != null && juchuFileAbsolutePath != null && !juchuFileAbsolutePath.isBlank()) {
            return registry.headerRowOneBasedFor(juchuFileAbsolutePath);
        }
        return JuchuHeaderAliasRegistry.DEFAULT_HEADER_ROW_ONE_BASED;
    }

    public static int resolveHeaderRowIndex(
            JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        return Math.max(0, resolveHeaderRowOneBased(registry, juchuFileAbsolutePath) - 1);
    }

    public static int resolveFirstDataRowIndex(
            JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        return resolveHeaderRowIndex(registry, juchuFileAbsolutePath) + 1;
    }

    /** 列ごとの不一致（ウィザード用）。採用列設定があればその列の見出しを検証する。 */
    public static List<JuchuHeaderMismatch> collectHeaderMismatches(
            Row headerRow, JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        List<JuchuHeaderMismatch> out = new ArrayList<>();
        if (headerRow == null) {
            return out;
        }
        for (Col col : Col.values()) {
            if (registry != null
                    && juchuFileAbsolutePath != null
                    && registry.isExcludedFromTransfer(juchuFileAbsolutePath, col)) {
                continue;
            }
            String actual =
                    readHeaderCell(
                            headerRow,
                            resolveTransferColumnIndex(col, registry, juchuFileAbsolutePath));
            String expected =
                    registry == null
                            ? col.primaryHeader()
                            : registry.expectedHeaderFor(juchuFileAbsolutePath, col);
            if (!headerMatches(col, actual, registry, juchuFileAbsolutePath)) {
                out.add(
                        new JuchuHeaderMismatch(
                                col,
                                expected,
                                actual,
                                actual.isBlank(),
                                resolveTransferColumnLetter(col, registry, juchuFileAbsolutePath)));
            }
        }
        return out;
    }

    /** フォーム転記項目の一覧（ウィザード用）。見出しは採用列から読む。 */
    public static List<JuchuHeaderMismatch> collectAllKnownColumns(
            Row headerRow, JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        List<JuchuHeaderMismatch> out = new ArrayList<>();
        if (headerRow == null) {
            return out;
        }
        for (Col col : Col.values()) {
            String actual =
                    readHeaderCell(
                            headerRow,
                            resolveTransferColumnIndex(col, registry, juchuFileAbsolutePath));
            String expected =
                    registry == null
                            ? col.primaryHeader()
                            : registry.expectedHeaderFor(juchuFileAbsolutePath, col);
            out.add(
                    new JuchuHeaderMismatch(
                            col,
                            expected,
                            actual,
                            actual.isBlank(),
                            resolveTransferColumnLetter(col, registry, juchuFileAbsolutePath)));
        }
        return out;
    }

    public static boolean isKnownColumnIndex(int columnIndex) {
        for (Col col : Col.values()) {
            if (col.columnIndex() == columnIndex) {
                return true;
            }
        }
        return false;
    }

    /** 既知列位置以外の行3見出し（転記定義外）。 */
    public static List<JuchuUnknownExcelColumn> collectUnknownExcelColumns(
            Row headerRow, JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        List<JuchuUnknownExcelColumn> out = new ArrayList<>();
        if (headerRow == null) {
            return out;
        }
        int scanExclusiveEnd = resolveHeaderPickScanExclusiveEnd(headerRow);
        for (int c = 0; c < scanExclusiveEnd; c++) {
            if (isKnownColumnIndex(c)) {
                continue;
            }
            String text = readHeaderCell(headerRow, c);
            if (text.isBlank()) {
                continue;
            }
            String letter = indexToColumnLetter(c);
            boolean ignored =
                    registry != null
                            && registry.isUnknownColumnIgnored(juchuFileAbsolutePath, letter);
            out.add(new JuchuUnknownExcelColumn(letter, c, text, ignored));
        }
        return out;
    }

    static boolean headerMatches(
            Col col,
            String actual,
            JuchuHeaderAliasRegistry registry,
            String juchuFileAbsolutePath) {
        List<String> extras =
                registry == null
                        ? List.of()
                        : registry.extraAliasesFor(juchuFileAbsolutePath, col);
        if (registry != null) {
            Optional<String> override =
                    registry.expectedOverrideFor(juchuFileAbsolutePath, col);
            if (override.isPresent()) {
                if (actual.isBlank()) {
                    return true;
                }
                if (normalizeHeader(actual).equals(normalizeHeader(override.get()))) {
                    return true;
                }
            }
        }
        return col.matchesHeader(actual, extras);
    }

    /** 行3の非空見出し一覧（ウィザードの Excel 見出し選択用）。 */
    public static List<ExcelHeaderPick> readExcelHeaderPicks(Row headerRow) {
        List<ExcelHeaderPick> picks = new ArrayList<>();
        if (headerRow == null) {
            return picks;
        }
        int scanExclusiveEnd = resolveHeaderPickScanExclusiveEnd(headerRow);
        for (int c = 0; c < scanExclusiveEnd; c++) {
            String text = readHeaderCell(headerRow, c);
            if (text.isBlank()) {
                continue;
            }
            picks.add(new ExcelHeaderPick(indexToColumnLetter(c), c, text));
        }
        return picks;
    }

    /**
     * 行3を左から走査し、非空見出しの採用候補に含める列の終端（exclusive）。
     * {@link #HEADER_PICK_EMPTY_RUN_STOP} 個連続で空の列が出た時点より右は除外する。
     */
    public static int resolveHeaderPickScanExclusiveEnd(Row headerRow) {
        if (headerRow == null) {
            return 0;
        }
        int lastCellNum = Math.max(headerRow.getLastCellNum(), 0);
        int layoutMax =
                Arrays.stream(Col.values()).mapToInt(Col::columnIndex).max().orElse(0) + 1;
        int provisional =
                Math.min(Math.max(lastCellNum, layoutMax), HEADER_PICK_MAX_SCAN_COLUMNS);

        int consecutiveEmpty = 0;
        boolean seenHeader = false;
        for (int c = 0; c < provisional; c++) {
            if (readHeaderCell(headerRow, c).isBlank()) {
                if (seenHeader) {
                    consecutiveEmpty++;
                    if (consecutiveEmpty >= HEADER_PICK_EMPTY_RUN_STOP) {
                        return c - HEADER_PICK_EMPTY_RUN_STOP + 1;
                    }
                }
            } else {
                seenHeader = true;
                consecutiveEmpty = 0;
            }
        }
        return provisional;
    }

    public record ExcelHeaderPick(String columnLetter, int columnIndex, String headerText) {
        public String displayLabel() {
            return columnLetter + "列: " + headerText;
        }
    }

    /** 行3の見出しセルへ期待見出しを書き込む（0-based 行 index は {@link #HEADER_ROW_INDEX}）。 */
    public static void writeHeaderCell(Row headerRow, Col col, String headerText) {
        if (headerRow == null || col == null) {
            return;
        }
        Cell cell = headerRow.getCell(col.columnIndex());
        if (cell == null) {
            cell = headerRow.createCell(col.columnIndex());
        }
        cell.setCellValue(headerText != null ? headerText : "");
    }

    /**
     * レイアウト定義に基づき db キー → 値 のマップを構築（読込用）。
     */
    public static Map<String, String> readDbValuesFromRow(Row dataRow) {
        return readDbValuesFromRow(dataRow, null, null);
    }

    public static Map<String, String> readDbValuesFromRow(
            Row dataRow, JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        Map<String, String> vals = new LinkedHashMap<>();
        if (dataRow == null) {
            return vals;
        }
        for (Col col : Col.values()) {
            if (registry != null
                    && juchuFileAbsolutePath != null
                    && registry.isExcludedFromTransfer(juchuFileAbsolutePath, col)) {
                continue;
            }
            String value =
                    readDataCell(
                            dataRow,
                            resolveTransferColumnIndex(col, registry, juchuFileAbsolutePath));
            vals.put(col.dbKey(), value);
            if (col == Col.HINMEI_1) {
                vals.put("原反品名", value);
            }
        }
        return vals;
    }

    /**
     * フォーム項目の転記・読込に使う列 index。
     * 列定義ウィザードで採用した {@code XX列: 見出し} があればその列、なければ {@link Col} 既定。
     */
    public static int resolveTransferColumnIndex(
            Col col, JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        if (col == null) {
            return 0;
        }
        if (registry != null && juchuFileAbsolutePath != null && !juchuFileAbsolutePath.isBlank()) {
            OptionalInt fromPick =
                    columnIndexFromPickDisplayLabel(
                            registry
                                    .expectedPickLabelFor(juchuFileAbsolutePath, col)
                                    .orElse(null));
            if (fromPick.isPresent()) {
                return fromPick.getAsInt();
            }
        }
        return col.columnIndex();
    }

    public static String resolveTransferColumnLetter(
            Col col, JuchuHeaderAliasRegistry registry, String juchuFileAbsolutePath) {
        return indexToColumnLetter(resolveTransferColumnIndex(col, registry, juchuFileAbsolutePath));
    }

    /** {@code BU列: 商品(製品)} 形式から列 index を得る。 */
    public static OptionalInt columnIndexFromPickDisplayLabel(String pickLabel) {
        if (pickLabel == null || pickLabel.isBlank()) {
            return OptionalInt.empty();
        }
        int colon = pickLabel.indexOf("列:");
        if (colon <= 0) {
            return OptionalInt.empty();
        }
        try {
            return OptionalInt.of(columnLetterToIndex(pickLabel.substring(0, colon).strip()));
        } catch (Exception ex) {
            return OptionalInt.empty();
        }
    }

    public static String readHeaderCell(Row headerRow, int columnIndex) {
        if (headerRow == null) {
            return "";
        }
        Cell cell = headerRow.getCell(columnIndex);
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try {
                    yield cell.getStringCellValue().trim();
                } catch (Exception ex) {
                    yield String.valueOf(cell.getNumericCellValue());
                }
            }
            default -> "";
        };
    }

    public static String readDataCell(Row dataRow, int columnIndex) {
        if (dataRow == null) {
            return "";
        }
        Cell cell = dataRow.getCell(columnIndex);
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    java.util.Date d = cell.getDateCellValue();
                    yield new java.text.SimpleDateFormat("yyyy-MM-dd").format(d);
                }
                double n = cell.getNumericCellValue();
                if (org.apache.poi.ss.usermodel.DateUtil.isValidExcelDate(n)
                        && n >= 25000.0
                        && n == Math.rint(n)) {
                    java.util.Date d = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(n);
                    yield new java.text.SimpleDateFormat("yyyy-MM-dd").format(d);
                }
                if (n == Math.rint(n)) {
                    yield String.valueOf((long) n);
                }
                yield String.valueOf(n);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try {
                    yield cell.getStringCellValue();
                } catch (Exception ex) {
                    try {
                        if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                            java.util.Date d = cell.getDateCellValue();
                            yield new java.text.SimpleDateFormat("yyyy-MM-dd").format(d);
                        }
                        double n = cell.getNumericCellValue();
                        if (org.apache.poi.ss.usermodel.DateUtil.isValidExcelDate(n)
                                && n >= 25000.0
                                && n == Math.rint(n)) {
                            java.util.Date d = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(n);
                            yield new java.text.SimpleDateFormat("yyyy-MM-dd").format(d);
                        }
                        if (n == Math.rint(n)) {
                            yield String.valueOf((long) n);
                        }
                        yield String.valueOf(n);
                    } catch (Exception ex2) {
                        yield "";
                    }
                }
            }
            default -> "";
        };
    }

    /** 品番-タイプ-幅X長 形式の製品名/原反名を生成。 */
    public static String buildSpecName(String part, String type, String width, String length) {
        return part.trim()
                + "-"
                + type.trim()
                + "-"
                + width.trim()
                + "X"
                + length.trim();
    }

    /**
     * 製品/原反 spec 文字列を品番・タイプ・幅・長さに分解。
     * {@code 20010-H600-1180X250} または {@code 20010-H600-1180-250} に対応。
     */
    public static String[] parseSpecName(String spec) {
        if (spec == null || spec.isBlank()) {
            return new String[] {"", "", "", ""};
        }
        String text = spec.trim();
        String[] parts = text.split("-");
        if (parts.length >= 3) {
            String part = parts[0];
            String type = parts[1];
            String dims = parts[2];
            for (int i = 3; i < parts.length; i++) {
                dims = dims + "-" + parts[i];
            }
            String[] wL = dims.split("X", 2);
            if (wL.length >= 2) {
                return new String[] {part, type, wL[0], wL[1]};
            }
            if (parts.length >= 4) {
                return new String[] {part, type, parts[2], parts[3]};
            }
            return new String[] {part, type, dims, ""};
        }
        return new String[] {text, "", "", ""};
    }

    /**
     * 原反ロール数 = 数量 ÷ 長さ(m) の整数部分（小数点以下切り捨て）。
     * 数量・長さのいずれかが空、または長さ≦0 のときは空。
     */
    public static java.util.OptionalInt computeRawRollCountFromQtyAndLength(String qtyText, String lengthText) {
        if (qtyText == null || qtyText.isBlank() || lengthText == null || lengthText.isBlank()) {
            return java.util.OptionalInt.empty();
        }
        double qty = parseLooseNumeric(qtyText);
        double length = parseLooseNumeric(lengthText);
        if (length <= 0.0) {
            return java.util.OptionalInt.empty();
        }
        return java.util.OptionalInt.of((int) Math.floor(qty / length));
    }

    private static double parseLooseNumeric(String text) {
        if (text == null || text.isBlank()) {
            return 0.0;
        }
        String trimmed = text.strip();
        String withoutGrouping = trimmed.replace(",", "").replace("，", "");
        java.util.regex.Matcher m =
                java.util.regex.Pattern.compile("[-+]?\\d*\\.\\d+|\\d+").matcher(withoutGrouping);
        if (m.find()) {
            return Double.parseDouble(m.group());
        }
        return 0.0;
    }
}
