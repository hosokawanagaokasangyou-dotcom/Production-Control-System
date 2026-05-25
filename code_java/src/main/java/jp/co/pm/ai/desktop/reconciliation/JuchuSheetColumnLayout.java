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

/**
 * 受注ファイル「受注ﾌｧｲﾙ」シート行3の列位置と見出し名の正本。
 * 転記・読込は列 index を優先し、見出しは検証用（別名許容）。
 */
public final class JuchuSheetColumnLayout {

    public static final int HEADER_ROW_INDEX = 2;

    public enum Col {
        IRAI_NO("A", "依頼No", "依頼Ｎｏ", "依頼NO"),
        NYURYOKU_KBN("B", "入力区分"),
        KAKO_KBN("C", "加工区分"),
        NYURYOKU_TANTO("D", "入力担当"),
        NYURYOKU_BI("E", "入力日"),
        HINMEI("G", "品名"),
        SEIHIN("H", "製品"),
        KON_TO_1("I", "梱-等1", "梱－等1"),
        IRO_1("J", "色1"),
        KUBUN_1("K", "区分1"),
        EDABAN("L", "枝番"),
        SURYO_1("M", "数量1"),
        EC_MEN("N", "EC面", "ＥＣ面"),
        TRIMMING("O", "トリミング", "ﾄﾘﾐﾝｸﾞ"),
        WARISU("P", "割数"),
        HINMEI_1("Q", "品名1"),
        GENPAN("R", "原反"),
        KON_TO("S", "梱-等", "梱－等"),
        IRO("T", "色"),
        KUBUN("U", "区分"),
        SURYO("V", "数量"),
        ZAIKO_BASHO("W", "在庫場所"),
        TONYU_BASHO("X", "投入場所"),
        KAKO_NAIYO("Z", "加工内容"),
        TOKKI_1("AA", "特記事項1"),
        TOKKI_2("AB", "特記事項2"),
        TOKKI_3("AC", "特記事項3"),
        YOTO("AD", "用途"),
        USER("AE", "ユーザー"),
        KIBO_NOKI("AF", "希望納期"),
        CHOSEI_NOKI("AG", "調整納期"),
        KAKOCHIN("AH", "加工賃"),
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

        public boolean matchesHeader(String actual) {
            if (actual == null || actual.isBlank()) {
                return false;
            }
            String normActual = normalizeHeader(actual);
            for (String alias : aliases) {
                if (normalizeHeader(alias).equals(normActual)) {
                    return true;
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
        List<String> warnings = new ArrayList<>();
        if (headerRow == null) {
            warnings.add("受注ﾌｧｲﾙ: 見出し行（行3）が存在しません。");
            return warnings;
        }
        for (Col col : Col.values()) {
            String actual = readHeaderCell(headerRow, col.columnIndex());
            if (actual.isBlank()) {
                warnings.add(
                        col.columnLetter()
                                + "列: 見出しが空です（期待: "
                                + col.primaryHeader()
                                + "）");
                continue;
            }
            if (!col.matchesHeader(actual)) {
                warnings.add(
                        col.columnLetter()
                                + "列: 期待「"
                                + col.primaryHeader()
                                + "」だが実際「"
                                + actual
                                + "」");
            }
        }
        return warnings;
    }

    /**
     * レイアウト定義に基づき db キー → 値 のマップを構築（読込用）。
     */
    public static Map<String, String> readDbValuesFromRow(Row dataRow) {
        Map<String, String> vals = new LinkedHashMap<>();
        if (dataRow == null) {
            return vals;
        }
        for (Col col : Col.values()) {
            String value = readDataCell(dataRow, col.columnIndex());
            vals.put(col.dbKey(), value);
            if (col == Col.HINMEI_1) {
                vals.put("原反品名", value);
            }
        }
        return vals;
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
}
