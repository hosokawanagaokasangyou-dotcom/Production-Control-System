package jp.co.pm.ai.desktop.io;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * Excel {@link Cell} to plain display text without pathological comma-between-every-digit formatting.
 */
public final class ExcelCellReadSupport {

    private static final DataFormatter CELL_FORMAT = new DataFormatter();

    private ExcelCellReadSupport() {}

    /**
     * Reads cell content for UI / tabular export. Numeric cells use plain digits (not Excel display format that
     * may insert commas between each digit under some locales / custom formats).
     */
    public static String cellToDisplayString(Cell cell) {
        if (cell == null) {
            return "";
        }
        try {
            CellType t = cell.getCellType();
            if (t == CellType.FORMULA) {
                t = cell.getCachedFormulaResultType();
            }
            switch (t) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return stripMidnightDateTimeSuffix(
                                normalizeCommaDigitArtifacts(
                                        CELL_FORMAT.formatCellValue(cell)));
                    }
                    double d = cell.getNumericCellValue();
                    if (!Double.isFinite(d)) {
                        return "";
                    }
                    double r = Math.rint(d);
                    if (Math.abs(d - r) < 1e-9 && Math.abs(r) <= (double) Long.MAX_VALUE) {
                        return Long.toString((long) r);
                    }
                    return Double.toString(d);
                case STRING:
                    return normalizeCommaDigitArtifacts(
                            cell.getRichStringCellValue() != null
                                    ? cell.getRichStringCellValue().getString()
                                    : "");
                case BOOLEAN:
                    return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                case BLANK:
                    return "";
                case ERROR:
                    return "";
                default:
                    return normalizeCommaDigitArtifacts(CELL_FORMAT.formatCellValue(cell));
            }
        } catch (RuntimeException ignored) {
            return normalizeCommaDigitArtifacts(CELL_FORMAT.formatCellValue(cell));
        }
    }

    /**
     * {@link DataFormatter} can emit comma between every digit (e.g. {@code 3,0,0,0}). When every comma-separated
     * segment is exactly one digit, join digits. Normal thousands like {@code 1,234} are unchanged.
     *
     * <p>Excel の表示形式が小数を付けると {@code 8,0,0,0.00} のように最後のセグメントだけが {@code 0.00} になり、従来の
     * 「全区画が1桁」判定に引っかからない。この場合は ASCII の {@code '.'} より後ろを小数部として除外してから整数部だけを結合する。
     */
    public static String normalizeCommaDigitArtifacts(String s) {
        if (s == null || s.isEmpty() || !s.contains(",")) {
            return s == null ? "" : s;
        }
        String t = s.trim();
        int dot = indexOfTrailingAsciiDecimalSeparator(t);
        String intPart = dot >= 0 ? t.substring(0, dot) : t;
        String fracPart = dot >= 0 ? t.substring(dot) : "";

        String collapsed = tryCollapseCommaSingleDigitInteger(intPart.trim());
        if (collapsed != null) {
            return collapsed + fracPart;
        }
        return s;
    }

    /**
     * Returns index of {@code '.'} that starts the fractional part (only ASCII digits after it), or {@code -1}.
     */
    private static int indexOfTrailingAsciiDecimalSeparator(String t) {
        for (int i = t.length() - 1; i >= 0; i--) {
            if (t.charAt(i) != '.') {
                continue;
            }
            String after = t.substring(i + 1);
            if (!after.isEmpty() && allAsciiDigits(after)) {
                return i;
            }
        }
        return -1;
    }

    private static boolean allAsciiDigits(String after) {
        for (int i = 0; i < after.length(); i++) {
            char ch = after.charAt(i);
            if (ch < '0' || ch > '9') {
                return false;
            }
        }
        return true;
    }

    /**
     * @return collapsed digits-only integer string, or {@code null} if the pattern is not per-digit commas
     */
    private static String tryCollapseCommaSingleDigitInteger(String intPart) {
        if (intPart.isEmpty() || !intPart.contains(",")) {
            return null;
        }
        boolean neg = intPart.startsWith("-");
        String u = neg ? intPart.substring(1).trim() : intPart;
        if (u.isEmpty() || !u.contains(",")) {
            return null;
        }
        String[] parts = u.split(",", -1);
        if (parts.length < 2) {
            return null;
        }
        for (String p : parts) {
            String q = p.trim();
            if (q.length() != 1 || !Character.isDigit(q.charAt(0))) {
                return null;
            }
        }
        StringBuilder sb = new StringBuilder(parts.length);
        for (String p : parts) {
            sb.append(p.trim());
        }
        String digits = sb.toString();
        return neg ? "-" + digits : digits;
    }

    /**
     * Excel {@link DataFormatter} が日付セルに付ける {@code yyyy-MM-dd 00:00:00} 等の深夜時刻を除き、日付部分だけ返す。
     * 非深夜の時刻はそのまま残す。
     */
    public static String stripMidnightDateTimeSuffix(String s) {
        if (s == null || s.isEmpty()) {
            return s == null ? "" : s;
        }
        String t = s.strip();
        if (t.length() < 16) {
            return s;
        }
        char sep = t.charAt(10);
        if (sep != ' ' && sep != 'T') {
            return s;
        }
        if (!looksLikeYmdDatePrefix(t)) {
            return s;
        }
        String timePart = t.substring(11).strip();
        if (!isMidnightTimePart(timePart)) {
            return s;
        }
        return t.substring(0, 10);
    }

    private static boolean looksLikeYmdDatePrefix(String t) {
        if (t.length() < 10) {
            return false;
        }
        char sep1 = t.charAt(4);
        char sep2 = t.charAt(7);
        if (sep1 != sep2 || (sep1 != '-' && sep1 != '/' && sep1 != '.')) {
            return false;
        }
        for (int i = 0; i < 10; i++) {
            if (i == 4 || i == 7) {
                continue;
            }
            if (!Character.isDigit(t.charAt(i))) {
                return false;
            }
        }
        return true;
    }

    private static boolean isMidnightTimePart(String timePart) {
        if (timePart.isEmpty()) {
            return false;
        }
        if ("00:00:00".equals(timePart) || "00:00".equals(timePart) || "0:00:00".equals(timePart)) {
            return true;
        }
        return timePart.startsWith("00:00:00.");
    }
}
