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
                        return normalizeCommaDigitArtifacts(CELL_FORMAT.formatCellValue(cell));
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
     */
    public static String normalizeCommaDigitArtifacts(String s) {
        if (s == null || s.isEmpty() || !s.contains(",")) {
            return s == null ? "" : s;
        }
        String t = s.trim();
        String[] parts = t.split(",", -1);
        if (parts.length < 2) {
            return s;
        }
        for (String p : parts) {
            String q = p.trim();
            if (q.length() != 1 || !Character.isDigit(q.charAt(0))) {
                return s;
            }
        }
        StringBuilder sb = new StringBuilder(parts.length);
        for (String p : parts) {
            sb.append(p.trim());
        }
        return sb.toString();
    }
}
