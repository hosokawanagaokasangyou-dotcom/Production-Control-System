package jp.co.pm.ai.desktop.reconciliation;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/** 依頼書セル文字列の共通整形ユーティリティ。 */
final class RequestFormCellTextUtil {

    /** 日付シリアルとみなす下限（おおよそ 1968 年以降）。 */
    private static final double MIN_PLAUSIBLE_DATE_SERIAL = 25000.0;

    private static final SimpleDateFormat PREVIEW_DATE_FORMAT =
            new SimpleDateFormat("yyyy/MM/dd", Locale.JAPAN);

    private RequestFormCellTextUtil() {}

    /**
     * 依頼書プレビュー・原本抽出向けにセル表示文字列を返す。
     * 日付書式のない Excel 日付シリアル（例: {@code 46171.0}）は {@code yyyy/MM/dd} に変換する。
     */
    static String formatCellDisplayText(
            Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) {
            return "";
        }
        String formatted;
        try {
            formatted =
                    evaluator != null
                            ? formatter.formatCellValue(cell, evaluator)
                            : formatter.formatCellValue(cell);
        } catch (RuntimeException ex) {
            formatted = formatter.formatCellValue(cell);
        }
        formatted = stripFormatLiteralQuotes(cell, formatted);
        return coerceUnformattedExcelDateDisplay(cell, formatted);
    }

    static String formatCellDisplayText(Cell cell, DataFormatter formatter) {
        return formatCellDisplayText(cell, formatter, null);
    }

    /**
     * Excel ユーザー定義書式のリテラル引用符（例: {@code m"月"d"日"}）は、POI の
     * {@code DataFormatter} が日付書式では除去せず {@code 6"月"3"日"} のように残してしまう。
     * 整形後文字列に残った書式リテラルのダブルクォートを取り除く。
     * 書式文字列にダブルクォートを含むセルに限定するため、文字列セル内の正当な引用符は対象外。
     */
    static String stripFormatLiteralQuotes(Cell cell, String formatted) {
        if (formatted == null || cell == null || formatted.indexOf('"') < 0) {
            return formatted;
        }
        CellStyle style = cell.getCellStyle();
        String fmt = style != null ? style.getDataFormatString() : null;
        if (fmt != null && fmt.indexOf('"') >= 0) {
            return formatted.replace("\"", "");
        }
        return formatted;
    }

    private static String coerceUnformattedExcelDateDisplay(Cell cell, String formatted) {
        if (cell == null || formatted == null || formatted.isBlank()) {
            return formatted != null ? formatted : "";
        }
        if (isDateFormattedCell(cell)) {
            return formatted;
        }
        Double serial = tryNumericCellValue(cell);
        if (serial == null
                || !DateUtil.isValidExcelDate(serial)
                || serial < MIN_PLAUSIBLE_DATE_SERIAL
                || !looksLikeRawExcelDateSerial(formatted.strip(), serial)) {
            return formatted;
        }
        Date date = DateUtil.getJavaDate(serial);
        synchronized (PREVIEW_DATE_FORMAT) {
            return PREVIEW_DATE_FORMAT.format(date);
        }
    }

    private static boolean isDateFormattedCell(Cell cell) {
        if (cell == null) {
            return false;
        }
        try {
            CellType type = cell.getCellType();
            if (type == CellType.FORMULA) {
                type = cell.getCachedFormulaResultType();
            }
            if (type != CellType.NUMERIC) {
                return false;
            }
            return DateUtil.isCellDateFormatted(cell);
        } catch (RuntimeException ex) {
            return false;
        }
    }

    private static boolean looksLikeRawExcelDateSerial(String formatted, double serial) {
        if (!formatted.matches("\\d{5,6}(\\.0+)?")) {
            return false;
        }
        return serial == Math.rint(serial);
    }

    private static Double tryNumericCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        try {
            CellType type = cell.getCellType();
            if (type == CellType.FORMULA) {
                type = cell.getCachedFormulaResultType();
            }
            if (type == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            }
        } catch (RuntimeException ignored) {
            // fall through
        }
        return null;
    }
}
