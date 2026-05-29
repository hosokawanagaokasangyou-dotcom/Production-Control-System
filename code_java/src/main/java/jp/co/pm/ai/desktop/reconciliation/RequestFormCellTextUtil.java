package jp.co.pm.ai.desktop.reconciliation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/** 依頼書セル文字列の共通整形ユーティリティ。 */
final class RequestFormCellTextUtil {

    private RequestFormCellTextUtil() {}

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
}
