package jp.co.pm.ai.desktop.ui;

import java.util.regex.Pattern;

import javafx.collections.ObservableList;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.util.Callback;

/**
 * Tabular sheet cell highlight (light green). Column naming follows planning / dispatch sheets.
 */
final class TabularCellHighlight {

    /**
     * Light green: TableCell needs both background and inner (TextFieldTableCell paints over inner).
     */
    static final String LIGHT_GREEN_STYLE =
            "-fx-background-color: #d4edd4; -fx-control-inner-background: #d4edd4;";

    /** Column title is a calendar date only, e.g. {@code 2026/04/30} (?H??????????`????t??). */
    private static final Pattern HEADER_YMD_SLASH = Pattern.compile("^\\d{4}/\\d{1,2}/\\d{1,2}$");

    private static final Pattern HEADER_YMD_DASH = Pattern.compile("^\\d{4}-\\d{1,2}-\\d{1,2}$");

    private TabularCellHighlight() {}

    static boolean isUnprocessedColumnHeader(String header) {
        if (header == null) {
            return false;
        }
        return "\u672a\u52a0\u5de5".equals(header.strip());
    }

    /**
     * Stage1 shaped-result preview: headers treated as date columns when painting green for numeric
     * values {@code > 0}.
     */
    static boolean isStage1DateColumnHeader(String header) {
        if (header == null) {
            return false;
        }
        String h = header.strip();
        if (h.isEmpty()) {
            return false;
        }
        if (h.contains("\u65e5\u4ed8")) {
            return true;
        }
        if (h.endsWith("\u7d0d\u671f")) {
            return true;
        }
        if (h.endsWith("\u6295\u5165\u65e5")) {
            return true;
        }
        if (h.endsWith("\u914d\u53f0\u65e5")) {
            return true;
        }
        if (h.endsWith("\u5b8c\u4e86\u65e5")) {
            return true;
        }
        if ("\u53d7\u6ce8\u65e5".equals(h)) {
            return true;
        }
        if ("\u30c7\u30fc\u30bf\u62bd\u51fa\u65e5".equals(h)) {
            return true;
        }
        if (h.startsWith("\u914d\u53f0\u6e08_")
                && (h.contains("\u52a0\u5de5\u958b\u59cb") || h.contains("\u52a0\u5de5\u7d42\u4e86"))) {
            return true;
        }
        if ("\u52a0\u5de5\u958b\u59cb\u65e5".equals(h)) {
            return true;
        }
        if ("\u8a08\u753b\u57fa\u6e96\u7d0d\u671f".equals(h)) {
            return true;
        }
        if (HEADER_YMD_SLASH.matcher(h).matches()) {
            return true;
        }
        if (HEADER_YMD_DASH.matcher(h).matches()) {
            return true;
        }
        return false;
    }

    /**
     * Parsed numeric value strictly greater than zero (commas / full-width comma stripped).
     */
    static boolean isStrictPositiveNumber(String raw) {
        if (raw == null) {
            return false;
        }
        String t = raw.strip();
        if (t.isEmpty()) {
            return false;
        }
        try {
            String n = t.replace("\u3000", "").replace(",", "").replace("\uff0c", "");
            double v = Double.parseDouble(n);
            return v > 0.0 && Double.isFinite(v);
        } catch (NumberFormatException ignored) {
            return false;
        }
    }

    static Callback<TableColumn<ObservableList<String>, String>, TableCell<ObservableList<String>, String>>
            stage1DateHighlightCellFactory(String columnTitle) {
        return column ->
                new TextFieldTableCell<ObservableList<String>, String>() {
                    @Override
                    public void updateItem(String item, boolean empty) {
                        super.updateItem(item, empty);
                        if (empty || item == null) {
                            setStyle("");
                            return;
                        }
                        if (isStage1DateColumnHeader(columnTitle) && isStrictPositiveNumber(item)) {
                            setStyle(LIGHT_GREEN_STYLE);
                        } else {
                            setStyle("");
                        }
                    }
                };
    }

    static Callback<TableColumn<ObservableList<String>, String>, TableCell<ObservableList<String>, String>>
            planInputUnprocessedHighlightCellFactory(String columnTitle) {
        return column ->
                new TextFieldTableCell<ObservableList<String>, String>() {
                    @Override
                    public void updateItem(String item, boolean empty) {
                        super.updateItem(item, empty);
                        if (empty || item == null) {
                            setStyle("");
                            return;
                        }
                        if (isUnprocessedColumnHeader(columnTitle) && isStrictPositiveNumber(item)) {
                            setStyle(LIGHT_GREEN_STYLE);
                        } else {
                            setStyle("");
                        }
                    }
                };
    }
}
