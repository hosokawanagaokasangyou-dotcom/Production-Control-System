package jp.co.pm.ai.desktop.ui;

import java.util.regex.Pattern;
import java.util.function.IntSupplier;

import javafx.collections.ObservableList;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.util.Callback;

import org.controlsfx.control.spreadsheet.SpreadsheetCell;

/**
 * Tabular sheet cell highlight (light green). Column naming follows planning / dispatch sheets.
 */
public final class TabularCellHighlight {

    /**
     * Light green: TableCell needs both background and inner (TextFieldTableCell paints over inner).
     */
    /** Exported for {@link org.controlsfx.control.spreadsheet.SpreadsheetCell} styling. */
    public static final String LIGHT_GREEN_STYLE =
            "-fx-background-color: #d4edd4; -fx-control-inner-background: #d4edd4;";

    /** Column title is a calendar date only, e.g. {@code 2026/04/30} （{@code yyyy/MM/dd} 形式の日付列のみ）。 */
    private static final Pattern HEADER_YMD_SLASH = Pattern.compile("^\\d{4}/\\d{1,2}/\\d{1,2}$");

    private static final Pattern HEADER_YMD_DASH = Pattern.compile("^\\d{4}-\\d{1,2}-\\d{1,2}$");

    private TabularCellHighlight() {}

    /** ControlsFX spreadsheet cell highlight for plan-input rules. */
    public static void applyPlanInputSpreadsheetHighlight(SpreadsheetCell cell, String columnTitle, String text) {
        if (cell == null) {
            return;
        }
        if (isUnprocessedColumnHeader(columnTitle) && isStrictPositiveNumber(text)) {
            cell.setStyle(LIGHT_GREEN_STYLE);
        } else {
            cell.setStyle("");
        }
    }

    /** ControlsFX spreadsheet cell highlight for Stage1 preview date columns. */
    public static void applyStage1SpreadsheetHighlight(SpreadsheetCell cell, String columnTitle, String text) {
        if (cell == null) {
            return;
        }
        if (isStage1DateColumnHeader(columnTitle) && isStrictPositiveNumber(text)) {
            cell.setStyle(LIGHT_GREEN_STYLE);
        } else {
            cell.setStyle("");
        }
    }

    static boolean isUnprocessedColumnHeader(String header) {
        if (header == null) {
            return false;
        }
        return "未加工".equals(header.strip());
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
        if (h.contains("日付")) {
            return true;
        }
        if (h.endsWith("納期")) {
            return true;
        }
        if (h.endsWith("投入日")) {
            return true;
        }
        if (h.endsWith("配台日")) {
            return true;
        }
        if (h.endsWith("完了日")) {
            return true;
        }
        if ("受注日".equals(h)) {
            return true;
        }
        if ("データ抽出日".equals(h)) {
            return true;
        }
        if (h.startsWith("配台済_")
                && (h.contains("加工開始") || h.contains("加工終了"))) {
            return true;
        }
        if ("加工開始日".equals(h)) {
            return true;
        }
        if ("計画基準納期".equals(h)) {
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
            String n = t.replace("　", "").replace(",", "").replace("，", "");
            double v = Double.parseDouble(n);
            return v > 0.0 && Double.isFinite(v);
        } catch (NumberFormatException ignored) {
            return false;
        }
    }

    public static Callback<TableColumn<ObservableList<String>, String>, TableCell<ObservableList<String>, String>>
            stage1DateHighlightCellFactory(String columnTitle) {
        return stage1DateHighlightCellFactory(columnTitle, null, () -> 0);
    }

    /**
     * @param headerColumnCount first {@code n} visual columns get {@link TableHeaderColumnStyle} tint unless date
     *     highlight wins
     */
    public static Callback<TableColumn<ObservableList<String>, String>, TableCell<ObservableList<String>, String>>
            stage1DateHighlightCellFactory(
                    String columnTitle,
                    TableView<ObservableList<String>> table,
                    IntSupplier headerColumnCount) {
        IntSupplier hc = headerColumnCount != null ? headerColumnCount : () -> 0;
        return column ->
                new TextFieldTableCell<ObservableList<String>, String>() {
                    @Override
                    public void updateItem(String item, boolean empty) {
                        super.updateItem(item, empty);
                        if (table == null) {
                            applyLegacyStyle(item, empty);
                            return;
                        }
                        if (empty || item == null) {
                            setStyle("");
                            TableHeaderColumnStyle.applyBodyCellTint(this, table, column, hc);
                            return;
                        }
                        if (isStage1DateColumnHeader(columnTitle) && isStrictPositiveNumber(item)) {
                            getStyleClass().remove(TableHeaderColumnStyle.HEADER_BODY_CELL_STYLE_CLASS);
                            setStyle(LIGHT_GREEN_STYLE);
                        } else {
                            setStyle("");
                            TableHeaderColumnStyle.applyBodyCellTint(this, table, column, hc);
                        }
                    }

                    private void applyLegacyStyle(String item, boolean empty) {
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

    public static Callback<TableColumn<ObservableList<String>, String>, TableCell<ObservableList<String>, String>>
            planInputUnprocessedHighlightCellFactory(String columnTitle) {
        return planInputUnprocessedHighlightCellFactory(columnTitle, null, () -> 0);
    }

    /** @see #stage1DateHighlightCellFactory(String, TableView, IntSupplier) */
    public static Callback<TableColumn<ObservableList<String>, String>, TableCell<ObservableList<String>, String>>
            planInputUnprocessedHighlightCellFactory(
                    String columnTitle,
                    TableView<ObservableList<String>> table,
                    IntSupplier headerColumnCount) {
        IntSupplier hc = headerColumnCount != null ? headerColumnCount : () -> 0;
        return column ->
                new TextFieldTableCell<ObservableList<String>, String>() {
                    @Override
                    public void updateItem(String item, boolean empty) {
                        super.updateItem(item, empty);
                        if (table == null) {
                            if (empty || item == null) {
                                setStyle("");
                                return;
                            }
                            if (isUnprocessedColumnHeader(columnTitle) && isStrictPositiveNumber(item)) {
                                setStyle(LIGHT_GREEN_STYLE);
                            } else {
                                setStyle("");
                            }
                            return;
                        }
                        if (empty || item == null) {
                            setStyle("");
                            TableHeaderColumnStyle.applyBodyCellTint(this, table, column, hc);
                            return;
                        }
                        if (isUnprocessedColumnHeader(columnTitle) && isStrictPositiveNumber(item)) {
                            getStyleClass().remove(TableHeaderColumnStyle.HEADER_BODY_CELL_STYLE_CLASS);
                            setStyle(LIGHT_GREEN_STYLE);
                        } else {
                            setStyle("");
                            TableHeaderColumnStyle.applyBodyCellTint(this, table, column, hc);
                        }
                    }
                };
    }
}
