package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.BitSet;
import java.util.List;
import java.util.Objects;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.scene.Node;
import javafx.scene.control.TableView;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * Bridges tabular {@link ObservableList} rows to ControlsFX {@link SpreadsheetView} / {@link GridBase}.
 */
public final class SpreadsheetTabularSupport {

    /** Grid row index reserved for ControlsFX column filters ({@link SpreadsheetView#setFilteredRow}). */
    public static final int SPREADSHEET_FILTER_ROW = 0;

    /** First grid row index that maps to {@link ObservableList} data row 0. */
    public static int spreadsheetFirstDataRowIndex() {
        return SPREADSHEET_FILTER_ROW + 1;
    }

    private SpreadsheetTabularSupport() {}

    /**
     * Enables ControlsFX per-column filters and row sorting (filter menu on column headers). Requires a filter
     * placeholder row in the {@link GridBase} at {@link #SPREADSHEET_FILTER_ROW}.
     */
    public static void applyColumnFilters(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        view.setFilteredRow(SPREADSHEET_FILTER_ROW);
        ObservableList<SpreadsheetColumn> cols = view.getColumns();
        for (int i = 0; i < cols.size(); i++) {
            cols.get(i).setFilter(new ExcelLikeSpreadsheetFilter(view, i));
        }
    }

    /**
     * Clears spreadsheet-wide row sort and per-column row filters (hidden rows), and resets filter menu sort labels.
     */
    public static void clearAllFiltersAndSort(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        view.setComparator(null);
        view.setHiddenRows(new BitSet());
        ExcelLikeSpreadsheetFilter.resetAllColumnSortMenus(view);
    }

    /** Applies leading column freeze from persisted \u898b\u51fa\u3057\u5217\u6570 (must run after {@link SpreadsheetView#setGrid}). */
    public static void applyFixedLeadingColumns(SpreadsheetView view, int headerColumnCount) {
        if (view == null) {
            return;
        }
        view.getFixedColumns().clear();
        int n = Math.max(0, headerColumnCount);
        if (n <= 0) {
            return;
        }
        ObservableList<SpreadsheetColumn> cols = view.getColumns();
        int limit = Math.min(n, cols.size());
        for (int i = 0; i < limit; i++) {
            SpreadsheetColumn col = cols.get(i);
            if (col.isColumnFixable()) {
                view.getFixedColumns().add(col);
            }
        }
    }

    public static void applyFixedLeadingColumnsLater(SpreadsheetView view, int headerColumnCount) {
        Platform.runLater(() -> applyFixedLeadingColumns(view, headerColumnCount));
    }

    /**
     * ControlsFX embeds an inner {@link TableView} whose default constrained resize policy can block or fight
     * interactive column resizing; unconstrained mode matches typical spreadsheet drag-to-resize expectations.
     */
    public static void applyUnconstrainedColumnResizePolicy(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        for (Node n : view.getChildrenUnmodifiable()) {
            if (n instanceof TableView<?> tv) {
                tv.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
                break;
            }
        }
        for (SpreadsheetColumn col : view.getColumns()) {
            col.setResizable(true);
        }
    }

    public static GridBase buildPlanInputGrid(
            List<String> headersRef, ObservableList<ObservableList<String>> rows, boolean editable) {
        int cols = headersRef.size();
        int rc = rows.size();
        int gridRowsTotal = rc + 1;
        GridBase grid = new GridBase(gridRowsTotal, cols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(headersRef);

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < cols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        int firstData = spreadsheetFirstDataRowIndex();
        for (int r = 0; r < rc; r++) {
            int gridRow = firstData + r;
            ObservableList<String> src = rows.get(r);
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            for (int c = 0; c < cols; c++) {
                String raw = c < src.size() && src.get(c) != null ? src.get(c) : "";
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw);
                cell.setEditable(editable);
                TabularCellHighlight.applyPlanInputSpreadsheetHighlight(cell, headersRef.get(c), raw);
                rowCells.add(cell);
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        return grid;
    }

    public static GridBase buildStage1PreviewGrid(List<String> headersRef, ObservableList<ObservableList<String>> rows) {
        int cols = headersRef.size();
        int rc = rows.size();
        int gridRowsTotal = rc + 1;
        GridBase grid = new GridBase(gridRowsTotal, cols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(headersRef);

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < cols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        int firstData = spreadsheetFirstDataRowIndex();
        for (int r = 0; r < rc; r++) {
            int gridRow = firstData + r;
            ObservableList<String> src = rows.get(r);
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            for (int c = 0; c < cols; c++) {
                String raw = c < src.size() && src.get(c) != null ? src.get(c) : "";
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw);
                cell.setEditable(false);
                TabularCellHighlight.applyStage1SpreadsheetHighlight(cell, headersRef.get(c), raw);
                rowCells.add(cell);
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        return grid;
    }

    /**
     * Read-only string grid with Excel-like column filters ({@link #applyColumnFilters}); no row highlighting.
     * Used for JSON-backed viewers such as {@code \u7d50\u679c_\u914d\u53f0\u8868.json}.
     */
    public static GridBase buildReadOnlyPlainGrid(
            List<String> headersRef, ObservableList<ObservableList<String>> rows) {
        int cols = headersRef.size();
        int rc = rows.size();
        int gridRowsTotal = rc + 1;
        GridBase grid = new GridBase(gridRowsTotal, cols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(headersRef);

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < cols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        int firstData = spreadsheetFirstDataRowIndex();
        for (int r = 0; r < rc; r++) {
            int gridRow = firstData + r;
            ObservableList<String> src = rows.get(r);
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            for (int c = 0; c < cols; c++) {
                String raw = c < src.size() && src.get(c) != null ? src.get(c) : "";
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw);
                cell.setEditable(false);
                rowCells.add(cell);
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        return grid;
    }

    public static javafx.event.EventHandler<org.controlsfx.control.spreadsheet.GridChange> newRowsSyncHandler(
            ObservableList<ObservableList<String>> rows, List<String> headersRef, int firstDataGridRow) {
        return ev -> {
            int r = ev.getRow();
            int c = ev.getColumn();
            if (r < firstDataGridRow || c < 0 || c >= headersRef.size()) {
                return;
            }
            int dataIndex = r - firstDataGridRow;
            if (dataIndex < 0 || dataIndex >= rows.size()) {
                return;
            }
            Object nv = ev.getNewValue();
            String s = nv != null ? Objects.toString(nv, "") : "";
            ObservableList<String> row = rows.get(dataIndex);
            while (row.size() <= c) {
                row.add("");
            }
            row.set(c, s);
        };
    }

    public static void applyColumnWidths(SpreadsheetView view, List<Double> widths, double defaultWidth) {
        if (view == null || widths == null) {
            return;
        }
        var cols = view.getColumns();
        for (int i = 0; i < cols.size(); i++) {
            double w = i < widths.size() ? widths.get(i) : defaultWidth;
            cols.get(i).setPrefWidth(w);
        }
    }
}
