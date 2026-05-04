package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.BitSet;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.BooleanSupplier;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.TablePosition;
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
     * Same as {@link #applyColumnFilters(SpreadsheetView)} but opens a modal dialog per column (\u9069\u7528 / OK /
     * \u30ad\u30e3\u30f3\u30bb\u30eb) via {@link DialogExcelLikeSpreadsheetFilter}.
     */
    public static void applyColumnFiltersWithDialog(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        view.setFilteredRow(SPREADSHEET_FILTER_ROW);
        ObservableList<SpreadsheetColumn> cols = view.getColumns();
        for (int i = 0; i < cols.size(); i++) {
            cols.get(i).setFilter(new DialogExcelLikeSpreadsheetFilter(view, i));
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
        DialogExcelLikeSpreadsheetFilter.resetAllColumnSortMenus(view);
    }

    /** Applies leading column freeze from persisted 見出し列数 (must run after {@link SpreadsheetView#setGrid}). */
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
    /**
     * ControlsFX {@link SpreadsheetView} の内側 {@link TableView}（行＝グリッドの1行）を返す。
     */
    @SuppressWarnings("unchecked")
    public static TableView<ObservableList<SpreadsheetCell>> findInnerTableView(SpreadsheetView view) {
        if (view == null) {
            return null;
        }
        for (Node n : view.getChildrenUnmodifiable()) {
            if (n instanceof TableView<?> tv) {
                return (TableView<ObservableList<SpreadsheetCell>>) tv;
            }
        }
        return null;
    }

    public static void applyUnconstrainedColumnResizePolicy(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        /*
         * SpreadsheetView may host more than one embedded TableView (e.g. fixed vs scrollable columns).
         * Applying UNCONSTRAINED only to the first child left the other at the default CONSTRAINED policy,
         * which blocks interactive column resize for part of the grid.
         */
        setUnconstrainedOnEmbeddedTableViews(view, 0);
        for (SpreadsheetColumn col : view.getColumns()) {
            col.setResizable(true);
        }
    }

    private static void setUnconstrainedOnEmbeddedTableViews(Node n, int depth) {
        if (n == null || depth > 12) {
            return;
        }
        if (n instanceof TableView<?> tv) {
            tv.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
            return;
        }
        if (n instanceof Parent p) {
            for (Node c : p.getChildrenUnmodifiable()) {
                setUnconstrainedOnEmbeddedTableViews(c, depth + 1);
            }
        }
    }

    /**
     * After any cell selection change, expands selection to full grid rows (all columns) for each data row that
     * had at least one selected cell, so the active row is visually continuous. Skips the filter row at
     * {@link #SPREADSHEET_FILTER_ROW}.
     *
     * <p>Requires {@link SpreadsheetView#getSelectionModel()}{@code .setSelectionMode(SelectionMode.MULTIPLE)}.
     *
     * @param skipFullRowExpansion when {@code true}, skips expanding selection (e.g. during plan-input row drag).
     */
    public static void installFullRowDataSelection(
            SpreadsheetView view, BooleanSupplier skipFullRowExpansion) {
        if (view == null) {
            return;
        }
        final boolean[] guard = {false};
        var sm = view.getSelectionModel();
        // SpreadsheetView exposes ObservableList<TablePosition> (raw); listener must accept TablePosition.
        sm.getSelectedCells()
                .addListener(
                        (ListChangeListener<? super TablePosition>)
                                change -> {
                                    if (guard[0]) {
                                        return;
                                    }
                                    if (skipFullRowExpansion != null && skipFullRowExpansion.getAsBoolean()) {
                                        return;
                                    }
                                    ObservableList<SpreadsheetColumn> cols = view.getColumns();
                                    if (cols.isEmpty()) {
                                        return;
                                    }
                                    ObservableList<? extends TablePosition> selected =
                                            sm.getSelectedCells();
                                    if (selected.isEmpty()) {
                                        return;
                                    }
                                    int firstData = spreadsheetFirstDataRowIndex();
                                    Set<Integer> rows = new HashSet<>();
                                    for (TablePosition p : selected) {
                                        int r = p.getRow();
                                        if (r >= firstData) {
                                            rows.add(r);
                                        }
                                    }
                                    if (rows.isEmpty()) {
                                        return;
                                    }
                                    int colCount = cols.size();
                                    boolean allFull = true;
                                    for (int r : rows) {
                                        int cnt = 0;
                                        for (TablePosition p : selected) {
                                            if (p.getRow() == r) {
                                                cnt++;
                                            }
                                        }
                                        if (cnt < colCount) {
                                            allFull = false;
                                            break;
                                        }
                                    }
                                    if (allFull) {
                                        return;
                                    }
                                    TablePosition focus = sm.getFocusedCell();
                                    int focusRow;
                                    if (focus != null && focus.getRow() >= firstData) {
                                        focusRow = focus.getRow();
                                    } else {
                                        focusRow = Collections.min(rows);
                                    }
                                    int focusCol =
                                            focus != null && focus.getColumn() >= 0 ? focus.getColumn() : 0;

                                    guard[0] = true;
                                    try {
                                        sm.clearSelection();
                                        SpreadsheetColumn firstCol = cols.get(0);
                                        SpreadsheetColumn lastCol = cols.get(cols.size() - 1);
                                        ArrayList<Integer> sorted = new ArrayList<>(rows);
                                        Collections.sort(sorted);
                                        for (int r : sorted) {
                                            sm.selectRange(r, firstCol, r, lastCol);
                                        }
                                        int fc = Math.min(Math.max(focusCol, 0), cols.size() - 1);
                                        sm.focus(focusRow, cols.get(fc));
                                    } finally {
                                        guard[0] = false;
                                    }
                                });
    }

    /** @see #installFullRowDataSelection(SpreadsheetView, BooleanSupplier) */
    public static void installFullRowDataSelection(SpreadsheetView view) {
        installFullRowDataSelection(view, null);
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
     * Used for JSON-backed viewers such as {@code 結果_配台表.json}.
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

    /**
     * Read-only grid with timeline / Gantt-style cell coloring (see {@link GanttScheduleStyle}).
     *
     * @param kind 設備ガント Excel 風など表現の選択
     */
    public static GridBase buildReadOnlyGanttGrid(
            List<String> headersRef,
            ObservableList<ObservableList<String>> rows,
            GanttSheetKind kind) {
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
            boolean sectionRow = false;
            if (!src.isEmpty() && src.get(0) != null) {
                String head = src.get(0);
                sectionRow = head.contains("■") || head.contains("▪");
            }
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            for (int c = 0; c < cols; c++) {
                String raw = c < src.size() && src.get(c) != null ? src.get(c) : "";
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw);
                cell.setEditable(false);
                GanttScheduleStyle.applyTimelineCell(
                        cell, c, headersRef.get(c), raw, sectionRow, r, kind);
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
