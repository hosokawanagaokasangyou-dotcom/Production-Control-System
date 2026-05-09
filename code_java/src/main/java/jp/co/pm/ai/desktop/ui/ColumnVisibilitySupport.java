package jp.co.pm.ai.desktop.ui;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Supplier;

import javafx.application.Platform;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.stage.Window;

import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * Applies persisted per-column visibility to {@link SpreadsheetView} / {@link TableView} and opens the dialog.
 */
public final class ColumnVisibilitySupport {

    private ColumnVisibilitySupport() {}

    public static void applyColumnVisibilityToSpreadsheet(
            SpreadsheetView view, List<String> headersAlignedToColumns, boolean[] visible) {
        if (view == null || headersAlignedToColumns == null || visible == null) {
            return;
        }
        List<SpreadsheetColumn> cols = view.getColumns();
        int n = Math.min(headersAlignedToColumns.size(), Math.min(cols.size(), visible.length));
        for (int i = 0; i < n; i++) {
            setSpreadsheetColumnInnerVisible(cols.get(i), visible[i]);
        }
        for (int i = n; i < cols.size(); i++) {
            setSpreadsheetColumnInnerVisible(cols.get(i), true);
        }
    }

    /** ControlsFX {@link SpreadsheetColumn} wraps a {@link TableColumn}; visibility is set on the inner column. */
    private static void setSpreadsheetColumnInnerVisible(SpreadsheetColumn wrapper, boolean visible) {
        if (wrapper == null) {
            return;
        }
        try {
            Field f = SpreadsheetColumn.class.getDeclaredField("column");
            f.setAccessible(true);
            Object tc = f.get(wrapper);
            if (tc instanceof TableColumn<?, ?> col) {
                col.setVisible(visible);
            }
        } catch (ReflectiveOperationException e) {
            throw new IllegalStateException("SpreadsheetColumn.column", e);
        }
    }

    /**
     * After {@link SpreadsheetView#setGrid}, inner columns may appear on a later layout pulse ? retries briefly.
     */
    public static void applyColumnVisibilityToSpreadsheetWhenReady(
            SpreadsheetView view,
            Supplier<List<String>> headersSupplier,
            Supplier<boolean[]> visibilitySupplier) {
        if (view == null || headersSupplier == null || visibilitySupplier == null) {
            return;
        }
        Runnable[] job = new Runnable[1];
        final int[] attempts = {0};
        job[0] =
                () -> {
                    attempts[0]++;
                    List<String> h = headersSupplier.get();
                    boolean[] v = visibilitySupplier.get();
                    if (h == null || v == null) {
                        return;
                    }
                    int expected = h.size();
                    int actual = view.getColumns().size();
                    if (actual < expected && attempts[0] < 48) {
                        Platform.runLater(job[0]);
                        return;
                    }
                    applyColumnVisibilityToSpreadsheet(view, h, v);
                };
        Platform.runLater(job[0]);
    }

    public static void applyColumnVisibilityToTableView(TableView<?> table, boolean[] visible) {
        if (table == null || visible == null) {
            return;
        }
        List<? extends TableColumn<?, ?>> cols = table.getColumns();
        int n = Math.min(cols.size(), visible.length);
        for (int i = 0; i < n; i++) {
            cols.get(i).setVisible(visible[i]);
        }
        for (int i = n; i < cols.size(); i++) {
            cols.get(i).setVisible(true);
        }
    }

    public static void openSpreadsheetColumnVisibilityDialog(
            Window owner,
            TableColumnOrderPersistence.TableId tableId,
            SpreadsheetView view,
            Supplier<List<String>> headersSupplier) {
        Objects.requireNonNull(view, "view");
        List<String> headers = headersSupplier != null ? headersSupplier.get() : List.of();
        if (headers == null || headers.isEmpty()) {
            return;
        }
        boolean[] vis = TableColumnOrderPersistence.loadColumnVisibility(tableId, headers.size());
        ColumnVisibilityDialog.show(owner, headers, vis)
                .ifPresent(
                        arr -> {
                            TableColumnOrderPersistence.saveColumnVisibility(tableId, arr);
                            applyColumnVisibilityToSpreadsheetWhenReady(
                                    view,
                                    () -> new ArrayList<>(headersSupplier.get()),
                                    () -> arr);
                        });
    }

    public static void openSpreadsheetColumnVisibilityDialogForScope(
            Window owner,
            String scopeKey,
            SpreadsheetView view,
            Supplier<List<String>> headersSupplier) {
        Objects.requireNonNull(view, "view");
        List<String> headers = headersSupplier != null ? headersSupplier.get() : List.of();
        if (headers == null || headers.isEmpty() || scopeKey == null || scopeKey.isEmpty()) {
            return;
        }
        boolean[] vis = TableColumnOrderPersistence.loadColumnVisibilityForScope(scopeKey, headers.size());
        ColumnVisibilityDialog.show(owner, headers, vis)
                .ifPresent(
                        arr -> {
                            TableColumnOrderPersistence.saveColumnVisibilityForScope(scopeKey, arr);
                            applyColumnVisibilityToSpreadsheetWhenReady(
                                    view,
                                    () -> new ArrayList<>(headersSupplier.get()),
                                    () -> arr);
                        });
    }

    public static void openTableViewColumnVisibilityDialog(
            Window owner, TableColumnOrderPersistence.TableId tableId, TableView<?> table) {
        if (table == null || tableId == null) {
            return;
        }
        List<String> titles = new ArrayList<>();
        for (TableColumn<?, ?> c : table.getColumns()) {
            String t = c.getText();
            titles.add(t != null ? t : "");
        }
        if (titles.isEmpty()) {
            return;
        }
        boolean[] vis = TableColumnOrderPersistence.loadColumnVisibility(tableId, titles.size());
        ColumnVisibilityDialog.show(owner, titles, vis)
                .ifPresent(
                        arr -> {
                            TableColumnOrderPersistence.saveColumnVisibility(tableId, arr);
                            applyColumnVisibilityToTableView(table, arr);
                        });
    }
}
