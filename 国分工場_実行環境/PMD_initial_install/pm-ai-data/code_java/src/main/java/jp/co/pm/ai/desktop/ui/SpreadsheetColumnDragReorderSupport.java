package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BooleanSupplier;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.application.Platform;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;

import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * ControlsFX {@link SpreadsheetView} sets {@link TableColumn#setReorderable(boolean)} false on inner
 * columns. Enables reorder on non-fixed columns and maps header drag permutations to logical column order.
 */
public final class SpreadsheetColumnDragReorderSupport {

    private static final String PROP_LISTENER = "pmSpreadsheetColumnReorderListener";

    /**
     * {@link #refreshAfterGridReady} が設定する先頭固定列数。{@link ColumnVisibilitySupport} が表示変更後に
     * {@link #updateColumnReorderFlags(SpreadsheetView, int)} を呼ぶために参照する。
     */
    public static final String PROP_LEADING_FIXED = "pmSpreadsheetReorderLeadingFixed";

    private SpreadsheetColumnDragReorderSupport() {}

    /**
     * Call after {@link SpreadsheetView#setGrid}. Enables drag reorder on embedded {@link TableView}
     * columns and invokes {@code onVisualOrderChanged} with header texts left-to-right when user
     * permutes columns.
     *
     * @param leadingFixedColumnCount leading columns that stay non-reorderable (pinned header count)
     */
    public static void refreshAfterGridReady(
            SpreadsheetView view,
            BooleanSupplier suppress,
            Supplier<List<String>> headersSupplier,
            int leadingFixedColumnCount,
            Consumer<List<String>> onVisualOrderChanged) {
        if (view == null || onVisualOrderChanged == null) {
            return;
        }
        Platform.runLater(
                () ->
                        refreshAfterGridReadyImpl(
                                view,
                                suppress,
                                headersSupplier,
                                leadingFixedColumnCount,
                                onVisualOrderChanged));
    }

    @SuppressWarnings("unchecked")
    private static void refreshAfterGridReadyImpl(
            SpreadsheetView view,
            BooleanSupplier suppress,
            Supplier<List<String>> headersSupplier,
            int leadingFixedColumnCount,
            Consumer<List<String>> onVisualOrderChanged) {
        List<String> headers = headersSupplier != null ? headersSupplier.get() : null;
        int expected = headers == null ? 0 : headers.size();
        if (expected <= 0) {
            return;
        }
        TableView<?> tv = findEmbeddedTableViewMatchingColumnCount(view, expected);
        if (tv == null) {
            return;
        }
        ObservableList<TableColumn<?, ?>> cols =
                (ObservableList<TableColumn<?, ?>>) (ObservableList<?>) tv.getColumns();
        Object prev = view.getProperties().get(PROP_LISTENER);
        if (prev instanceof ListChangeListener) {
            try {
                cols.removeListener((ListChangeListener<? super TableColumn<?, ?>>) prev);
            } catch (Exception ignored) {
            }
        }
        int fixed = Math.max(0, leadingFixedColumnCount);
        view.getProperties().put(PROP_LEADING_FIXED, fixed);
        updateColumnReorderFlags(view, fixed);
        ListChangeListener<TableColumn<?, ?>> listener =
                c -> {
                    while (c.next()) {
                        if (!c.wasPermutated()) {
                            continue;
                        }
                        if (suppress != null && suppress.getAsBoolean()) {
                            continue;
                        }
                        List<String> h = headersSupplier != null ? headersSupplier.get() : List.of();
                        if (h.size() != cols.size()) {
                            continue;
                        }
                        List<String> visual = new ArrayList<>(cols.size());
                        for (TableColumn<?, ?> col : cols) {
                            String t = col.getText();
                            visual.add(t != null ? t : "");
                        }
                        if (!isSameMultiset(visual, h)) {
                            continue;
                        }
                        onVisualOrderChanged.accept(visual);
                    }
                };
        cols.addListener(listener);
        view.getProperties().put(PROP_LISTENER, listener);
    }

    /**
     * 先頭 {@code leadingFixedColumnCount} 列はドラッグ並べ替え不可。それ以外は列が表示されているときのみ
     * 並べ替え可能（非表示列はヘッダが無いため対象外）。
     */
    public static void updateColumnReorderFlags(SpreadsheetView view, int leadingFixedColumnCount) {
        if (view == null) {
            return;
        }
        int expected = view.getColumns().size();
        if (expected <= 0) {
            return;
        }
        TableView<?> tv = findEmbeddedTableViewMatchingColumnCount(view, expected);
        if (tv == null) {
            return;
        }
        @SuppressWarnings("unchecked")
        ObservableList<TableColumn<?, ?>> cols =
                (ObservableList<TableColumn<?, ?>>) (ObservableList<?>) tv.getColumns();
        int fixed = Math.max(0, leadingFixedColumnCount);
        for (int i = 0; i < cols.size(); i++) {
            TableColumn<?, ?> col = cols.get(i);
            col.setReorderable(i >= fixed && col.isVisible());
        }
    }

    private static boolean isSameMultiset(List<String> a, List<String> b) {
        if (a.size() != b.size()) {
            return false;
        }
        Map<String, Integer> freq = new HashMap<>();
        for (String s : a) {
            freq.merge(s != null ? s : "", 1, Integer::sum);
        }
        for (String s : b) {
            String k = s != null ? s : "";
            Integer n = freq.get(k);
            if (n == null || n <= 0) {
                return false;
            }
            if (n == 1) {
                freq.remove(k);
            } else {
                freq.put(k, n - 1);
            }
        }
        return freq.isEmpty();
    }

    private static TableView<?> findEmbeddedTableViewMatchingColumnCount(SpreadsheetView view, int expected) {
        return scanForTableView(view, expected, 0);
    }

    private static TableView<?> scanForTableView(Node n, int expected, int depth) {
        if (n == null || depth > 24) {
            return null;
        }
        if (n instanceof TableView<?> tv) {
            if (tv.getColumns().size() == expected) {
                return tv;
            }
        }
        if (n instanceof Parent p) {
            for (Node c : p.getChildrenUnmodifiable()) {
                TableView<?> r = scanForTableView(c, expected, depth + 1);
                if (r != null) {
                    return r;
                }
            }
        }
        return null;
    }
}
