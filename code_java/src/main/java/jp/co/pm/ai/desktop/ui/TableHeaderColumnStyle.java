package jp.co.pm.ai.desktop.ui;

import java.util.function.IntSupplier;

import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;

/**
 * Marks the first {@code n} <em>visual</em> columns (left-to-right order in {@link TableView#getColumns()}) as
 * header columns: distinct column header chrome and a light tint on body cells.
 *
 * <p>Body tint uses {@link #HEADER_BODY_CELL_STYLE_CLASS} plus {@code pm-ai-desktop.css} so selected rows keep theme
 * contrast (inline styles would paint white text on a pale band).
 */
public final class TableHeaderColumnStyle {

    /** Applied to body {@link javafx.scene.control.TableCell}s in the header band; colors come from CSS. */
    public static final String HEADER_BODY_CELL_STYLE_CLASS = "pm-header-body-cell";

    /** Style for {@link TableColumn} header region. */
    static final String HEADER_COLUMN_CHROME_STYLE =
            "-fx-background-color: derive(-fx-base, -10%); -fx-font-weight: bold;";

    public static final String HEADER_COLUMN_STYLE_CLASS = "pm-header-column";

    private TableHeaderColumnStyle() {}

    /**
     * Applies header-column chrome to the first {@code headerColumnCount} columns. Safe when column count is smaller
     * than {@code n}.
     */
    public static void applyToTableColumns(TableView<?> table, int headerColumnCount) {
        if (table == null) {
            return;
        }
        int n = Math.max(0, headerColumnCount);
        var cols = table.getColumns();
        int limit = Math.min(n, cols.size());
        for (int i = 0; i < cols.size(); i++) {
            TableColumn<?, ?> c = cols.get(i);
            if (i < limit) {
                if (!c.getStyleClass().contains(HEADER_COLUMN_STYLE_CLASS)) {
                    c.getStyleClass().add(HEADER_COLUMN_STYLE_CLASS);
                }
                c.setStyle(HEADER_COLUMN_CHROME_STYLE);
            } else {
                c.getStyleClass().remove(HEADER_COLUMN_STYLE_CLASS);
                c.setStyle(null);
            }
        }
    }

    /**
     * Body cell tint for dynamic tables: uses current visual index of {@code column} and {@code headerColumnCount}.
     */
    public static void applyBodyCellTint(
            TableCell<?, ?> cell,
            TableView<?> table,
            TableColumn<?, ?> column,
            IntSupplier headerColumnCount) {
        if (cell == null || table == null || column == null || headerColumnCount == null) {
            return;
        }
        int idx = table.getColumns().indexOf(column);
        int n = Math.max(0, headerColumnCount.getAsInt());
        int cap = Math.min(n, table.getColumns().size());
        if (idx >= 0 && idx < cap) {
            if (!cell.getStyleClass().contains(HEADER_BODY_CELL_STYLE_CLASS)) {
                cell.getStyleClass().add(HEADER_BODY_CELL_STYLE_CLASS);
            }
        } else {
            cell.getStyleClass().remove(HEADER_BODY_CELL_STYLE_CLASS);
            cell.setStyle("");
        }
    }

    /**
     * Same as {@link #applyBodyCellTint} but preserves {@code styleWhenNotHeader} when the cell is not in the header
     * column band (e.g. existing highlight).
     */
    public static void applyBodyCellTintPreservingAlternate(
            TableCell<?, ?> cell,
            TableView<?> table,
            TableColumn<?, ?> column,
            IntSupplier headerColumnCount,
            String styleWhenNotHeader) {
        if (cell == null || table == null || column == null || headerColumnCount == null) {
            return;
        }
        int idx = table.getColumns().indexOf(column);
        int n = Math.max(0, headerColumnCount.getAsInt());
        int cap = Math.min(n, table.getColumns().size());
        if (idx >= 0 && idx < cap) {
            if (!cell.getStyleClass().contains(HEADER_BODY_CELL_STYLE_CLASS)) {
                cell.getStyleClass().add(HEADER_BODY_CELL_STYLE_CLASS);
            }
        } else {
            cell.getStyleClass().remove(HEADER_BODY_CELL_STYLE_CLASS);
            cell.setStyle(styleWhenNotHeader != null ? styleWhenNotHeader : "");
        }
    }
}
