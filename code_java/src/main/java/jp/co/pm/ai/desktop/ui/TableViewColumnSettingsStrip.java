package jp.co.pm.ai.desktop.ui;

import java.util.concurrent.atomic.AtomicInteger;

import javafx.collections.ListChangeListener;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.layout.HBox;

/**
 * Per-table column UI: resize policy (flex last) + reset widths. Place directly above the
 * {@link TableView} it controls.
 */
public final class TableViewColumnSettingsStrip {

    private TableViewColumnSettingsStrip() {}

    /**
     * @param table the table this strip controls
     * @param resetToDefaults reapplies design-time column widths (and min widths if the callback does so)
     * @param flexLastColumnInitially same as {@link TableView#CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN}
     */
    public static HBox create(
            TableView<?> table, Runnable resetToDefaults, boolean flexLastColumnInitially) {
        return create(table, resetToDefaults, flexLastColumnInitially, null, null, null);
    }

    /**
     * Same as {@link #create(TableView, Runnable, boolean)} plus optional leading visual columns as header columns
     * (persisted under {@link TableColumnOrderPersistence#saveHeaderColumnCount}).
     *
     * @param headerColumnCountHolder receives and publishes {@code n}; used by cell factories via {@link
     *     AtomicInteger#get()}
     */
    public static HBox create(
            TableView<?> table,
            Runnable resetToDefaults,
            boolean flexLastColumnInitially,
            TableColumnOrderPersistence.TableId tableId,
            AtomicInteger headerColumnCountHolder) {
        return create(table, resetToDefaults, flexLastColumnInitially, tableId, headerColumnCountHolder, null);
    }

    /**
     * @param columnVisibility opens column visibility UI ({@code null} = hide button); typically persists via
     *     {@link ColumnVisibilitySupport#openTableViewColumnVisibilityDialog}
     */
    public static HBox create(
            TableView<?> table,
            Runnable resetToDefaults,
            boolean flexLastColumnInitially,
            TableColumnOrderPersistence.TableId tableId,
            AtomicInteger headerColumnCountHolder,
            Runnable columnVisibility) {
        CheckBox flex = new CheckBox("最終列を伸縮");
        flex.setSelected(flexLastColumnInitially);
        applyResizePolicy(table, flex.isSelected());
        flex.selectedProperty()
                .addListener(
                        (obs, a, b) -> {
                            if (b != null) {
                                applyResizePolicy(table, b);
                            }
                        });
        Button reset = new Button("列幅を既定に");
        reset.setOnAction(
                e -> {
                    if (resetToDefaults != null) {
                        resetToDefaults.run();
                    }
                });

        Button visibility =
                new Button(
                        "\u5217\u306e\u8868\u793a");
        visibility.setOnAction(
                e -> {
                    if (columnVisibility != null) {
                        columnVisibility.run();
                    }
                });
        visibility.setManaged(columnVisibility != null);
        visibility.setVisible(columnVisibility != null);

        Runnable refreshHeaderColumns =
                () -> {
                    if (tableId != null && headerColumnCountHolder != null) {
                        int n = Math.max(0, headerColumnCountHolder.get());
                        TableHeaderColumnStyle.applyToTableColumns(table, n);
                        table.refresh();
                    }
                };

        if (tableId != null && headerColumnCountHolder != null) {
            int initial = TableColumnOrderPersistence.loadHeaderColumnCount(tableId);
            headerColumnCountHolder.set(initial);
            Spinner<Integer> headerSpinner =
                    new Spinner<>(
                            new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 999, initial));
            headerSpinner.setEditable(true);
            headerSpinner
                    .valueProperty()
                    .addListener(
                            (obs, o, v) -> {
                                if (v == null) {
                                    return;
                                }
                                headerColumnCountHolder.set(Math.max(0, v));
                                TableHeaderColumnStyle.applyToTableColumns(table, headerColumnCountHolder.get());
                                TableColumnOrderPersistence.saveHeaderColumnCount(
                                        tableId, headerColumnCountHolder.get());
                                table.refresh();
                            });
            table.getColumns()
                    .addListener(
                            (ListChangeListener<TableColumn<?, ?>>)
                                    c -> {
                                        while (c.next()) {
                                            // structural or reorder
                                        }
                                        refreshHeaderColumns.run();
                                    });
            javafx.application.Platform.runLater(refreshHeaderColumns);
            HBox h =
                    new HBox(
                            8,
                            new Label("列設定"),
                            flex,
                            new Label("見出し列数"),
                            headerSpinner,
                            visibility,
                            reset);
            h.setStyle("-fx-alignment: CENTER_LEFT;");
            return h;
        }

        HBox h = new HBox(8, new Label("列設定"), flex, visibility, reset);
        h.setStyle("-fx-alignment: CENTER_LEFT;");
        return h;
    }

    private static void applyResizePolicy(TableView<?> table, boolean flexLast) {
        if (table == null) {
            return;
        }
        table.setColumnResizePolicy(
                flexLast
                        ? TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN
                        : TableView.UNCONSTRAINED_RESIZE_POLICY);
    }
}
