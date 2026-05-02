package jp.co.pm.ai.desktop.ui;

import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
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
        CheckBox flex = new CheckBox("\u6700\u7d42\u5217\u3092\u4f38\u7e2e");
        flex.setSelected(flexLastColumnInitially);
        applyResizePolicy(table, flex.isSelected());
        flex.selectedProperty()
                .addListener(
                        (obs, a, b) -> {
                            if (b != null) {
                                applyResizePolicy(table, b);
                            }
                        });
        Button reset = new Button("\u5217\u5e45\u3092\u65e2\u5b9a\u306b");
        reset.setOnAction(
                e -> {
                    if (resetToDefaults != null) {
                        resetToDefaults.run();
                    }
                });
        HBox h = new HBox(8, new Label("\u5217\u8a2d\u5b9a"), flex, reset);
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
