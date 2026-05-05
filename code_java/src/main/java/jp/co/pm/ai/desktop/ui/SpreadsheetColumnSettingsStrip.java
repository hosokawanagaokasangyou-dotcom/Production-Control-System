package jp.co.pm.ai.desktop.ui;

import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;

import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.HBox;

/**
 * Column settings for ControlsFX spreadsheet tabs: default-width reset + persisted 見出し列数
 * (fixed leading columns are applied in the tab controller after {@code setGrid}).
 */
public final class SpreadsheetColumnSettingsStrip {

    private SpreadsheetColumnSettingsStrip() {}

    /**
     * @param onLeadingColumnCountChanged called after count is saved; typically rebuilds grid and reapplies freeze
     * @param reorderColumns opens column reorder UI ({@code null} = hide button)
     */
    public static HBox create(
            Runnable resetColumnWidths,
            TableColumnOrderPersistence.TableId tableId,
            AtomicInteger headerColumnCountHolder,
            Consumer<Integer> onLeadingColumnCountChanged,
            Runnable reorderColumns) {
        int initial = TableColumnOrderPersistence.loadHeaderColumnCount(tableId);
        headerColumnCountHolder.set(initial);
        Spinner<Integer> headerSpinner =
                new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 999, initial));
        headerSpinner.setEditable(true);
        headerSpinner
                .valueProperty()
                .addListener(
                        (obs, o, v) -> {
                            if (v == null) {
                                return;
                            }
                            int hv = Math.max(0, v);
                            headerColumnCountHolder.set(hv);
                            TableColumnOrderPersistence.saveHeaderColumnCount(tableId, hv);
                            onLeadingColumnCountChanged.accept(hv);
                        });
        Button reorder =
                new Button(
                        "列の並べ替え");
        reorder.setOnAction(e -> {
            if (reorderColumns != null) {
                reorderColumns.run();
            }
        });
        reorder.setManaged(reorderColumns != null);
        reorder.setVisible(reorderColumns != null);

        Button reset = new Button("列幅を既定に");
        reset.setOnAction(
                e -> {
                    if (resetColumnWidths != null) {
                        resetColumnWidths.run();
                    }
                });
        HBox h =
                new HBox(
                        8,
                        new Label("列設定"),
                        new Label("見出し列数"),
                        headerSpinner,
                        reorder,
                        reset);
        h.setStyle("-fx-alignment: CENTER_LEFT;");
        return h;
    }

    /**
     * 計画結果 JSON ビューアなど、{@link TableColumnOrderPersistence#planResultViewerSheetScopeKey} 単位で見出し列数を保存する。
     */
    public static HBox createForScope(
            Runnable resetColumnWidths,
            String sheetScopeKey,
            AtomicInteger headerColumnCountHolder,
            Consumer<Integer> onLeadingColumnCountChanged,
            Runnable reorderColumns) {
        int initial = TableColumnOrderPersistence.loadHeaderColumnCountForScope(sheetScopeKey);
        headerColumnCountHolder.set(initial);
        Spinner<Integer> headerSpinner =
                new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 999, initial));
        headerSpinner.setEditable(true);
        headerSpinner
                .valueProperty()
                .addListener(
                        (obs, o, v) -> {
                            if (v == null) {
                                return;
                            }
                            int hv = Math.max(0, v);
                            headerColumnCountHolder.set(hv);
                            TableColumnOrderPersistence.saveHeaderColumnCountForScope(sheetScopeKey, hv);
                            onLeadingColumnCountChanged.accept(hv);
                        });
        Button reorder =
                new Button(
                        "列の並べ替え");
        reorder.setOnAction(e -> {
            if (reorderColumns != null) {
                reorderColumns.run();
            }
        });
        reorder.setManaged(reorderColumns != null);
        reorder.setVisible(reorderColumns != null);

        Button reset = new Button("列幅を既定に");
        reset.setOnAction(
                e -> {
                    if (resetColumnWidths != null) {
                        resetColumnWidths.run();
                    }
                });
        HBox h =
                new HBox(
                        8,
                        new Label("列設定"),
                        new Label("見出し列数"),
                        headerSpinner,
                        reorder,
                        reset);
        h.setStyle("-fx-alignment: CENTER_LEFT;");
        return h;
    }
}
