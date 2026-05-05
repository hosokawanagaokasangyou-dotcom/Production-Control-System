package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.DialogPane;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

/**
 * Chooses a left-to-right column order for ControlsFX {@link org.controlsfx.control.spreadsheet.SpreadsheetView}
 * data (permutation of original column indices).
 */
public final class SpreadsheetColumnReorderDialog {

    private SpreadsheetColumnReorderDialog() {}

    /**
     * @param owner parent window
     * @param headers current column headers (may contain duplicates)
     * @return permutation: visual position {@code j} uses original column index {@code result.get(j)}
     */
    public static Optional<List<Integer>> show(Window owner, List<String> headers) {
        if (headers == null || headers.isEmpty()) {
            return Optional.empty();
        }
        int n = headers.size();
        ObservableList<Integer> perm = FXCollections.observableArrayList();
        for (int i = 0; i < n; i++) {
            perm.add(i);
        }

        Dialog<List<Integer>> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.WINDOW_MODAL);
        dialog.setTitle("列の並べ替え");
        DialogPane pane = dialog.getDialogPane();
        pane.getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);

        ListView<Integer> list = new ListView<>(perm);
        list.setPrefHeight(Math.min(420, 28 * n + 40));
        list.setCellFactory(
                lv ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(Integer oldIdx, boolean empty) {
                                super.updateItem(oldIdx, empty);
                                if (empty || oldIdx == null) {
                                    setText(null);
                                } else {
                                    String t =
                                            oldIdx < headers.size() && headers.get(oldIdx) != null
                                                    ? headers.get(oldIdx)
                                                    : "";
                                    setText((oldIdx + 1) + ": " + t);
                                }
                            }
                        });

        Button up = new Button("↑ 上へ");
        Button down = new Button("↓ 下へ");
        up.setMaxWidth(Double.MAX_VALUE);
        down.setMaxWidth(Double.MAX_VALUE);
        up.setOnAction(
                e -> {
                    int i = list.getSelectionModel().getSelectedIndex();
                    if (i > 0) {
                        Integer a = perm.get(i - 1);
                        Integer b = perm.get(i);
                        perm.set(i - 1, b);
                        perm.set(i, a);
                        list.getSelectionModel().select(i - 1);
                    }
                });
        down.setOnAction(
                e -> {
                    int i = list.getSelectionModel().getSelectedIndex();
                    if (i >= 0 && i < perm.size() - 1) {
                        Integer a = perm.get(i);
                        Integer b = perm.get(i + 1);
                        perm.set(i, b);
                        perm.set(i + 1, a);
                        list.getSelectionModel().select(i + 1);
                    }
                });

        Label hint =
                new Label(
                        "表示順は左から。↑↓で入れ替え。");
        hint.setWrapText(true);

        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(8));
        GridPane.setHgrow(list, Priority.ALWAYS);
        GridPane.setVgrow(list, Priority.ALWAYS);
        grid.add(hint, 0, 0, 2, 1);
        grid.add(list, 0, 1);
        VBox side = new VBox(6, up, down);
        grid.add(side, 1, 1);

        pane.setContent(new VBox(8, grid));

        dialog.setResultConverter(
                btn -> {
                    if (btn == ButtonType.OK) {
                        return new ArrayList<>(perm);
                    }
                    return null;
                });

        return dialog.showAndWait();
    }
}
