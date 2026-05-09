package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

import javafx.geometry.Insets;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.DialogPane;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

/**
 * Per-column visibility editor for spreadsheet / {@link javafx.scene.control.TableView} columns (title-aligned).
 */
public final class ColumnVisibilityDialog {

    private ColumnVisibilityDialog() {}

    /**
     * @param columnTitles headers left-to-right (duplicate titles: last-wins on save)
     * @param visibleInitial aligned length to titles; {@code null} or length mismatch means all visible
     */
    public static Optional<boolean[]> show(
            Window owner, List<String> columnTitles, boolean[] visibleInitial) {
        if (columnTitles == null || columnTitles.isEmpty()) {
            return Optional.empty();
        }
        int n = columnTitles.size();
        boolean[] state = new boolean[n];
        Arrays.fill(state, true);
        if (visibleInitial != null && visibleInitial.length == n) {
            System.arraycopy(visibleInitial, 0, state, 0, n);
        }
        Dialog<boolean[]> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.initModality(Modality.WINDOW_MODAL);
        dialog.setTitle("\u5217\u306e\u8868\u793a");
        DialogPane pane = dialog.getDialogPane();
        pane.getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);

        List<CheckBox> boxes = new ArrayList<>(n);
        VBox box = new VBox(6);
        box.setPadding(new Insets(8));
        for (int i = 0; i < n; i++) {
            String t = columnTitles.get(i) != null ? columnTitles.get(i) : "";
            CheckBox cb = new CheckBox((i + 1) + ": " + t);
            cb.setSelected(state[i]);
            boxes.add(cb);
            box.getChildren().add(cb);
        }
        Label hint =
                new Label(
                        "\u8868\u793a\u3059\u308b\u5217\u306b\u30c1\u30a7\u30c3\u30af\u3092\u5165\u308c\u3066\u304f\u3060\u3055\u3044\u3002"
                                + " \u5c11\u306a\u304f\u3068\u30821\u5217\u306f\u8868\u793a\u3059\u308b\u5fc5\u8981\u304c\u3042\u308a\u307e\u3059\u3002");
        hint.setWrapText(true);
        ScrollPane sp = new ScrollPane(new VBox(8, hint, box));
        sp.setFitToWidth(true);
        sp.setPrefViewportHeight(Math.min(420, 28 * n + 80));
        pane.setContent(sp);

        Button okBtn = (Button) pane.lookupButton(ButtonType.OK);
        if (okBtn != null) {
            okBtn.addEventFilter(
                    javafx.event.ActionEvent.ACTION,
                    ev -> {
                        int visCount = 0;
                        for (CheckBox cb : boxes) {
                            if (cb.isSelected()) {
                                visCount++;
                            }
                        }
                        if (visCount == 0) {
                            ev.consume();
                            javafx.scene.control.Alert a =
                                    new javafx.scene.control.Alert(
                                            javafx.scene.control.Alert.AlertType.WARNING);
                            a.initOwner(owner);
                            a.setTitle("\u5217\u306e\u8868\u793a");
                            a.setHeaderText(null);
                            a.setContentText(
                                    "\u6700\u4f4e1\u5217\u306f\u8868\u793a\u3059\u308b\u5fc5\u8981\u304c\u3042\u308a\u307e\u3059\u3002");
                            a.showAndWait();
                        }
                    });
        }

        dialog.setResultConverter(
                btn -> {
                    if (btn != ButtonType.OK) {
                        return null;
                    }
                    boolean[] out = new boolean[n];
                    for (int i = 0; i < n; i++) {
                        out[i] = boxes.get(i).isSelected();
                    }
                    int cnt = 0;
                    for (boolean b : out) {
                        if (b) {
                            cnt++;
                        }
                    }
                    if (cnt == 0) {
                        return null;
                    }
                    return out;
                });

        return dialog.showAndWait();
    }
}
