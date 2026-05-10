package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.DialogPane;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

/**
 * Per-column visibility editor for spreadsheet / {@link javafx.scene.control.TableView} columns (title-aligned).
 *
 * <p>Japanese UI strings use {@code \\uXXXX} so they survive toolchains where source encoding is not UTF-8.
 */
public final class ColumnVisibilityDialog {

    private ColumnVisibilityDialog() {}

    /**
     * @param columnTitles headers left-to-right (duplicate titles: last-wins on save)
     * @param visibleInitial aligned length to titles; {@code null} or length mismatch means all visible
     */
    public static Optional<boolean[]> show(
            Window owner, List<String> columnTitles, boolean[] visibleInitial) {
        return show(owner, columnTitles, visibleInitial, null);
    }

    /**
     * @param mandatoryVisible same length as titles; {@code true} means column must stay visible (checkbox
     *     disabled). {@code null} or length mismatch: same as {@link #show(Window, List, boolean[])} without
     *     locking.
     */
    public static Optional<boolean[]> show(
            Window owner,
            List<String> columnTitles,
            boolean[] visibleInitial,
            boolean[] mandatoryVisible) {
        Objects.requireNonNull(columnTitles, "columnTitles");
        if (columnTitles.isEmpty()) {
            return Optional.empty();
        }
        int n = columnTitles.size();
        boolean[] mandatory = normalizeMandatoryMask(mandatoryVisible, n);
        boolean[] state = new boolean[n];
        Arrays.fill(state, true);
        if (visibleInitial != null && visibleInitial.length == n) {
            System.arraycopy(visibleInitial, 0, state, 0, n);
        }
        if (mandatory != null) {
            for (int i = 0; i < n; i++) {
                if (mandatory[i]) {
                    state[i] = true;
                }
            }
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
        box.setPadding(new Insets(0));
        for (int i = 0; i < n; i++) {
            String t = columnTitles.get(i) != null ? columnTitles.get(i) : "";
            boolean locked = mandatory != null && mandatory[i];
            CheckBox cb =
                    new CheckBox(
                            (i + 1)
                                    + ": "
                                    + t
                                    + (locked
                                            ? " (\u30ed\u30b8\u30c3\u30af\u5fc5\u9801\u30fb\u975e\u8868\u793a\u306b\u3067\u304d\u307e\u305b\u3093)"
                                            : ""));
            cb.setSelected(state[i]);
            if (locked) {
                // Disable alone can fail to block toggles in some cases; keep selection pinned.
                cb.selectedProperty()
                        .addListener(
                                (obs, was, now) -> {
                                    if (Boolean.FALSE.equals(now)) {
                                        Platform.runLater(() -> cb.setSelected(true));
                                    }
                                });
                cb.setDisable(true);
            }
            boxes.add(cb);
            box.getChildren().add(cb);
        }

        Label searchCaption = new Label("\u5217\u540d:");
        TextField searchField = new TextField();
        searchField.setPromptText("\u5217\u540d\u3067\u691c\u7d22");
        HBox.setHgrow(searchField, Priority.ALWAYS);
        HBox searchRow = new HBox(8);
        searchRow.setAlignment(Pos.CENTER_LEFT);
        searchRow.getChildren().addAll(searchCaption, searchField);

        Button selectAllButton =
                new Button("\u5168\u3066\u306e\u5217\u3092\u9078\u629e");
        Button clearAllButton =
                new Button("\u5168\u3066\u306e\u5217\u306e\u9078\u629e\u3092\u89e3\u9664");
        HBox bulkRow = new HBox(8);
        bulkRow.setAlignment(Pos.CENTER_LEFT);
        bulkRow.getChildren().addAll(selectAllButton, clearAllButton);

        Runnable applySearchFilter =
                () -> {
                    String raw = searchField.getText();
                    for (int i = 0; i < n; i++) {
                        CheckBox cb = boxes.get(i);
                        String title = columnTitles.get(i) != null ? columnTitles.get(i) : "";
                        boolean show = columnTitleMatchesSearch(title, raw);
                        cb.setVisible(show);
                        cb.setManaged(show);
                    }
                };
        searchField.textProperty().addListener((obs, a, b) -> applySearchFilter.run());

        selectAllButton.setOnAction(
                e -> {
                    for (CheckBox cb : boxes) {
                        cb.setSelected(true);
                    }
                });
        clearAllButton.setOnAction(
                e -> {
                    for (int i = 0; i < boxes.size(); i++) {
                        CheckBox cb = boxes.get(i);
                        if (mandatory != null && i < mandatory.length && mandatory[i]) {
                            continue;
                        }
                        cb.setSelected(false);
                    }
                });

        Label hint =
                new Label(
                        "\u8868\u793a\u3059\u308b\u5217\u306b\u30c1\u30a7\u30c3\u30af\u3092\u5165\u308c\u3066\u304f\u3060\u3055\u3044\u3002"
                                + " \u5c11\u306a\u304f\u3068\u30821\u5217\u306f\u8868\u793a\u3059\u308b\u5fc5\u8981\u304c\u3042\u308a\u307e\u3059\u3002"
                                + (mandatory != null
                                        ? " \u30ed\u30b8\u30c3\u30af\u5fc5\u9801\u306e\u5217\u306f\u975e\u8868\u793a\u306b\u3067\u304d\u307e\u305b\u3093\u3002"
                                        : ""));
        hint.setWrapText(true);
        VBox scrollBody = new VBox(8, hint, box);
        ScrollPane sp = new ScrollPane(scrollBody);
        sp.setFitToWidth(true);
        sp.setPrefViewportHeight(Math.min(420, 28 * n + 80));

        VBox rootLayout = new VBox(8);
        rootLayout.setPadding(new Insets(8));
        rootLayout.getChildren().addAll(searchRow, bulkRow, sp);
        pane.setContent(rootLayout);

        applySearchFilter.run();

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
                        boolean sel = boxes.get(i).isSelected();
                        if (mandatory != null && mandatory[i]) {
                            sel = true;
                        }
                        out[i] = sel;
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

    /**
     * Pads or truncates {@code raw} to {@code columnCount}. {@code null} stays {@code null}; otherwise returns a
     * new array of length {@code columnCount} (missing indices are {@code false}).
     */
    static boolean[] normalizeMandatoryMask(boolean[] raw, int columnCount) {
        if (raw == null || columnCount <= 0) {
            return null;
        }
        boolean[] out = new boolean[columnCount];
        int copy = Math.min(columnCount, raw.length);
        System.arraycopy(raw, 0, out, 0, copy);
        return out;
    }

    /**
     * Empty or blank query shows all rows; otherwise substring match on title (case-folded for ASCII).
     */
    static boolean columnTitleMatchesSearch(String columnTitle, String queryRaw) {
        if (queryRaw == null || queryRaw.isBlank()) {
            return true;
        }
        String q = queryRaw.strip();
        String t = columnTitle != null ? columnTitle : "";
        if (t.contains(q)) {
            return true;
        }
        return t.toLowerCase(Locale.ROOT).contains(q.toLowerCase(Locale.ROOT));
    }
}
