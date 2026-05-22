package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TableView;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.TableRow;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo;

/**
 * 段階2直前: 加工途中タスクの翌日配台量 (m) を表形式で一括入力する。
 */
public final class Stage2InProgressNextDayDispatchDialog {

    private static final double M_EPS = 1e-6;

    private Stage2InProgressNextDayDispatchDialog() {}

    public static final class Row {
        private final String taskId;
        private final String process;
        private final String machineName;
        private final double actualDoneM;
        private final double remainingM;
        private final StringProperty nextDayDispatchM = new SimpleStringProperty();

        public Row(
                String taskId,
                String process,
                String machineName,
                double actualDoneM,
                double remainingM,
                double defaultNextDayM) {
            this.taskId = taskId != null ? taskId : "";
            this.process = process != null ? process : "";
            this.machineName = machineName != null ? machineName : "";
            this.actualDoneM = actualDoneM;
            this.remainingM = remainingM;
            this.nextDayDispatchM.set(formatM(defaultNextDayM));
        }

        public String taskId() {
            return taskId;
        }

        public String process() {
            return process;
        }

        public String machineName() {
            return machineName;
        }

        public double actualDoneM() {
            return actualDoneM;
        }

        public double remainingM() {
            return remainingM;
        }

        public StringProperty nextDayDispatchMProperty() {
            return nextDayDispatchM;
        }

        Stage2InProgressNextDayDispatchIo.Entry toEntry(double nextDayM) {
            return new Stage2InProgressNextDayDispatchIo.Entry(
                    taskId, process, machineName, nextDayM);
        }
    }

    /**
     * @return 確定時は各行の入力値。キャンセル時は empty。
     */
    public static Optional<List<Stage2InProgressNextDayDispatchIo.Entry>> prompt(
            Window owner, List<Row> rows) {
        if (rows == null || rows.isEmpty()) {
            return Optional.of(List.of());
        }

        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle("加工途中タスク — 翌日の配台量");
        dialog.setHeaderText(
                "実加工数が入っている行について、翌日に配台する量 (m) を指定してください。"
                        + " 0 の行は段階2の配台対象から外します。");

        Label hint =
                new Label(
                        "列は依頼NO・機械名・実績・残量・翌日配台のみです。"
                                + " 初期値は配台使用残数量（残量）です。"
                                + " 翌日配台は 0 以上・残量以下。OK で未確定の入力も反映します。");
        hint.setWrapText(true);
        hint.setStyle("-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 22%);");

        TableView<Row> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        table.setEditable(true);
        table.setPrefHeight(Math.min(420, 56 + rows.size() * 28.0));

        TableColumn<Row, String> cTask = new TableColumn<>("依頼NO");
        cTask.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().taskId()));
        cTask.setEditable(false);
        cTask.setPrefWidth(88);

        TableColumn<Row, String> cMach = new TableColumn<>("機械名");
        cMach.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().machineName()));
        cMach.setEditable(false);
        cMach.setPrefWidth(140);

        TableColumn<Row, String> cDone = new TableColumn<>("実加工");
        cDone.setCellValueFactory(
                cd -> new SimpleStringProperty(formatM(cd.getValue().actualDoneM()) + " m"));
        cDone.setEditable(false);
        cDone.setStyle("-fx-alignment: CENTER-RIGHT;");
        cDone.setPrefWidth(72);

        TableColumn<Row, String> cRem = new TableColumn<>("残量");
        cRem.setCellValueFactory(
                cd -> new SimpleStringProperty(formatM(cd.getValue().remainingM()) + " m"));
        cRem.setEditable(false);
        cRem.setStyle("-fx-alignment: CENTER-RIGHT;");
        cRem.setPrefWidth(72);

        TableColumn<Row, String> cNext = new TableColumn<>("翌日配台 (m)");
        cNext.setCellValueFactory(cd -> cd.getValue().nextDayDispatchMProperty());
        cNext.setCellFactory(TextFieldTableCell.forTableColumn());
        cNext.setOnEditCommit(
                ev -> {
                    if (ev.getNewValue() != null) {
                        ev.getRowValue().nextDayDispatchMProperty().set(ev.getNewValue());
                    }
                });
        cNext.setEditable(true);
        cNext.setPrefWidth(96);

        table.getColumns().setAll(cTask, cMach, cDone, cRem, cNext);

        VBox content = new VBox(10, hint, table);
        VBox.setVgrow(table, Priority.ALWAYS);
        content.setPadding(new Insets(4, 0, 0, 0));
        dialog.getDialogPane().setContent(content);
        dialog.getDialogPane().setPrefWidth(520);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.getDialogPane()
                .lookupButton(ButtonType.OK)
                .addEventFilter(
                        javafx.event.ActionEvent.ACTION,
                        ev -> {
                            commitPendingTableCellEdit(table);
                            for (Row r : rows) {
                                Optional<Double> parsed = parseMeters(r.nextDayDispatchMProperty().get());
                                if (parsed.isEmpty()) {
                                    ev.consume();
                                    showParseError(dialog, r);
                                    return;
                                }
                                double next = parsed.get();
                                if (next < -M_EPS) {
                                    ev.consume();
                                    showNegativeError(dialog, r);
                                    return;
                                }
                                if (next > r.remainingM() + M_EPS) {
                                    ev.consume();
                                    showExceedsRemainingError(dialog, r, next);
                                    return;
                                }
                            }
                        });

        Optional<ButtonType> result = dialog.showAndWait();
        if (result.isEmpty() || result.get() != ButtonType.OK) {
            return Optional.empty();
        }
        List<Stage2InProgressNextDayDispatchIo.Entry> out = new ArrayList<>(rows.size());
        for (Row r : rows) {
            double next =
                    Stage2InProgressNextDayDispatchIo.sanitizeMeters(
                            parseMeters(r.nextDayDispatchMProperty().get()).orElse(0.0));
            out.add(r.toEntry(Math.max(0.0, next)));
        }
        return Optional.of(out);
    }

    /**
     * 編集中セルが Enter 未押下でも OK 時に {@link Row#nextDayDispatchMProperty()} へ反映する。
     */
    private static void commitPendingTableCellEdit(TableView<Row> table) {
        TablePosition<Row, ?> editing = table.getEditingCell();
        if (editing == null) {
            return;
        }
        int rowIdx = editing.getRow();
        if (rowIdx < 0 || rowIdx >= table.getItems().size()) {
            table.edit(-1, null);
            return;
        }
        Row row = table.getItems().get(rowIdx);
        String committed = null;
        if (table.getScene() != null) {
            Node focus = table.getScene().getFocusOwner();
            if (focus instanceof TextField tf) {
                committed = tf.getText();
            }
        }
        if (committed == null) {
            for (Node node : table.lookupAll(".table-row-cell")) {
                if (node instanceof TableRow<?> tr && tr.getIndex() == rowIdx) {
                    TextField tf = findTextFieldIn(tr);
                    if (tf != null) {
                        committed = tf.getText();
                        break;
                    }
                }
            }
        }
        if (committed != null) {
            row.nextDayDispatchMProperty().set(committed);
        }
        table.edit(-1, null);
    }

    private static TextField findTextFieldIn(Parent parent) {
        for (Node child : parent.getChildrenUnmodifiable()) {
            if (child instanceof TextField tf) {
                return tf;
            }
            if (child instanceof Parent p) {
                TextField nested = findTextFieldIn(p);
                if (nested != null) {
                    return nested;
                }
            }
        }
        return null;
    }

    private static Optional<Double> parseMeters(String raw) {
        if (raw == null || raw.isBlank()) {
            return Optional.of(0.0);
        }
        try {
            return Optional.of(Double.parseDouble(raw.strip().replace(",", "")));
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }

    private static void showParseError(Dialog<?> dialog, Row row) {
        showValidationError(
                dialog,
                "入力エラー",
                "翌日配台 (m) に数値を入力してください",
                "依頼NO "
                        + row.taskId()
                        + " / "
                        + row.machineName()
                        + " の値が不正です。");
    }

    private static void showNegativeError(Dialog<?> dialog, Row row) {
        showValidationError(
                dialog,
                "入力エラー",
                "翌日配台 (m) は 0 以上で入力してください",
                "依頼NO "
                        + row.taskId()
                        + " / "
                        + row.machineName()
                        + " の値が負です。");
    }

    private static void showExceedsRemainingError(Dialog<?> dialog, Row row, double nextM) {
        showValidationError(
                dialog,
                "整合性エラー",
                "翌日配台 (m) は残量以下にしてください",
                String.format(
                        Locale.ROOT,
                        "依頼NO %s / %s: 翌日配台 %s m、残量 %s m",
                        row.taskId(),
                        row.machineName(),
                        formatM(nextM),
                        formatM(row.remainingM())));
    }

    private static void showValidationError(
            Dialog<?> dialog, String title, String header, String content) {
        Dialog<Void> err = new Dialog<>();
        err.initOwner(dialog.getDialogPane().getScene().getWindow());
        err.initModality(Modality.WINDOW_MODAL);
        err.setTitle(title);
        err.setHeaderText(header);
        err.setContentText(content);
        err.getDialogPane().getButtonTypes().setAll(ButtonType.OK);
        err.showAndWait();
    }

    private static String formatM(double v) {
        if (Math.abs(v - Math.rint(v)) <= 1e-9) {
            return String.valueOf((long) Math.rint(v));
        }
        return String.format(java.util.Locale.ROOT, "%.3f", v).replaceAll("0+$", "").replaceAll("\\.$", "");
    }
}
