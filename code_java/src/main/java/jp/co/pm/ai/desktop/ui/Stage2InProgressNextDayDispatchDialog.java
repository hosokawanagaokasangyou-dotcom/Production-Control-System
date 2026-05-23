package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.dispatch.DispatchInteractiveRollUnitSupport;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.Stage2InProgressNextDayRollInput;
import jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;

/**
 * 段階2直前: 加工途中タスクの翌日配台量をロール本数で一括入力する。
 */
public final class Stage2InProgressNextDayDispatchDialog {

    private Stage2InProgressNextDayDispatchDialog() {}

    public static final class Row {
        private final String taskId;
        private final String process;
        private final String machineName;
        private final double actualDoneM;
        private final double remainingM;
        private final Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo;
        private final StringProperty nextDayRollCount = new SimpleStringProperty();

        public Row(
                String taskId,
                String process,
                String machineName,
                double actualDoneM,
                double remainingM,
                Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo) {
            this.taskId = taskId != null ? taskId : "";
            this.process = process != null ? process : "";
            this.machineName = machineName != null ? machineName : "";
            this.actualDoneM = actualDoneM;
            this.remainingM = remainingM;
            this.unitInfo =
                    unitInfo != null
                            ? unitInfo
                            : new Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM(
                                    0.0, 0.0, 0.0, false);
            int defaultRolls =
                    Stage2InProgressNextDayRollInput.defaultRollCount(
                            remainingM, this.unitInfo.unitM());
            this.nextDayRollCount.set(String.valueOf(defaultRolls));
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

        public Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo() {
            return unitInfo;
        }

        public double unitM() {
            return unitInfo.unitM();
        }

        public int maxRolls() {
            return Stage2InProgressNextDayRollInput.maxRolls(remainingM, unitM());
        }

        public StringProperty nextDayRollCountProperty() {
            return nextDayRollCount;
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
                "実加工数が入っている行について、翌日に配台するロール数を指定してください。"
                        + " 0 の行は段階2の配台対象から外します。");

        Label hint =
                new Label(
                        "配台計画手動修正タブと同様、配台ロール単位 (m) の整数倍で配台します。"
                                + " 初期値は残量以内の最大ロール本数です。"
                                + " 翌日配台は 0 以上・残量に収まるロール整数倍のみ。OK で未確定の入力も反映します。");
        hint.setWrapText(true);
        hint.setStyle("-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 22%);");

        TableView<Row> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        table.setEditable(true);
        table.setPrefHeight(Math.min(420, 56 + rows.size() * 28.0));

        TableColumn<Row, String> cTask = new TableColumn<>("依頼NO");
        cTask.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().taskId()));
        cTask.setEditable(false);
        cTask.setPrefWidth(80);

        TableColumn<Row, String> cMach = new TableColumn<>("機械名");
        cMach.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().machineName()));
        cMach.setEditable(false);
        cMach.setPrefWidth(120);

        TableColumn<Row, String> cDone = new TableColumn<>("実加工");
        cDone.setCellValueFactory(
                cd -> new SimpleStringProperty(formatM(cd.getValue().actualDoneM()) + " m"));
        cDone.setEditable(false);
        cDone.setStyle("-fx-alignment: CENTER-RIGHT;");
        cDone.setPrefWidth(68);

        TableColumn<Row, String> cRem = new TableColumn<>("残量");
        cRem.setCellValueFactory(
                cd -> new SimpleStringProperty(formatM(cd.getValue().remainingM()) + " m"));
        cRem.setEditable(false);
        cRem.setStyle("-fx-alignment: CENTER-RIGHT;");
        cRem.setPrefWidth(68);

        TableColumn<Row, String> cUnit = new TableColumn<>("1ロール");
        cUnit.setCellValueFactory(
                cd -> {
                    Row r = cd.getValue();
                    if (r.unitM() <= 1e-9) {
                        return new SimpleStringProperty("—");
                    }
                    return new SimpleStringProperty(
                            ResultDispatchNormalizer.formatQty(r.unitM()) + " m");
                });
        cUnit.setEditable(false);
        cUnit.setStyle("-fx-alignment: CENTER-RIGHT;");
        cUnit.setPrefWidth(64);

        TableColumn<Row, String> cRolls = new TableColumn<>("翌日(ロール)");
        cRolls.setCellValueFactory(cd -> cd.getValue().nextDayRollCountProperty());
        cRolls.setCellFactory(TextFieldTableCell.forTableColumn());
        cRolls.setOnEditCommit(
                ev -> {
                    if (ev.getNewValue() != null) {
                        ev.getRowValue().nextDayRollCountProperty().set(ev.getNewValue());
                    }
                });
        cRolls.setEditable(true);
        cRolls.setStyle("-fx-alignment: CENTER-RIGHT;");
        cRolls.setPrefWidth(72);

        TableColumn<Row, String> cMeters = new TableColumn<>("換算(m)");
        cMeters.setCellValueFactory(
                cd -> {
                    Row r = cd.getValue();
                    Optional<Integer> rolls =
                            Stage2InProgressNextDayRollInput.parseNonNegativeRollCount(
                                    r.nextDayRollCountProperty().get());
                    int n = rolls.orElse(0);
                    return new SimpleStringProperty(
                            Stage2InProgressNextDayRollInput.formatConvertedMetersPreview(
                                    n, r.unitM()));
                });
        cMeters.setEditable(false);
        cMeters.setStyle("-fx-alignment: CENTER-RIGHT;");
        cMeters.setPrefWidth(72);

        table.getColumns().setAll(cTask, cMach, cDone, cRem, cUnit, cRolls, cMeters);
        table.getItems().forEach(r -> r.nextDayRollCountProperty().addListener((o, a, b) -> table.refresh()));

        VBox content = new VBox(10, hint, table);
        VBox.setVgrow(table, Priority.ALWAYS);
        content.setPadding(new Insets(4, 0, 0, 0));
        dialog.getDialogPane().setContent(content);
        dialog.getDialogPane().setPrefWidth(640);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.getDialogPane()
                .lookupButton(ButtonType.OK)
                .addEventFilter(
                        javafx.event.ActionEvent.ACTION,
                        ev -> {
                            commitPendingTableCellEdit(table);
                            for (Row r : rows) {
                                Optional<String> err =
                                        Stage2InProgressNextDayRollInput.validateRollInput(
                                                r.nextDayRollCountProperty().get(),
                                                r.remainingM(),
                                                r.unitInfo());
                                if (err.isPresent()) {
                                    ev.consume();
                                    showValidationError(
                                            dialog, "入力エラー", err.get(), rowDetail(r));
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
            int rolls =
                    Stage2InProgressNextDayRollInput.parseNonNegativeRollCount(
                                    r.nextDayRollCountProperty().get())
                            .orElse(0);
            double next =
                    Stage2InProgressNextDayRollInput.resolveNextDayMeters(
                                    rolls, r.remainingM(), r.unitM())
                            .orElse(0.0);
            out.add(r.toEntry(Math.max(0.0, next)));
        }
        return Optional.of(out);
    }

    private static String rowDetail(Row row) {
        String unitLine =
                row.unitM() > 1e-9
                        ? DispatchInteractiveRollUnitSupport.rollUnitDialogHeader(
                                row.remainingM(), row.unitInfo(), row.taskId() + " / " + row.machineName())
                        : "依頼NO "
                                + row.taskId()
                                + " / "
                                + row.machineName()
                                + "\n配台ロール単位 (m) を決定できません。";
        return unitLine;
    }

    /**
     * 編集中セルが Enter 未押下でも OK 時に {@link Row#nextDayRollCountProperty()} へ反映する。
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
            row.nextDayRollCountProperty().set(committed);
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
