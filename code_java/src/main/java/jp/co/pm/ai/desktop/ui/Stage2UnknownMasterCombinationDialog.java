package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.CheckBoxTableCell;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.Stage2UnknownMasterCombinationPrompt.PromptBundle;
import jp.co.pm.ai.desktop.Stage2UnknownMasterCombinationPrompt.UnknownPair;

/**
 * 段階2: master「組み合わせ表」に無い工程+機械について、配台不要にするかユーザーに確認する。
 */
public final class Stage2UnknownMasterCombinationDialog {

    private Stage2UnknownMasterCombinationDialog() {}

    public record Result(List<UnknownPair> markExclude) {}

    public static final class Row {
        private final String process;
        private final String machine;
        private final String sampleTaskId;
        private final BooleanProperty markExclude = new SimpleBooleanProperty(true);

        Row(UnknownPair src) {
            this.process = src.process();
            this.machine = src.machine();
            this.sampleTaskId = src.sampleTaskId() != null ? src.sampleTaskId() : "";
        }

        public String process() {
            return process;
        }

        public String machine() {
            return machine;
        }

        public String sampleTaskId() {
            return sampleTaskId;
        }

        public BooleanProperty markExcludeProperty() {
            return markExclude;
        }

        UnknownPair toPair() {
            return new UnknownPair(process, machine, sampleTaskId);
        }
    }

    public static Optional<Result> prompt(Window owner, PromptBundle bundle) {
        if (bundle == null || bundle.empty()) {
            return Optional.empty();
        }
        List<Row> rows = new ArrayList<>();
        for (UnknownPair p : bundle.pairs()) {
            rows.add(new Row(p));
        }

        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle("段階2 — マスタ未登録の工程+機械");
        dialog.setHeaderText(
                "PM_AI_MASTER_WORKBOOK の「組み合わせ表」に存在しない工程名+機械名があります。"
                        + " 配台不要の可能性が高い行にチェックを付け、OK で配台不要 JSON と計画タスク入力へ反映します。");

        Label hint =
                new Label(
                        "「配台不要にする」のチェックを外した行は段階2の配台対象のままです。"
                                + " キャンセルすると段階2実行を中止します。");
        hint.setWrapText(true);
        hint.setStyle("-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 22%);");

        TableView<Row> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        table.setEditable(true);
        table.setPrefHeight(Math.min(360, 56 + rows.size() * 28.0));

        TableColumn<Row, String> cProc = new TableColumn<>("工程名");
        cProc.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().process()));
        cProc.setEditable(false);

        TableColumn<Row, String> cMach = new TableColumn<>("機械名");
        cMach.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().machine()));
        cMach.setEditable(false);

        TableColumn<Row, String> cTask = new TableColumn<>("依頼NO例");
        cTask.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().sampleTaskId()));
        cTask.setEditable(false);
        cTask.setPrefWidth(88);

        TableColumn<Row, Boolean> cMark = new TableColumn<>("配台不要にする");
        cMark.setCellValueFactory(cd -> cd.getValue().markExcludeProperty());
        cMark.setCellFactory(CheckBoxTableCell.forTableColumn(cMark));
        cMark.setEditable(true);
        cMark.setPrefWidth(96);

        table.getColumns().setAll(cProc, cMach, cTask, cMark);

        ScrollPane scroll = new ScrollPane(table);
        scroll.setFitToWidth(true);
        VBox root = new VBox(10, hint, scroll);
        VBox.setVgrow(scroll, Priority.ALWAYS);
        root.setPadding(new Insets(4, 0, 0, 0));
        dialog.getDialogPane().setContent(root);
        dialog.getDialogPane().setPrefWidth(760);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);

        Optional<ButtonType> choice = dialog.showAndWait();
        if (choice.isEmpty() || choice.get() != ButtonType.OK) {
            return Optional.empty();
        }
        List<UnknownPair> marked = new ArrayList<>();
        for (Row r : rows) {
            if (r.markExcludeProperty().get()) {
                marked.add(r.toPair());
            }
        }
        return Optional.of(new Result(marked));
    }
}
