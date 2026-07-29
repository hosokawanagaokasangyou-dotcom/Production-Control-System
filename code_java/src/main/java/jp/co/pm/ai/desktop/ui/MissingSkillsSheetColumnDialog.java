package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javafx.beans.property.SimpleStringProperty;
import javafx.geometry.Insets;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.PlanTasksMissingSkillsColumnPrompt.MissingPair;
import jp.co.pm.ai.desktop.PlanTasksMissingSkillsColumnPrompt.PromptBundle;

/**
 * master「skills」シートに工程+機械の列が無く配台できない組み合わせを表示する。
 */
public final class MissingSkillsSheetColumnDialog {

    private MissingSkillsSheetColumnDialog() {}

    public static final class Row {
        private final String process;
        private final String machine;
        private final String sampleTaskId;

        Row(MissingPair src) {
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
    }

    /**
     * @param allowContinue {@code true} のとき「続行」ボタンを出す（段階2実行前）。{@code false} は警告のみ（段階1完了時）。
     * @return 続行可能モードで「続行」を選んだとき {@code true}。警告のみモードでは OK で {@code true}。
     */
    public static Optional<Boolean> prompt(Window owner, PromptBundle bundle, boolean allowContinue) {
        if (bundle == null || bundle.empty()) {
            return Optional.empty();
        }
        List<Row> rows = new ArrayList<>();
        for (MissingPair p : bundle.pairs()) {
            rows.add(new Row(p));
        }

        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle(allowContinue ? "段階2 — skills シート未登録" : "段階1 — skills シート未登録");
        dialog.setHeaderText(
                "PM_AI_MASTER_WORKBOOK の「skills」シートに、次の工程名+機械名の列がありません。"
                        + " OP/AS スキルが割り当てられないため、段階2では配台されません。");

        Label hint =
                new Label(
                        allowContinue
                                ? "master の skills シート 1 行目（工程名）・2 行目（機械名）に列を追加し、"
                                        + "メンバー行へ OP/AS を設定してから段階2を実行してください。"
                                        + " 「続行」は配台されない行が残る可能性があります。"
                                : "段階2実行前に master の skills シートへ列を追加し、メンバー行へ OP/AS を設定してください。");
        hint.setWrapText(true);
        hint.setStyle("-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 22%);");

        TableView<Row> table = new TableView<>();
        table.getItems().addAll(rows);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        table.setEditable(false);
        table.setPrefHeight(Math.min(360, 56 + rows.size() * 28.0));

        TableColumn<Row, String> cProc = new TableColumn<>("工程名");
        cProc.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().process()));

        TableColumn<Row, String> cMach = new TableColumn<>("機械名");
        cMach.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().machine()));

        TableColumn<Row, String> cTask = new TableColumn<>("依頼NO例");
        cTask.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().sampleTaskId()));
        cTask.setPrefWidth(88);

        table.getColumns().setAll(cProc, cMach, cTask);

        ScrollPane scroll = new ScrollPane(table);
        scroll.setFitToWidth(true);
        VBox root = new VBox(10, hint, scroll);
        VBox.setVgrow(scroll, Priority.ALWAYS);
        root.setPadding(new Insets(4, 0, 0, 0));
        dialog.getDialogPane().setContent(root);
        dialog.getDialogPane().setPrefWidth(720);

        ButtonType continueBtn = new ButtonType("続行");
        if (allowContinue) {
            dialog.getDialogPane().getButtonTypes().setAll(continueBtn, ButtonType.CANCEL);
        } else {
            dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK);
        }

        Optional<ButtonType> choice = dialog.showAndWait();
        if (choice.isEmpty()) {
            return Optional.of(false);
        }
        if (allowContinue) {
            return Optional.of(choice.get() == continueBtn);
        }
        return Optional.of(choice.get() == ButtonType.OK);
    }
}
