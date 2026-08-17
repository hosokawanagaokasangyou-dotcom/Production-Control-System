package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

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
import javafx.scene.control.cell.ComboBoxTableCell;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.Stage1EcSideUnknownPrompt.PromptBundle;
import jp.co.pm.ai.desktop.Stage1EcSideUnknownPrompt.UnknownIrai;
import jp.co.pm.ai.desktop.reconciliation.EcSideClassification;

/**
 * 段階1完了後: EC面区分が「不明」の依頼NOについて、両面EC/片面ECをユーザーに選ばせる。
 */
public final class Stage1EcSideUnknownDialog {

    private Stage1EcSideUnknownDialog() {}

    public static final class Row {
        private final String iraiNo;
        private final StringProperty choice =
                new SimpleStringProperty(EcSideClassification.DOUBLE_SIDED);

        Row(UnknownIrai src) {
            this.iraiNo = src.iraiNo() != null ? src.iraiNo() : "";
        }

        public String iraiNo() {
            return iraiNo;
        }

        public StringProperty choiceProperty() {
            return choice;
        }

        Stage1EcSideUnknownDialogResult.Selection toSelection() {
            return new Stage1EcSideUnknownDialogResult.Selection(iraiNo, choice.get());
        }
    }

    public static Optional<Stage1EcSideUnknownDialogResult> prompt(
            Window owner, PromptBundle bundle) {
        if (bundle == null || bundle.empty()) {
            return Optional.empty();
        }
        List<Row> rows = new ArrayList<>();
        for (UnknownIrai item : bundle.items()) {
            rows.add(new Row(item));
        }

        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle("EC面区分の選択");
        dialog.setHeaderText("受注ファイルから EC面区分を判定できない依頼があります。");

        Label hint =
                new Label(
                        "各依頼NOについて「両面EC」または「片面EC」を選んで OK してください。"
                                + " キャンセルすると「不明」のまま残ります。");
        hint.setWrapText(true);

        TableView<Row> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setEditable(true);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

        TableColumn<Row, String> cIrai = new TableColumn<>("依頼NO");
        cIrai.setCellValueFactory(d -> new SimpleStringProperty(d.getValue().iraiNo()));
        cIrai.setEditable(false);
        cIrai.setPrefWidth(160);

        TableColumn<Row, String> cChoice = new TableColumn<>("EC面区分");
        cChoice.setCellValueFactory(d -> d.getValue().choiceProperty());
        cChoice.setCellFactory(
                col ->
                        new ComboBoxTableCell<>(
                                FXCollections.observableArrayList(
                                        EcSideClassification.DOUBLE_SIDED,
                                        EcSideClassification.SINGLE_SIDED)));
        cChoice.setEditable(true);
        cChoice.setPrefWidth(140);

        table.getColumns().setAll(cIrai, cChoice);

        ScrollPane scroll = new ScrollPane(table);
        scroll.setFitToWidth(true);
        VBox root = new VBox(10, hint, scroll);
        VBox.setVgrow(scroll, Priority.ALWAYS);
        root.setPadding(new Insets(4, 0, 0, 0));
        dialog.getDialogPane().setContent(root);
        dialog.getDialogPane().setPrefWidth(520);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);

        Optional<ButtonType> choice = dialog.showAndWait();
        if (choice.isEmpty() || choice.get() != ButtonType.OK) {
            return Optional.empty();
        }
        List<Stage1EcSideUnknownDialogResult.Selection> selections = new ArrayList<>();
        for (Row r : rows) {
            selections.add(r.toSelection());
        }
        return Optional.of(new Stage1EcSideUnknownDialogResult(selections));
    }
}
