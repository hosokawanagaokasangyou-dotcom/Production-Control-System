package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Optional;

import javafx.application.Platform;
import javafx.event.ActionEvent;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

/** 「担当OP_限定」用の検索付き複数選択ダイアログ。 */
public final class LimitedOperatorChecklistDialog {

    private LimitedOperatorChecklistDialog() {}

    public static Optional<String> edit(
            Window owner,
            List<String> candidates,
            String currentValue,
            double anchorScreenX,
            double anchorScreenY) {
        List<String> initial = LimitedOperatorJsonCodec.decode(currentValue);
        LimitedOperatorSelectionModel model =
                new LimitedOperatorSelectionModel(candidates, initial);

        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.WINDOW_MODAL);
        dialog.setTitle("担当OP_限定");
        dialog.setHeaderText(null);

        TextField search = new TextField();
        search.setPromptText("メンバー名を検索");
        ListView<String> list = new ListView<>();
        list.setPrefHeight(360);
        Label selectionOrder = new Label();
        Label validationMessage = new Label();
        validationMessage.setWrapText(true);
        validationMessage.setStyle("-fx-text-fill: #b91c1c; -fx-font-weight: bold;");

        Runnable refreshStatus =
                () -> {
                    selectionOrder.setText(
                            model.selectedNames().isEmpty()
                                    ? "選択: なし"
                                    : "選択順: " + String.join(" → ", model.selectedNames()));
                    List<String> invalid = model.selectedOutOfCandidateNames();
                    validationMessage.setText(
                            invalid.isEmpty()
                                    ? ""
                                    : "資格外/候補外の既存値です。確定するにはチェック解除してください: "
                                            + String.join(", ", invalid));
                };
        Runnable refreshList =
                () -> {
                    list.getItems().setAll(model.filteredDisplayNames(search.getText()));
                    list.refresh();
                    refreshStatus.run();
                };
        list.setCellFactory(
                ignored ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(String name, boolean empty) {
                                super.updateItem(name, empty);
                                if (empty || name == null) {
                                    setGraphic(null);
                                    setText(null);
                                    return;
                                }
                                boolean candidate = model.isCandidate(name);
                                CheckBox check =
                                        new CheckBox(
                                                candidate
                                                        ? name
                                                        : name + "（資格外/候補外の既存値）");
                                check.setSelected(model.isSelected(name));
                                if (!candidate) {
                                    check.setStyle(
                                            "-fx-text-fill: #b91c1c; -fx-font-weight: bold;");
                                    check.setDisable(!check.isSelected());
                                }
                                check.selectedProperty()
                                        .addListener(
                                                (obs, before, selected) -> {
                                                    model.setSelected(name, selected);
                                                    if (!candidate && !selected) {
                                                        check.setDisable(true);
                                                    }
                                                    refreshStatus.run();
                                                });
                                setGraphic(check);
                                setText(null);
                            }
                        });

        Button selectAll = new Button("全選択");
        selectAll.setOnAction(
                e -> {
                    model.selectAll(model.filteredCandidates(search.getText()));
                    refreshList.run();
                });
        Button clearAll = new Button("全解除");
        clearAll.setOnAction(
                e -> {
                    model.clearAll();
                    refreshList.run();
                });
        HBox actions = new HBox(8, selectAll, clearAll);
        actions.setAlignment(Pos.CENTER_LEFT);

        search.textProperty().addListener((obs, before, after) -> refreshList.run());
        VBox content =
                new VBox(8, search, actions, list, selectionOrder, validationMessage);
        VBox.setVgrow(list, Priority.ALWAYS);
        dialog.getDialogPane().setContent(content);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.getDialogPane().setPrefSize(430, 520);
        dialog.getDialogPane()
                .lookupButton(ButtonType.OK)
                .addEventFilter(
                        ActionEvent.ACTION,
                        event -> {
                            try {
                                model.validateConfirmable();
                            } catch (IllegalStateException ex) {
                                validationMessage.setText(ex.getMessage());
                                event.consume();
                            }
                        });

        dialog.setOnShown(
                e ->
                        Platform.runLater(
                                () -> {
                                    Window window = dialog.getDialogPane().getScene().getWindow();
                                    window.setX(anchorScreenX - 60);
                                    window.setY(anchorScreenY - 80);
                                    search.requestFocus();
                                }));
        refreshList.run();

        Optional<ButtonType> result = dialog.showAndWait();
        if (result.isPresent() && result.get() == ButtonType.OK) {
            model.validateConfirmable();
            return Optional.of(LimitedOperatorJsonCodec.encode(model.selectedNames()));
        }
        return Optional.empty();
    }
}
