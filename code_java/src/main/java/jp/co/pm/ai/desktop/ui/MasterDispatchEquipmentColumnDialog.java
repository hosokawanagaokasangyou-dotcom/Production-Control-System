package jp.co.pm.ai.desktop.ui;

import java.util.Optional;

import javafx.geometry.Insets;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.stage.Window;

/** 配台マスタ skills / need / speed へ工程名+機械名の列を追加する入力。 */
public final class MasterDispatchEquipmentColumnDialog {

    private MasterDispatchEquipmentColumnDialog() {}

    public record Result(String process, String machine) {}

    public static Optional<Result> prompt(Window owner) {
        Dialog<Result> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle("設備列を追加");
        dialog.setHeaderText("skills / need / speed に同じ工程名+機械名の列を追加します。");

        TextField processField = new TextField();
        processField.setPromptText("例: 分割");
        TextField machineField = new TextField();
        machineField.setPromptText("例: LAC/EC機");

        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(12, 16, 8, 16));
        grid.add(new Label("工程名"), 0, 0);
        grid.add(processField, 1, 0);
        grid.add(new Label("機械名"), 0, 1);
        grid.add(machineField, 1, 1);
        dialog.getDialogPane().setContent(grid);

        ButtonType ok = new ButtonType("追加", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(ok, ButtonType.CANCEL);
        dialog.setResultConverter(
                btn -> {
                    if (btn != ok) {
                        return null;
                    }
                    String p = processField.getText() != null ? processField.getText().strip() : "";
                    String m = machineField.getText() != null ? machineField.getText().strip() : "";
                    if (p.isEmpty() || m.isEmpty()) {
                        return null;
                    }
                    return new Result(p, m);
                });
        return dialog.showAndWait();
    }
}
