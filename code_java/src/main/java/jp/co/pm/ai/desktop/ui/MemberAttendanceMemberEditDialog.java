package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Optional;

import javafx.geometry.Insets;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.stage.Window;

/** メンバー勤怠名簿の追加・編集ダイアログ。 */
public final class MemberAttendanceMemberEditDialog {

    public static final String ROLE_POST = "後加工";
    public static final String ROLE_LOGISTICS = "物流";

    private MemberAttendanceMemberEditDialog() {}

    public record Result(String name, String primaryRole) {}

    public static Optional<Result> showAdd(Window owner) {
        return show(owner, "メンバー追加", null, null);
    }

    public static Optional<Result> showEdit(
            Window owner, String currentName, String currentRole) {
        return show(owner, "メンバー編集", currentName, currentRole);
    }

    private static Optional<Result> show(
            Window owner,
            String title,
            String initialName,
            String initialRole) {
        Dialog<Result> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle(title);
        dialog.setHeaderText("氏名と主担当を入力してください。");

        TextField nameField = new TextField(initialName != null ? initialName : "");
        nameField.setPromptText("氏名");
        ComboBox<String> roleBox = new ComboBox<>();
        roleBox.getItems().addAll(List.of(ROLE_POST, ROLE_LOGISTICS));
        String role =
                initialRole != null && !initialRole.isBlank()
                        ? initialRole.strip()
                        : ROLE_POST;
        if (!role.equals(ROLE_LOGISTICS)) {
            role = ROLE_POST;
        }
        roleBox.getSelectionModel().select(role);

        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(12, 16, 8, 16));
        grid.add(new Label("氏名"), 0, 0);
        grid.add(nameField, 1, 0);
        grid.add(new Label("主担当"), 0, 1);
        grid.add(roleBox, 1, 1);
        dialog.getDialogPane().setContent(grid);

        ButtonType ok = new ButtonType("OK", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(ok, ButtonType.CANCEL);

        dialog.setResultConverter(
                btn -> {
                    if (btn != ok) {
                        return null;
                    }
                    String name = nameField.getText().trim();
                    if (name.isEmpty()) {
                        return null;
                    }
                    String selected =
                            roleBox.getSelectionModel().getSelectedItem();
                    return new Result(
                            name,
                            selected != null && selected.equals(ROLE_LOGISTICS)
                                    ? ROLE_LOGISTICS
                                    : ROLE_POST);
                });

        return dialog.showAndWait();
    }
}
