package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.util.Optional;

import javafx.geometry.Insets;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.stage.Window;

/** メンバー勤怠の職場異動（異動日以降は名簿表示・配台から外す）。 */
public final class MemberAttendanceTransferDialog {

    private MemberAttendanceTransferDialog() {}

    public record Result(LocalDate inactiveFrom, boolean cancelledTransfer) {}

    public static Optional<Result> show(
            Window owner, String memberName, LocalDate currentInactiveFrom) {
        Dialog<Result> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle("職場異動");
        dialog.setHeaderText("異動日以降、この工場の名簿・配台から外れます。過去の勤怠は残ります。");

        TextField nameField = new TextField(memberName != null ? memberName : "");
        nameField.setEditable(false);
        DatePicker datePicker = new DatePicker(
                currentInactiveFrom != null ? currentInactiveFrom : LocalDate.now());

        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(12, 16, 8, 16));
        grid.add(new Label("氏名"), 0, 0);
        grid.add(nameField, 1, 0);
        grid.add(new Label("異動日"), 0, 1);
        grid.add(datePicker, 1, 1);
        dialog.getDialogPane().setContent(grid);

        ButtonType ok = new ButtonType("OK", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().add(ok);
        if (currentInactiveFrom != null) {
            ButtonType cancelTransfer =
                    new ButtonType("異動取消", ButtonBar.ButtonData.LEFT);
            dialog.getDialogPane().getButtonTypes().add(cancelTransfer);
        }
        dialog.getDialogPane().getButtonTypes().add(ButtonType.CANCEL);

        dialog.setResultConverter(
                btn -> {
                    if (btn == ok) {
                        LocalDate d = datePicker.getValue();
                        if (d == null) {
                            return null;
                        }
                        return new Result(d, false);
                    }
                    if (btn != null && "異動取消".equals(btn.getText())) {
                        return new Result(null, true);
                    }
                    return null;
                });

        return dialog.showAndWait();
    }
}
