package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.util.List;
import java.util.Optional;

import javafx.geometry.Insets;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.layout.GridPane;
import javafx.stage.Window;

/** 職場異動したメンバーを復帰日以降に名簿・配台へ戻す。 */
public final class MemberAttendanceReturnDialog {

    private MemberAttendanceReturnDialog() {}

    public record Result(String name, LocalDate returnedOn, boolean cancelledTransfer) {}

    public static Optional<Result> show(
            Window owner,
            List<String> transferredNames,
            String initialName,
            LocalDate initialReturnedOn,
            LocalDate inactiveFrom) {
        if (transferredNames == null || transferredNames.isEmpty()) {
            return Optional.empty();
        }
        Dialog<Result> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle("職場復帰");
        dialog.setHeaderText("復帰日以降、この工場の名簿・配台に戻します。異動中の勤怠は残ります。");

        ComboBox<String> nameBox = new ComboBox<>();
        nameBox.getItems().addAll(transferredNames);
        String pick =
                initialName != null && transferredNames.contains(initialName)
                        ? initialName
                        : transferredNames.get(0);
        nameBox.getSelectionModel().select(pick);

        DatePicker datePicker = new DatePicker();
        datePicker.setValue(
                initialReturnedOn != null
                        ? initialReturnedOn
                        : defaultReturnDate(inactiveFrom));

        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(12, 16, 8, 16));
        grid.add(new Label("氏名"), 0, 0);
        grid.add(nameBox, 1, 0);
        grid.add(new Label("復帰日"), 0, 1);
        grid.add(datePicker, 1, 1);
        dialog.getDialogPane().setContent(grid);

        ButtonType ok = new ButtonType("OK", ButtonBar.ButtonData.OK_DONE);
        ButtonType cancelTransfer = new ButtonType("異動取消", ButtonBar.ButtonData.LEFT);
        dialog.getDialogPane().getButtonTypes().addAll(ok, cancelTransfer, ButtonType.CANCEL);

        dialog.setResultConverter(
                btn -> {
                    String name = nameBox.getSelectionModel().getSelectedItem();
                    if (name == null || name.isBlank()) {
                        return null;
                    }
                    if (btn == ok) {
                        LocalDate d = datePicker.getValue();
                        if (d == null) {
                            return null;
                        }
                        return new Result(name, d, false);
                    }
                    if (btn != null && "異動取消".equals(btn.getText())) {
                        return new Result(name, null, true);
                    }
                    return null;
                });

        return dialog.showAndWait();
    }

    private static LocalDate defaultReturnDate(LocalDate inactiveFrom) {
        LocalDate today = LocalDate.now();
        if (inactiveFrom != null && today.isBefore(inactiveFrom)) {
            return inactiveFrom;
        }
        return today;
    }
}
