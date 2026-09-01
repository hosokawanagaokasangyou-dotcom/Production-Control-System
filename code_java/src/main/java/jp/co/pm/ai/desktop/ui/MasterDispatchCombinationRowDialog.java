package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Optional;

import javafx.geometry.Insets;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.layout.GridPane;
import javafx.stage.Window;
import javafx.util.StringConverter;

/** 組み合わせ表へ工程名+機械名の行を追加する。 */
public final class MasterDispatchCombinationRowDialog {

    private MasterDispatchCombinationRowDialog() {}

    public record Result(String process, String machine, boolean autoFillSkillMembers) {}

    public static Optional<Result> prompt(Window owner, List<String[]> equipmentPairs) {
        Dialog<Result> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle("組み合わせ行を追加");
        dialog.setHeaderText("skills にある工程名+機械名で、組み合わせ表に行を追加します。既存の工程+機械は重複追加しません。");

        List<String[]> pairs = equipmentPairs != null ? equipmentPairs : List.of();
        ComboBox<String[]> combo = new ComboBox<>();
        combo.getItems().addAll(pairs);
        combo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(String[] v) {
                        if (v == null || v.length < 2) {
                            return "";
                        }
                        return v[0] + " / " + v[1];
                    }

                    @Override
                    public String[] fromString(String s) {
                        return null;
                    }
                });
        combo.setMaxWidth(Double.MAX_VALUE);
        if (!pairs.isEmpty()) {
            combo.getSelectionModel().selectFirst();
        }

        CheckBox autoFill =
                new CheckBox("スキル（OP/AS）のあるメンバーを自動で入れる");
        autoFill.setSelected(true);
        autoFill.setWrapText(true);
        autoFill.setMaxWidth(Double.MAX_VALUE);

        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(12, 16, 8, 16));
        grid.add(new Label("工程名 / 機械名"), 0, 0);
        grid.add(combo, 1, 0);
        grid.add(autoFill, 0, 1, 2, 1);
        dialog.getDialogPane().setContent(grid);

        ButtonType ok = new ButtonType("追加", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(ok, ButtonType.CANCEL);
        dialog.getDialogPane().lookupButton(ok).setDisable(pairs.isEmpty());
        dialog.setResultConverter(
                btn -> {
                    if (btn != ok) {
                        return null;
                    }
                    String[] sel = combo.getSelectionModel().getSelectedItem();
                    if (sel == null || sel.length < 2) {
                        return null;
                    }
                    String p = sel[0] != null ? sel[0].strip() : "";
                    String m = sel[1] != null ? sel[1].strip() : "";
                    if (p.isEmpty() || m.isEmpty()) {
                        return null;
                    }
                    return new Result(p, m, autoFill.isSelected());
                });
        return dialog.showAndWait();
    }
}
