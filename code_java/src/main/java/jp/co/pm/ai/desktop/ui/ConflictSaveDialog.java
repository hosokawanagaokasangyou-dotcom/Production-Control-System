package jp.co.pm.ai.desktop.ui;

import java.util.Optional;

import javafx.geometry.Insets;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.layout.VBox;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.io.conflict.ConflictSaveChoice;

/** 保存前の外部変更競合ダイアログ（強調ヘッダ + スクロール要約）。 */
public final class ConflictSaveDialog {

    private ConflictSaveDialog() {}

    public static ConflictSaveChoice show(Window owner, String screenTitle, String summary) {
        Dialog<ConflictSaveChoice> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle("外部変更の検出");
        dialog.setHeaderText(
                "⚠ 競合 — "
                        + (screenTitle != null && !screenTitle.isBlank() ? screenTitle : "保存先")
                        + " が他で変更されています");

        Label caption = new Label("変更の要約");
        TextArea area = new TextArea(summary != null ? summary : "");
        area.setEditable(false);
        area.setWrapText(true);
        area.setPrefRowCount(12);
        area.setPrefWidth(560);
        Label note = new Label("「再読込して更新」を選ぶと、いまの未保存編集は破棄されます。");
        note.setWrapText(true);

        VBox body = new VBox(8, caption, area, note);
        body.setPadding(new Insets(8));
        dialog.getDialogPane().setContent(body);
        dialog.getDialogPane().setStyle("-fx-background-color: #fff8f8;");

        ButtonType reload = new ButtonType("再読込して更新", ButtonBar.ButtonData.LEFT);
        ButtonType overwrite = new ButtonType("強制上書き", ButtonBar.ButtonData.OTHER);
        ButtonType cancel = new ButtonType("キャンセル", ButtonBar.ButtonData.CANCEL_CLOSE);
        dialog.getDialogPane().getButtonTypes().setAll(reload, overwrite, cancel);
        dialog.setResultConverter(
                bt -> {
                    if (bt == reload) {
                        return ConflictSaveChoice.RELOAD;
                    }
                    if (bt == overwrite) {
                        return ConflictSaveChoice.OVERWRITE;
                    }
                    return ConflictSaveChoice.CANCEL;
                });
        Optional<ConflictSaveChoice> ans = dialog.showAndWait();
        return ans.orElse(ConflictSaveChoice.CANCEL);
    }
}
