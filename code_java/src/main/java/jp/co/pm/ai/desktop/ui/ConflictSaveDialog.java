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
        dialog.setTitle("保存の確認");
        dialog.setHeaderText(
                (screenTitle != null && !screenTitle.isBlank() ? screenTitle : "保存先")
                        + " を保存する前に、ファイルの違いを確認してください");

        Label caption = new Label("いまの状態");
        caption.setStyle("-fx-text-fill: #1f2937;");
        String guide =
                "\n\nボタンの意味\n"
                        + "・再読込して更新 … 画面の未保存の編集を捨て、今のファイルを表示します。保存はしません。\n"
                        + "・強制上書き … 今の画面の内容でファイルを上書きします。ファイル側の変更は残しません。\n"
                        + "・キャンセル … 保存を中止します。画面もファイルもそのままです。";
        TextArea area = new TextArea((summary != null ? summary : "") + guide);
        area.setEditable(false);
        area.setWrapText(true);
        area.setPrefRowCount(14);
        area.setPrefWidth(560);
        area.setStyle("-fx-control-inner-background: #ffffff; -fx-text-fill: #111827;");
        Label note = new Label("迷ったときは「キャンセル」です。画面の編集を残したまま、保存だけ止めます。");
        note.setWrapText(true);
        note.setStyle("-fx-text-fill: #1f2937;");

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
