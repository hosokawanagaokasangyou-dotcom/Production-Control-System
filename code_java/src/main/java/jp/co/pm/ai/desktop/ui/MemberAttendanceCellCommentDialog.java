package jp.co.pm.ai.desktop.ui;

import java.util.function.Consumer;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

/** メンバー勤怠セル向けコメント入力ダイアログ。 */
public final class MemberAttendanceCellCommentDialog {

    private MemberAttendanceCellCommentDialog() {}

    public static void show(
            Stage owner,
            String member,
            String dateKey,
            String initialComment,
            Consumer<String> onSave) {
        Stage stage = new Stage(StageStyle.DECORATED);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle(member + " — " + dateKey + " コメント");

        Label hint = new Label("セルに付けるメモ（占有・勤怠区分とは別に保存されます）");
        hint.setWrapText(true);

        TextArea area = new TextArea(initialComment != null ? initialComment : "");
        area.setWrapText(true);
        area.setPrefRowCount(5);
        VBox.setVgrow(area, Priority.ALWAYS);

        Button deleteBtn = new Button("コメントを削除");
        deleteBtn.setDisable(initialComment == null || initialComment.isBlank());
        deleteBtn.setOnAction(
                e -> {
                    if (onSave != null) {
                        onSave.accept("");
                    }
                    stage.close();
                });

        Button okBtn = new Button("OK");
        okBtn.setDefaultButton(true);
        okBtn.setOnAction(
                e -> {
                    if (onSave != null) {
                        onSave.accept(area.getText());
                    }
                    stage.close();
                });

        Button cancelBtn = new Button("キャンセル");
        cancelBtn.setCancelButton(true);
        cancelBtn.setOnAction(e -> stage.close());

        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox buttons = new HBox(8, deleteBtn, spacer, okBtn, cancelBtn);
        buttons.setAlignment(Pos.CENTER_RIGHT);

        VBox root = new VBox(10, hint, area, buttons);
        root.setPadding(new Insets(14));
        root.setPrefWidth(420);
        root.setPrefHeight(220);
        stage.setScene(new javafx.scene.Scene(root));
        stage.showAndWait();
    }
}
