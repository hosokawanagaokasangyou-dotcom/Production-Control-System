package jp.co.pm.ai.desktop.ui;

import java.util.concurrent.ThreadLocalRandom;
import java.util.function.UnaryOperator;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

/** 誤操作防止: 表示した4桁数字の入力を要求する確認ダイアログ。 */
public final class FourDigitConfirmationDialog {

    private FourDigitConfirmationDialog() {}

    public static boolean confirm(Stage owner, String title, String message) {
        return confirm(owner, title, message, "実行");
    }

    /**
     * @param owner 親ステージ（null 可）
     * @param title ダイアログタイトル
     * @param message 警告本文
     * @param confirmButtonText OK ボタン文言
     * @return 番号が一致して OK されたとき true
     */
    public static boolean confirm(
            Stage owner, String title, String message, String confirmButtonText) {
        int code = ThreadLocalRandom.current().nextInt(9000) + 1000;
        String expected = Integer.toString(code);

        Stage stage = new Stage(StageStyle.DECORATED);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle(title);

        Label messageLabel = new Label(message);
        messageLabel.setWrapText(true);

        Label codeHint = new Label("確認のため、下記の4桁を入力してください:");
        Label codeLabel = new Label(expected);
        codeLabel.getStyleClass().add("pm-four-digit-confirmation-code");

        TextField input = new TextField();
        input.setPromptText("4桁の数字");
        input.setMaxWidth(Double.MAX_VALUE);
        UnaryOperator<TextFormatter.Change> filter =
                change -> {
                    String t = change.getControlNewText();
                    if (t.isEmpty()) {
                        return change;
                    }
                    if (!t.matches("\\d{0,4}")) {
                        return null;
                    }
                    return change;
                };
        input.setTextFormatter(new TextFormatter<>(filter));

        final boolean[] confirmed = {false};

        String okLabel =
                confirmButtonText != null && !confirmButtonText.isBlank()
                        ? confirmButtonText
                        : "実行";
        Button okBtn = new Button(okLabel);
        okBtn.setDefaultButton(true);
        okBtn.setDisable(true);
        okBtn.setOnAction(
                e -> {
                    confirmed[0] = true;
                    stage.close();
                });

        Button cancelBtn = new Button("キャンセル");
        cancelBtn.setCancelButton(true);
        cancelBtn.setOnAction(e -> stage.close());

        input.textProperty()
                .addListener(
                        (obs, o, n) ->
                                okBtn.setDisable(
                                        n == null || !expected.equals(n.trim())));

        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox buttons = new HBox(8, spacer, okBtn, cancelBtn);
        buttons.setAlignment(Pos.CENTER_RIGHT);

        VBox root = new VBox(10, messageLabel, codeHint, codeLabel, input, buttons);
        root.setPadding(new Insets(14));
        root.setPrefWidth(400);
        stage.setScene(new javafx.scene.Scene(root));
        stage.showAndWait();
        return confirmed[0];
    }
}
