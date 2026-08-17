package jp.co.pm.ai.desktop.ui;

import java.util.concurrent.atomic.AtomicReference;
import java.util.function.UnaryOperator;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;
import javafx.stage.WindowEvent;

/**
 * 終了ゲート用。12桁一致で終了確認へ進む。アプリ終了のキャンセルはできないが、確認へ戻れる。
 */
public final class SevenDigitChallengeDialog {

    public enum Outcome {
        CONFIRMED,
        RETURN_TO_CHECK
    }

    private SevenDigitChallengeDialog() {}

    public static Outcome showAndConfirm(Window owner, String code) {
        return showAndConfirm(owner, code, "", "");
    }

    public static Outcome showAndConfirm(
            Window owner, String code, String detail, String dialogBody) {
        String expected = code != null ? code : SevenDigitChallenge.generate();
        AtomicReference<Outcome> outcome = new AtomicReference<>();

        Stage stage = new Stage();
        stage.initModality(Modality.APPLICATION_MODAL);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.setTitle("配台計画と加工計画が不一致 — 終了確認");
        stage.setOnCloseRequest(WindowEvent::consume);

        Label warn =
                new Label(
                        "段階2後、配台計画とアラジン加工計画が揃っていません。\n"
                                + (detail != null && !detail.isBlank() ? "状態: " + detail + "\n" : "")
                                + "内容を確認してから終了してください。この窓の✕とEscは使えません。");
        warn.setWrapText(true);

        TextArea body = new TextArea(dialogBody != null ? dialogBody : "");
        body.setEditable(false);
        body.setWrapText(true);
        body.setPrefRowCount(8);
        body.setVisible(dialogBody != null && !dialogBody.isBlank());
        body.setManaged(body.isVisible());

        Label hint =
                new Label("納期管理ビューの「配台結果」で同一化チェック（ローカル最新）を実行できます。");
        hint.setWrapText(true);

        Label digits = new Label(expected);
        digits.setStyle("-fx-font-size: 28px; -fx-font-weight: bold;");

        TextField input = new TextField();
        input.setPromptText("12桁を入力");
        input.setPrefColumnCount(14);
        UnaryOperator<TextFormatter.Change> digitsOnly =
                change -> {
                    String next = change.getControlNewText();
                    if (next == null
                            || next.matches("[0-9]{0," + SevenDigitChallenge.DIGIT_COUNT + "}")) {
                        return change;
                    }
                    return null;
                };
        input.setTextFormatter(new TextFormatter<>(digitsOnly));

        Label mismatch = new Label("");
        mismatch.setStyle("-fx-text-fill: #b00020;");

        Button ok = new Button("12桁を入力して終了確認へ");
        ok.setDefaultButton(true);
        ok.setOnAction(
                e -> {
                    if (SevenDigitChallenge.matches(expected, input.getText())) {
                        outcome.set(Outcome.CONFIRMED);
                        stage.setOnCloseRequest(null);
                        stage.close();
                    } else {
                        input.clear();
                        mismatch.setText("数字が一致しません。");
                        input.requestFocus();
                    }
                });

        Button back = new Button("確認に戻る（アプリは閉じない）");
        back.setOnAction(
                e -> {
                    outcome.set(Outcome.RETURN_TO_CHECK);
                    stage.setOnCloseRequest(null);
                    stage.close();
                });

        HBox buttons = new HBox(10, back, ok);
        buttons.setAlignment(Pos.CENTER);

        VBox box = new VBox(12, warn, body, hint, digits, input, mismatch, buttons);
        box.setPadding(new Insets(16));
        box.setAlignment(Pos.CENTER);
        VBox.setVgrow(body, Priority.ALWAYS);
        Scene scene = new Scene(box, 720, 520);
        if (owner != null && owner.getScene() != null) {
            scene.getStylesheets().setAll(owner.getScene().getStylesheets());
        }
        scene.addEventFilter(
                KeyEvent.KEY_PRESSED,
                e -> {
                    if (e.getCode() == KeyCode.ESCAPE) {
                        e.consume();
                    }
                });
        stage.setScene(scene);
        input.requestFocus();
        stage.showAndWait();
        return outcome.get() != null ? outcome.get() : Outcome.RETURN_TO_CHECK;
    }
}
