package jp.co.pm.ai.desktop.ui;

import java.util.concurrent.atomic.AtomicBoolean;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;
import javafx.stage.WindowEvent;

/**
 * 終了ゲート用。表示した7桁と一致するまで閉じられない。
 */
public final class SevenDigitChallengeDialog {

    private SevenDigitChallengeDialog() {}

    /**
     * 7桁を表示して入力を待つ。一致したら true。キャンセル不可。
     */
    public static boolean showAndConfirm(Window owner, String code) {
        String expected = code != null ? code : SevenDigitChallenge.generate();
        AtomicBoolean confirmed = new AtomicBoolean(false);

        Stage stage = new Stage();
        stage.initModality(Modality.APPLICATION_MODAL);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.setTitle("同一化未整合 — 終了確認");
        stage.setOnCloseRequest(WindowEvent::consume);

        Label warn =
                new Label(
                        "段階2後、配台計画と加工計画が同一ではありません。\n"
                                + "下の7桁を入力すると終了確認へ進めます。キャンセルはできません。");
        warn.setWrapText(true);

        Label digits = new Label(expected);
        digits.setStyle("-fx-font-size: 36px; -fx-font-weight: bold; -fx-letter-spacing: 4px;");

        TextField input = new TextField();
        input.setPromptText("7桁を入力");
        input.setPrefColumnCount(10);

        Label mismatch = new Label("");
        mismatch.setStyle("-fx-text-fill: #b00020;");

        Button ok = new Button("確認");
        ok.setDefaultButton(true);
        ok.setOnAction(
                e -> {
                    if (SevenDigitChallenge.matches(expected, input.getText())) {
                        confirmed.set(true);
                        stage.setOnCloseRequest(null);
                        stage.close();
                    } else {
                        input.clear();
                        mismatch.setText("数字が一致しません。");
                        input.requestFocus();
                    }
                });

        VBox box = new VBox(12, warn, digits, input, mismatch, ok);
        box.setPadding(new Insets(16));
        box.setAlignment(Pos.CENTER);
        Scene scene = new Scene(box, 480, 280);
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
        stage.setResizable(false);
        input.requestFocus();
        stage.showAndWait();
        return confirmed.get();
    }
}
