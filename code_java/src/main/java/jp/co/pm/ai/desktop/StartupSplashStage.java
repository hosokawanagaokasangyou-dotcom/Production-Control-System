package jp.co.pm.ai.desktop;

import java.util.concurrent.atomic.AtomicLong;

import javafx.application.Platform;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.stage.WindowEvent;

/**
 * Simple splash until main window FXML is loaded and initialized.
 *
 * <p>Japanese labels use Unicode code point escapes in string literals so this source file is pure
 * ASCII. That
 * avoids Windows builds failing with "cannot map to UTF-8" when a copy of this file is saved in
 * Shift_JIS (CP932) by an editor, while {@code javac} is invoked with {@code -encoding UTF-8}.
 */
final class StartupSplashStage {

    private static final String SPLASH_FONT_STACK =
            "\"Yu Gothic UI\", \"Meiryo UI\", Meiryo, \"Segoe UI\", \"Noto Sans CJK JP\", sans-serif";

    private StartupSplashStage() {}

    /**
     * Creates and shows the splash. Must run on the JavaFX application thread.
     *
     * <p>{@code outVisibleSinceNanos} が非 null のとき、最初にウィンドウが表示されたとみなせる時刻（ナノ秒）を
     * 一度だけ格納する。{@link javafx.stage.WindowEvent#WINDOW_SHOWN} またはその直後のパルスで設定する。
     *
     * @param outVisibleSinceNanos 表示開始時刻を格納するコンテナ（必要なければ {@code null}）
     * @return the stage; close it when the main window is ready
     */
    static Stage createAndShow(AtomicLong outVisibleSinceNanos) {
        Stage stage = new Stage();
        stage.initStyle(StageStyle.UNDECORATED);
        stage.initModality(Modality.APPLICATION_MODAL);
        stage.setAlwaysOnTop(true);
        // Japanese UI strings below: ASCII source via Unicode escapes (see class Javadoc).
        stage.setTitle("\u8d77\u52d5\u4e2d");

        Label title = new Label("\u5de5\u7a0b\u7ba1\u7406 AI \u914d\u53f0");
        title.setStyle(
                "-fx-font-family: "
                        + SPLASH_FONT_STACK
                        + "; -fx-font-size: 18px; -fx-font-weight: bold;");

        Label sub = new Label("\u8d77\u52d5\u3057\u3066\u3044\u307e\u3059...");
        sub.setStyle(
                "-fx-font-family: "
                        + SPLASH_FONT_STACK
                        + "; -fx-font-size: 13px;");

        ProgressIndicator busy = new ProgressIndicator();
        busy.setPrefSize(48, 48);
        busy.setMaxSize(48, 48);

        VBox root = new VBox(20, title, busy, sub);
        root.setAlignment(Pos.CENTER);
        root.setStyle(
                "-fx-font-family: "
                        + SPLASH_FONT_STACK
                        + ";"
                        + " -fx-background-color: linear-gradient(to bottom, #f7f7f7, #e6e6e6);"
                        + " -fx-padding: 32px 48px;"
                        + " -fx-background-radius: 8px;");
        root.setPrefWidth(420);
        root.setPrefHeight(240);

        Scene scene = new Scene(root);
        stage.setScene(scene);
        stage.setResizable(false);
        stage.centerOnScreen();
        if (outVisibleSinceNanos != null) {
            stage.addEventHandler(
                    WindowEvent.WINDOW_SHOWN,
                    e -> outVisibleSinceNanos.compareAndSet(0L, System.nanoTime()));
        }
        stage.show();
        raiseToFront(stage);
        if (outVisibleSinceNanos != null) {
            Platform.runLater(() -> outVisibleSinceNanos.compareAndSet(0L, System.nanoTime()));
        }
        return stage;
    }

    /** Moves splash forward after OS focus steal or other Stage creation. */
    static void raiseToFront(Stage splash) {
        if (splash == null) {
            return;
        }
        splash.toFront();
        splash.requestFocus();
    }
}
