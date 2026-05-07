package jp.co.pm.ai.desktop;

import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

/** 本体ウィンドウのFXML読込・初期化が終わるまで表示する簡易スプラッシュ。 */
final class StartupSplashStage {

    private StartupSplashStage() {}

    /**
     * スプラッシュを生成して表示する。呼び出しは JavaFX アプリケーションスレッド上。
     *
     * @return 後から {@link Stage#close()} するステージ
     */
    static Stage createAndShow() {
        Stage stage = new Stage();
        stage.initStyle(StageStyle.UNDECORATED);
        stage.setAlwaysOnTop(true);
        stage.setTitle("起動中");

        Label title = new Label("工程管理 AI 配台");
        title.setStyle("-fx-font-size: 18px; -fx-font-weight: bold;");

        Label sub = new Label("起動しています…");
        sub.setStyle("-fx-font-size: 13px;");

        ProgressIndicator busy = new ProgressIndicator();
        busy.setPrefSize(48, 48);
        busy.setMaxSize(48, 48);

        VBox root = new VBox(20, title, busy, sub);
        root.setAlignment(Pos.CENTER);
        root.setStyle(
                "-fx-background-color: linear-gradient(to bottom, #f7f7f7, #e6e6e6);"
                        + " -fx-padding: 32px 48px;"
                        + " -fx-background-radius: 8px;");
        root.setPrefWidth(420);
        root.setPrefHeight(240);

        Scene scene = new Scene(root);
        stage.setScene(scene);
        stage.setResizable(false);
        stage.centerOnScreen();
        stage.show();
        return stage;
    }
}
