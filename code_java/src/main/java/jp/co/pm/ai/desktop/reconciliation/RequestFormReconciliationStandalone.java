package jp.co.pm.ai.desktop.reconciliation;

import java.util.Map;

import javafx.application.Application;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

/** 依頼書照合 UI の単体起動用（メインシェル外）。 */
public final class RequestFormReconciliationStandalone extends Application {

    @Override
    public void start(Stage primaryStage) {
        ReconciliationApp app = new ReconciliationApp();
        Parent root = app.buildEmbeddedRoot(primaryStage, null, Map.of());
        primaryStage.setTitle(
                ReconciliationApp.embeddedWindowTitle(Map.of()).replace(
                        "統合管理データベース (JavaFX版)", "受注一括照合・対比型入力支援システム"));
        Scene scene = new Scene(root, 1560, 700);
        primaryStage.setScene(scene);
        primaryStage.show();
        app.onEmbeddedTabActivated(Map.of());
        primaryStage.setOnCloseRequest(e -> app.onEmbeddedTabDeactivated());
    }

    public static void main(String[] args) {
        launch(args);
    }
}
