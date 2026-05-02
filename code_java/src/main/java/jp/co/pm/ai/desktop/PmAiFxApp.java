package jp.co.pm.ai.desktop;

import java.awt.GraphicsEnvironment;
import java.nio.charset.StandardCharsets;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

/**
 * JavaFX エントリ — UI レイアウトは FXML（{@code jp/co/pm/ai/desktop/fxml/MainShell.fxml}）、ロジックは
 * {@link MainShellController}。
 */
public class PmAiFxApp extends Application {

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("\u5de5\u7a0b\u7ba1\u7406 AI \u914d\u53f0 \u2014 JavaFX MVP");

        try {
            FXMLLoader loader =
                    new FXMLLoader(
                            PmAiFxApp.class.getResource("/jp/co/pm/ai/desktop/fxml/MainShell.fxml"));
            loader.setCharset(StandardCharsets.UTF_8);
            loader.setControllerFactory(
                    clazz -> {
                        if (clazz == MainShellController.class) {
                            return new MainShellController(primaryStage);
                        }
                        try {
                            return clazz.getDeclaredConstructor().newInstance();
                        } catch (Exception e) {
                            throw new IllegalStateException(e);
                        }
                    });
            Parent root = loader.load();
            MainShellController shell = loader.getController();
            Scene scene = new Scene(root, 960, 640);
            primaryStage.setScene(scene);
            primaryStage.show();
            shell.appendBootMessage();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static void main(String[] args) {
        System.setProperty("file.encoding", "UTF-8");
        if (GraphicsEnvironment.isHeadless()) {
            System.err.println(
                    "[PmAiFxApp] No graphical display (headless). "
                            + "Run on Windows desktop, or on WSL set DISPLAY for JavaFX (e.g. WSLg / VcXsrv). "
                            + "Do not run javafx:run from SSH without X forwarding.");
            System.exit(2);
        }
        launch(args);
    }
}
