package jp.co.pm.ai.desktop;

import java.awt.GraphicsEnvironment;
import java.nio.charset.StandardCharsets;
import java.util.Locale;

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

    /**
     * Prism のパイプライン順（Toolkit 初期化前）。既定はソフトウェア描画（{@code sw}）。Canvas（{@code NGCanvas}）と GPU
     * の組み合わせで {@code RTTexture} が null になり得るため。
     *
     * <p>GPU を試す: JVM {@code -Dpm.ai.javafx.prism.gpu=true} または環境変数 {@code PM_AI_JAVAFX_PRISM_GPU=1}。そのときのみ
     * OS 別の GPU 優先 {@code prism.order} を適用する。opt-in しない場合は既存の {@code prism.order} が無ければ {@code sw}。
     */
    private static void configurePrismOrderEarly() {
        if (prismGpuOptIn()) {
            applyPrismGpuPipelineOrder();
            return;
        }
        if (System.getProperty("prism.order") != null) {
            return;
        }
        System.setProperty("prism.order", "sw");
    }

    private static boolean prismGpuOptIn() {
        if (Boolean.getBoolean("pm.ai.javafx.prism.gpu")) {
            return true;
        }
        String env = System.getenv("PM_AI_JAVAFX_PRISM_GPU");
        return env != null && ("1".equals(env) || "true".equalsIgnoreCase(env));
    }

    private static void applyPrismGpuPipelineOrder() {
        String os = System.getProperty("os.name", "").toLowerCase(Locale.ROOT);
        if (os.contains("windows")) {
            System.setProperty("prism.order", "d3d,es2,sw");
        } else if (os.contains("mac")) {
            System.setProperty("prism.order", "metal,es2,sw");
        } else {
            System.setProperty("prism.order", "es2,sw");
        }
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("工程管理 AI 配台 — JavaFX MVP");

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
            scene.getStylesheets()
                    .add(
                            PmAiFxApp.class
                                    .getResource("/jp/co/pm/ai/desktop/css/pm-ai-desktop.css")
                                    .toExternalForm());
            primaryStage.setScene(scene);
            shell.finishStartup(scene);
            primaryStage.show();
            shell.appendBootMessage();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static void main(String[] args) {
        configurePrismOrderEarly();
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
