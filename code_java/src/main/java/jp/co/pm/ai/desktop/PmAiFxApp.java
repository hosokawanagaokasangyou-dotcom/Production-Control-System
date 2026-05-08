package jp.co.pm.ai.desktop;

import java.awt.GraphicsEnvironment;
import java.nio.charset.StandardCharsets;
import java.util.Locale;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.StartupCrashLog;
import jp.co.pm.ai.desktop.runtime.JvmMemoryMonitor;
import jp.co.pm.ai.desktop.runtime.WindowsLauncherUserDir;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * JavaFX エントリ — UI レイアウトは FXML（{@code jp/co/pm/ai/desktop/fxml/MainShell.fxml}）、ロジックは
 * {@link MainShellController}。
 *
 * <p>Prism は通常、起動時に {@link GpuProbeLauncher} で Canvas＋GPU パイプラインを別 JVM が試し、合格時のみ本体 JVM で GPU
 * を有効にする。強制や省略は {@code pm.ai.javafx.prism.*} を参照。
 */
public class PmAiFxApp extends Application {

    static {
        System.setProperty("file.encoding", "UTF-8");
        try {
            StartupCrashLog.append("PmAiFxApp: static initializer (before main)");
        } catch (Throwable ignored) {
            /* ログクラス初期化失敗でも本体クラスは読み込む */
        }
    }

    /**
     * Toolkit 初期化前に呼ぶ。既定は別プロセスの GPU プローブに従い {@code prism.order} を決める。
     *
     * <ul>
     *   <li>{@code -Dpm.ai.javafx.prism.skipGpuProbe=true} … プローブせず {@link #applyLegacyPrismConfiguration()}
     *   <li>{@code -Dpm.ai.javafx.prism.gpu=true} または {@code PM_AI_JAVAFX_PRISM_GPU=1} … プローブ省略で GPU 優先
     * </ul>
     */
    private static void configurePrismAfterProbe() {
        if (Boolean.getBoolean("pm.ai.javafx.prism.skipGpuProbe")) {
            applyLegacyPrismConfiguration();
            return;
        }
        if (prismGpuOptIn()) {
            applyPrismGpuPipelineOrder();
            PrismGpuBootstrapStatus.recordGpuOptIn();
            return;
        }
        boolean ok = GpuProbeLauncher.runGpuCanvasProbe();
        if (ok) {
            applyPrismGpuPipelineOrder();
            PrismGpuBootstrapStatus.recordGpuAfterProbe();
        } else {
            System.setProperty("prism.order", "sw");
        }
    }

    /** プローブ無効時の従来どおりの設定（opt-in GPU または JVM の prism.order / 既定 sw）。 */
    private static void applyLegacyPrismConfiguration() {
        if (prismGpuOptIn()) {
            applyPrismGpuPipelineOrder();
            PrismGpuBootstrapStatus.recordGpuOptIn();
            return;
        }
        PrismGpuBootstrapStatus.recordLegacyNoProbe();
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

        Stage splash = StartupSplashStage.createAndShow();
        // loader.load() が同じパルスで走るとスプラッシュが描画されないため、次のパルスで本体を構築する
        Platform.runLater(
                () -> {
                    try {
                        StartupSplashStage.raiseToFront(splash);
                        MainShellController shell = bootstrapMainWindow(primaryStage);
                        // primary にシーンを載せた直後、前面が移ることがあるため閉じる直前にも前面化する
                        StartupSplashStage.raiseToFront(splash);
                        primaryStage.show();
                        splash.close();
                        shell.appendBootMessage();
                    } catch (Exception e) {
                        splash.close();
                        throw new RuntimeException(e);
                    }
                });
    }

    private static MainShellController bootstrapMainWindow(Stage primaryStage) throws Exception {
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
        TableColumnOrderPersistence.materializeBundledDefaultsIfStoreMissing();
        Parent root = loader.load();
        MainShellController shell = loader.getController();
        Scene scene = new Scene(root, 1800, 850);
        scene.getStylesheets()
                .add(
                        PmAiFxApp.class
                                .getResource("/jp/co/pm/ai/desktop/css/pm-ai-desktop.css")
                                .toExternalForm());
        primaryStage.setScene(scene);
        shell.finishStartup(scene);
        return shell;
    }

    public static void main(String[] args) {
        WindowsLauncherUserDir.alignWithPackagedLauncherIfWindows();
        StartupCrashLog.installUncaughtExceptionLogging();
        StartupCrashLog.append("main: begin user.dir=" + System.getProperty("user.dir"));
        if (GraphicsEnvironment.isHeadless()) {
            String msg =
                    "[PmAiFxApp] No graphical display (headless). "
                            + "Run on Windows desktop, or on WSL set DISPLAY for JavaFX (e.g. WSLg / VcXsrv). "
                            + "Do not run javafx:run from SSH without X forwarding.";
            StartupCrashLog.append(msg);
            System.err.println(msg);
            System.exit(2);
        }
        try {
            configurePrismAfterProbe();
            StartupCrashLog.append(
                    "main: after configurePrism prism.order="
                            + System.getProperty("prism.order", ""));
            JvmMemoryMonitor.startFromMain();
            launch(args);
        } catch (Throwable t) {
            StartupCrashLog.appendThrowable("main: launch failed", t);
            throw t;
        }
    }
}
