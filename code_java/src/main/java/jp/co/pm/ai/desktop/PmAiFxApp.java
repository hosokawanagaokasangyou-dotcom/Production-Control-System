package jp.co.pm.ai.desktop;

import java.awt.GraphicsEnvironment;
import java.nio.charset.StandardCharsets;
import java.util.Locale;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicLong;

import javafx.animation.PauseTransition;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;
import javafx.stage.WindowEvent;
import javafx.util.Duration;

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
     *   <li>{@code -Dpm.ai.javafx.prism.gpu=true} または {@code PM_AI_JAVAFX_PRISM_GPU=1} … GPU 優先（無名
     *       CLASSPATH のときは起動前に鏡像 GPU プローブに合格した場合のみ HW 順を適用。不合格なら SW）
     *   <li>{@code -Dpm.ai.javafx.prism.allowHwOnClasspath=true} … 無名 CLASSPATH かつ通常起動（上記 gpu 以外）で、GPU
     *       試験合格後に HW パイプラインを試す（既定は Canvas 安定のため SW）
     *   <li>{@code -Dpm.ai.javafx.prism.forceSwOnClasspath=true} … 無名 CLASSPATH のとき、GPU 試験合格後も必ず SW
     *   <li>{@code -Dpm.ai.javafx.prism.probeSplitOpenJfx=true} … GPU 子プロセスだけ従来どおり OpenJFX を
     *       {@code --module-path} に切り出す（無名 CLASSPATH 親との鏡像プローブが子起動に失敗するときの切り戻し）
     * </ul>
     */
    private static void configurePrismAfterProbe() {
        if (Boolean.getBoolean("pm.ai.javafx.prism.skipGpuProbe")) {
            applyLegacyPrismConfiguration();
            return;
        }
        if (applyPrismWhenGpuOptIn()) {
            return;
        }
        boolean ok = GpuProbeLauncher.runGpuCanvasProbe(javaFxRuntimeOnUnnamedClasspath());
        if (ok) {
            /*
             * 無名 CLASSPATH では Canvas＋HW の組み合わせが環境によって不安定になり得る。既定は SW。
             * GPU を試すには allowHwOnClasspath または gpu／環境変数オプトイン（オプトインは上で鏡像プローブ済み）。
             */
            boolean unnamed = javaFxRuntimeOnUnnamedClasspath();
            if (unnamed && Boolean.getBoolean("pm.ai.javafx.prism.forceSwOnClasspath")) {
                System.setProperty("prism.order", "sw");
                PrismGpuBootstrapStatus.recordSoftwareClasspathOpenJfx("forceSw");
            } else if (unnamed && !Boolean.getBoolean("pm.ai.javafx.prism.allowHwOnClasspath")) {
                System.setProperty("prism.order", "sw");
                PrismGpuBootstrapStatus.recordSoftwareClasspathOpenJfx("defaultSw");
            } else {
                applyPrismGpuPipelineOrder();
                PrismGpuBootstrapStatus.recordGpuAfterProbe();
            }
        } else {
            System.setProperty("prism.order", "sw");
            PrismGpuBootstrapStatus.recordSoftwareAfterProbe("GPU テスト不合格");
        }
    }

    /** プローブ無効時の従来どおりの設定（opt-in GPU または JVM の prism.order / 既定 sw）。 */
    private static void applyLegacyPrismConfiguration() {
        if (applyPrismWhenGpuOptIn()) {
            return;
        }
        PrismGpuBootstrapStatus.recordLegacyNoProbe();
        if (System.getProperty("prism.order") != null) {
            return;
        }
        System.setProperty("prism.order", "sw");
    }

    /**
     * {@code pm.ai.javafx.prism.gpu} / {@code PM_AI_JAVAFX_PRISM_GPU} 指定時に HW 順を適用する。無名 CLASSPATH のときは鏡像
     * GPU プローブに不合格なら SW に落とす。
     *
     * @return オプトインを処理して呼び出し元が return すべきなら true
     */
    private static boolean applyPrismWhenGpuOptIn() {
        if (!prismGpuOptIn()) {
            return false;
        }
        if (javaFxRuntimeOnUnnamedClasspath()) {
            boolean gpuOk = GpuProbeLauncher.runGpuCanvasProbe(true);
            if (!gpuOk) {
                System.setProperty("prism.order", "sw");
                PrismGpuBootstrapStatus.recordSoftwareAfterProbe(
                        "GPU opt-in: 無名 CLASSPATH 鏡像プローブ不合格");
                return true;
            }
        }
        applyPrismGpuPipelineOrder();
        PrismGpuBootstrapStatus.recordGpuOptIn();
        return true;
    }

    private static boolean prismGpuOptIn() {
        if (Boolean.getBoolean("pm.ai.javafx.prism.gpu")) {
            return true;
        }
        String env = System.getenv("PM_AI_JAVAFX_PRISM_GPU");
        return env != null && ("1".equals(env) || "true".equalsIgnoreCase(env));
    }

    /** OpenJFX が名前付きモジュールではなく CLASSPATH（無名）で解決されているか。 */
    private static boolean javaFxRuntimeOnUnnamedClasspath() {
        try {
            return !Node.class.getModule().isNamed();
        } catch (Throwable ignored) {
            return false;
        }
    }

    private static void applyPrismGpuPipelineOrder() {
        String os = System.getProperty("os.name", "").toLowerCase(Locale.ROOT);
        if (os.contains("windows")) {
            /*
             * javafx:run の CLASSPATH（無名モジュール）構成では D3D 先行だと Canvas 描画で
             * NGCanvas$RenderBuf.validate が RTTexture null の NPE になる事例がある（ターミナル再現）。
             * ES2 を先に試し、必要なら D3D へ退避可能な順序にする。
             */
            System.setProperty("prism.order", "es2,d3d,sw");
        } else if (os.contains("mac")) {
            System.setProperty("prism.order", "metal,es2,sw");
        } else {
            System.setProperty("prism.order", "es2,sw");
        }
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("工程管理 AI 配台");

        AtomicLong splashVisibleSinceNanos = new AtomicLong();
        StartupSplashStage.createAndShow(
                splashVisibleSinceNanos,
                splash -> {
                    try {
                        AtomicLong mainWindowPaintedNanos = new AtomicLong();
                        AtomicBoolean splashCloseScheduled = new AtomicBoolean();

                        StartupSplashStage.raiseToFront(splash);
                        MainShellController shell = bootstrapMainWindow(primaryStage);
                        // primary にシーンを載せた直後、前面が移ることがあるため閉じる直前にも前面化する
                        StartupSplashStage.raiseToFront(splash);

                        Runnable markMainPaintedAndScheduleClose =
                                () -> {
                                    mainWindowPaintedNanos.compareAndSet(
                                            0L, System.nanoTime());
                                    scheduleSplashCloseAfterMainPainted(
                                            splash,
                                            splashVisibleSinceNanos,
                                            mainWindowPaintedNanos,
                                            shell,
                                            splashCloseScheduled);
                                };
                        // メインを一度表示したあと、レイアウト・初回描画が終わるまで待ってから閉じる
                        primaryStage.addEventHandler(
                                WindowEvent.WINDOW_SHOWN,
                                e ->
                                        Platform.runLater(
                                                () ->
                                                        Platform.runLater(
                                                                markMainPaintedAndScheduleClose)));
                        primaryStage.show();
                        Platform.runLater(
                                () ->
                                        Platform.runLater(markMainPaintedAndScheduleClose));
                    } catch (Exception e) {
                        splash.close();
                        throw new RuntimeException(e);
                    }
                });
    }

    private static final long SPLASH_MIN_VISIBLE_NANOS = 3_000_000_000L;

    /**
     * メインウィンドウの初回表示・レイアウト後（{@link WindowEvent#WINDOW_SHOWN} のあと 2 パルス）と、スプラッシュの最低表示時間の
     * いずれか遅い方まで待ってから閉じる。
     */
    private static void scheduleSplashCloseAfterMainPainted(
            Stage splash,
            AtomicLong splashVisibleSinceNanos,
            AtomicLong mainWindowPaintedNanos,
            MainShellController shell,
            AtomicBoolean splashCloseScheduled) {
        long painted = mainWindowPaintedNanos.get();
        if (painted == 0L) {
            return;
        }
        if (!splashCloseScheduled.compareAndSet(false, true)) {
            return;
        }
        long since = splashVisibleSinceNanos.get();
        if (since == 0L) {
            since = System.nanoTime();
        }
        long earliestCloseNanos = Math.max(since + SPLASH_MIN_VISIBLE_NANOS, painted);
        long waitNs = earliestCloseNanos - System.nanoTime();
        Runnable finish =
                () -> {
                    splash.close();
                    Stage main = shell.primaryStageForDialogs();
                    /* モーダルスプラッシュ解除直後は OS がフォーカスを他へ逃がすことがあるため、次パルスで前面化 */
                    Platform.runLater(
                            () ->
                                    Platform.runLater(
                                            () -> {
                                                if (main != null && main.isShowing()) {
                                                    main.toFront();
                                                    main.requestFocus();
                                                }
                                                shell.appendBootMessage();
                                            }));
                };
        if (waitNs <= 0) {
            finish.run();
            return;
        }
        double millis = waitNs / 1_000_000.0;
        PauseTransition pause = new PauseTransition(Duration.millis(millis));
        pause.setOnFinished(e -> finish.run());
        pause.play();
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
        Scene scene = new Scene(root);
        scene.getStylesheets()
                .add(
                        PmAiFxApp.class
                                .getResource("/jp/co/pm/ai/desktop/css/pm-ai-desktop.css")
                                .toExternalForm());
        MainShellController.debugLogParentsWithExactChildCount(root, 19, shell.snapshotUiEnv());
        primaryStage.setScene(scene);
        shell.finishStartup(scene);
        Platform.runLater(shell::reapplyMainShellTabContentManagedAfterSceneAttach);
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
                            + "Do not run the desktop from Maven/exec without X forwarding when using SSH.";
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
