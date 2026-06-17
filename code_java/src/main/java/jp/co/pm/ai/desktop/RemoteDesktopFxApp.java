package jp.co.pm.ai.desktop;

import java.awt.GraphicsEnvironment;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicLong;

import javafx.animation.PauseTransition;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;
import javafx.stage.WindowEvent;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.StartupCrashLog;
import jp.co.pm.ai.desktop.runtime.JvmMemoryMonitor;
import jp.co.pm.ai.desktop.runtime.WindowsLauncherUserDir;

/** リモートデスクトップ専用ポータブル（{@code rpa_luncher_release} の PmAiRpaLuncher.exe）の JavaFX エントリ。 */
public final class RemoteDesktopFxApp extends Application {

    static {
        RemoteDesktopStandaloneBootstrap.activate();
        System.setProperty("file.encoding", "UTF-8");
        if (System.getProperty("prism.order") == null) {
            System.setProperty("prism.order", "sw");
        }
        try {
            StartupCrashLog.append("RemoteDesktopFxApp: static initializer");
        } catch (Throwable ignored) {
        }
    }

    private static final long SPLASH_MIN_VISIBLE_NANOS = 3_000_000_000L;

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle(RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE);

        AtomicLong splashVisibleSinceNanos = new AtomicLong();
        StartupSplashStage.createAndShow(
                StartupSplashBranding.REMOTE_DESKTOP_LAUNCHER,
                splashVisibleSinceNanos,
                splash -> {
                    try {
                        AtomicLong mainWindowPaintedNanos = new AtomicLong();
                        AtomicBoolean splashCloseScheduled = new AtomicBoolean();

                        StartupSplashStage.raiseToFront(splash);
                        RemoteDesktopShellController shell = bootstrapMainWindow(primaryStage);
                        StartupSplashStage.raiseToFront(splash);

                        Runnable markMainPaintedAndScheduleClose =
                                () -> {
                                    mainWindowPaintedNanos.compareAndSet(0L, System.nanoTime());
                                    scheduleSplashCloseAfterMainPainted(
                                            splash,
                                            splashVisibleSinceNanos,
                                            mainWindowPaintedNanos,
                                            shell,
                                            splashCloseScheduled);
                                };
                        primaryStage.addEventHandler(
                                WindowEvent.WINDOW_SHOWN,
                                e ->
                                        Platform.runLater(
                                                () ->
                                                        Platform.runLater(
                                                                markMainPaintedAndScheduleClose)));
                        primaryStage.show();
                        Platform.runLater(
                                () -> Platform.runLater(markMainPaintedAndScheduleClose));
                    } catch (Exception e) {
                        splash.close();
                        throw new RuntimeException(e);
                    }
                });
    }

    private static void scheduleSplashCloseAfterMainPainted(
            Stage splash,
            AtomicLong splashVisibleSinceNanos,
            AtomicLong mainWindowPaintedNanos,
            RemoteDesktopShellController shell,
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
                    Platform.runLater(
                            () ->
                                    Platform.runLater(
                                            () -> {
                                                if (main != null && main.isShowing()) {
                                                    applyStartupFullScreen(main);
                                                    main.toFront();
                                                    main.requestFocus();
                                                }
                                            }));
                };
        if (waitNs <= 0) {
            finish.run();
            return;
        }
        PauseTransition pause = new PauseTransition(Duration.millis(waitNs / 1_000_000.0));
        pause.setOnFinished(e -> finish.run());
        pause.play();
    }

    /**
     * 起動直後のアプリウィンドウを画面全体に広げる（RDP セッションのフルスクリーン設定とは別）。
     * 排他フルスクリーンは mstsc 全画面セッションの Z 順・HWND 探索を妨げるため最大化のみ使う。
     */
    private static void applyStartupFullScreen(Stage stage) {
        stage.setFullScreen(false);
        stage.setMaximized(true);
    }

    private static RemoteDesktopShellController bootstrapMainWindow(Stage primaryStage)
            throws Exception {
        FXMLLoader loader =
                new FXMLLoader(
                        RemoteDesktopFxApp.class.getResource(
                                "/jp/co/pm/ai/desktop/fxml/RemoteDesktopShell.fxml"));
        loader.setCharset(StandardCharsets.UTF_8);
        loader.setControllerFactory(
                clazz -> {
                    if (clazz == RemoteDesktopShellController.class) {
                        return new RemoteDesktopShellController(primaryStage);
                    }
                    try {
                        return clazz.getDeclaredConstructor().newInstance();
                    } catch (Exception e) {
                        throw new IllegalStateException(e);
                    }
                });
        Parent root = loader.load();
        RemoteDesktopShellController shell = loader.getController();
        Scene scene = new Scene(root);
        scene.getStylesheets()
                .add(
                        RemoteDesktopFxApp.class
                                .getResource("/jp/co/pm/ai/desktop/css/pm-ai-desktop.css")
                                .toExternalForm());
        shell.finishStartup(scene);
        Platform.runLater(() -> primaryStage.setScene(scene));
        return shell;
    }

    public static void main(String[] args) {
        WindowsLauncherUserDir.alignWithPackagedLauncherIfWindows();
        StartupCrashLog.installUncaughtExceptionLogging();
        StartupCrashLog.append("RemoteDesktopFxApp.main begin");
        if (GraphicsEnvironment.isHeadless()) {
            String msg =
                    "[RemoteDesktopFxApp] No graphical display (headless). Run on Windows desktop.";
            StartupCrashLog.append(msg);
            System.err.println(msg);
            System.exit(2);
        }
        try {
            JvmMemoryMonitor.startFromMain();
            launch(args);
        } catch (Throwable t) {
            StartupCrashLog.appendThrowable("RemoteDesktopFxApp.main failed", t);
            throw t;
        }
    }
}
