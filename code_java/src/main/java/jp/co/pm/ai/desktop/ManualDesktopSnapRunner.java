package jp.co.pm.ai.desktop;

import java.awt.image.BufferedImage;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.embed.swing.SwingFXUtils;
import javafx.scene.Scene;
import javafx.scene.image.WritableImage;
import javafx.stage.Stage;
import javafx.util.Duration;

/**
 * Dev-only: when {@code pm.ai.desktop.manual.snap.tabKeys} is set, selects main-shell tabs in order,
 * sizes the stage, snapshots each scene to PNG under the configured output directory.
 */
public final class ManualDesktopSnapRunner {

    private ManualDesktopSnapRunner() {}

    /** Called after splash teardown (e.g. end of {@link MainShellController#appendBootMessage()}). */
    public static void scheduleIfRequested(MainShellController shell) {
        String keys = System.getProperty("pm.ai.desktop.manual.snap.tabKeys");
        if (keys == null || keys.isBlank()) {
            return;
        }
        Platform.runLater(() -> runSequence(shell, keys));
    }

    private static void runSequence(MainShellController shell, String keysComma) {
        List<MainShellTabId> ids = parseKeys(keysComma);
        if (ids.isEmpty()) {
            System.err.println("[manual-snap] No valid MainShellTabId keys in: " + keysComma);
            maybeExitAfterEmpty();
            return;
        }
        Stage stage = shell.primaryStageForDialogs();
        if (stage == null) {
            System.err.println("[manual-snap] primary stage is null.");
            maybeExitAfterEmpty();
            return;
        }
        Path outDir;
        try {
            outDir = resolveOutputDir();
            Files.createDirectories(outDir);
        } catch (IOException e) {
            e.printStackTrace();
            maybeExitAfterEmpty();
            return;
        }
        runStep(shell, stage, outDir, ids, 0);
    }

    private static List<MainShellTabId> parseKeys(String keysComma) {
        List<MainShellTabId> ids = new ArrayList<>();
        for (String raw : keysComma.split(",")) {
            String t = raw.trim();
            if (t.isEmpty()) {
                continue;
            }
            MainShellTabId id = MainShellTabId.fromKey(t);
            if (id == null) {
                System.err.println("[manual-snap] Unknown tab key (skipped): " + t);
            } else {
                ids.add(id);
            }
        }
        return ids;
    }

    private static void runStep(
            MainShellController shell,
            Stage stage,
            Path outDir,
            List<MainShellTabId> ids,
            int index) {
        if (index >= ids.size()) {
            finishSnapSession();
            return;
        }
        MainShellTabId id = ids.get(index);
        shell.selectMainShellTab(id);
        PauseTransition pause = new PauseTransition(Duration.millis(pauseMillis()));
        pause.setOnFinished(
                e -> {
                    try {
                        applyStageSize(stage);
                        snapshotToPng(stage, outDir.resolve(id.key() + ".png"));
                    } catch (IOException ex) {
                        ex.printStackTrace();
                    }
                    PauseTransition next = new PauseTransition(Duration.millis(50));
                    next.setOnFinished(ev -> runStep(shell, stage, outDir, ids, index + 1));
                    next.play();
                });
        pause.play();
    }

    private static void applyStageSize(Stage stage) {
        double w = parseDoubleProp("pm.ai.desktop.manual.snap.stageWidth", 1800);
        double h = parseDoubleProp("pm.ai.desktop.manual.snap.stageHeight", 900);
        stage.setWidth(w);
        stage.setHeight(h);
    }

    private static double parseDoubleProp(String key, double defaultVal) {
        String v = System.getProperty(key);
        if (v == null || v.isBlank()) {
            return defaultVal;
        }
        try {
            return Double.parseDouble(v.trim());
        } catch (NumberFormatException e) {
            return defaultVal;
        }
    }

    private static long pauseMillis() {
        String v = System.getProperty("pm.ai.desktop.manual.snap.pauseMillis");
        if (v == null || v.isBlank()) {
            return 1500L;
        }
        try {
            return Long.parseLong(v.trim());
        } catch (NumberFormatException e) {
            return 1500L;
        }
    }

    private static void snapshotToPng(Stage stage, Path targetFile) throws IOException {
        Scene scene = stage.getScene();
        if (scene == null) {
            throw new IOException("Stage has no scene");
        }
        WritableImage img = scene.snapshot(null);
        BufferedImage swingImg = SwingFXUtils.fromFXImage(img, null);
        ImageIO.write(swingImg, "png", targetFile.toFile());
        System.out.println("[manual-snap] Wrote " + targetFile.toAbsolutePath());
    }

    private static Path resolveOutputDir() {
        String prop = System.getProperty("pm.ai.desktop.manual.snap.outputDir");
        if (prop != null && !prop.isBlank()) {
            return Paths.get(prop.trim()).toAbsolutePath().normalize();
        }
        Path cwd = Paths.get(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        Path parent = cwd.getParent();
        if (parent != null && cwd.getFileName().toString().equalsIgnoreCase("code_java")) {
            return parent.resolve("manual").resolve("snap-out").normalize();
        }
        return cwd.resolve("manual").resolve("snap-out").normalize();
    }

    private static void finishSnapSession() {
        if (!Boolean.getBoolean("pm.ai.desktop.manual.snap.exitAfter")) {
            return;
        }
        Platform.exit();
        new Thread(
                        () -> {
                            try {
                                Thread.sleep(400);
                            } catch (InterruptedException ignored) {
                                Thread.currentThread().interrupt();
                            }
                            System.exit(0);
                        },
                        "manual-snap-exit")
                .start();
    }

    private static void maybeExitAfterEmpty() {
        if (Boolean.getBoolean("pm.ai.desktop.manual.snap.exitAfter")) {
            finishSnapSession();
        }
    }
}
