package jp.co.pm.ai.desktop;

import java.util.Locale;
import java.util.concurrent.atomic.AtomicReference;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.scene.Scene;
import javafx.scene.canvas.Canvas;
import javafx.scene.canvas.GraphicsContext;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.SnapshotParameters;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.util.Duration;

/**
 * 別 JVM で JavaFX Canvas を GPU Prism 経路で描画し、{@code NGCanvas}/{@code RTTexture} 系の例外が出ないか確認する。
 *
 * <p>終了コード: 0 合格 / 1 不合格。
 */
public final class JavaFxGpuProbeApp extends Application {

    private static final AtomicReference<Throwable> asyncFailure = new AtomicReference<>();

    @Override
    public void start(Stage stage) {
        Thread.setDefaultUncaughtExceptionHandler(
                (t, ex) -> {
                    asyncFailure.compareAndSet(null, ex);
                    try {
                        Platform.exit();
                    } catch (Throwable ignored) {
                        // ignore
                    }
                });

        try {
            Canvas canvas = new Canvas(320, 240);
            GraphicsContext gc = canvas.getGraphicsContext2D();
            gc.setFill(Color.web("#404060"));
            gc.fillRect(0, 0, 320, 240);
            gc.setStroke(Color.LIGHTGRAY);
            gc.strokeText("gpu probe", 40, 120);

            StackPane root = new StackPane(canvas);
            Scene scene = new Scene(root);
            stage.initStyle(StageStyle.UNDECORATED);
            stage.setScene(scene);
            stage.setWidth(320);
            stage.setHeight(240);
            if (System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("windows")) {
                stage.setX(-4000);
                stage.setY(-4000);
            } else {
                stage.setX(0);
                stage.setY(0);
            }
            stage.show();

            SnapshotParameters snapParams = new SnapshotParameters();
            WritableImage img =
                    new WritableImage(
                            (int) Math.ceil(canvas.getWidth()),
                            (int) Math.ceil(canvas.getHeight()));

            Timeline timeline =
                    new Timeline(
                            new KeyFrame(
                                    Duration.ZERO,
                                    e -> {
                                        try {
                                            canvas.snapshot(snapParams, img);
                                        } catch (Throwable err) {
                                            asyncFailure.compareAndSet(null, err);
                                        }
                                    }),
                            new KeyFrame(
                                    Duration.millis(120),
                                    e -> {
                                        try {
                                            canvas.snapshot(snapParams, img);
                                        } catch (Throwable err) {
                                            asyncFailure.compareAndSet(null, err);
                                        }
                                    }),
                            new KeyFrame(
                                    Duration.millis(900),
                                    e -> {
                                        try {
                                            canvas.snapshot(snapParams, img);
                                        } catch (Throwable err) {
                                            asyncFailure.compareAndSet(null, err);
                                        }
                                    }),
                            new KeyFrame(Duration.millis(1700), e -> Platform.exit()));
            timeline.play();
        } catch (Throwable ex) {
            asyncFailure.compareAndSet(null, ex);
            Platform.exit();
        }
    }

    public static void main(String[] args) {
        launch(args);
        Throwable fail = asyncFailure.get();
        if (fail != null) {
            fail.printStackTrace(System.err);
            System.exit(1);
        }
        System.exit(0);
    }
}
