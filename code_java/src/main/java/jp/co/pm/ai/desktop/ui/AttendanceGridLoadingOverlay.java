package jp.co.pm.ai.desktop.ui;

import java.util.Locale;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.geometry.Pos;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.util.Duration;

/** 勤怠グリッド上の読込・保存オーバーレイ（スピナー＋経過秒でフリーズと区別）。 */
public final class AttendanceGridLoadingOverlay extends StackPane {

    private static final String DEFAULT_MESSAGE = "読込中";

    private final Region backdrop;
    private final VBox contentBox;
    private final ProgressIndicator indicator = new ProgressIndicator();
    private final Label messageLabel = new Label(DEFAULT_MESSAGE);
    private final Label elapsedLabel = new Label();
    private final Timeline tick;
    private long startMs = 0L;
    private String baseMessage = DEFAULT_MESSAGE;
    private int dotPhase = 0;
    private boolean processing = false;
    private boolean attentionOnly = false;

    public AttendanceGridLoadingOverlay(String backdropStyleClass) {
        backdrop = new Region();
        backdrop.getStyleClass().add(backdropStyleClass);
        backdrop.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);

        indicator.setPrefSize(40, 40);
        indicator.setMaxSize(40, 40);
        messageLabel.getStyleClass().add("pm-attendance-grid-loading-message");
        elapsedLabel.getStyleClass().add("pm-attendance-grid-loading-elapsed");

        contentBox = new VBox(10, indicator, messageLabel, elapsedLabel);
        contentBox.setAlignment(Pos.CENTER);
        contentBox.getStyleClass().add("pm-attendance-grid-loading-box");
        contentBox.setMaxWidth(360);

        getChildren().addAll(backdrop, contentBox);
        setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        setVisible(false);
        setMouseTransparent(true);

        tick =
                new Timeline(
                        new KeyFrame(
                                Duration.millis(400),
                                e -> {
                                    updateElapsed();
                                    updateDots();
                                }));
        tick.setCycleCount(Timeline.INDEFINITE);
    }

    public void setLoading(boolean loading) {
        setLoading(loading, null);
    }

    public void setLoading(boolean loading, String message) {
        processing = loading;
        if (loading) {
            attentionOnly = false;
            setMessage(message);
            startMs = System.currentTimeMillis();
            updateElapsed();
            updateDots();
            indicator.setVisible(true);
            messageLabel.setVisible(true);
            elapsedLabel.setVisible(true);
            contentBox.setVisible(true);
            tick.play();
        } else {
            tick.stop();
        }
        applyVisibility();
    }

    /** 段階2未準備など、暗転のみ（スピナーなし・クリック透過）。 */
    public void setAttentionOnly(boolean visible) {
        attentionOnly = visible;
        if (visible) {
            processing = false;
            tick.stop();
            contentBox.setVisible(false);
        }
        applyVisibility();
    }

    public void setMessage(String message) {
        baseMessage =
                message != null && !message.isBlank()
                        ? message.strip().replaceAll("[.．…]+$", "")
                        : DEFAULT_MESSAGE;
        dotPhase = 0;
        messageLabel.setText(baseMessage + "…");
    }

    private void applyVisibility() {
        if (processing) {
            setVisible(true);
            setMouseTransparent(false);
            return;
        }
        if (attentionOnly) {
            setVisible(true);
            setMouseTransparent(true);
            return;
        }
        setVisible(false);
        setMouseTransparent(true);
    }

    private void updateElapsed() {
        double sec = (System.currentTimeMillis() - startMs) / 1000.0;
        elapsedLabel.setText(String.format(Locale.ROOT, "経過 %.1f 秒（処理中）", sec));
    }

    private void updateDots() {
        dotPhase = (dotPhase + 1) % 4;
        String dots = ".".repeat(dotPhase);
        messageLabel.setText(baseMessage + "…" + dots);
    }
}
