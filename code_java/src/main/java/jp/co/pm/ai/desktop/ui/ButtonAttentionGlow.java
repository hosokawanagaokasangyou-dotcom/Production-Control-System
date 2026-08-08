package jp.co.pm.ai.desktop.ui;

import javafx.animation.KeyFrame;
import javafx.animation.KeyValue;
import javafx.animation.Timeline;
import javafx.scene.control.Button;
import javafx.scene.effect.DropShadow;
import javafx.scene.paint.Color;
import javafx.util.Duration;

/** ボタンにパルス状のグローを付け、{@link #stop()} まで目立たせる。 */
public final class ButtonAttentionGlow {

    static final String STYLE_CLASS = "pm-aladdin-entry-export-attention";

    private static final Color GLOW_COLOR = Color.web("#38bdf8");
    private static final double RADIUS_MIN = 6.0;
    private static final double RADIUS_MAX = 22.0;

    private final Button target;
    private Timeline timeline;
    private DropShadow shadow;

    public ButtonAttentionGlow(Button target) {
        this.target = target;
    }

    /** 未起動ならグローアニメーションを開始する。 */
    public void startIfIdle() {
        if (target == null || timeline != null) {
            return;
        }
        startGlowTimeline();
    }

    /** グローを確実に表示する（既に動作中なら一度止めて再開）。 */
    public void ensureActive() {
        if (target == null) {
            return;
        }
        stop();
        startGlowTimeline();
    }

    private void startGlowTimeline() {
        if (!target.getStyleClass().contains(STYLE_CLASS)) {
            target.getStyleClass().add(STYLE_CLASS);
        }
        shadow = new DropShadow();
        shadow.setColor(GLOW_COLOR);
        shadow.setRadius(RADIUS_MIN);
        shadow.setSpread(0.35);
        target.setEffect(shadow);
        timeline =
                new Timeline(
                        new KeyFrame(
                                Duration.ZERO,
                                new KeyValue(shadow.radiusProperty(), RADIUS_MIN)),
                        new KeyFrame(
                                Duration.millis(850),
                                new KeyValue(shadow.radiusProperty(), RADIUS_MAX)));
        timeline.setAutoReverse(true);
        timeline.setCycleCount(Timeline.INDEFINITE);
        timeline.play();
    }

    /** グローを止め通常表示に戻す。 */
    public void stop() {
        if (timeline != null) {
            timeline.stop();
            timeline = null;
        }
        if (target != null) {
            target.setEffect(null);
            target.getStyleClass().remove(STYLE_CLASS);
        }
        shadow = null;
    }

    public static void stopAll(ButtonAttentionGlow... glows) {
        if (glows == null) {
            return;
        }
        for (ButtonAttentionGlow glow : glows) {
            if (glow != null) {
                glow.stop();
            }
        }
    }
}
