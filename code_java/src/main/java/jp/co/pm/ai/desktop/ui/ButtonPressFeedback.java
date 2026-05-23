package jp.co.pm.ai.desktop.ui;

import javafx.animation.PauseTransition;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.MenuButton;
import javafx.scene.effect.DropShadow;
import javafx.scene.effect.Effect;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.input.MouseEvent;
import javafx.scene.paint.Color;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.audio.UiClickSound;

/**
 * ボタン押下時に短いクリック音と視覚フラッシュを、重い {@code @FXML} 処理より先に出す。
 *
 * <p>{@link Scene} にイベントフィルタを1回だけ登録する。二次ダイアログの {@link Scene} にも {@link #installOnScene} を呼ぶ。
 */
public final class ButtonPressFeedback {

    static final String SCENE_INSTALLED_PROP = "pmButtonPressFeedbackInstalled";
    static final String FLASH_ACTIVE_PROP = "pmButtonPressFlashActive";
    static final String FLASH_STYLE_CLASS = "pm-button-press-flash";

    private static final Duration FLASH_HOLD = Duration.millis(130);

    private ButtonPressFeedback() {}

    public static void installOnScene(Scene scene) {
        if (scene == null) {
            return;
        }
        if (Boolean.TRUE.equals(scene.getProperties().get(SCENE_INSTALLED_PROP))) {
            return;
        }
        scene.getProperties().put(SCENE_INSTALLED_PROP, Boolean.TRUE);

        scene.addEventFilter(
                MouseEvent.MOUSE_PRESSED,
                e -> {
                    if (!e.isPrimaryButtonDown()) {
                        return;
                    }
                    Button button = findButton(e.getPickResult().getIntersectedNode());
                    if (button != null) {
                        trigger(button);
                    }
                });

        scene.addEventFilter(
                KeyEvent.KEY_PRESSED,
                e -> {
                    if (e.getCode() != KeyCode.SPACE && e.getCode() != KeyCode.ENTER) {
                        return;
                    }
                    Button button = findButton(scene.getFocusOwner());
                    if (button != null) {
                        trigger(button);
                    }
                });
    }

    /** 音 → 視覚フラッシュ（処理本体より先）。 */
    public static void trigger(Button button) {
        if (button == null || button.isDisabled()) {
            return;
        }
        UiClickSound.playClick();
        flashButton(button);
    }

    private static void flashButton(Button button) {
        if (Boolean.TRUE.equals(button.getProperties().get(FLASH_ACTIVE_PROP))) {
            return;
        }
        button.getProperties().put(FLASH_ACTIVE_PROP, Boolean.TRUE);

        double priorTy = button.getTranslateY();
        Effect priorEffect = button.getEffect();
        boolean hadFlashClass = button.getStyleClass().contains(FLASH_STYLE_CLASS);
        if (!hadFlashClass) {
            button.getStyleClass().add(FLASH_STYLE_CLASS);
        }

        button.setTranslateY(priorTy + 2.0);
        if (priorEffect instanceof DropShadow ds) {
            DropShadow pressed = new DropShadow();
            pressed.setColor(ds.getColor());
            pressed.setRadius(Math.max(4, ds.getRadius() * 0.45));
            pressed.setSpread(Math.max(0, ds.getSpread() * 0.5));
            pressed.setOffsetX(ds.getOffsetX());
            pressed.setOffsetY(Math.max(0.5, ds.getOffsetY() * 0.35));
            button.setEffect(pressed);
        }

        PauseTransition hold = new PauseTransition(FLASH_HOLD);
        hold.setOnFinished(
                ev -> {
                    button.setTranslateY(priorTy);
                    button.setEffect(priorEffect);
                    if (!hadFlashClass) {
                        button.getStyleClass().remove(FLASH_STYLE_CLASS);
                    }
                    button.getProperties().remove(FLASH_ACTIVE_PROP);
                });
        hold.play();
    }

    private static Button findButton(Node node) {
        Node n = node;
        while (n != null) {
            if (n instanceof Button b && !(n instanceof MenuButton)) {
                return b;
            }
            n = n.getParent();
        }
        return null;
    }
}
