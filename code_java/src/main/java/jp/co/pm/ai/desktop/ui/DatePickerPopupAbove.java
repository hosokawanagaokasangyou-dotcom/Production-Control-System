package jp.co.pm.ai.desktop.ui;

import java.lang.reflect.Method;
import java.util.List;

import javafx.application.Platform;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.geometry.Bounds;
import javafx.geometry.Rectangle2D;
import javafx.scene.control.ComboBoxBase;
import javafx.scene.control.DatePicker;
import javafx.scene.control.PopupControl;
import javafx.scene.control.Skin;
import javafx.scene.control.skin.ComboBoxPopupControl;
import javafx.stage.Screen;
import javafx.stage.Window;

/**
 * {@link DatePicker} のカレンダーポップアップを、余白があれば入力欄の上方へ出す。
 * 依頼書入力の投入日など、下段フィールドを隠さないため。
 */
public final class DatePickerPopupAbove {

    private DatePickerPopupAbove() {}

    public static void install(DatePicker picker) {
        if (picker == null) {
            return;
        }
        picker.addEventHandler(
                ComboBoxBase.ON_SHOWN,
                evt ->
                        Platform.runLater(
                                () -> {
                                    placeAbove(picker);
                                    Platform.runLater(() -> placeAbove(picker));
                                }));
    }

    /**
     * 画面上端をはみ出さない範囲で上方配置し、収まらないときだけ下方へ戻す。
     *
     * @return ポップアップのスクリーン Y
     */
    public static double computePopupY(
            double fieldMinY,
            double fieldMaxY,
            double popupHeight,
            double screenMinY,
            double screenMaxY) {
        if (popupHeight <= 0) {
            return fieldMaxY;
        }
        double aboveY = fieldMinY - popupHeight;
        if (aboveY >= screenMinY) {
            return aboveY;
        }
        if (fieldMaxY + popupHeight <= screenMaxY) {
            return fieldMaxY;
        }
        return Math.max(screenMinY, aboveY);
    }

    static void placeAbove(DatePicker picker) {
        if (picker == null || !picker.isShowing()) {
            return;
        }
        PopupControl popup = popupFor(picker);
        if (popup == null) {
            return;
        }
        Bounds fieldBounds = picker.localToScreen(picker.getLayoutBounds());
        if (fieldBounds == null) {
            return;
        }
        double popupHeight = popupHeight(popup);
        if (popupHeight <= 0) {
            listenOnceForHeight(picker, popup);
            return;
        }
        Rectangle2D screen = visualBoundsFor(fieldBounds);
        popup.setY(
                computePopupY(
                        fieldBounds.getMinY(),
                        fieldBounds.getMaxY(),
                        popupHeight,
                        screen.getMinY(),
                        screen.getMaxY()));
    }

    private static void listenOnceForHeight(DatePicker picker, PopupControl popup) {
        ChangeListener<Number> listener =
                new ChangeListener<>() {
                    @Override
                    public void changed(
                            ObservableValue<? extends Number> obs, Number oldH, Number newH) {
                        if (newH != null && newH.doubleValue() > 0) {
                            popup.heightProperty().removeListener(this);
                            placeAbove(picker);
                        }
                    }
                };
        popup.heightProperty().addListener(listener);
    }

    private static double popupHeight(PopupControl popup) {
        double height = popup.getHeight();
        if (height > 0) {
            return height;
        }
        height = popup.prefHeight(-1);
        if (height > 0) {
            return height;
        }
        if (popup.getScene() != null && popup.getScene().getRoot() != null) {
            height = popup.getScene().getRoot().prefHeight(-1);
            if (height > 0) {
                return height;
            }
        }
        return 0;
    }

    private static Rectangle2D visualBoundsFor(Bounds fieldBounds) {
        List<Screen> screens =
                Screen.getScreensForRectangle(
                        fieldBounds.getMinX(),
                        fieldBounds.getMinY(),
                        Math.max(1, fieldBounds.getWidth()),
                        Math.max(1, fieldBounds.getHeight()));
        if (!screens.isEmpty()) {
            return screens.get(0).getVisualBounds();
        }
        return Screen.getPrimary().getVisualBounds();
    }

    private static PopupControl popupFor(DatePicker picker) {
        Skin<?> skin = picker.getSkin();
        if (skin instanceof ComboBoxPopupControl<?> popupSkin) {
            try {
                Method method = ComboBoxPopupControl.class.getDeclaredMethod("getPopup");
                method.setAccessible(true);
                Object result = method.invoke(popupSkin);
                if (result instanceof PopupControl popup) {
                    return popup;
                }
            } catch (ReflectiveOperationException ex) {
                // リフレクション不可時は表示中ウィンドウから拾う
            }
        }
        return showingPopupOwnedBy(picker);
    }

    private static PopupControl showingPopupOwnedBy(DatePicker picker) {
        for (Window window : Window.getWindows()) {
            if (window instanceof PopupControl popup
                    && popup.isShowing()
                    && popup.getOwnerNode() == picker) {
                return popup;
            }
        }
        return null;
    }
}
