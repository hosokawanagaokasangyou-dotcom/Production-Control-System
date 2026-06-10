package jp.co.pm.ai.desktop.ui;

import javafx.application.Platform;
import javafx.geometry.Bounds;
import javafx.scene.Node;
import javafx.scene.control.ComboBox;
import javafx.scene.control.PopupControl;
import javafx.scene.control.Skin;
import javafx.scene.control.skin.ComboBoxListViewSkin;
import javafx.scene.control.skin.ComboBoxPopupControl;
import javafx.scene.layout.Region;
import javafx.stage.Window;

import java.lang.reflect.Method;

/**
 * {@link ComboBox} のドロップダウンをコンボ右端に揃え、左側へ広げる。
 * 右ペインのプレビュー等へ張り出す候補リスト向け。
 */
public final class ComboBoxPopupRightAlign {

    private static final String REQUEST_FORM_RECONCILIATION_ROOT_STYLE =
            "pm-request-form-reconciliation-root";

    private ComboBoxPopupRightAlign() {}

    public static void install(ComboBox<?> combo) {
        if (combo == null) {
            return;
        }
        combo.addEventHandler(ComboBox.ON_SHOWN, evt -> Platform.runLater(() -> align(combo)));
    }

    private static void align(ComboBox<?> combo) {
        if (!combo.isShowing()) {
            return;
        }
        ComboBoxListViewSkin<?> listSkin = listViewSkin(combo);
        PopupControl popup = popupFor(listSkin);
        if (popup == null) {
            return;
        }
        Node popupContent = listSkin != null ? listSkin.getPopupContent() : null;
        Bounds comboBounds = combo.localToScreen(combo.getLayoutBounds());
        if (comboBounds == null) {
            return;
        }
        double leftBound = leftClampScreenX(combo);
        double maxWidth = Math.max(0, comboBounds.getMaxX() - leftBound);
        if (maxWidth > 0) {
            constrainPopupWidth(popup, popupContent, maxWidth);
        }
        double popupWidth = popupWidth(popup, popupContent);
        if (popupWidth <= 0) {
            return;
        }
        if (maxWidth > 0 && popupWidth > maxWidth) {
            popupWidth = maxWidth;
        }
        double comboWidth = comboBounds.getWidth();
        if (popupWidth <= comboWidth) {
            return;
        }
        double desiredX = comboBounds.getMaxX() - popupWidth;
        if (desiredX < leftBound) {
            desiredX = leftBound;
        }
        popup.setX(desiredX);
    }

    private static double leftClampScreenX(ComboBox<?> combo) {
        Node node = combo;
        while (node != null) {
            if (node.getStyleClass().contains(REQUEST_FORM_RECONCILIATION_ROOT_STYLE)) {
                Bounds bounds = node.localToScreen(node.getLayoutBounds());
                if (bounds != null) {
                    return bounds.getMinX();
                }
                break;
            }
            node = node.getParent();
        }
        Window window = combo.getScene() != null ? combo.getScene().getWindow() : null;
        return window != null ? window.getX() : 0;
    }

    private static void constrainPopupWidth(PopupControl popup, Node popupContent, double maxWidth) {
        popup.setMaxWidth(maxWidth);
        if (popupContent != null) {
            applyMaxWidth(popupContent, maxWidth);
            Node listView = popupContent.lookup(".list-view");
            if (listView != null) {
                applyMaxWidth(listView, maxWidth);
            }
        }
    }

    private static void applyMaxWidth(Node node, double maxWidth) {
        if (node instanceof Region region) {
            region.setMaxWidth(maxWidth);
            double pref = region.prefWidth(-1);
            if (pref > 0 && pref > maxWidth) {
                region.setPrefWidth(maxWidth);
            }
        }
    }

    private static double popupWidth(PopupControl popup, Node popupContent) {
        double width = popup.getWidth();
        if (width > 0) {
            return width;
        }
        width = popup.prefWidth(-1);
        if (width > 0) {
            return width;
        }
        if (popupContent != null) {
            width = popupContent.prefWidth(-1);
            if (width > 0) {
                return width;
            }
            Bounds bounds = popupContent.getBoundsInLocal();
            if (bounds != null && bounds.getWidth() > 0) {
                return bounds.getWidth();
            }
        }
        return 0;
    }

    private static ComboBoxListViewSkin<?> listViewSkin(ComboBox<?> combo) {
        Skin<?> skin = combo.getSkin();
        if (skin instanceof ComboBoxListViewSkin<?> listSkin) {
            return listSkin;
        }
        return null;
    }

    private static PopupControl popupFor(ComboBoxListViewSkin<?> listSkin) {
        if (listSkin == null) {
            return null;
        }
        try {
            Method method = ComboBoxPopupControl.class.getDeclaredMethod("getPopup");
            method.setAccessible(true);
            return (PopupControl) method.invoke(listSkin);
        } catch (ReflectiveOperationException ex) {
            return null;
        }
    }
}
