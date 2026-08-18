package jp.co.pm.ai.desktop.ui;

import javafx.geometry.Pos;
import javafx.scene.control.Label;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;

/** ソース拡張子不正時に表を暗転し、エラーメッセージを重ねる。 */
public final class SourceExtensionErrorOverlay {

    public static final String STYLE_CLASS = "source-extension-error-overlay";

    private SourceExtensionErrorOverlay() {}

    public static void show(StackPane host, String message) {
        clear(host);
        if (host == null) {
            return;
        }
        Label label = new Label(message != null ? message : "ソースファイルの拡張子が不正です。");
        label.setWrapText(true);
        label.setMaxWidth(520);
        label.getStyleClass().add("source-extension-error-overlay-label");
        VBox box = new VBox(label);
        box.setAlignment(Pos.CENTER);
        box.getStyleClass().add(STYLE_CLASS);
        box.setMouseTransparent(false);
        StackPane.setAlignment(box, Pos.CENTER);
        host.getChildren().add(box);
    }

    public static void clear(StackPane host) {
        if (host == null) {
            return;
        }
        host.getChildren()
                .removeIf(
                        n ->
                                n.getStyleClass() != null
                                        && n.getStyleClass().contains(STYLE_CLASS));
    }

    public static boolean isShowing(StackPane host) {
        if (host == null) {
            return false;
        }
        return host.getChildren().stream()
                .anyMatch(
                        n ->
                                n.getStyleClass() != null
                                        && n.getStyleClass().contains(STYLE_CLASS));
    }
}
