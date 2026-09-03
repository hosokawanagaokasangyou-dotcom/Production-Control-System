package jp.co.pm.ai.desktop.ui;

import javafx.application.Platform;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.Button;
import javafx.scene.control.Labeled;
import javafx.scene.control.TabPane;
import javafx.scene.text.Text;

/**
 * JavaFX 26 以降、{@link Labeled} の {@code -fx-text-fill} と子 {@link Text}（LabeledText）の {@code
 * -fx-fill} が一致せず、ホバーや再レイアウトまで文字が黒のまま残ることがある。インライン {@code -fx-fill}
 * で明示する。
 */
public final class LabeledTextFillSupport {

    /** テーマ連動の中間文字色（ダークテーマでは明るい色に再定義される）。 */
    public static final String THEME_MID = "-fx-mid-text-color";

    public static final String THEME_LIGHT = "-fx-light-text-color";

    private LabeledTextFillSupport() {}

    /**
     * ノード配下の {@link Text} に {@code -fx-fill}、{@link Labeled}（{@link Button} 除く）に {@code
     * -fx-text-fill} をインライン指定する。
     */
    public static void applyCssColorToken(Node root, String cssColorToken) {
        if (root == null || cssColorToken == null || cssColorToken.isBlank()) {
            return;
        }
        String token = cssColorToken.strip();
        applyCssColorTokenRecursive(root, token);
    }

    private static void applyCssColorTokenRecursive(Node root, String token) {
        if (root instanceof Text textNode) {
            textNode.setStyle("-fx-fill: " + token + ";");
        } else if (root instanceof Labeled labeled && !(root instanceof Button)) {
            labeled.setStyle("-fx-text-fill: " + token + ";");
        }
        if (root instanceof Parent parent) {
            for (Node child : parent.getChildrenUnmodifiable()) {
                applyCssColorTokenRecursive(child, token);
            }
        }
    }

    /** {@link Button} 配下の LabeledText に fill を強制する（ボタン本体の他インラインは触らない）。 */
    public static void applyToButton(Button button, String cssColorToken) {
        if (button == null || cssColorToken == null || cssColorToken.isBlank()) {
            return;
        }
        String token = cssColorToken.strip();
        applyCssColorTokenRecursive(button, token);
        /* Skin 内の LabeledText は children に出ないことがあるため lookup も行う */
        Node labeledText = button.lookup(".text");
        if (labeledText != null) {
            labeledText.setStyle("-fx-fill: " + token + ";");
        }
    }

    /**
     * TabPane 見出し行の各タブセルへ fill を適用する。Skin 未準備時は取りこぼすため {@link
     * Platform#runLater} で再試行する。
     */
    public static void applyToTabPaneHeaders(TabPane pane, String cssColorToken) {
        if (pane == null || cssColorToken == null || cssColorToken.isBlank()) {
            return;
        }
        String token = cssColorToken.strip();
        Runnable op =
                () -> {
                    Node headersRegion = pane.lookup(".headers-region");
                    if (!(headersRegion instanceof Parent parent)) {
                        return;
                    }
                    for (Node child : parent.getChildrenUnmodifiable()) {
                        if (child.getStyleClass().contains("tab")) {
                            applyCssColorToken(child, token);
                        }
                    }
                };
        op.run();
        Platform.runLater(op);
        Platform.runLater(() -> Platform.runLater(op));
    }
}
