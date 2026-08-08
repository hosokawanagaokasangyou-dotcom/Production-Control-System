package jp.co.pm.ai.desktop.ui;

import javafx.scene.Node;
import javafx.scene.control.Label;

/** JavaFX ノードの表示切替。 */
public final class JavaFxNodeVisibility {

    private JavaFxNodeVisibility() {}

    public static void apply(Node node, boolean visible) {
        if (node == null) {
            return;
        }
        node.setVisible(visible);
        node.setManaged(visible);
    }

    public static void applyPlanningStageBadgePolicyNoop(Label badge, java.util.Map<String, String> ui) {
        // 段階3 バッジ非表示ポリシーは廃止
    }
}
