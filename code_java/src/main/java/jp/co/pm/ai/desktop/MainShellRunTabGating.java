package jp.co.pm.ai.desktop;

import java.util.function.Predicate;

import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;

/** 段階処理中の操作可否を、ネストされたメインシェルタブへ再帰適用する。 */
final class MainShellRunTabGating {

    private MainShellRunTabGating() {}

    static void apply(TabPane pane, boolean busy, Predicate<Tab> operableLeaf) {
        if (pane == null || operableLeaf == null) {
            return;
        }
        for (Tab tab : pane.getTabs()) {
            boolean operable = isOperable(tab, operableLeaf);
            tab.setDisable(busy && !operable);
            if (tab.getContent() instanceof TabPane inner) {
                apply(inner, busy, operableLeaf);
            }
        }
    }

    static boolean isOperable(Tab tab, Predicate<Tab> operableLeaf) {
        if (tab == null || operableLeaf == null) {
            return false;
        }
        if (tab.getContent() instanceof TabPane inner) {
            return inner.getTabs().stream()
                    .anyMatch(child -> isOperable(child, operableLeaf));
        }
        return operableLeaf.test(tab);
    }
}
