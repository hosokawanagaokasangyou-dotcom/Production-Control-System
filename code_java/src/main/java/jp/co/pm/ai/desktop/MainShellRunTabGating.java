package jp.co.pm.ai.desktop;

import java.util.function.Predicate;

import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;

/** 段階処理中の操作可否を、ネストされたメインシェルタブへ再帰適用する。 */
final class MainShellRunTabGating {

    private MainShellRunTabGating() {}

    static void apply(TabPane pane, boolean busy, Predicate<Tab> operableLeaf) {
        apply(pane, busy, operableLeaf, null);
    }

    /**
     * @param preferredLeaf busy 中に先に選択しておく葉タブ（例: 実行・ログ）。無効化時の
     *     TabPane 自動遷移がリモート等へ寄るのを防ぐ。
     */
    static void apply(
            TabPane pane, boolean busy, Predicate<Tab> operableLeaf, Tab preferredLeaf) {
        if (pane == null || operableLeaf == null) {
            return;
        }
        if (busy && preferredLeaf != null) {
            // 無効化前に希望タブへ寄せる（選択中タブが disable されると別タブへ自動遷移するため）
            selectInTree(pane, preferredLeaf);
        }
        applyDisableRecursive(pane, busy, operableLeaf);
        if (busy && preferredLeaf != null) {
            // disable 後の自動遷移（操作可能なリモート等）を打ち消し、実行・ログへ戻す
            selectInTree(pane, preferredLeaf);
        }
    }

    private static void applyDisableRecursive(
            TabPane pane, boolean busy, Predicate<Tab> operableLeaf) {
        for (Tab tab : pane.getTabs()) {
            boolean operable = isOperable(tab, operableLeaf);
            tab.setDisable(busy && !operable);
            if (tab.getContent() instanceof TabPane inner) {
                applyDisableRecursive(inner, busy, operableLeaf);
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

    static Tab effectiveLeaf(Tab rootSelected) {
        if (rootSelected == null) {
            return null;
        }
        if (rootSelected.getContent() instanceof TabPane inner) {
            Tab innerSel = inner.getSelectionModel().getSelectedItem();
            if (innerSel != null) {
                return effectiveLeaf(innerSel);
            }
            if (!inner.getTabs().isEmpty()) {
                return effectiveLeaf(inner.getTabs().getFirst());
            }
            return null;
        }
        return rootSelected;
    }

    static boolean selectInTree(TabPane pane, Tab target) {
        if (pane == null || target == null) {
            return false;
        }
        for (Tab t : pane.getTabs()) {
            if (t == target) {
                pane.getSelectionModel().select(t);
                return true;
            }
        }
        for (Tab t : pane.getTabs()) {
            if (t.getContent() instanceof TabPane inner) {
                if (selectInTree(inner, target)) {
                    pane.getSelectionModel().select(t);
                    return true;
                }
            }
        }
        return false;
    }
}
