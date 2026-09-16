package jp.co.pm.ai.desktop.config;

import java.util.ArrayList;
import java.util.List;

import jp.co.pm.ai.desktop.MainShellTabId;

/**
 * 既存セッションに欠落している {@code kouchin} を {@code processingTrend} リーフ直前へ挿入する。
 * ルールの「DEFAULT 末尾追加」の意図的例外。
 */
public final class MainShellKouchinTabPlacement {

    public static final String KOUCHIN = MainShellTabId.KOUCHIN.key();
    public static final String PROCESSING_TREND = MainShellTabId.PROCESSING_TREND.key();

    private MainShellKouchinTabPlacement() {}

    /** キー列。欠落時のみ processingTrend 直前へ挿入。無ければ末尾。 */
    public static List<String> insertMissingKey(List<String> keys) {
        if (keys == null || keys.isEmpty()) {
            return List.of();
        }
        if (keys.contains(KOUCHIN)) {
            return List.copyOf(keys);
        }
        List<String> out = new ArrayList<>(keys);
        int idx = out.indexOf(PROCESSING_TREND);
        if (idx >= 0) {
            out.add(idx, KOUCHIN);
        } else {
            out.add(KOUCHIN);
        }
        return List.copyOf(out);
    }

    /** レイアウトツリー。グループ内を再帰し processingTrend 直前へ挿入。 */
    public static List<MainShellTabLayoutNode> insertMissingLeaf(List<MainShellTabLayoutNode> top) {
        if (top == null || top.isEmpty()) {
            return List.of();
        }
        if (containsKey(top, KOUCHIN)) {
            return List.copyOf(top);
        }
        List<MainShellTabLayoutNode> out = new ArrayList<>(top);
        if (insertInList(out)) {
            return List.copyOf(out);
        }
        out.add(MainShellTabLayoutNode.tabNode(KOUCHIN, "#c0504d"));
        return List.copyOf(out);
    }

    private static boolean containsKey(List<MainShellTabLayoutNode> nodes, String key) {
        for (MainShellTabLayoutNode n : nodes) {
            if (n.isTab() && key.equals(n.id())) {
                return true;
            }
            if (n.isGroup() && containsKey(n.children(), key)) {
                return true;
            }
        }
        return false;
    }

    private static boolean insertInList(List<MainShellTabLayoutNode> nodes) {
        for (int i = 0; i < nodes.size(); i++) {
            MainShellTabLayoutNode n = nodes.get(i);
            if (n.isTab() && PROCESSING_TREND.equals(n.id())) {
                nodes.add(i, MainShellTabLayoutNode.tabNode(KOUCHIN, "#c0504d"));
                return true;
            }
            if (n.isGroup()) {
                List<MainShellTabLayoutNode> ch = new ArrayList<>(n.children());
                if (insertInList(ch)) {
                    nodes.set(i, MainShellTabLayoutNode.groupNode(n.title(), n.colorHex(), ch));
                    return true;
                }
            }
        }
        return false;
    }
}
