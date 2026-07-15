package jp.co.pm.ai.desktop.config;

import java.util.ArrayList;
import java.util.List;

import jp.co.pm.ai.desktop.MainShellTabId;

/** 非表示中も段階3メインタブの位置・色・所属グループを保持するレイアウト補助。 */
public final class Stage3MainShellLayout {

    private static final String STAGE3_KEY = MainShellTabId.PLAN_INPUT_STAGE3.key();

    private Stage3MainShellLayout() {}

    public static List<MainShellTabLayoutNode> withoutStage3(
            List<MainShellTabLayoutNode> nodes) {
        List<MainShellTabLayoutNode> out = new ArrayList<>();
        if (nodes == null) {
            return List.of();
        }
        for (MainShellTabLayoutNode node : nodes) {
            if (node == null) {
                continue;
            }
            if (node.isTab()) {
                if (!STAGE3_KEY.equals(node.id())) {
                    out.add(node);
                }
                continue;
            }
            List<MainShellTabLayoutNode> children = withoutStage3(node.children());
            if (!children.isEmpty()) {
                out.add(MainShellTabLayoutNode.groupNode(node.title(), node.colorHex(), children));
            }
        }
        return List.copyOf(out);
    }

    public static List<MainShellTabLayoutNode> mergeHiddenTab(
            List<MainShellTabLayoutNode> visible, List<MainShellTabLayoutNode> complete) {
        List<MainShellTabLayoutNode> base =
                visible == null ? List.of() : List.copyOf(visible);
        if (containsStage3(base)) {
            return base;
        }
        Location location = locate(complete, List.of(), 0);
        if (location == null) {
            return base;
        }
        if (location.groupPath().isEmpty()) {
            return insert(base, location.topIndex(), location.node());
        }
        boolean[] inserted = {false};
        List<MainShellTabLayoutNode> nested =
                insertIntoGroupPath(
                        base,
                        location.groupPath(),
                        0,
                        location.childIndex(),
                        location.node(),
                        inserted);
        return inserted[0]
                ? nested
                : insert(nested, location.topIndex(), location.node());
    }

    private static Location locate(
            List<MainShellTabLayoutNode> nodes, List<String> groupPath, int topIndex) {
        if (nodes == null) {
            return null;
        }
        for (int i = 0; i < nodes.size(); i++) {
            MainShellTabLayoutNode node = nodes.get(i);
            int rootIndex = groupPath.isEmpty() ? i : topIndex;
            if (node.isTab() && STAGE3_KEY.equals(node.id())) {
                return new Location(groupPath, rootIndex, i, node);
            }
            if (node.isGroup()) {
                List<String> childPath = new ArrayList<>(groupPath);
                childPath.add(node.title());
                Location found =
                        locate(node.children(), List.copyOf(childPath), rootIndex);
                if (found != null) {
                    return found;
                }
            }
        }
        return null;
    }

    private static List<MainShellTabLayoutNode> insertIntoGroupPath(
            List<MainShellTabLayoutNode> nodes,
            List<String> path,
            int depth,
            int childIndex,
            MainShellTabLayoutNode stage3,
            boolean[] inserted) {
        List<MainShellTabLayoutNode> out = new ArrayList<>(nodes.size());
        for (MainShellTabLayoutNode node : nodes) {
            if (!inserted[0]
                    && node.isGroup()
                    && node.title().equals(path.get(depth))) {
                List<MainShellTabLayoutNode> children;
                if (depth == path.size() - 1) {
                    children = insert(node.children(), childIndex, stage3);
                    inserted[0] = true;
                } else {
                    children =
                            insertIntoGroupPath(
                                    node.children(),
                                    path,
                                    depth + 1,
                                    childIndex,
                                    stage3,
                                    inserted);
                }
                out.add(MainShellTabLayoutNode.groupNode(node.title(), node.colorHex(), children));
            } else {
                out.add(node);
            }
        }
        return List.copyOf(out);
    }

    private static List<MainShellTabLayoutNode> insert(
            List<MainShellTabLayoutNode> nodes, int index, MainShellTabLayoutNode node) {
        List<MainShellTabLayoutNode> out = new ArrayList<>(nodes);
        out.add(Math.max(0, Math.min(index, out.size())), node);
        return List.copyOf(out);
    }

    private static boolean containsStage3(List<MainShellTabLayoutNode> nodes) {
        for (MainShellTabLayoutNode node : nodes) {
            if (node.isTab() && STAGE3_KEY.equals(node.id())) {
                return true;
            }
            if (node.isGroup() && containsStage3(node.children())) {
                return true;
            }
        }
        return false;
    }

    private record Location(
            List<String> groupPath,
            int topIndex,
            int childIndex,
            MainShellTabLayoutNode node) {}
}
