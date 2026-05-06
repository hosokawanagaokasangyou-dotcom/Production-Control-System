package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.EnumSet;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TreeItem;
import javafx.scene.control.TreeView;
import javafx.scene.paint.Color;

import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.MainShellTabLayoutNode;

/**
 * メインシェルタブの入れ子と色を編集する専用タブ。
 */
public final class MainShellTabOrganizerTabController {

    @FXML
    private TreeView<OrgRow> treeView;

    @FXML
    private ColorPicker colorPicker;

    private MainShellController shell;

    @FXML
    private void initialize() {
        if (colorPicker != null) {
            colorPicker.setValue(Color.web("#4a90d9"));
        }
        if (treeView != null) {
            treeView.setShowRoot(false);
            treeView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        reloadTreeFromShell();
    }

    private void reloadTreeFromShell() {
        if (treeView == null || shell == null) {
            return;
        }
        List<MainShellTabLayoutNode> layout = shell.snapshotMainShellTabLayoutNodes();
        TreeItem<OrgRow> invisibleRoot = new TreeItem<>(OrgRow.placeholder());
        if (layout.isEmpty()) {
            for (MainShellTabId id : shell.defaultMainShellTabIds()) {
                if (id == MainShellTabId.TAB_ORGANIZER) {
                    continue;
                }
                invisibleRoot.getChildren().add(leafItem(OrgRow.tab(id, "")));
            }
        } else {
            for (MainShellTabLayoutNode n : layout) {
                TreeItem<OrgRow> ti = treeItemForLayoutNode(n);
                if (ti != null) {
                    invisibleRoot.getChildren().add(ti);
                }
            }
        }
        treeView.setRoot(invisibleRoot);
        expandAll(invisibleRoot);
    }

    private static void expandAll(TreeItem<OrgRow> n) {
        n.setExpanded(true);
        for (TreeItem<OrgRow> c : n.getChildren()) {
            expandAll(c);
        }
    }

    private TreeItem<OrgRow> treeItemForLayoutNode(MainShellTabLayoutNode n) {
        if (n == null) {
            return null;
        }
        if (n.isTab()) {
            MainShellTabId id = MainShellTabId.fromKey(n.id());
            if (id == null || id == MainShellTabId.TAB_ORGANIZER) {
                return null;
            }
            return leafItem(OrgRow.tab(id, n.colorHex()));
        }
        if (n.isGroup()) {
            TreeItem<OrgRow> g = new TreeItem<>(OrgRow.group(n.title(), n.colorHex()));
            for (MainShellTabLayoutNode c : n.children()) {
                TreeItem<OrgRow> ch = treeItemForLayoutNode(c);
                if (ch != null) {
                    g.getChildren().add(ch);
                }
            }
            return g;
        }
        return null;
    }

    private static TreeItem<OrgRow> leafItem(OrgRow row) {
        TreeItem<OrgRow> t = new TreeItem<>(Objects.requireNonNull(row));
        t.setExpanded(true);
        return t;
    }

    @FXML
    private void onAddGroup() {
        if (treeView == null || treeView.getRoot() == null) {
            return;
        }
        TreeItem<OrgRow> g = new TreeItem<>(OrgRow.group("新規グループ", ""));
        g.setExpanded(true);
        treeView.getRoot().getChildren().add(g);
    }

    @FXML
    private void onGroupSelection() {
        if (treeView == null || treeView.getRoot() == null) {
            return;
        }
        ObservableList<TreeItem<OrgRow>> sel = treeView.getSelectionModel().getSelectedItems();
        if (sel == null || sel.size() < 2) {
            alert(AlertType.INFORMATION, "タブ行を2つ以上選んでください（グループ行は選べません）。");
            return;
        }
        List<TreeItem<OrgRow>> tabItems = new ArrayList<>();
        for (TreeItem<OrgRow> ti : sel) {
            if (ti == null || ti.getValue() == null) {
                continue;
            }
            OrgRow r = ti.getValue();
            if (r.kind == OrgRow.Kind.TAB) {
                tabItems.add(ti);
            }
        }
        if (tabItems.size() < 2) {
            alert(AlertType.INFORMATION, "タブ行を2つ以上選んでください。");
            return;
        }
        TreeItem<OrgRow> parentCheck = tabItems.getFirst().getParent();
        for (TreeItem<OrgRow> ti : tabItems) {
            if (ti.getParent() != parentCheck) {
                alert(AlertType.WARNING, "同じ階層のタブだけをまとめられます。");
                return;
            }
        }
        ObservableList<TreeItem<OrgRow>> siblings = parentCheck.getChildren();
        TreeItem<OrgRow> group =
                new TreeItem<>(OrgRow.group("新規グループ", ""));
        group.setExpanded(true);
        int firstIdx = siblings.indexOf(tabItems.getFirst());
        if (firstIdx < 0) {
            return;
        }
        siblings.add(firstIdx, group);
        for (TreeItem<OrgRow> ti : tabItems) {
            siblings.remove(ti);
            group.getChildren().add(ti);
        }
        treeView.getSelectionModel().clearSelection();
        treeView.getSelectionModel().select(group);
    }

    @FXML
    private void onApplySelectedColor() {
        if (treeView == null || colorPicker == null) {
            return;
        }
        Color c = colorPicker.getValue();
        if (c == null) {
            return;
        }
        String hex = toHexRgb(c);
        ObservableList<TreeItem<OrgRow>> sel = treeView.getSelectionModel().getSelectedItems();
        if (sel == null || sel.isEmpty()) {
            return;
        }
        for (TreeItem<OrgRow> ti : sel) {
            if (ti != null && ti.getValue() != null) {
                ti.getValue().colorHex = hex;
            }
        }
        treeView.refresh();
    }

    @FXML
    private void onClearSelectedColor() {
        if (treeView == null) {
            return;
        }
        ObservableList<TreeItem<OrgRow>> sel = treeView.getSelectionModel().getSelectedItems();
        if (sel == null || sel.isEmpty()) {
            return;
        }
        for (TreeItem<OrgRow> ti : sel) {
            if (ti != null && ti.getValue() != null) {
                ti.getValue().colorHex = "";
            }
        }
        treeView.refresh();
    }

    @FXML
    private void onResetFlat() {
        if (shell == null) {
            return;
        }
        shell.restoreDefaultFlatMainShellTabLayout();
        DesktopSessionStateStore.save(shell.collectDesktopSessionSnapshot());
        reloadTreeFromShell();
    }

    @FXML
    private void onApplyLayout() {
        if (shell == null || treeView == null || treeView.getRoot() == null) {
            return;
        }
        List<MainShellTabLayoutNode> layout = new ArrayList<>();
        for (TreeItem<OrgRow> ch : treeView.getRoot().getChildren()) {
            MainShellTabLayoutNode n = layoutNodeFromTreeItem(ch);
            if (n != null) {
                layout.add(n);
            }
        }
        if (!validateAllTabsOnce(layout)) {
            alert(
                    AlertType.WARNING,
                    "すべての作業タブをちょうど1回ずつ使う必要があります（不足・重複があります）。");
            return;
        }
        shell.applyMainShellTabLayoutFromOrganizer(layout);
        reloadTreeFromShell();
    }

    private static boolean validateAllTabsOnce(List<MainShellTabLayoutNode> top) {
        Set<String> seen = new HashSet<>();
        Set<String> required = new HashSet<>();
        for (MainShellTabId id : MainShellTabId.values()) {
            if (id != MainShellTabId.TAB_ORGANIZER) {
                required.add(id.key());
            }
        }
        for (MainShellTabLayoutNode n : top) {
            collectLeafKeys(n, seen);
        }
        return seen.size() == required.size() && seen.containsAll(required);
    }

    private static void collectLeafKeys(MainShellTabLayoutNode n, Set<String> out) {
        if (n.isTab()) {
            out.add(n.id());
            return;
        }
        for (MainShellTabLayoutNode c : n.children()) {
            collectLeafKeys(c, out);
        }
    }

    private MainShellTabLayoutNode layoutNodeFromTreeItem(TreeItem<OrgRow> ti) {
        if (ti == null || ti.getValue() == null) {
            return null;
        }
        OrgRow r = ti.getValue();
        if (r.kind == OrgRow.Kind.TAB) {
            return MainShellTabLayoutNode.tabNode(r.tabId.key(), nz(r.colorHex));
        }
        List<MainShellTabLayoutNode> ch = new ArrayList<>();
        for (TreeItem<OrgRow> c : ti.getChildren()) {
            MainShellTabLayoutNode n = layoutNodeFromTreeItem(c);
            if (n != null) {
                ch.add(n);
            }
        }
        String title = r.groupTitle != null && !r.groupTitle.isBlank() ? r.groupTitle : "グループ";
        return MainShellTabLayoutNode.groupNode(title, nz(r.colorHex), ch);
    }

    private static String nz(String s) {
        return s != null ? s.strip() : "";
    }

    private static String toHexRgb(Color c) {
        int r = (int) Math.round(c.getRed() * 255);
        int g = (int) Math.round(c.getGreen() * 255);
        int b = (int) Math.round(c.getBlue() * 255);
        r = Math.max(0, Math.min(255, r));
        g = Math.max(0, Math.min(255, g));
        b = Math.max(0, Math.min(255, b));
        return String.format("#%02x%02x%02x", r, g, b);
    }

    private static void alert(AlertType type, String msg) {
        Alert a = new Alert(type, msg);
        a.setHeaderText(null);
        a.showAndWait();
    }

    /** ツリー1行分（タブまたはグループ）。 */
    static final class OrgRow {
        enum Kind {
            TAB,
            GROUP
        }

        final Kind kind;
        MainShellTabId tabId;
        String groupTitle;
        String colorHex;

        private OrgRow(Kind kind, MainShellTabId tabId, String groupTitle, String colorHex) {
            this.kind = kind;
            this.tabId = tabId;
            this.groupTitle = groupTitle != null ? groupTitle : "";
            this.colorHex = colorHex != null ? colorHex : "";
        }

        static OrgRow placeholder() {
            return new OrgRow(Kind.GROUP, null, "", "");
        }

        static OrgRow tab(MainShellTabId id, String colorHex) {
            return new OrgRow(Kind.TAB, Objects.requireNonNull(id), "", nz(colorHex));
        }

        static OrgRow group(String title, String colorHex) {
            return new OrgRow(Kind.GROUP, null, title != null ? title : "", nz(colorHex));
        }

        String formatDisplay(MainShellController shell) {
            if (kind == Kind.GROUP) {
                String t = groupTitle != null && !groupTitle.isBlank() ? groupTitle : "グループ";
                String c = colorHex != null && !colorHex.isBlank() ? "  [" + colorHex + "]" : "";
                return "[グループ] " + t + c;
            }
            String base = shell != null ? shell.mainShellTabTitle(tabId) : tabId.name();
            String c = colorHex != null && !colorHex.isBlank() ? "  [" + colorHex + "]" : "";
            return base + c;
        }
    }

    /** {@link TreeView} 用セルファクトリはコントローラ初期化後にシェルへバインドしてから設定する。 */
    void installTreeCellFactory() {
        if (treeView == null) {
            return;
        }
        treeView.setCellFactory(
                tv ->
                        new javafx.scene.control.TreeCell<>() {
                            @Override
                            protected void updateItem(OrgRow row, boolean empty) {
                                super.updateItem(row, empty);
                                if (empty || row == null) {
                                    setText(null);
                                    return;
                                }
                                setText(row.formatDisplay(shell));
                            }
                        });
    }
}
