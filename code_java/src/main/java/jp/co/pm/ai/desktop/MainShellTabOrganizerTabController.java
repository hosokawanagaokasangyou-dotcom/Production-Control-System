package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.HashSet;
import java.util.IdentityHashMap;

import javafx.application.Platform;
import javafx.beans.value.ChangeListener;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.Cursor;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.Label;
import javafx.scene.control.MultipleSelectionModel;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.control.TreeCell;
import javafx.scene.control.TreeItem;
import javafx.scene.control.TreeView;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.HBox;
import javafx.scene.Node;
import javafx.scene.layout.Pane;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.shape.Line;
import javafx.scene.shape.Polyline;
import javafx.scene.input.DataFormat;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.paint.Color;
import javafx.scene.text.TextAlignment;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.MainShellTabLayoutNode;

/**
 * メインシェルタブの入れ子と色を編集する専用タブ。
 */
public final class MainShellTabOrganizerTabController {

    private static final DataFormat ROW_MOVE_MARKER =
            new DataFormat("application/x-pm-main-shell-tab-org-move");

    /** ブラウザ型タブ列：やや細め・高めで横に多く並べる */
    private static final double PREVIEW_STRIP_CHIP_W = 132.0;
    private static final double PREVIEW_STRIP_CHIP_H = 36.0;
    private static final double PREVIEW_STRIP_OVERLAP = 50.0;

    private static final double TREE_INDENT_STEP = 22.0;
    private static final double TREE_CONNECTOR_ROW_H = 30.0;

    @FXML
    private TreeView<OrgRow> treeView;

    @FXML
    private ScrollPane previewStripScroll;

    @FXML
    private HBox previewStripBox;

    private final IdentityHashMap<TreeItem<OrgRow>, StackPane> previewStripChipByTreeItem =
            new IdentityHashMap<>();

    @FXML
    private ColorPicker colorPicker;

    @FXML
    private TextField groupNameField;

    @FXML
    private TextField tabAliasField;

    @FXML
    private CheckBox headerGlowCheck;

    private MainShellController shell;

    private boolean suppressOrganizerChromeListeners;

    private boolean headerGlowListenerHooked;

    /** ドラッグ開始セルと {@link Dragboard} を対応付けるための作業領域（ツリー行の移動用）。 */
    private TreeItem<OrgRow> dragSourceItem;

    private ChangeListener<TreeItem<OrgRow>> treeSelectionListener;

    private boolean treeSelectionRefreshHooked;

    @FXML
    private void initialize() {
        if (colorPicker != null) {
            colorPicker.setValue(Color.web("#4a90d9"));
        }
        if (treeView != null) {
            treeView.setShowRoot(false);
            treeView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        }
        if (groupNameField != null) {
            groupNameField.setDisable(true);
            groupNameField.setOnAction(e -> onApplyGroupName());
        }
        if (tabAliasField != null) {
            tabAliasField.setDisable(true);
            tabAliasField.setOnAction(e -> onApplyTabAlias());
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        if (treeView != null && treeSelectionListener == null) {
            treeSelectionListener =
                    (obs, prev, cur) -> {
                        syncOrganizerSideFields();
                        syncPreviewStripChipHighlight();
                    };
            treeView.getSelectionModel().selectedItemProperty().addListener(treeSelectionListener);
        }
        if (treeView != null && !treeSelectionRefreshHooked) {
            treeSelectionRefreshHooked = true;
            treeView.getSelectionModel()
                    .getSelectedItems()
                    .addListener(
                            (ListChangeListener<TreeItem<OrgRow>>)
                                    c -> {
                                        if (treeView != null) {
                                            treeView.refresh();
                                        }
                                    });
        }
        installHeaderGlowControls();
        reloadTreeFromShell();
    }

    /** セッション復元後にチェック状態だけシェルと揃える。 */
    void syncHeaderGlowCheckFromShell() {
        if (headerGlowCheck == null || shell == null) {
            return;
        }
        suppressOrganizerChromeListeners = true;
        try {
            headerGlowCheck.setSelected(shell.isMainShellTabOrganizerHeaderGlowEnabled());
        } finally {
            suppressOrganizerChromeListeners = false;
        }
    }

    private void installHeaderGlowControls() {
        if (headerGlowCheck != null && shell != null && !headerGlowListenerHooked) {
            headerGlowListenerHooked = true;
            headerGlowCheck
                    .selectedProperty()
                    .addListener(
                            (obs, prev, selected) -> {
                                if (suppressOrganizerChromeListeners) {
                                    return;
                                }
                                shell.setMainShellTabOrganizerHeaderGlowEnabled(selected);
                                shell.refreshMainShellTabHeaderChromeFromStoredColors();
                                if (treeView != null) {
                                    treeView.refresh();
                                }
                                rebuildPreviewStrip();
                                DesktopSessionState snap = shell.collectDesktopSessionSnapshot();
                                DesktopSessionStateStore.save(snap);
                            });
        }
        syncHeaderGlowCheckFromShell();
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
        syncOrganizerSideFields();
        rebuildPreviewStrip();
    }

    private void syncOrganizerSideFields() {
        syncGroupNameField();
        syncTabAliasField();
    }

    private void syncGroupNameField() {
        if (groupNameField == null || treeView == null) {
            return;
        }
        TreeItem<OrgRow> sel = treeView.getSelectionModel().getSelectedItem();
        if (sel != null
                && sel.getValue() != null
                && sel.getValue().kind == OrgRow.Kind.GROUP) {
            groupNameField.setDisable(false);
            groupNameField.setText(sel.getValue().groupTitle);
        } else {
            groupNameField.setDisable(true);
            groupNameField.clear();
        }
    }

    private void syncTabAliasField() {
        if (tabAliasField == null || treeView == null || shell == null) {
            return;
        }
        ObservableList<TreeItem<OrgRow>> multi =
                treeView.getSelectionModel().getSelectedItems();
        if (multi != null
                && multi.size() == 1
                && multi.getFirst() != null
                && multi.getFirst().getValue() != null
                && multi.getFirst().getValue().kind == OrgRow.Kind.TAB) {
            MainShellTabId tid = multi.getFirst().getValue().tabId;
            tabAliasField.setDisable(false);
            tabAliasField.setText(shell.mainShellTabTitleAliasStored(tid));
            tabAliasField.setPromptText("既定: " + shell.mainShellTabBaselineTitle(tid));
            return;
        }
        tabAliasField.setDisable(true);
        tabAliasField.clear();
        tabAliasField.setPromptText("");
    }

    @FXML
    private void onApplyGroupName() {
        if (treeView == null || groupNameField == null) {
            return;
        }
        TreeItem<OrgRow> sel = treeView.getSelectionModel().getSelectedItem();
        if (sel == null
                || sel.getValue() == null
                || sel.getValue().kind != OrgRow.Kind.GROUP) {
            alert(AlertType.INFORMATION, "グループ行を1つ選んでください。");
            return;
        }
        String t = groupNameField.getText() != null ? groupNameField.getText().strip() : "";
        OrgRow prev = sel.getValue();
        sel.setValue(OrgRow.group(t, prev != null ? prev.colorHex : ""));
    }

    @FXML
    private void onApplyTabAlias() {
        if (shell == null || treeView == null || tabAliasField == null) {
            return;
        }
        TreeItem<OrgRow> sel = treeView.getSelectionModel().getSelectedItem();
        if (sel == null
                || sel.getValue() == null
                || sel.getValue().kind != OrgRow.Kind.TAB) {
            alert(AlertType.INFORMATION, "タブ行を1つ選んでください。");
            return;
        }
        String raw = tabAliasField.getText();
        shell.setMainShellTabDisplayAlias(sel.getValue().tabId, raw);
        DesktopSessionStateStore.save(shell.collectDesktopSessionSnapshot());
        treeView.refresh();
        syncTabAliasField();
        rebuildPreviewStrip();
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
        rebuildPreviewStrip();
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
        List<TreeItem<OrgRow>> keep = new ArrayList<>(sel);
        for (TreeItem<OrgRow> ti : keep) {
            replaceRowColorHex(ti, hex);
        }
        if (shell != null && treeView.getRoot() != null) {
            shell.syncMainShellTabHeaderColorsFromOrganizerTree(treeView.getRoot());
        }
        Platform.runLater(() -> restoreOrganizerTreeSelection(keep));
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
        List<TreeItem<OrgRow>> keep = new ArrayList<>(sel);
        for (TreeItem<OrgRow> ti : keep) {
            replaceRowColorHex(ti, "");
        }
        if (shell != null && treeView.getRoot() != null) {
            shell.syncMainShellTabHeaderColorsFromOrganizerTree(treeView.getRoot());
        }
        Platform.runLater(() -> restoreOrganizerTreeSelection(keep));
    }

    /** {@link TreeItem#setValue} やシェル同期のあとでも選択が維持されるようにする。 */
    private void restoreOrganizerTreeSelection(List<TreeItem<OrgRow>> items) {
        if (treeView == null || items == null || items.isEmpty()) {
            return;
        }
        MultipleSelectionModel<TreeItem<OrgRow>> sm = treeView.getSelectionModel();
        sm.clearSelection();
        for (TreeItem<OrgRow> ti : items) {
            if (ti != null) {
                sm.select(ti);
            }
        }
        treeView.requestFocus();
        rebuildPreviewStrip();
    }

    /**
     * {@link OrgRow} のフィールドを直接書き換えると {@link TreeItem} が変更を検知せず行が描画更新されないことがあるため、
     * 置き換え後の行で {@link TreeItem#setValue} する。
     */
    private static void replaceRowColorHex(TreeItem<OrgRow> ti, String hex) {
        if (ti == null) {
            return;
        }
        OrgRow r = ti.getValue();
        if (r == null) {
            return;
        }
        String h = hex != null ? hex : "";
        if (r.kind == OrgRow.Kind.TAB) {
            ti.setValue(OrgRow.tab(r.tabId, h));
        } else {
            ti.setValue(OrgRow.group(r.groupTitle, h));
        }
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
        mergePendingOrganizerFieldsIntoModel();
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
                    "すべての作業タブをちょうど1回ずつ使う必要があります（不足・重複があります）。\n"
                            + leafKeyMismatchDetail(layout));
            return;
        }
        if (!shell.applyMainShellTabLayoutFromOrganizer(layout)) {
            alert(
                    AlertType.WARNING,
                    "メイン画面上部への反映に失敗しました。タブキーの一覧が一致していません。\n"
                            + leafKeyMismatchDetail(layout));
            return;
        }
        reloadTreeFromShell();
    }

    /**
     * 「名前を反映」「別名を反映」を押し忘れたときでも、構成適用で入力欄の内容を取り込む。
     */
    private void mergePendingOrganizerFieldsIntoModel() {
        if (treeView == null || shell == null) {
            return;
        }
        if (groupNameField != null && !groupNameField.isDisable()) {
            TreeItem<OrgRow> sel = treeView.getSelectionModel().getSelectedItem();
            if (sel != null
                    && sel.getValue() != null
                    && sel.getValue().kind == OrgRow.Kind.GROUP) {
                String t = groupNameField.getText() != null ? groupNameField.getText().strip() : "";
                OrgRow r = sel.getValue();
                if (r != null && r.kind == OrgRow.Kind.GROUP) {
                    sel.setValue(OrgRow.group(t, r.colorHex));
                }
            }
        }
        if (tabAliasField != null && !tabAliasField.isDisable()) {
            TreeItem<OrgRow> sel = treeView.getSelectionModel().getSelectedItem();
            if (sel != null
                    && sel.getValue() != null
                    && sel.getValue().kind == OrgRow.Kind.TAB) {
                shell.setMainShellTabDisplayAlias(sel.getValue().tabId, tabAliasField.getText());
            }
        }
    }

    private static String leafKeyMismatchDetail(List<MainShellTabLayoutNode> top) {
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
        Set<String> missing = new HashSet<>(required);
        missing.removeAll(seen);
        Set<String> extra = new HashSet<>(seen);
        extra.removeAll(required);
        StringBuilder sb = new StringBuilder();
        if (!missing.isEmpty()) {
            sb.append("不足キー: ").append(missing).append("\n");
        }
        if (!extra.isEmpty()) {
            sb.append("余分キー: ").append(extra).append("\n");
        }
        if (seen.size() != required.size() && missing.isEmpty() && extra.isEmpty()) {
            sb.append("(リーフ数 ").append(seen.size()).append(" / 期待 ").append(required.size()).append(")");
        }
        return sb.length() > 0 ? sb.toString().strip() : "";
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

        /** ツリー右側テキスト（プレビューチップとは別）—色コードはチップ側で示すため省略。 */
        String treeDetailWithoutHex(MainShellController shellCtl) {
            if (kind == Kind.GROUP) {
                String t = groupTitle != null && !groupTitle.isBlank() ? groupTitle : "グループ";
                return "[グループ] " + t;
            }
            return shellCtl != null ? shellCtl.mainShellTabTitle(tabId) : tabId.name();
        }

        /** ツリー色ピル／ブラウザ型プレビュー用の短い見出し（[グループ] プレフィックスなし）。 */
        String treePillPrimaryLabel(MainShellController shellCtl) {
            if (kind == Kind.GROUP) {
                return groupTitle != null && !groupTitle.isBlank() ? groupTitle : "グループ";
            }
            return shellCtl != null ? shellCtl.mainShellTabTitle(tabId) : tabId.name();
        }
    }

    private void rebuildPreviewStrip() {
        if (previewStripBox == null || treeView == null || treeView.getRoot() == null || shell == null) {
            return;
        }
        previewStripBox.getChildren().clear();
        previewStripChipByTreeItem.clear();
        previewStripBox.setSpacing(-PREVIEW_STRIP_OVERLAP);
        List<TreeItem<OrgRow>> tabs = new ArrayList<>();
        collectTabTreeItemsInOrder(treeView.getRoot(), tabs);
        for (TreeItem<OrgRow> tabItem : tabs) {
            StackPane chip = createPreviewStripChip(tabItem);
            previewStripChipByTreeItem.put(tabItem, chip);
            previewStripBox.getChildren().add(chip);
        }
        syncPreviewStripChipHighlight();
        if (previewStripScroll != null) {
            Platform.runLater(
                    () -> {
                        previewStripScroll.applyCss();
                        previewStripScroll.layout();
                    });
        }
    }

    private static void collectTabTreeItemsInOrder(TreeItem<OrgRow> node, List<TreeItem<OrgRow>> out) {
        if (node == null || node.getValue() == null) {
            return;
        }
        OrgRow r = node.getValue();
        if (r.kind == OrgRow.Kind.TAB) {
            out.add(node);
        }
        for (TreeItem<OrgRow> ch : node.getChildren()) {
            collectTabTreeItemsInOrder(ch, out);
        }
    }

    private StackPane createPreviewStripChip(TreeItem<OrgRow> tabItem) {
        OrgRow row = tabItem.getValue();
        StackPane chip = new StackPane();
        chip.setCursor(Cursor.HAND);
        chip.setPrefSize(PREVIEW_STRIP_CHIP_W, PREVIEW_STRIP_CHIP_H);
        chip.setMinSize(PREVIEW_STRIP_CHIP_W, PREVIEW_STRIP_CHIP_H);
        chip.setMaxHeight(PREVIEW_STRIP_CHIP_H);
        chip.getStyleClass().setAll("pm-org-strip-chip", "pm-org-strip-browser-tab");

        String fullTitle = row.treeDetailWithoutHex(shell);
        String pillTitle = row.treePillPrimaryLabel(shell);
        Label lab = new Label(pillTitle);
        lab.setWrapText(false);
        lab.setTextAlignment(TextAlignment.CENTER);
        lab.setAlignment(Pos.CENTER);
        lab.setMaxWidth(PREVIEW_STRIP_CHIP_W - 10);
        lab.getStyleClass().add("pm-org-strip-chip-label");
        Tooltip.install(chip, new Tooltip(fullTitle));

        String hex = row.colorHex;
        if (hex != null && !hex.isBlank()) {
            chip.setStyle(
                    shell.tabOrganizerPreviewChipSurfaceStyle(hex)
                            + " -fx-background-radius: 10 10 0 0; -fx-border-radius: 10 10 0 0;");
            // 添付イメージどおりブラウザタブは白文字基調（メインシェル選択タブの見えに寄せる）
            lab.setStyle("-fx-text-fill: #f8fafc; -fx-font-size: 10.5px; -fx-font-weight: bold;");
        } else {
            chip.getStyleClass().add("pm-org-tab-preview-chip-empty");
            chip.setStyle("");
            lab.setStyle(
                    "-fx-font-size: 10.5px; -fx-font-weight: bold; -fx-text-fill: "
                            + shell.tabOrganizerPreviewChipLabelTextFill("")
                            + ";");
        }
        chip.getChildren().setAll(lab);

        chip.setOnMouseClicked(
                e -> {
                    if (treeView == null) {
                        return;
                    }
                    treeView.getSelectionModel().clearSelection();
                    treeView.getSelectionModel().select(tabItem);
                    int rowIdx = treeView.getRow(tabItem);
                    if (rowIdx >= 0) {
                        treeView.scrollTo(rowIdx);
                    }
                    e.consume();
                });
        return chip;
    }

    private void syncPreviewStripChipHighlight() {
        if (previewStripChipByTreeItem.isEmpty() || treeView == null) {
            return;
        }
        ObservableList<TreeItem<OrgRow>> sel = treeView.getSelectionModel().getSelectedItems();
        Set<TreeItem<OrgRow>> selected =
                sel == null || sel.isEmpty() ? Set.of() : new HashSet<>(sel);
        for (Map.Entry<TreeItem<OrgRow>, StackPane> e : previewStripChipByTreeItem.entrySet()) {
            boolean on = selected.contains(e.getKey());
            StackPane chip = e.getValue();
            if (on) {
                if (!chip.getStyleClass().contains("pm-org-strip-chip-selected")) {
                    chip.getStyleClass().add("pm-org-strip-chip-selected");
                }
            } else {
                chip.getStyleClass().remove("pm-org-strip-chip-selected");
            }
        }
    }

    /** ルート直下から見て各階層で「末っ子か」を末端から順に並べたリスト（接続線描画用）。 */
    private static List<Boolean> lastAncestorChain(
            TreeItem<OrgRow> item, TreeItem<OrgRow> invisibleRoot) {
        LinkedList<Boolean> chain = new LinkedList<>();
        TreeItem<OrgRow> n = item;
        while (n != null && n.getParent() != null && n.getParent() != invisibleRoot) {
            TreeItem<OrgRow> p = n.getParent();
            ObservableList<TreeItem<OrgRow>> sibs = p.getChildren();
            boolean last = !sibs.isEmpty() && sibs.indexOf(n) == sibs.size() - 1;
            chain.addFirst(last);
            n = p;
        }
        return new ArrayList<>(chain);
    }

    /** ツリー接続線（縦＋最深段の L 字）。幅は深さ×{@link #TREE_INDENT_STEP}。 */
    private static Pane buildTreeConnectorPane(List<Boolean> chain, double rowH) {
        Pane pane = new Pane();
        int depth = chain.size();
        double w = depth * TREE_INDENT_STEP;
        pane.setPrefWidth(w);
        pane.setMinWidth(w);
        if (depth <= 0) {
            return pane;
        }
        double h = Math.max(26, rowH);
        double midY = h * 0.42;
        Color lc = Color.web("#5ba9ff");
        for (int lev = 0; lev < depth; lev++) {
            double x = lev * TREE_INDENT_STEP + TREE_INDENT_STEP / 2.0;
            boolean lastAtLevel = chain.get(lev);
            if (lev < depth - 1) {
                double yEnd = lastAtLevel ? midY : h;
                Line vl = new Line(x, 0, x, yEnd);
                vl.setStroke(lc);
                vl.setStrokeWidth(1.35);
                pane.getChildren().add(vl);
            } else {
                Polyline elbow = new Polyline(x, 0, x, midY, x + 13, midY);
                elbow.setStroke(lc);
                elbow.setFill(Color.TRANSPARENT);
                elbow.setStrokeWidth(1.35);
                pane.getChildren().add(elbow);
            }
        }
        return pane;
    }

    /** {@link TreeView} 用セルファクトリはコントローラ初期化後にシェルへバインドしてから設定する。 */
    void installTreeCellFactory() {
        if (treeView == null) {
            return;
        }
        treeView.setCellFactory(tv -> createDnDTreeCell());
    }

    private TreeCell<OrgRow> createDnDTreeCell() {
        TreeCell<OrgRow> cell =
                new TreeCell<>() {
                    @Override
                    protected void updateItem(OrgRow row, boolean empty) {
                        super.updateItem(row, empty);
                        getStyleClass().removeIf(st -> st.startsWith("pm-org-depth-"));
                        if (empty || row == null) {
                            setText(null);
                            setGraphic(null);
                            return;
                        }
                        TreeView<OrgRow> tv = getTreeView();
                        TreeItem<OrgRow> ti = getTreeItem();
                        TreeItem<OrgRow> invisibleRoot = tv != null ? tv.getRoot() : null;
                        List<Boolean> chain =
                                invisibleRoot != null && ti != null
                                        ? lastAncestorChain(ti, invisibleRoot)
                                        : List.of();
                        int depth = chain.size();
                        getStyleClass().add("pm-org-depth-" + Math.min(depth, 12));

                        Node indentGraphic;
                        if (depth > 0) {
                            indentGraphic = buildTreeConnectorPane(chain, TREE_CONNECTOR_ROW_H);
                        } else {
                            Region z = new Region();
                            z.setMinWidth(0);
                            z.setPrefWidth(0);
                            z.setMaxWidth(0);
                            indentGraphic = z;
                        }

                        StackPane pill = new StackPane();
                        pill.setMinHeight(30);
                        pill.setPrefHeight(32);
                        pill.getStyleClass().setAll("pm-org-tree-pill");
                        String pillTitle = row.treePillPrimaryLabel(shell);
                        Label lab = new Label(pillTitle);
                        lab.setWrapText(false);
                        lab.setMaxWidth(Double.MAX_VALUE);
                        lab.setAlignment(Pos.CENTER_LEFT);
                        lab.getStyleClass().add("pm-org-tree-pill-label");
                        String hx = row.colorHex;
                        if (hx != null && !hx.isBlank() && shell != null) {
                            pill.setStyle(
                                    shell.tabOrganizerPreviewChipSurfaceStyle(hx)
                                            + " -fx-background-radius: 8 8 8 8; -fx-border-radius: 8 8 8 8;");
                            lab.setStyle(
                                    "-fx-text-fill: #f8fafc; -fx-font-size: 11px; -fx-font-weight: bold;");
                            pill.getStyleClass().remove("pm-org-tree-pill-empty");
                        } else {
                            pill.setStyle("");
                            pill.getStyleClass().add("pm-org-tree-pill-empty");
                            lab.setStyle(
                                    "-fx-font-size: 11px; -fx-font-weight: bold; -fx-text-fill: "
                                            + (shell != null
                                                    ? shell.tabOrganizerPreviewChipLabelTextFill("")
                                                    : "#94a3b8")
                                            + ";");
                        }
                        Tooltip.install(
                                pill,
                                new Tooltip(row.treeDetailWithoutHex(shell)));
                        pill.getChildren().setAll(lab);
                        HBox.setHgrow(pill, Priority.ALWAYS);

                        boolean sel =
                                tv != null
                                        && tv.getSelectionModel()
                                                .getSelectedItems()
                                                .contains(ti);
                        pill.getStyleClass().remove("pm-org-tree-pill-selected");
                        if (sel) {
                            pill.getStyleClass().add("pm-org-tree-pill-selected");
                        }

                        HBox rowBox = new HBox(6);
                        rowBox.setAlignment(Pos.CENTER_LEFT);
                        rowBox.getChildren().addAll(indentGraphic, pill);
                        setText(null);
                        setGraphic(rowBox);
                    }
                };
        cell.setOnDragDetected(
                event -> {
                    TreeItem<OrgRow> dragged = cell.getTreeItem();
                    if (dragged == null || dragged.getValue() == null) {
                        return;
                    }
                    Dragboard db = cell.startDragAndDrop(TransferMode.MOVE);
                    dragSourceItem = dragged;
                    ClipboardContent cc = new ClipboardContent();
                    cc.put(ROW_MOVE_MARKER, "move");
                    db.setContent(cc);
                    event.consume();
                });
        cell.setOnDragOver(
                event -> {
                    if (event.getGestureSource() != cell
                            && event.getDragboard().hasContent(ROW_MOVE_MARKER)
                            && dragSourceItem != null) {
                        TreeItem<OrgRow> target = cell.getTreeItem();
                        if (canAcceptDrop(dragSourceItem, target)) {
                            event.acceptTransferModes(TransferMode.MOVE);
                        }
                    }
                    event.consume();
                });
        cell.setOnDragDropped(
                event -> {
                    Dragboard db = event.getDragboard();
                    boolean success = false;
                    if (db.hasContent(ROW_MOVE_MARKER) && dragSourceItem != null) {
                        TreeItem<OrgRow> target = cell.getTreeItem();
                        success = performDrop(dragSourceItem, target);
                    }
                    event.setDropCompleted(success);
                    event.consume();
                });
        cell.setOnDragDone(
                event -> {
                    dragSourceItem = null;
                    event.consume();
                });
        return cell;
    }

    /**
     * {@code target} がグループならその子の末尾へ、タブなら同一親リストのその手前へ移動する。
     */
    private boolean performDrop(TreeItem<OrgRow> source, TreeItem<OrgRow> target) {
        if (!canAcceptDrop(source, target)) {
            return false;
        }
        OrgRow tv = target.getValue();
        TreeItem<OrgRow> oldParent = source.getParent();
        if (oldParent == null) {
            return false;
        }

        boolean intoGroup = tv.kind == OrgRow.Kind.GROUP;

        TreeItem<OrgRow> newParent;
        int insertIndex;

        if (intoGroup) {
            newParent = target;
            insertIndex = -1;
        } else {
            newParent = target.getParent();
            if (newParent == null) {
                return false;
            }
            insertIndex = newParent.getChildren().indexOf(target);
            if (insertIndex < 0) {
                return false;
            }
        }

        int oldIndex = oldParent.getChildren().indexOf(source);
        if (oldIndex < 0) {
            return false;
        }

        oldParent.getChildren().remove(source);

        if (intoGroup) {
            insertIndex = newParent.getChildren().size();
        } else {
            if (oldParent == newParent && oldIndex < insertIndex) {
                insertIndex--;
            }
        }

        newParent.getChildren().add(insertIndex, source);

        treeView.getSelectionModel().clearSelection();
        treeView.getSelectionModel().select(source);
        treeView.refresh();
        syncOrganizerSideFields();
        rebuildPreviewStrip();
        return true;
    }

    private static boolean canAcceptDrop(TreeItem<OrgRow> source, TreeItem<OrgRow> target) {
        if (source == null || target == null || source == target) {
            return false;
        }
        if (target.getValue() == null) {
            return false;
        }
        return !isStrictDescendant(source, target);
    }

    /** {@code target} が {@code ancestor} の（自身を含む）配下にあれば真。 */
    private static boolean isStrictDescendant(TreeItem<OrgRow> ancestor, TreeItem<OrgRow> node) {
        TreeItem<OrgRow> p = node;
        while (p != null) {
            if (p == ancestor) {
                return true;
            }
            p = p.getParent();
        }
        return false;
    }
}
