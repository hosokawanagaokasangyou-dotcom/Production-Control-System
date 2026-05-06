package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

import javafx.application.Platform;
import javafx.beans.value.ChangeListener;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.Label;
import javafx.scene.control.MultipleSelectionModel;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Slider;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.control.TreeItem;
import javafx.scene.control.TreeView;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.HBox;
import javafx.scene.Node;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.input.DataFormat;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.paint.Color;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.MainShellTabLayoutNode;

/**
 * メインシェルタブの入れ子と色を編集する専用タブ。
 */
public final class MainShellTabOrganizerTabController {

    private static final DataFormat ROW_MOVE_MARKER =
            new DataFormat("application/x-pm-main-shell-tab-org-move");

    /** 親の右の子行同士の縦隙間（重ならないようにフル高さで積む）。 */
    private static final double CHILD_STRIP_VERTICAL_GAP = 4.0;

    /** 子行ごとに左へ積み上げるインデント（階段状の「半歩ずらし」用、横方向のみ）。 */
    private static final double CHILD_STRIP_HORIZONTAL_STAGGER = 8.0;

    @FXML
    private TreeView<OrgRow> treeView;

    @FXML
    private VBox organizerVisualRoot;

    @FXML
    private ColorPicker colorPicker;

    @FXML
    private TextField groupNameField;

    @FXML
    private TextField tabAliasField;

    @FXML
    private CheckBox headerGlowCheck;

    @FXML
    private Slider headerGlowStrengthSlider;

    @FXML
    private Label headerGlowStrengthLabel;

    private MainShellController shell;

    private boolean suppressOrganizerChromeListeners;

    private boolean headerGlowListenerHooked;

    private boolean headerGlowStrengthListenerHooked;

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
                        rebuildOrganizerVisualTree();
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
                                        syncOrganizerSideFields();
                                        rebuildOrganizerVisualTree();
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

    /** グローのオンオフ・強度スライダーをシェル値と同期する（セッション復元・ツリー選択時）。 */
    void syncHeaderGlowControlsFromShell() {
        syncHeaderGlowCheckFromShell();
        syncHeaderGlowStrengthFromShell();
    }

    private void syncHeaderGlowStrengthFromShell() {
        if (headerGlowStrengthSlider == null || shell == null) {
            return;
        }
        suppressOrganizerChromeListeners = true;
        try {
            headerGlowStrengthSlider.setValue(shell.getMainShellTabOrganizerHeaderGlowStrength() * 100.0);
            updateHeaderGlowStrengthLabelText();
        } finally {
            suppressOrganizerChromeListeners = false;
        }
    }

    private void updateHeaderGlowStrengthLabelText() {
        if (headerGlowStrengthLabel != null && headerGlowStrengthSlider != null) {
            headerGlowStrengthLabel.setText(
                    String.format(Locale.JAPAN, "%.0f%%", headerGlowStrengthSlider.getValue()));
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
                                rebuildOrganizerVisualTree();
                                DesktopSessionState snap = shell.collectDesktopSessionSnapshot();
                                DesktopSessionStateStore.save(snap);
                            });
        }
        installHeaderGlowStrengthSlider();
        syncHeaderGlowControlsFromShell();
    }

    private void installHeaderGlowStrengthSlider() {
        if (headerGlowStrengthSlider == null || shell == null || headerGlowStrengthListenerHooked) {
            return;
        }
        headerGlowStrengthListenerHooked = true;
        headerGlowStrengthSlider
                .valueProperty()
                .addListener(
                        (obs, prev, n) -> {
                            if (suppressOrganizerChromeListeners) {
                                return;
                            }
                            shell.setMainShellTabOrganizerHeaderGlowStrength(n.doubleValue() / 100.0);
                            shell.refreshMainShellTabHeaderChromeFromStoredColors();
                            rebuildOrganizerVisualTree();
                            updateHeaderGlowStrengthLabelText();
                            DesktopSessionStateStore.save(shell.collectDesktopSessionSnapshot());
                        });
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
        rebuildOrganizerVisualTree();
    }

    private void syncOrganizerSideFields() {
        syncGroupNameField();
        syncTabAliasField();
        syncColorPickerFromSelection();
        syncHeaderGlowControlsFromShell();
    }

    /** 単一選択時、見出し色ピッカーを選択行の {@link OrgRow#colorHex} に合わせる。 */
    private void syncColorPickerFromSelection() {
        if (colorPicker == null || treeView == null) {
            return;
        }
        ObservableList<TreeItem<OrgRow>> multi =
                treeView.getSelectionModel().getSelectedItems();
        if (multi == null || multi.size() != 1) {
            return;
        }
        TreeItem<OrgRow> ti = multi.getFirst();
        if (ti == null || ti.getValue() == null) {
            return;
        }
        String hx = ti.getValue().colorHex;
        hx = hx != null ? hx.strip() : "";
        try {
            if (!hx.isBlank()) {
                colorPicker.setValue(Color.web(hx));
            } else {
                colorPicker.setValue(Color.web("#4a90d9"));
            }
        } catch (IllegalArgumentException ex) {
            colorPicker.setValue(Color.web("#4a90d9"));
        }
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
        rebuildOrganizerVisualTree();
        syncTabAliasField();
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
        // TreeView は非表示で可視一覧は organizerVisualRoot を手組みするため、モデルだけ追加では画面が更新されない。
        // 選択を変えてリスナー経由で rebuild する（ドラッグ移動完了時と同様）。
        treeView.getSelectionModel().clearSelection();
        treeView.getSelectionModel().select(g);
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
        rebuildOrganizerVisualTree();
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

        /** ツールチップ等用の行説明（グループは [グループ] 付き）。 */
        String treeDetailWithoutHex(MainShellController shellCtl) {
            if (kind == Kind.GROUP) {
                String t = groupTitle != null && !groupTitle.isBlank() ? groupTitle : "グループ";
                return "[グループ] " + t;
            }
            return shellCtl != null ? shellCtl.mainShellTabTitle(tabId) : tabId.name();
        }

        /** ツリー色ピル用の短い見出し（[グループ] プレフィックスなし）。 */
        String treePillPrimaryLabel(MainShellController shellCtl) {
            if (kind == Kind.GROUP) {
                return groupTitle != null && !groupTitle.isBlank() ? groupTitle : "グループ";
            }
            return shellCtl != null ? shellCtl.mainShellTabTitle(tabId) : tabId.name();
        }
    }

    /** {@link TreeView} はモデル専用。可視レイアウトは {@link #organizerVisualRoot} を再構築する。 */
    void installTreeCellFactory() {
        rebuildOrganizerVisualTree();
    }

    private void rebuildOrganizerVisualTree() {
        if (organizerVisualRoot == null || treeView == null) {
            return;
        }
        organizerVisualRoot.getChildren().clear();
        TreeItem<OrgRow> root = treeView.getRoot();
        if (root == null) {
            return;
        }
        for (TreeItem<OrgRow> ch : root.getChildren()) {
            organizerVisualRoot.getChildren().add(buildOrganizerRow(ch));
        }
    }

    private Node buildOrganizerRow(TreeItem<OrgRow> item) {
        if (item == null || item.getValue() == null) {
            return new Region();
        }
        OrgRow r = item.getValue();
        if (r.kind == OrgRow.Kind.TAB) {
            return createPillForTreeItem(item);
        }
        HBox row = new HBox(8);
        row.setAlignment(Pos.TOP_LEFT);
        row.setFillHeight(false);

        Button disclosure = new Button(item.isExpanded() ? "▼" : "▶");
        disclosure.setFocusTraversable(false);
        disclosure.getStyleClass().add("pm-org-tree-disclosure");
        disclosure.setMinSize(Region.USE_PREF_SIZE, Region.USE_PREF_SIZE);
        disclosure.setMaxSize(Region.USE_PREF_SIZE, Region.USE_PREF_SIZE);
        disclosure.setOnAction(
                e -> {
                    item.setExpanded(!item.isExpanded());
                    rebuildOrganizerVisualTree();
                });

        StackPane pill = createPillForTreeItem(item);
        row.getChildren().addAll(disclosure, pill);
        attachDropTarget(row, item);

        if (item.isExpanded() && !item.getChildren().isEmpty()) {
            row.getChildren().add(buildStaggeredChildStrip(item));
        }
        return row;
    }

    /**
     * グループの右側に子を縦に並べる（縦は重ならない）。行ごとに左インデントを増やし、軽い階段状にする。
     */
    private VBox buildStaggeredChildStrip(TreeItem<OrgRow> groupItem) {
        VBox col = new VBox(CHILD_STRIP_VERTICAL_GAP);
        col.setFillWidth(false);
        col.setAlignment(Pos.TOP_LEFT);
        int i = 0;
        for (TreeItem<OrgRow> ch : groupItem.getChildren()) {
            Node n =
                    ch.getValue() != null && ch.getValue().kind == OrgRow.Kind.GROUP
                            ? buildOrganizerRow(ch)
                            : createPillForTreeItem(ch);
            HBox line = new HBox(0);
            line.setAlignment(Pos.CENTER_LEFT);
            line.setFillHeight(false);
            double indent = i * CHILD_STRIP_HORIZONTAL_STAGGER;
            if (indent > 0) {
                Region lead = new Region();
                lead.setMinWidth(indent);
                lead.setPrefWidth(indent);
                lead.setMaxWidth(indent);
                line.getChildren().add(lead);
            }
            line.getChildren().add(n);
            col.getChildren().add(line);
            i++;
        }
        attachDropTarget(col, groupItem);
        return col;
    }

    private StackPane createPillForTreeItem(TreeItem<OrgRow> item) {
        OrgRow row = item.getValue();
        StackPane pill = new StackPane();
        pill.setMinSize(Region.USE_PREF_SIZE, Region.USE_PREF_SIZE);
        pill.setMaxSize(Region.USE_PREF_SIZE, Region.USE_PREF_SIZE);
        pill.getStyleClass().setAll("pm-org-tree-pill");
        Label lab = new Label(row.treePillPrimaryLabel(shell));
        lab.setWrapText(false);
        lab.getStyleClass().add("pm-org-tree-pill-label");
        String hx = row.colorHex;
        if (hx != null && !hx.isBlank() && shell != null) {
            String fill = shell.tabOrganizerPreviewChipLabelTextFill(hx);
            lab.setStyle(
                    "-fx-text-fill: "
                            + fill
                            + "; -fx-font-size: 11px; -fx-font-weight: bold;");
            pill.getStyleClass().remove("pm-org-tree-pill-empty");
        } else {
            pill.getStyleClass().add("pm-org-tree-pill-empty");
            lab.setStyle(
                    "-fx-font-size: 11px; -fx-font-weight: bold; -fx-text-fill: "
                            + (shell != null
                                    ? shell.tabOrganizerPreviewChipLabelTextFill("")
                                    : "#94a3b8")
                            + ";");
        }
        Tooltip.install(pill, new Tooltip(row.treeDetailWithoutHex(shell)));
        pill.getChildren().setAll(lab);
        wirePillInteractions(pill, item);
        updatePillSelectionStyle(pill, item);
        return pill;
    }

    private void wirePillInteractions(StackPane pill, TreeItem<OrgRow> item) {
        pill.setOnMouseClicked(
                ev -> {
                    if (treeView == null) {
                        return;
                    }
                    MultipleSelectionModel<TreeItem<OrgRow>> sm = treeView.getSelectionModel();
                    if (ev.isControlDown()) {
                        if (sm.getSelectedItems().contains(item)) {
                            int row = treeView.getRow(item);
                            if (row >= 0) {
                                sm.clearSelection(row);
                            }
                        } else {
                            sm.select(item);
                        }
                    } else {
                        sm.clearSelection();
                        sm.select(item);
                    }
                    rebuildOrganizerVisualTree();
                    syncOrganizerSideFields();
                    ev.consume();
                });
        pill.setOnDragDetected(
                ev -> {
                    if (item.getValue() == null) {
                        return;
                    }
                    Dragboard db = pill.startDragAndDrop(TransferMode.MOVE);
                    dragSourceItem = item;
                    ClipboardContent cc = new ClipboardContent();
                    cc.put(ROW_MOVE_MARKER, "move");
                    db.setContent(cc);
                    ev.consume();
                });
        pill.setOnDragOver(
                ev -> {
                    if (ev.getGestureSource() != pill
                            && ev.getDragboard().hasContent(ROW_MOVE_MARKER)
                            && dragSourceItem != null
                            && canAcceptDrop(dragSourceItem, item)) {
                        ev.acceptTransferModes(TransferMode.MOVE);
                    }
                    ev.consume();
                });
        pill.setOnDragDropped(
                ev -> {
                    Dragboard db = ev.getDragboard();
                    boolean ok = false;
                    if (db.hasContent(ROW_MOVE_MARKER) && dragSourceItem != null) {
                        ok = performDrop(dragSourceItem, item);
                    }
                    ev.setDropCompleted(ok);
                    ev.consume();
                });
        pill.setOnDragDone(ev -> dragSourceItem = null);
    }

    private void attachDropTarget(Node node, TreeItem<OrgRow> target) {
        node.setOnDragOver(
                ev -> {
                    if (ev.getGestureSource() != node
                            && ev.getDragboard().hasContent(ROW_MOVE_MARKER)
                            && dragSourceItem != null
                            && canAcceptDrop(dragSourceItem, target)) {
                        ev.acceptTransferModes(TransferMode.MOVE);
                    }
                    ev.consume();
                });
        node.setOnDragDropped(
                ev -> {
                    Dragboard db = ev.getDragboard();
                    boolean ok = false;
                    if (db.hasContent(ROW_MOVE_MARKER) && dragSourceItem != null) {
                        ok = performDrop(dragSourceItem, target);
                    }
                    ev.setDropCompleted(ok);
                    ev.consume();
                });
        node.setOnDragDone(ev -> dragSourceItem = null);
    }

    /**
     * 選択枠はインラインで指定する（色付きピルの {@code tabOrganizerTreePillSurfaceStyle} が
     * {@code -fx-border-color} を直指定しており、CSS の {@code .pm-org-tree-pill-selected} より優先されるため）。
     */
    private void updatePillSelectionStyle(StackPane pill, TreeItem<OrgRow> item) {
        if (treeView == null || item == null || item.getValue() == null) {
            return;
        }
        OrgRow row = item.getValue();
        boolean sel = treeView.getSelectionModel().getSelectedItems().contains(item);
        StringBuilder sb = new StringBuilder();
        if (row.colorHex != null && !row.colorHex.isBlank() && shell != null) {
            sb.append(shell.tabOrganizerTreePillSurfaceStyle(row.colorHex.strip()));
        }
        if (sel) {
            if (sb.length() > 0) {
                sb.append(' ');
            }
            sb.append("-fx-border-color: -fx-accent; -fx-border-width: 2;");
        }
        pill.setStyle(sb.toString().strip());
        pill.getStyleClass().remove("pm-org-tree-pill-selected");
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
        rebuildOrganizerVisualTree();
        syncOrganizerSideFields();
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
