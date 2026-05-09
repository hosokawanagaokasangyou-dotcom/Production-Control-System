package jp.co.pm.ai.desktop;

import java.net.URL;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import javafx.application.Platform;
import javafx.beans.value.ChangeListener;
import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.Toggle;
import javafx.scene.control.ToggleGroup;
import javafx.scene.control.TreeItem;
import javafx.scene.control.TreeView;
import javafx.scene.web.WebEngine;
import javafx.scene.web.WebView;

import jp.co.pm.ai.desktop.config.AppVersionInfo;

/** バンドル HTML のオンラインマニュアル（{@code OnlineManualTab.fxml}）。 */
public final class OnlineManualTabController {

    static final class ManualChapter {

        final String anchor;

        final String title;

        ManualChapter(String anchor, String title) {
            this.anchor = anchor;
            this.title = title;
        }

        @Override
        public String toString() {
            return title;
        }
    }

    private enum ManualDepth {
        BEGINNER("beginner.html"),
        INTERMEDIATE("intermediate.html"),
        ADVANCED("advanced.html"),
        DEVELOPER("developer.html");

        final String fileName;

        ManualDepth(String fileName) {
            this.fileName = fileName;
        }
    }

    private static final String MANUAL_BASE = "/jp/co/pm/ai/desktop/manual/";

    private static final String PM_DESKTOP_TAB_SCHEME = "pm-desktop://tab/";

    private static final List<ManualChapter> CHAPTERS =
            List.of(
                    new ManualChapter("ch01", "はじめに・全体像"),
                    new ManualChapter("ch02", "実行・ログ"),
                    new ManualChapter("ch03", "環境設定（環境変数・メモリ・グローバル）"),
                    new ManualChapter("ch04", "配台計画_タスク入力"),
                    new ManualChapter("ch05", "マスタ読込サマリ"),
                    new ManualChapter("ch06", "段階1（実行とログの見方）"),
                    new ManualChapter("ch07", "段階1 成形結果"),
                    new ManualChapter("ch08", "段階2（計画シミュレーション）"),
                    new ManualChapter("ch09", "計画結果ビューア・結果_配台表"),
                    new ManualChapter("ch10", "配台計画手動修正"),
                    new ManualChapter("ch11", "設備ガント・納期管理ビュー・オペレーターカード"),
                    new ManualChapter("ch12", "特別ルール・配台不要ルール・実績DATA"));

    private MainShellController shell;

    @FXML
    private ToggleGroup manualDepthToggleGroup;

    @FXML
    private RadioButton depthBeginner;

    @FXML
    private RadioButton depthIntermediate;

    @FXML
    private RadioButton depthAdvanced;

    @FXML
    private RadioButton depthDeveloper;

    @FXML
    private TreeView<ManualChapter> chapterTree;

    @FXML
    private WebView manualWebView;

    @FXML
    private Label manualVersionLabel;

    private WebEngine webEngine;

    private ManualDepth currentDepth = ManualDepth.BEGINNER;

    private String currentAnchor = "ch01";

    /** カスタムスキーム処理後に {@link WebEngine#load(String)} へ戻す直近のマニュアル URL。 */
    private volatile String lastLoadedManualUrl;

    private final ChangeListener<String> locationListener =
            (obs, oldLoc, newLoc) -> {
                if (newLoc == null || !newLoc.startsWith(PM_DESKTOP_TAB_SCHEME)) {
                    return;
                }
                String rest = newLoc.substring(PM_DESKTOP_TAB_SCHEME.length());
                int q = rest.indexOf('?');
                if (q >= 0) {
                    rest = rest.substring(0, q);
                }
                int hash = rest.indexOf('#');
                if (hash >= 0) {
                    rest = rest.substring(0, hash);
                }
                String key = rest.trim();
                MainShellTabId id = MainShellTabId.fromKey(key);
                if (shell != null && id != null && id != MainShellTabId.TAB_ORGANIZER) {
                    Platform.runLater(
                            () -> {
                                shell.selectMainShellTab(id);
                                String restore =
                                        lastLoadedManualUrl != null
                                                ? lastLoadedManualUrl
                                                : buildManualUrl();
                                webEngine.load(restore);
                            });
                }
            };

    @FXML
    private void initialize() {
        webEngine = manualWebView.getEngine();
        depthBeginner.setUserData(ManualDepth.BEGINNER);
        depthIntermediate.setUserData(ManualDepth.INTERMEDIATE);
        depthAdvanced.setUserData(ManualDepth.ADVANCED);
        depthDeveloper.setUserData(ManualDepth.DEVELOPER);

        TreeItem<ManualChapter> invisibleRoot = new TreeItem<>();
        for (ManualChapter ch : CHAPTERS) {
            invisibleRoot.getChildren().add(new TreeItem<>(ch));
        }
        invisibleRoot.setExpanded(true);
        chapterTree.setRoot(invisibleRoot);
        chapterTree.setShowRoot(false);

        chapterTree
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, prev, item) -> {
                            if (item != null && item.getValue() != null) {
                                currentAnchor = item.getValue().anchor;
                                loadCurrentManualPage();
                            }
                        });

        manualDepthToggleGroup
                .selectedToggleProperty()
                .addListener(
                        (obs, oldT, newT) -> {
                            ManualDepth d = toggleToDepth(newT);
                            if (d != null) {
                                currentDepth = d;
                                loadCurrentManualPage();
                            }
                        });

        webEngine.locationProperty().addListener(locationListener);

        Platform.runLater(
                () -> {
                    syncDepthFromToggle();
                    selectTreeChapter(currentAnchor);
                    refreshVersionLabel();
                });
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        refreshVersionLabel();
    }

    private void syncDepthFromToggle() {
        ManualDepth d = toggleToDepth(manualDepthToggleGroup.getSelectedToggle());
        if (d != null) {
            currentDepth = d;
        }
    }

    private static ManualDepth toggleToDepth(Toggle t) {
        if (t == null) {
            return null;
        }
        Object u = t.getUserData();
        return u instanceof ManualDepth md ? md : null;
    }

    private void selectTreeChapter(String anchor) {
        if (chapterTree == null || anchor == null) {
            return;
        }
        TreeItem<ManualChapter> root = chapterTree.getRoot();
        if (root == null) {
            return;
        }
        for (TreeItem<ManualChapter> child : root.getChildren()) {
            ManualChapter v = child.getValue();
            if (v != null && anchor.equals(v.anchor)) {
                chapterTree.getSelectionModel().select(child);
                return;
            }
        }
    }

    private void loadCurrentManualPage() {
        String url = buildManualUrl();
        lastLoadedManualUrl = url;
        URL resource = OnlineManualTabController.class.getResource(MANUAL_BASE + currentDepth.fileName);
        if (resource == null) {
            webEngine.loadContent(
                    "<html><meta charset=\"UTF-8\"><body><p>マニュアル HTML が見つかりません: "
                            + escapeHtml(currentDepth.fileName)
                            + "</p></body></html>");
            return;
        }
        webEngine.load(url);
    }

    private String buildManualUrl() {
        URL base = OnlineManualTabController.class.getResource(MANUAL_BASE + currentDepth.fileName);
        if (base == null) {
            return "";
        }
        String u = base.toExternalForm();
        if (currentAnchor != null && !currentAnchor.isBlank()) {
            u += "#" + currentAnchor;
        }
        return u;
    }

    private static String escapeHtml(String s) {
        if (s == null) {
            return "";
        }
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;");
    }

    private void refreshVersionLabel() {
        if (manualVersionLabel == null) {
            return;
        }
        Map<String, String> ui =
                shell != null
                        ? shell.snapshotUiEnvForManual()
                        : Map.of();
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        String v = AppVersionInfo.resolveDisplayedVersion(cwd, ui);
        manualVersionLabel.setText("アプリ版: " + v + "（問い合わせ時に記載してください）");
    }
}
