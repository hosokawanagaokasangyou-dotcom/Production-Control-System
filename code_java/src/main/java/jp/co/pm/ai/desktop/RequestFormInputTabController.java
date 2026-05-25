package jp.co.pm.ai.desktop;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Window;

import java.nio.file.Path;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.reconciliation.ReconciliationApp;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPreviewBadgeConfig;

/**
 * 加工依頼書の照合・対比型入力（{@code jp.co.pm.ai.desktop.reconciliation}）をメインシェルに載せる。
 */
public final class RequestFormInputTabController {

    private static final String TAB_LOADING_STYLE_CLASS = "request-form-tab-loading";

    private MainShellController shell;
    private ReconciliationApp reconciliationApp;
    private boolean embeddedBuilt;
    private Parent tabLoadingPane;

    @FXML
    private StackPane contentHost;

    @FXML
    private void initialize() {
        // 遅延構築は bindShell / タブ選択後に行う（起動時の FXML 負荷を抑える）
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    /** メインシェルで当該タブが初めて実体化されたあとに呼ぶ。 */
    void onMainShellTabSelected() {
        showTabLoadingIfNeeded();
        Platform.runLater(this::activateAfterTabPainted);
    }

    /** 依頼書入力タブから離れたとき（バックグラウンド監視停止）。 */
    void onMainShellTabDeselected() {
        if (reconciliationApp != null) {
            reconciliationApp.onEmbeddedTabDeactivated();
        }
    }

    private void activateAfterTabPainted() {
        ensureEmbeddedMounted();
        if (reconciliationApp == null || shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        reconciliationApp.onEmbeddedTabActivated(ui);
        logRequestFormPaths(ui);
    }

    private void showTabLoadingIfNeeded() {
        if (contentHost == null || embeddedBuilt) {
            return;
        }
        if (tabLoadingPane == null) {
            tabLoadingPane = buildTabLoadingPane();
        }
        if (contentHost.getChildren().isEmpty()
                || !contentHost.getChildren().contains(tabLoadingPane)) {
            contentHost.getChildren().setAll(tabLoadingPane);
        }
    }

    private static Parent buildTabLoadingPane() {
        VBox box = new VBox(14);
        box.setAlignment(Pos.CENTER);
        box.getStyleClass().add(TAB_LOADING_STYLE_CLASS);
        ProgressIndicator indicator = new ProgressIndicator();
        indicator.setMaxSize(48, 48);
        Label label = new Label("依頼書入力を準備しています…");
        label.getStyleClass().add("request-form-tab-loading-label");
        box.getChildren().addAll(indicator, label);
        return box;
    }

    private void ensureEmbeddedMounted() {
        if (embeddedBuilt || contentHost == null) {
            return;
        }
        showTabLoadingIfNeeded();
        reconciliationApp = new ReconciliationApp();
        reconciliationApp.setOriginalDirChangeHandler(
                path -> {
                    if (shell != null) {
                        shell.updateEnvTabValue(
                                AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, path);
                    }
                });
        reconciliationApp.setJuchuFileChangeHandler(
                path -> {
                    if (shell != null) {
                        shell.updateEnvTabValue(
                                AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE, path);
                    }
                });
        Window host =
                contentHost.getScene() != null ? contentHost.getScene().getWindow() : null;
        if (host == null && shell != null) {
            host = shell.primaryStageForDialogs();
        }
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        Path repoRoot = shell != null ? AppPaths.resolveRepoRoot(ui) : null;
        Parent root = reconciliationApp.buildEmbeddedRoot(host, repoRoot, ui);
        reconciliationApp.setPreviewBadgeConfigSupplier(
                () ->
                        shell != null
                                ? shell.requestFormPreviewBadgeConfig()
                                : RequestFormPreviewBadgeConfig.defaults());
        contentHost.getChildren().setAll(root);
        embeddedBuilt = true;
        contentHost
                .sceneProperty()
                .addListener(
                        (obs, prev, scene) -> {
                            if (reconciliationApp != null && scene != null) {
                                reconciliationApp.updateHostWindow(scene.getWindow());
                            }
                        });
    }

    /** デザインタブ変更後に依頼書プレビュー上部バッジを再描画する。 */
    void refreshPreviewBadgeAppearance() {
        if (reconciliationApp != null) {
            reconciliationApp.refreshPreviewBadgeAppearance();
        }
    }

    private void logRequestFormPaths(Map<String, String> ui) {
        if (shell == null) {
            return;
        }
        String prefix = embeddedBuilt ? "" : "依頼書入力 UI 初期化。";
        shell.appendLog(
                "[request-form-input] "
                        + prefix
                        + "依頼書原本="
                        + AppPaths.resolveRequestFormOriginalDir(ui)
                        + " アラジンマスタ="
                        + AppPaths.resolveAladdinMasterDir(ui)
                        + " 受注ファイル="
                        + AppPaths.resolveRequestFormJuchuFile(ui).map(Path::toString).orElse(""));
    }
}
