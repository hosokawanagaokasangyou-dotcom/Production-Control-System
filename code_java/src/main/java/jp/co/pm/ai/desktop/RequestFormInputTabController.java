package jp.co.pm.ai.desktop;

import javafx.fxml.FXML;
import javafx.scene.Parent;
import javafx.scene.layout.StackPane;
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

    private MainShellController shell;
    private ReconciliationApp reconciliationApp;
    private boolean embeddedBuilt;

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
    void ensureEmbeddedMounted() {
        if (embeddedBuilt || contentHost == null) {
            return;
        }
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
        logRequestFormPaths(ui);
        contentHost
                .sceneProperty()
                .addListener(
                        (obs, prev, scene) -> {
                            if (reconciliationApp != null && scene != null) {
                                reconciliationApp.updateHostWindow(scene.getWindow());
                            }
                        });
    }

    /** メインシェルで当該タブを選択するたびに環境変数を再読込する。 */
    void onMainShellTabSelected() {
        ensureEmbeddedMounted();
        if (reconciliationApp == null || shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        reconciliationApp.configureFromUiEnv(ui);
        reconciliationApp.reloadMasterProductListFromDisk();
        logRequestFormPaths(ui);
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
