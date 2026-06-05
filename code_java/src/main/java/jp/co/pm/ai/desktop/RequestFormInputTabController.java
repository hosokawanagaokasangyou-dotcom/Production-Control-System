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

import jp.co.pm.ai.desktop.reconciliation.RequestFormComboChoices;
import jp.co.pm.ai.desktop.reconciliation.RequestFormFeedLocPlanSync;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.reconciliation.JuchuHeaderAliasRegistry;
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
    private RequestFormComboChoices pendingComboChoices = RequestFormComboChoices.empty();

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
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        reconciliationApp = new ReconciliationApp();
        FactorySite factorySite = GlobalInitSettingTarget.load();
        reconciliationApp.configureJuchuHeaderAliasRegistry(
                JuchuHeaderAliasRegistry.loadForFactory(factorySite, ui));
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
        Path repoRoot = shell != null ? AppPaths.resolveRepoRoot(ui) : null;
        Parent root = reconciliationApp.buildEmbeddedRoot(host, repoRoot, ui);
        if (pendingComboChoices != null) {
            reconciliationApp.applyComboChoices(pendingComboChoices.mergedWithDefaults());
        }
        reconciliationApp.setPreviewBadgeConfigSupplier(
                () ->
                        shell != null
                                ? shell.requestFormPreviewBadgeConfig()
                                : RequestFormPreviewBadgeConfig.defaults());
        reconciliationApp.setPlanningPipelineStageBusyChecker(
                () -> shell != null && shell.isPlanningPipelineStageRunning());
        contentHost.getChildren().setAll(root);
        embeddedBuilt = true;
        if (shell != null) {
            shell.refreshInnerTabHeaderColorsForShellTab(MainShellTabId.REQUEST_FORM_INPUT);
        }
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

    void refreshSessionInputTantoLabel() {
        if (reconciliationApp != null) {
            reconciliationApp.refreshSessionInputTantoLabel();
        }
    }

    /**
     * 段階1正常終了後: 計画データの原反投入場所を依頼書入力の投入場所候補へ追記する。
     *
     * @return 追加分（失敗時は -1、追記なしは 0）
     */
    int mergeFeedLocFromStage1Plan(Map<String, String> ui) {
        List<String> fromPlan;
        try {
            fromPlan = RequestFormFeedLocPlanSync.collectDistinctFeedLocations(ui);
        } catch (IOException ex) {
            return -1;
        }
        if (fromPlan.isEmpty()) {
            return 0;
        }
        if (reconciliationApp != null) {
            return reconciliationApp.mergeFeedLocOptionsFromPlanning(fromPlan);
        }
        List<String> before = pendingComboChoices.optionsFor(RequestFormComboChoices.KEY_FEED_LOC);
        List<String> merged = RequestFormFeedLocPlanSync.mergeDistinctFeedLocations(before, fromPlan);
        int added = RequestFormFeedLocPlanSync.countNewValues(before, merged);
        if (added <= 0) {
            return 0;
        }
        java.util.LinkedHashMap<String, java.util.List<String>> map =
                new java.util.LinkedHashMap<>(pendingComboChoices.asMap());
        map.put(RequestFormComboChoices.KEY_FEED_LOC, merged);
        pendingComboChoices =
                RequestFormComboChoices.of(map, pendingComboChoices.fieldDefaultsAsMap());
        return added;
    }

    void applyComboChoicesFromSession(RequestFormComboChoices choices) {
        pendingComboChoices =
                choices != null ? choices : RequestFormComboChoices.empty();
        if (reconciliationApp != null) {
            reconciliationApp.applyComboChoices(pendingComboChoices.mergedWithDefaults());
        }
    }

    RequestFormComboChoices snapshotComboChoices() {
        if (reconciliationApp != null) {
            return reconciliationApp.snapshotComboChoices();
        }
        return pendingComboChoices != null
                ? pendingComboChoices
                : RequestFormComboChoices.empty();
    }

    /** 依頼書入力設定をサマリ Excel 同フォルダへ保存する。 */
    void persistInputSettings() {
        if (reconciliationApp != null) {
            reconciliationApp.persistInputSettings();
        }
    }

    JuchuHeaderAliasRegistry snapshotJuchuHeaderAliasRegistry() {
        if (reconciliationApp != null) {
            return reconciliationApp.juchuHeaderAliasRegistry();
        }
        return JuchuHeaderAliasRegistry.loadForFactory(
                GlobalInitSettingTarget.load(), shell != null ? shell.snapshotUiEnv() : Map.of());
    }

    void reloadJuchuHeaderAliasRegistry(
            FactorySite site, Map<String, String> ui, boolean restoreFromInitSetting) {
        FactorySite effective = site != null ? site : GlobalInitSettingTarget.load();
        Map<String, String> env = ui != null ? ui : Map.of();
        JuchuHeaderAliasRegistry registry =
                JuchuHeaderAliasRegistry.loadForFactory(effective, env);
        if (restoreFromInitSetting) {
            Path initSetting = DesktopSessionStateStore.factoryJuchuHeaderAliasesPath(env, effective);
            if (initSetting != null && Files.isRegularFile(initSetting)) {
                try {
                    registry.replaceFromInitSetting(initSetting);
                } catch (IOException ignored) {
                }
            }
        }
        if (reconciliationApp != null) {
            reconciliationApp.configureJuchuHeaderAliasRegistry(registry);
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
