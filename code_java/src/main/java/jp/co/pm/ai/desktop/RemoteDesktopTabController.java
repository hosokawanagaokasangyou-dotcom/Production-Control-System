package jp.co.pm.ai.desktop;

import java.util.Map;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.RdpPreviewSettings;
import jp.co.pm.ai.desktop.reconciliation.RequestFormRemoteDesktopPane;

/** メインシェル「リモートデスクトップ」タブ。 */
public final class RemoteDesktopTabController {

    private static final String TAB_LOADING_STYLE_CLASS = "request-form-tab-loading";

    private DesktopShellHost shell;
    private boolean contentBuilt;
    private Runnable remoteDesktopOperatorContextRefresh;

    @FXML
    private StackPane contentHost;

    @FXML
    private void initialize() {
        // 初回選択まで遅延構築
    }

    void bindShell(DesktopShellHost shell) {
        this.shell = shell;
    }

    /** 操作者変更後に RPA設定 ini のパス表示と共有ファイル読込を同期する。 */
    void refreshForSessionOperatorChange() {
        if (remoteDesktopOperatorContextRefresh != null) {
            remoteDesktopOperatorContextRefresh.run();
        }
    }

    /** 工場切替後: 走査結果を破棄し、タブ選択時まで再走査を遅延する。 */
    void onFactorySiteChanged(boolean lightweight) {
        refreshForSessionOperatorChange();
    }

    void onMainShellTabSelected() {
        if (contentBuilt) {
            return;
        }
        showLoadingPane();
        Thread preload =
                new Thread(
                        () -> {
                            String error = preloadClasses();
                            Platform.runLater(() -> mountContent(error));
                        },
                        "remote-desktop-tab-mount");
        preload.setDaemon(true);
        preload.start();
    }

    private static String preloadClasses() {
        try {
            Class.forName("jp.co.pm.ai.desktop.reconciliation.RequestFormRemoteDesktopPane");
            return null;
        } catch (Throwable ex) {
            return ex.getMessage() != null ? ex.getMessage() : ex.toString();
        }
    }

    private void mountContent(String preloadError) {
        if (contentHost == null || contentBuilt) {
            return;
        }
        if (preloadError != null) {
            Label err =
                    new Label(
                            "リモートデスクトップの読み込みに失敗しました。\n"
                                    + preloadError
                                    + "\n\n対処: code_java で .\\mvnw.cmd clean compile を実行してから再起動してください。");
            err.setWrapText(true);
            err.setMaxWidth(560);
            err.getStyleClass().add("request-form-tab-loading-label");
            contentHost.getChildren().setAll(err);
            contentBuilt = true;
            return;
        }

        Window owner =
                contentHost.getScene() != null ? contentHost.getScene().getWindow() : null;
        if (owner == null && shell != null) {
            owner = shell.primaryStageForDialogs();
        }

        RequestFormRemoteDesktopPane.TabContent built =
                RequestFormRemoteDesktopPane.buildTabContent(
                        owner,
                        new RequestFormRemoteDesktopPane.Context(
                                () -> shell != null ? shell.snapshotUiEnv() : Map.of(),
                                path -> {
                                    if (shell != null) {
                                        shell.updateEnvTabValue(
                                                AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE, path);
                                    }
                                },
                                path -> {
                                    if (shell != null) {
                                        shell.updateEnvTabValue(
                                                AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM, path);
                                    }
                                },
                                args -> {
                                    if (shell != null) {
                                        shell.updateEnvTabValue(
                                                AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS, args);
                                    }
                                },
                                previewFlag -> {
                                    if (shell != null) {
                                        shell.updateEnvTabValue(
                                                RdpPreviewSettings.KEY_PM_AI_RDP_PREVIEW_IN_TAB,
                                                previewFlag);
                                    }
                                },
                                profileNumber -> {
                                    if (shell != null) {
                                        shell.updateEnvTabValue(
                                                AppPaths.KEY_PM_AI_RDP_LAUNCH_PROFILE_NUMBER,
                                                String.valueOf(profileNumber));
                                    }
                                },
                                path -> {
                                    if (shell != null) {
                                        shell.updateEnvTabValue(
                                                AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR, path);
                                    }
                                },
                                msg -> {
                                    if (shell != null) {
                                        shell.appendLog(msg);
                                    }
                                }));

        built.root().getStyleClass().add("pm-rdp-form-tab-container");
        contentHost.getChildren().setAll(built.root());
        remoteDesktopOperatorContextRefresh = built.onSessionOperatorChanged();
        built.scheduleInitialRefresh().run();
        contentBuilt = true;
    }

    private void showLoadingPane() {
        if (contentHost == null || contentBuilt) {
            return;
        }
        VBox box = new VBox(14);
        box.setAlignment(Pos.CENTER);
        box.getStyleClass().add(TAB_LOADING_STYLE_CLASS);
        ProgressIndicator indicator = new ProgressIndicator();
        indicator.setMaxSize(48, 48);
        Label label = new Label("リモートデスクトップを準備しています…");
        label.getStyleClass().add("request-form-tab-loading-label");
        box.getChildren().addAll(indicator, label);
        contentHost.getChildren().setAll(box);
    }
}
