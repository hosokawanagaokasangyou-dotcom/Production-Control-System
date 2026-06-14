package jp.co.pm.ai.desktop;

import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.application.Platform;
import javafx.animation.PauseTransition;
import javafx.collections.FXCollections;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.PasswordField;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import javafx.util.Duration;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.audio.UiClickSound;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.DesktopTheme;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.PortableBundleProfile;
import jp.co.pm.ai.desktop.config.PortableBundleSelfUpdateService;
import jp.co.pm.ai.desktop.config.RemoteDesktopEnvRows;
import jp.co.pm.ai.desktop.ui.ButtonPressFeedback;

/** リモートデスクトップ配布用アプリの最小シェル。 */
public final class RemoteDesktopShellController implements DesktopShellHost, EnvTabShellHost {

    private final Stage primaryStage;
    private ObservableList<EnvVarRow> envRows;

    private Scene primaryScene;
    private DesktopTheme pendingTheme = DesktopTheme.LIGHT;

    private boolean operatorUserAdminTabUnlocked;
    private final AtomicBoolean suppressOperatorUserAdminTabGuard = new AtomicBoolean(false);
    private final AtomicBoolean suppressEnvSessionPersistence = new AtomicBoolean(false);
    private Tab lastEffectiveTab;

    private final PauseTransition uiEnvSaveDebounce = new PauseTransition(Duration.millis(400));

    @FXML private ComboBox<DesktopTheme> themeCombo;
    @FXML private Label sessionOperatorToolbarLabel;
    @FXML private Button changeSessionOperatorToolbarButton;
    @FXML private Button changeSessionDepartmentToolbarButton;
    @FXML private TabPane tabPane;
    @FXML private Tab remoteDesktopTab;
    @FXML private Tab envSettingsTab;
    @FXML private Tab operatorUserManagementTab;
    @FXML private TextArea logTextArea;

    @FXML private RemoteDesktopTabController remoteDesktopTabContentController;
    @FXML private EnvTabController envTabController;
    @FXML private OperatorUserManagementTabController operatorUserManagementTabContentController;

    RemoteDesktopShellController(Stage primaryStage) {
        this.primaryStage = primaryStage;
    }

    @FXML
    private void initialize() {
        envRows = FXCollections.observableArrayList();
        DesktopSessionState session = DesktopSessionStateStore.load();
        RemoteDesktopEnvRows.bootstrap(envRows, session);
        FactoryOperatorUserStore.configureForCurrentApp(
                RemoteDesktopEnvRows.collectMap(envRows), FactorySite.RDP_LAUNCHER);

        if (remoteDesktopTabContentController != null) {
            remoteDesktopTabContentController.bindShell(this);
        }
        if (operatorUserManagementTabContentController != null) {
            operatorUserManagementTabContentController.bindShell(this);
        }

        pendingTheme = DesktopTheme.fromStored(session.uiTheme());
        installUiEnvAutoSave();

        tabPane
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, prev, now) -> {
                            if (now == remoteDesktopTab && remoteDesktopTabContentController != null) {
                                Platform.runLater(remoteDesktopTabContentController::onMainShellTabSelected);
                            }
                            emitTabNavigationGuard(prev, now);
                        });

        primaryStage.setOnShown(
                e -> {
                    if (tabPane.getSelectionModel().getSelectedItem() == null) {
                        tabPane.getSelectionModel().selectFirst();
                    }
                    Tab selected = tabPane.getSelectionModel().getSelectedItem();
                    if (selected == remoteDesktopTab && remoteDesktopTabContentController != null) {
                        remoteDesktopTabContentController.onMainShellTabSelected();
                    }
                    lastEffectiveTab = selected;
                    Platform.runLater(this::promptOperatorAtStartupAfterSplash);
                });
    }

    void finishStartup(Scene scene) {
        this.primaryScene = scene;
        ButtonPressFeedback.installOnScene(scene);
        UiClickSound.warmUp(snapshotUiEnv());
        if (envTabController != null) {
            envTabController.bindShell(this);
        }
        if (themeCombo == null) {
            Platform.runLater(this::maybePortableBundleSelfUpdate);
            return;
        }
        themeCombo.getItems().setAll(DesktopTheme.values());
        themeCombo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(DesktopTheme t) {
                        return t != null ? t.displayLabel() : "";
                    }

                    @Override
                    public DesktopTheme fromString(String s) {
                        return DesktopTheme.fromDisplayLabel(s);
                    }
                });
        DesktopTheme initial = pendingTheme != null ? pendingTheme : DesktopTheme.LIGHT;
        initial.applyTo(scene);
        themeCombo.setValue(initial);
        themeCombo
                .valueProperty()
                .addListener(
                        (obs, was, now) -> {
                            if (now == null) {
                                return;
                            }
                            pendingTheme = now;
                            now.applyTo(scene);
                            persistSessionQuietly();
                        });
        refreshOperatorUserPresentation();
        Platform.runLater(this::maybePortableBundleSelfUpdate);
    }

    private void installUiEnvAutoSave() {
        uiEnvSaveDebounce.setOnFinished(
                e -> {
                    if (!suppressEnvSessionPersistence.get()) {
                        persistSessionQuietly();
                    }
                    FactoryOperatorUserStore.configureForCurrentApp(
                            snapshotUiEnv(), FactorySite.RDP_LAUNCHER);
                });
        Runnable schedule = () -> uiEnvSaveDebounce.playFromStart();
        envRows.addListener(
                (ListChangeListener<EnvVarRow>)
                        c -> {
                            while (c.next()) {
                                if (c.wasAdded()) {
                                    for (EnvVarRow row : c.getAddedSubList()) {
                                        hookEnvRowForAutoSave(row, schedule);
                                    }
                                }
                            }
                            if (!suppressEnvSessionPersistence.get()) {
                                schedule.run();
                            }
                        });
        for (EnvVarRow row : envRows) {
            hookEnvRowForAutoSave(row, schedule);
        }
    }

    private static void hookEnvRowForAutoSave(EnvVarRow row, Runnable schedule) {
        row.nameProperty().addListener((o, a, b) -> schedule.run());
        row.valueProperty().addListener((o, a, b) -> schedule.run());
        row.descriptionProperty().addListener((o, a, b) -> schedule.run());
    }

    private void maybePortableBundleSelfUpdate() {
        PortableBundleSelfUpdateService.maybeSelfUpdateOnStartup(
                PortableBundleProfile.RDP_LAUNCHER,
                AppPaths.KEY_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR,
                snapshotUiEnv(),
                primaryStage,
                this::appendLog);
    }

    private void promptOperatorAtStartupAfterSplash() {
        requireOperatorSelectionForFactory(FactorySite.RDP_LAUNCHER, true);
        refreshOperatorUserPresentation();
    }

    @FXML
    private void onChangeSessionOperatorToolbarAction() {
        requireOperatorSelectionForFactory(FactorySite.RDP_LAUNCHER, false);
        refreshOperatorUserPresentation();
    }

    @FXML
    private void onChangeSessionDepartmentToolbarAction() {
        FactoryOperatorUserStore.clearSessionOperatorName();
        FactoryOperatorUserStore.clearSessionRdpDepartmentKey();
        requireOperatorSelectionForFactory(FactorySite.RDP_LAUNCHER, false);
        refreshOperatorUserPresentation();
    }

    private void emitTabNavigationGuard(Tab prev, Tab now) {
        if (now == operatorUserManagementTab
                && !operatorUserAdminTabUnlocked
                && !suppressOperatorUserAdminTabGuard.get()) {
            if (!promptOperatorUserAdminTabUnlock()) {
                suppressOperatorUserAdminTabGuard.set(true);
                try {
                    if (prev != null && prev != operatorUserManagementTab) {
                        tabPane.getSelectionModel().select(prev);
                    } else {
                        tabPane.getSelectionModel().select(remoteDesktopTab);
                    }
                } finally {
                    suppressOperatorUserAdminTabGuard.set(false);
                }
                return;
            }
            operatorUserAdminTabUnlocked = true;
        }
        lastEffectiveTab = now;
    }

    private boolean promptOperatorUserAdminTabUnlock() {
        if (primaryStage == null) {
            return false;
        }
        Dialog<ButtonType> dialog = new Dialog<>();
        prepareDialogForMainTheme(dialog);
        dialog.setTitle("ユーザー管理者");
        dialog.setHeaderText(null);
        Label hint =
                new Label(
                        "ユーザー管理者タブを開くには、ユーザー名 "
                                + FactoryOperatorUserStore.ADMIN_TAB_USERNAME
                                + " と管理者パスワードを入力してください。");
        hint.setWrapText(true);
        TextField userField = new TextField();
        userField.setPromptText(FactoryOperatorUserStore.ADMIN_TAB_USERNAME);
        PasswordField pf = new PasswordField();
        pf.setPromptText("管理者パスワード");
        VBox box =
                new VBox(
                        8,
                        hint,
                        new Label("ユーザー名:"),
                        userField,
                        new Label("パスワード:"),
                        pf);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        return dialog.showAndWait()
                .filter(bt -> bt == ButtonType.OK)
                .filter(
                        bt ->
                                FactoryOperatorUserStore.verifyAdminTabAccess(
                                        userField.getText(), pf.getText()))
                .isPresent();
    }

    private void persistSessionQuietly() {
        String themeId =
                themeCombo != null && themeCombo.getValue() != null
                        ? themeCombo.getValue().storedId()
                        : pendingTheme != null
                                ? pendingTheme.storedId()
                                : DesktopTheme.LIGHT.storedId();
        DesktopSessionStateStore.patchUiEnvRowsAndTheme(
                RemoteDesktopEnvRows.snapshot(envRows), themeId);
    }

    @Override
    public Stage getPrimaryStage() {
        return primaryStage;
    }

    @Override
    public ObservableList<EnvVarRow> getEnvRows() {
        return envRows;
    }

    @Override
    public void confirmAndResetEnvRowsToDefaults() {
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        prepareDialogForMainTheme(alert);
        alert.setTitle("環境変数を初期値に戻す");
        alert.setHeaderText(null);
        alert.setContentText(
                "ui_ref_env_defaults.json の既定行に戻します。"
                        + " RDP ランチャー向けの掲示板共有既定も空欄へ再適用します。"
                        + " 続行しますか？");
        if (alert.showAndWait().orElse(ButtonType.CANCEL) != ButtonType.OK) {
            return;
        }
        suppressEnvSessionPersistence.set(true);
        try {
            RemoteDesktopEnvRows.populateFromUiRef(envRows);
            RemoteDesktopEnvRows.applyRdpLauncherEmptyDefaults(envRows);
        } finally {
            suppressEnvSessionPersistence.set(false);
        }
        persistSessionQuietly();
        FactoryOperatorUserStore.configureForCurrentApp(snapshotUiEnv(), FactorySite.RDP_LAUNCHER);
        appendLog("[env] 環境変数を ui_ref 既定に戻しました。");
    }

    @Override
    public void addMissingReferenceEnvRows() {
        RemoteDesktopEnvRows.mergeMissingFromUiRef(envRows);
    }

    @Override
    public Map<String, String> snapshotUiEnv() {
        Map<String, String> base = RemoteDesktopEnvRows.collectMap(envRows);
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        if (operator.isBlank()) {
            return Map.copyOf(base);
        }
        Map<String, String> merged = new HashMap<>(base);
        merged.put(AppPaths.KEY_PM_AI_OPERATOR_USER, operator);
        return Map.copyOf(merged);
    }

    @Override
    public void updateEnvTabValue(String envKey, String value) {
        if (envKey == null || envKey.isBlank()) {
            return;
        }
        RemoteDesktopEnvRows.syncRowValue(envRows, envKey.strip(), value);
        if (!suppressEnvSessionPersistence.get()) {
            uiEnvSaveDebounce.playFromStart();
        }
    }

    @Override
    public void appendLog(String line) {
        if (line == null || line.isBlank() || logTextArea == null) {
            return;
        }
        String existing = logTextArea.getText();
        logTextArea.appendText((existing.isEmpty() ? "" : "\n") + line);
    }

    @Override
    public Stage primaryStageForDialogs() {
        return primaryStage;
    }

    @Override
    public void requireOperatorSelectionForFactory(FactorySite site, boolean startup) {
        OperatorUserSelectionSupport.requireOperatorSelectionForFactory(this, site, startup);
    }

    @Override
    public void refreshOperatorUserPresentation() {
        if (sessionOperatorToolbarLabel != null) {
            String op = FactoryOperatorUserStore.sessionOperatorName();
            String dept = FactoryOperatorUserStore.sessionRdpDepartmentKey();
            String text =
                    op.isBlank()
                            ? "操作者: （未選択）"
                            : "操作者: " + op;
            if (!dept.isBlank()) {
                text += "　部署: " + dept;
            }
            sessionOperatorToolbarLabel.setText(text);
        }
        if (operatorUserManagementTabContentController != null) {
            operatorUserManagementTabContentController.refreshPresentationQuietly();
        }
    }

    @Override
    public void prepareDialogForMainTheme(Dialog<?> dialog) {
        if (dialog == null) {
            return;
        }
        if (primaryStage != null && primaryStage.getScene() != null) {
            dialog.initOwner(primaryStage);
        }
        if (primaryScene != null) {
            dialog.getDialogPane()
                    .getStylesheets()
                    .setAll(primaryScene.getStylesheets());
        }
    }

    @Override
    public void showWarningDialog(String title, String message) {
        showThemedAlert(Alert.AlertType.WARNING, title, message);
    }

    @Override
    public void showInformationDialog(String title, String message) {
        showThemedAlert(Alert.AlertType.INFORMATION, title, message);
    }

    private void showThemedAlert(Alert.AlertType type, String title, String message) {
        Alert alert = new Alert(type);
        prepareDialogForMainTheme(alert);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message != null ? message : "");
        alert.showAndWait();
    }
}
