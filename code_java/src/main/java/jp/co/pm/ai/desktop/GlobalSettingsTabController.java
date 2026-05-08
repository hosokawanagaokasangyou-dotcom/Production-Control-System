package jp.co.pm.ai.desktop;

import java.util.Optional;

import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.TextInputDialog;

import jp.co.pm.ai.desktop.config.InitSettingPersistence;

/** Global settings tab (factory UI reset and saving package defaults to init_setting). */
public final class GlobalSettingsTabController {

    @FXML
    private Button resetUiButton;

    @FXML
    private Button saveDefaultsButton;

    private MainShellController shell;

    void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    @FXML
    private void onResetUiToFactoryAction() {
        if (shell == null) {
            return;
        }
        shell.performGlobalUiFactoryReset();
    }

    @FXML
    private void onSavePackageDefaultsAction() {
        if (shell == null) {
            return;
        }
        TextInputDialog dialog = new TextInputDialog();
        if (shell.primaryStageForDialogs() != null) {
            dialog.initOwner(shell.primaryStageForDialogs());
        }
        dialog.setTitle("??");
        dialog.setHeaderText(null);
        dialog.setContentText("?????????111??????????");
        Optional<String> ans = dialog.showAndWait();
        if (ans.isEmpty() || !"111".equals(ans.get().trim())) {
            return;
        }
        try {
            InitSettingPersistence.savePackageDefaults(
                    shell.snapshotUiEnv(), shell.snapshotDesktopSessionForExport());
            Alert ok = new Alert(AlertType.INFORMATION);
            if (shell.primaryStageForDialogs() != null) {
                ok.initOwner(shell.primaryStageForDialogs());
            }
            ok.setTitle("??");
            ok.setHeaderText(null);
            ok.setContentText(
                    "???????? init_setting ????????session_defaults.json, table_column_defaults.json??");
            ok.showAndWait();
        } catch (Exception ex) {
            Alert err = new Alert(AlertType.ERROR);
            if (shell.primaryStageForDialogs() != null) {
                err.initOwner(shell.primaryStageForDialogs());
            }
            err.setTitle("???");
            err.setHeaderText(null);
            err.setContentText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            err.showAndWait();
        }
    }
}
