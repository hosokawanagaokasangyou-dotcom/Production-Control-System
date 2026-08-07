package jp.co.pm.ai.desktop;

import java.util.Map;
import java.util.Optional;

import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TextInputDialog;

import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.config.InitSettingPaths;
import jp.co.pm.ai.desktop.config.InitSettingPersistence;
import jp.co.pm.ai.desktop.ui.FactorySiteComboPresentation;

/** Global settings tab (factory UI reset and saving package defaults to init_setting). */
public final class GlobalSettingsTabController {

    @FXML
    private Button resetUiButton;

    @FXML
    private Button saveDefaultsButton;

    @FXML
    private ComboBox<FactorySite> initSettingTargetCombo;

    @FXML
    private CheckBox tableRowHoverDimmingCheckBox;

    private MainShellController shell;

    /** {@link ComboBox#setValue} によるリスナー発火を抑止（起動時の Scene 未設定ダイアログ回避）。 */
    private boolean suppressInitSettingTargetComboEvents;

    void bindShell(MainShellController shell) {
        this.shell = shell;
        wireInitSettingTargetCombo();
        syncTableRowHoverDimmingCheckbox();
    }

    void syncTableRowHoverDimmingCheckbox() {
        if (tableRowHoverDimmingCheckBox == null || shell == null) {
            return;
        }
        tableRowHoverDimmingCheckBox.setSelected(shell.tableRowHoverDimmingEnabled());
    }

    @FXML
    private void onTableRowHoverDimmingChanged() {
        if (shell == null || tableRowHoverDimmingCheckBox == null) {
            return;
        }
        shell.setTableRowHoverDimmingEnabled(tableRowHoverDimmingCheckBox.isSelected());
        shell.persistGlobalDesktopSession();
    }

    /** 環境タブの工場プリセットなどで {@link GlobalInitSettingTarget} が変わったあと、コンボをディスクに合わせる。 */
    void refreshInitSettingTargetComboFromStore() {
        if (initSettingTargetCombo == null) {
            return;
        }
        FactorySite disk = GlobalInitSettingTarget.load();
        if (initSettingTargetCombo.getValue() != disk) {
            setInitSettingTargetComboValueSilently(disk);
        }
    }

    private void setInitSettingTargetComboValueSilently(FactorySite site) {
        suppressInitSettingTargetComboEvents = true;
        try {
            initSettingTargetCombo.setValue(site);
        } finally {
            suppressInitSettingTargetComboEvents = false;
        }
    }

    private void wireInitSettingTargetCombo() {
        if (initSettingTargetCombo == null) {
            return;
        }
        initSettingTargetCombo.getItems().setAll(FactorySite.dispatchProductionSites());
        FactorySiteComboPresentation.wire(
                initSettingTargetCombo, () -> shell != null ? shell.snapshotUiEnv() : Map.of());
        initSettingTargetCombo
                .valueProperty()
                .addListener(
                        (obs, oldV, newV) -> {
                            if (suppressInitSettingTargetComboEvents || newV == null || shell == null) {
                                return;
                            }
                            Map<String, String> ui = shell.snapshotUiEnv();
                            if (!FactorySiteComboPresentation.isSelectable(newV, ui)) {
                                refreshInitSettingTargetComboFromStore();
                                return;
                            }
                            shell.switchActiveFactorySite(newV);
                        });
        setInitSettingTargetComboValueSilently(GlobalInitSettingTarget.load());
    }

    void refreshInitSettingTargetComboPresentation() {
        if (initSettingTargetCombo != null) {
            initSettingTargetCombo.requestLayout();
            initSettingTargetCombo.getSelectionModel().select(initSettingTargetCombo.getValue());
        }
    }

    void setInitSettingTargetComboDisabled(boolean disabled) {
        if (initSettingTargetCombo != null) {
            initSettingTargetCombo.setDisable(disabled);
        }
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
        dialog.setTitle("確認");
        dialog.setHeaderText(null);
        dialog.setContentText(
                "現在の UI とセッションをパッケージ既定としてリポジトリの init_setting に書き出します。"
                        + "確認のため「111」と入力してください。");
        Optional<String> ans = dialog.showAndWait();
        if (ans.isEmpty() || !"111".equals(ans.get().trim())) {
            return;
        }
        try {
            shell.preparePackageDefaultsExport();
            InitSettingPersistence.savePackageDefaults(
                    shell.snapshotUiEnv(),
                    shell.snapshotDesktopSessionForExport(),
                    GlobalInitSettingTarget.load(),
                    shell.snapshotJuchuHeaderAliasRegistryForExport());
            Alert ok = new Alert(AlertType.INFORMATION);
            if (shell.primaryStageForDialogs() != null) {
                ok.initOwner(shell.primaryStageForDialogs());
            }
            ok.setTitle("完了");
            ok.setHeaderText(null);
            FactorySite t = GlobalInitSettingTarget.load();
            ok.setContentText(
                    "書き出しました。init_setting に "
                            + InitSettingPaths.sessionDefaultsFileForFactory(t)
                            + "（メイン／子タブ・依頼書フォーム候補含む）、"
                            + InitSettingPaths.tableColumnDefaultsFileForFactory(t)
                            + "（列・行高・納期管理ビュー等）、"
                            + InitSettingPaths.juchuHeaderAliasesFileForFactory(t)
                            + "（列定義ウィザード）を出力しました。");
            ok.showAndWait();
        } catch (Exception ex) {
            Alert err = new Alert(AlertType.ERROR);
            if (shell.primaryStageForDialogs() != null) {
                err.initOwner(shell.primaryStageForDialogs());
            }
            err.setTitle("エラー");
            err.setHeaderText(null);
            err.setContentText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            err.showAndWait();
        }
    }
}
