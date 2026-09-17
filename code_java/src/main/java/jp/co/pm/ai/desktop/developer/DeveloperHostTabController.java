package jp.co.pm.ai.desktop.developer;

import javafx.fxml.FXML;
import javafx.scene.control.TabPane;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.OperatorActionLogTabController;

/**
 * 「開発」メインタブ。子タブ「実行時エラー」「操作ログ」「バックアップ」。
 */
public class DeveloperHostTabController {

    @FXML private TabPane innerTabPane;
    @FXML private DeveloperRuntimeErrorTabController runtimeErrorTabController;
    @FXML private OperatorActionLogTabController operatorActionLogTabController;
    @FXML private DeveloperShareBackupTabController shareBackupTabController;

    @FXML
    private void initialize() {
        if (innerTabPane != null) {
            innerTabPane.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
            innerTabPane
                    .getSelectionModel()
                    .selectedIndexProperty()
                    .addListener((obs, a, b) -> notifyActiveChildSelected());
        }
    }

    public void bindShell(MainShellController shell) {
        if (runtimeErrorTabController != null) {
            runtimeErrorTabController.bindShell(shell);
        }
        if (operatorActionLogTabController != null) {
            operatorActionLogTabController.bindShell(shell);
        }
        if (shareBackupTabController != null) {
            shareBackupTabController.bindShell(shell);
        }
    }

    public void onMainShellTabSelected() {
        notifyActiveChildSelected();
    }

    private void notifyActiveChildSelected() {
        int idx = innerTabPane != null ? innerTabPane.getSelectionModel().getSelectedIndex() : 0;
        if (runtimeErrorTabController != null && idx == 0) {
            runtimeErrorTabController.onMainShellTabSelected();
        }
        if (operatorActionLogTabController != null && idx == 1) {
            operatorActionLogTabController.onMainShellTabSelected();
        }
        if (shareBackupTabController != null && idx == 2) {
            shareBackupTabController.onMainShellTabSelected();
        }
    }
}
