package jp.co.pm.ai.desktop.developer;

import javafx.fxml.FXML;
import javafx.scene.control.TabPane;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.OperatorActionLogTabController;

/**
 * 「開発」メインタブ。子タブ「実行時エラー」「操作ログ」。
 */
public class DeveloperHostTabController {

    @FXML private TabPane innerTabPane;
    @FXML private DeveloperRuntimeErrorTabController runtimeErrorTabController;
    @FXML private OperatorActionLogTabController operatorActionLogTabController;

    public void bindShell(MainShellController shell) {
        if (runtimeErrorTabController != null) {
            runtimeErrorTabController.bindShell(shell);
        }
        if (operatorActionLogTabController != null) {
            operatorActionLogTabController.bindShell(shell);
        }
        if (innerTabPane != null) {
            innerTabPane.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
            innerTabPane
                    .getSelectionModel()
                    .selectedIndexProperty()
                    .addListener((obs, a, b) -> onMainShellTabSelected());
        }
    }

    public void onMainShellTabSelected() {
        if (runtimeErrorTabController != null) {
            runtimeErrorTabController.onMainShellTabSelected();
        }
        if (operatorActionLogTabController != null) {
            operatorActionLogTabController.onMainShellTabSelected();
        }
    }
}
