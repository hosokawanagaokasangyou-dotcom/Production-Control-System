package jp.co.pm.ai.desktop;

import javafx.fxml.FXML;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;

/**
 * 「加工トレンド」メインタブのホスト。子タブ「加工量」「加工賃」を内包する。
 */
public class ProcessingTrendHostTabController {

    @FXML private TabPane innerTabPane;
    @FXML private ProcessingTrendTabController volumeTabController;
    @FXML private ProcessingFeeTrendTabController feeTabController;

    @FXML
    private void initialize() {
        if (innerTabPane != null) {
            innerTabPane
                    .getSelectionModel()
                    .selectedItemProperty()
                    .addListener((obs, o, n) -> notifyActiveChildSelected());
        }
    }

    public void bindShell(MainShellController shell) {
        if (volumeTabController != null) {
            volumeTabController.bindShell(shell);
        }
        if (feeTabController != null) {
            feeTabController.bindShell(shell);
        }
    }

    public void onMainShellTabSelected() {
        notifyActiveChildSelected();
    }

    public void onMainShellTabDeselected() {
        if (volumeTabController != null) {
            volumeTabController.onMainShellTabDeselected();
        }
        if (feeTabController != null) {
            feeTabController.onMainShellTabDeselected();
        }
    }

    public ProcessingTrendTabController volumeTab() {
        return volumeTabController;
    }

    public ProcessingFeeTrendTabController feeTab() {
        return feeTabController;
    }

    private void notifyActiveChildSelected() {
        if (innerTabPane == null) {
            return;
        }
        int idx = innerTabPane.getSelectionModel().getSelectedIndex();
        if (idx == 1) {
            if (volumeTabController != null) {
                volumeTabController.onMainShellTabDeselected();
            }
            if (feeTabController != null) {
                feeTabController.onMainShellTabSelected();
            }
        } else {
            if (feeTabController != null) {
                feeTabController.onMainShellTabDeselected();
            }
            if (volumeTabController != null) {
                volumeTabController.onMainShellTabSelected();
            }
        }
    }
}
