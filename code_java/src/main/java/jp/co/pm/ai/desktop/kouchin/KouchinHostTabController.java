package jp.co.pm.ai.desktop.kouchin;

import javafx.fxml.FXML;
import javafx.scene.control.TabPane;
import jp.co.pm.ai.desktop.MainShellController;

/**
 * 「後加工工賃」メインタブのホスト。子タブ「検証」「月次トレンド」「参照先」。
 */
public class KouchinHostTabController {

    @FXML private TabPane innerTabPane;
    @FXML private KouchinVerifyTabController verifyTabController;
    @FXML private KouchinTrendTabController trendTabController;
    @FXML private KouchinSourcesTabController sourcesTabController;

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
        if (verifyTabController != null) {
            verifyTabController.bindShell(shell, this);
        }
        if (trendTabController != null) {
            trendTabController.bindShell(shell, this);
        }
        if (sourcesTabController != null) {
            sourcesTabController.bindShell(shell, this);
        }
    }

    public void onMainShellTabSelected() {
        notifyActiveChildSelected();
    }

    public void onMainShellTabDeselected() {
        if (verifyTabController != null) {
            verifyTabController.onMainShellTabDeselected();
        }
        if (trendTabController != null) {
            trendTabController.onMainShellTabDeselected();
        }
        if (sourcesTabController != null) {
            sourcesTabController.onMainShellTabDeselected();
        }
    }

    public boolean hasUnappliedSourceEdits() {
        return sourcesTabController != null && sourcesTabController.hasUnappliedEdits();
    }

    public void onSourcesApplied() {
        if (verifyTabController != null) {
            verifyTabController.reloadDiscovery();
        }
    }

    public KouchinVerifyTabController verifyTab() {
        return verifyTabController;
    }

    public KouchinTrendTabController trendTab() {
        return trendTabController;
    }

    public KouchinSourcesTabController sourcesTab() {
        return sourcesTabController;
    }

    private void notifyActiveChildSelected() {
        if (innerTabPane == null) {
            return;
        }
        int idx = innerTabPane.getSelectionModel().getSelectedIndex();
        if (verifyTabController != null) {
            if (idx == 0) {
                verifyTabController.onMainShellTabSelected();
            } else {
                verifyTabController.onMainShellTabDeselected();
            }
        }
        if (trendTabController != null) {
            if (idx == 1) {
                trendTabController.onMainShellTabSelected();
            } else {
                trendTabController.onMainShellTabDeselected();
            }
        }
        if (sourcesTabController != null) {
            if (idx == 2) {
                sourcesTabController.onMainShellTabSelected();
            } else {
                sourcesTabController.onMainShellTabDeselected();
            }
        }
    }
}
