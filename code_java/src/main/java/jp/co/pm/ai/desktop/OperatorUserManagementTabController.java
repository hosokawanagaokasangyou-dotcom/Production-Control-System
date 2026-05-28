package jp.co.pm.ai.desktop;

import java.util.Optional;

import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;

import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

/** 工場別の配台システム操作者名（起動時選択肢）の管理タブ。 */
public final class OperatorUserManagementTabController {

    private MainShellController shell;

    @FXML
    private Label factoryLabel;

    @FXML
    private Label sessionOperatorLabel;

    @FXML
    private Button changeSessionOperatorButton;

    @FXML
    private TextField newNameField;

    @FXML
    private Button addNameButton;

    @FXML
    private Button removeNameButton;

    @FXML
    private Button resetDefaultsButton;

    @FXML
    private Button refreshButton;

    @FXML
    private ListView<String> nameListView;

    @FXML
    private void initialize() {
        refreshPresentation();
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        refreshPresentation();
    }

    void refreshPresentationQuietly() {
        refreshPresentation();
    }

    @FXML
    private void onRefreshAction() {
        refreshPresentation();
    }

    @FXML
    private void onAddNameAction() {
        if (shell == null) {
            return;
        }
        String name = newNameField != null && newNameField.getText() != null
                ? newNameField.getText().strip()
                : "";
        if (name.isBlank()) {
            warn("追加", "名前を入力してください。");
            return;
        }
        try {
            FactorySite site = GlobalInitSettingTarget.load();
            FactoryOperatorUserStore.addName(site, name);
            if (newNameField != null) {
                newNameField.clear();
            }
            refreshPresentation();
            shell.appendLog("[operator-user] 名前を追加: " + name + " （" + site.displayLabelJa() + "）");
        } catch (Exception ex) {
            warn("追加", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onRemoveNameAction() {
        if (shell == null || nameListView == null) {
            return;
        }
        String sel = nameListView.getSelectionModel().getSelectedItem();
        if (sel == null || sel.isBlank()) {
            warn("削除", "削除する名前を一覧から選んでください。");
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("削除の確認");
        confirm.setHeaderText(null);
        confirm.setContentText("「" + sel + "」をこの工場の一覧から削除しますか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            FactorySite site = GlobalInitSettingTarget.load();
            FactoryOperatorUserStore.removeName(site, sel);
            refreshPresentation();
            shell.appendLog("[operator-user] 名前を削除: " + sel + " （" + site.displayLabelJa() + "）");
            if (FactoryOperatorUserStore.sessionOperatorName().isBlank()) {
                shell.requireOperatorSelectionForFactory(site, false);
            }
        } catch (Exception ex) {
            warn("削除", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onResetDefaultsAction() {
        if (shell == null) {
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("既定へ戻す");
        confirm.setHeaderText(null);
        confirm.setContentText(
                "この工場の名前一覧を既定（砂田・古家・図司・細川）に戻します。よろしいですか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            FactorySite site = GlobalInitSettingTarget.load();
            FactoryOperatorUserStore.resetNamesToDefaults(site);
            refreshPresentation();
            shell.appendLog("[operator-user] 名前一覧を既定に戻しました（" + site.displayLabelJa() + "）");
            if (FactoryOperatorUserStore.sessionOperatorName().isBlank()) {
                shell.requireOperatorSelectionForFactory(site, false);
            }
        } catch (Exception ex) {
            warn("既定へ戻す", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onChangeSessionOperatorAction() {
        if (shell == null) {
            return;
        }
        shell.requireOperatorSelectionForFactory(GlobalInitSettingTarget.load(), false);
        refreshPresentation();
    }

    private void refreshPresentation() {
        FactorySite site = GlobalInitSettingTarget.load();
        if (factoryLabel != null) {
            factoryLabel.setText("対象工場: " + site.displayLabelJa());
        }
        if (sessionOperatorLabel != null) {
            String op = FactoryOperatorUserStore.sessionOperatorName();
            sessionOperatorLabel.setText(
                    op.isBlank()
                            ? "現在の操作者: （未選択）"
                            : "現在の操作者: " + op);
        }
        if (nameListView != null) {
            try {
                nameListView.getItems().setAll(FactoryOperatorUserStore.namesForFactory(site));
            } catch (Exception ex) {
                nameListView.getItems().clear();
            }
        }
        if (shell != null) {
            shell.refreshMainRunTabOperatorLabel();
        }
    }

    private void warn(String title, String msg) {
        Alert a = new Alert(AlertType.WARNING);
        a.setTitle(title);
        a.setHeaderText(null);
        a.setContentText(msg);
        if (shell != null && shell.primaryStageForDialogs() != null) {
            a.initOwner(shell.primaryStageForDialogs());
        }
        a.showAndWait();
    }
}
