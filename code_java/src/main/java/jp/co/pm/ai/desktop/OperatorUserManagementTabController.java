package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.io.FactoryOperatorUserBackupStore;
import jp.co.pm.ai.desktop.io.FactoryOperatorUserBackupStore.FactoryOperatorUserBackupEntry;

/** 工場別の配台システム操作者名と PIN（4～10 桁）の管理タブ（管理者パスワードで開く）。 */
public final class OperatorUserManagementTabController {

    private static final DateTimeFormatter BACKUP_TS =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm").withZone(ZoneId.systemDefault());

    static final class OperatorRow {
        private final SimpleStringProperty name = new SimpleStringProperty();
        private final SimpleStringProperty pinStatus = new SimpleStringProperty();

        OperatorRow(String name, String pinStatus) {
            this.name.set(name);
            this.pinStatus.set(pinStatus);
        }

        String getName() {
            return name.get();
        }

        SimpleStringProperty nameProperty() {
            return name;
        }

        String getPinStatus() {
            return pinStatus.get();
        }

        SimpleStringProperty pinStatusProperty() {
            return pinStatus;
        }
    }

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
    private Button issuePinButton;

    @FXML
    private Button unlockPinButton;

    @FXML
    private Button refreshButton;

    @FXML
    private TableView<OperatorRow> operatorTableView;

    @FXML
    private TextField backupLabelField;

    @FXML
    private Button createBackupButton;

    @FXML
    private Button restoreBackupButton;

    @FXML
    private Button refreshBackupButton;

    @FXML
    private ListView<FactoryOperatorUserBackupEntry> backupListView;

    @FXML
    private Label backupStorePathLabel;

    @FXML
    private void initialize() {
        if (operatorTableView != null) {
            TableColumn<OperatorRow, String> nameCol = new TableColumn<>("名前");
            nameCol.setCellValueFactory(new PropertyValueFactory<>("name"));
            nameCol.setPrefWidth(180);
            TableColumn<OperatorRow, String> pinCol = new TableColumn<>("PIN");
            pinCol.setCellValueFactory(new PropertyValueFactory<>("pinStatus"));
            pinCol.setPrefWidth(120);
            operatorTableView.getColumns().setAll(nameCol, pinCol);
            operatorTableView.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        }
        if (backupListView != null) {
            backupListView.setCellFactory(
                    lv ->
                            new ListCell<>() {
                                @Override
                                protected void updateItem(FactoryOperatorUserBackupEntry item, boolean empty) {
                                    super.updateItem(item, empty);
                                    if (empty || item == null) {
                                        setText(null);
                                        return;
                                    }
                                    String ts = BACKUP_TS.format(Instant.ofEpochMilli(item.createdAtMillis()));
                                    String by =
                                            item.createdByOperator() != null
                                                            && !item.createdByOperator().isBlank()
                                                    ? item.createdByOperator()
                                                    : "（不明）";
                                    setText(ts + "  —  " + item.displayLabel() + "  [作成: " + by + "]");
                                }
                            });
        }
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
        if (shell == null || operatorTableView == null) {
            return;
        }
        OperatorRow sel = operatorTableView.getSelectionModel().getSelectedItem();
        if (sel == null || sel.getName().isBlank()) {
            warn("削除", "削除する名前を一覧から選んでください。");
            return;
        }
        String name = sel.getName();
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("削除の確認");
        confirm.setHeaderText(null);
        confirm.setContentText("「" + name + "」をこの工場の一覧から削除しますか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            FactorySite site = GlobalInitSettingTarget.load();
            FactoryOperatorUserStore.removeName(site, name);
            refreshPresentation();
            shell.appendLog("[operator-user] 名前を削除: " + name + " （" + site.displayLabelJa() + "）");
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
    private void onIssuePinAction() {
        if (shell == null || operatorTableView == null) {
            return;
        }
        OperatorRow sel = operatorTableView.getSelectionModel().getSelectedItem();
        if (sel == null || sel.getName().isBlank()) {
            warn("PIN 発行", "PIN を発行するユーザーを一覧から選んでください。");
            return;
        }
        String name = sel.getName();
        FactorySite site = GlobalInitSettingTarget.load();
        boolean reissue;
        try {
            reissue = FactoryOperatorUserStore.hasPin(site, name);
        } catch (IOException ex) {
            warn("PIN 発行", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle(reissue ? "PIN 再発行" : "PIN 発行");
        confirm.setHeaderText(null);
        confirm.setContentText(
                reissue
                        ? "「" + name + "」の PIN を再発行します。旧 PIN は使えなくなります。よろしいですか？"
                        : "「" + name + "」に PIN（" + FactoryOperatorUserStore.pinLengthRangeDescriptionJa() + "）を新規発行します。よろしいですか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            String pin = FactoryOperatorUserStore.issuePin(site, name);
            refreshPresentation();
            shell.appendLog(
                    "[operator-user] PIN を"
                            + (reissue ? "再発行" : "発行")
                            + ": "
                            + name
                            + " （"
                            + site.displayLabelJa()
                            + "）");
            Alert info = new Alert(AlertType.INFORMATION);
            info.setTitle(reissue ? "PIN 再発行完了" : "PIN 発行完了");
            info.setHeaderText(null);
            info.setContentText(
                    "操作者「"
                            + name
                            + "」の PIN は "
                            + pin
                            + " です。\n"
                            + "この画面を閉じると再表示できません。必ず控えてください。");
            if (shell.primaryStageForDialogs() != null) {
                info.initOwner(shell.primaryStageForDialogs());
            }
            info.showAndWait();
        } catch (Exception ex) {
            warn(reissue ? "PIN 再発行" : "PIN 発行", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onUnlockPinAction() {
        if (shell == null || operatorTableView == null) {
            return;
        }
        OperatorRow sel = operatorTableView.getSelectionModel().getSelectedItem();
        if (sel == null || sel.getName().isBlank()) {
            warn("ロック解除", "ロック解除するユーザーを一覧から選んでください。");
            return;
        }
        String name = sel.getName();
        FactorySite site = GlobalInitSettingTarget.load();
        try {
            if (!FactoryOperatorUserStore.isPinLocked(site, name)) {
                warn("ロック解除", "「" + name + "」は PIN ロックされていません。");
                return;
            }
        } catch (IOException ex) {
            warn("ロック解除", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("PIN ロック解除");
        confirm.setHeaderText(null);
        confirm.setContentText("「" + name + "」の PIN ロックを解除します。よろしいですか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            FactoryOperatorUserStore.unlockPin(site, name);
            refreshPresentation();
            shell.appendLog(
                    "[operator-user] PIN ロック解除: "
                            + name
                            + " （"
                            + site.displayLabelJa()
                            + "）");
        } catch (Exception ex) {
            warn("ロック解除", ex.getMessage() != null ? ex.getMessage() : ex.toString());
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

    @FXML
    private void onCreateBackupAction() {
        if (shell == null) {
            return;
        }
        String label =
                backupLabelField != null && backupLabelField.getText() != null
                        ? backupLabelField.getText().strip()
                        : "";
        try {
            var ui = shell.snapshotUiEnv();
            FactoryOperatorUserBackupEntry created =
                    FactoryOperatorUserBackupStore.createManualBackup(ui, label);
            if (backupLabelField != null) {
                backupLabelField.clear();
            }
            refreshBackupList(ui);
            refreshPresentation();
            shell.appendLog(
                    "[operator-user-backup] 手動バックアップ: "
                            + created.displayLabel()
                            + " ("
                            + created.id()
                            + ")");
        } catch (Exception ex) {
            warn("バックアップ", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onRestoreBackupAction() {
        if (shell == null || backupListView == null) {
            return;
        }
        FactoryOperatorUserBackupEntry sel = backupListView.getSelectionModel().getSelectedItem();
        if (sel == null) {
            warn("復元", "復元するバックアップを一覧から選んでください。");
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("復元の確認");
        confirm.setHeaderText(null);
        confirm.setContentText(
                "選択したバックアップ（"
                        + sel.displayLabel()
                        + "）で現行のユーザー管理ファイル（"
                        + AppPaths.FACTORY_OPERATOR_USERS_BIN
                        + "）を上書きします。\n\n続行しますか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            var ui = shell.snapshotUiEnv();
            FactoryOperatorUserBackupStore.restoreFromBackup(sel, ui);
            refreshPresentation();
            shell.appendLog("[operator-user-backup] 復元しました: " + sel.displayLabel());
            Alert done = new Alert(AlertType.INFORMATION);
            done.setTitle("復元完了");
            done.setHeaderText(null);
            done.setContentText("バックアップからユーザー管理情報を復元しました。");
            if (shell.primaryStageForDialogs() != null) {
                done.initOwner(shell.primaryStageForDialogs());
            }
            done.showAndWait();
        } catch (Exception ex) {
            warn("復元", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onRefreshBackupAction() {
        if (shell == null) {
            return;
        }
        refreshBackupList(shell.snapshotUiEnv());
    }

    private void refreshBackupList(java.util.Map<String, String> ui) {
        if (backupListView != null) {
            backupListView.getItems().setAll(FactoryOperatorUserBackupStore.loadIndex(ui));
        }
        if (backupStorePathLabel != null) {
            backupStorePathLabel.setText(
                    "バックアップ先: "
                            + FactoryOperatorUserBackupStore.resolveBackupsRoot(ui)
                            + "　保持上限 "
                            + FactoryOperatorUserBackupStore.MAX_BACKUP_GENERATIONS
                            + " 世代");
        }
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
        if (operatorTableView != null) {
            try {
                List<String> names = FactoryOperatorUserStore.namesForFactory(site);
                List<OperatorRow> rows = new ArrayList<>();
                for (String name : names) {
                    rows.add(
                            new OperatorRow(
                                    name, FactoryOperatorUserStore.pinStatusLabel(site, name)));
                }
                operatorTableView.setItems(FXCollections.observableArrayList(rows));
            } catch (IOException ex) {
                operatorTableView.setItems(FXCollections.observableArrayList());
            }
        }
        if (shell != null) {
            shell.refreshMainRunTabOperatorLabel();
            refreshBackupList(shell.snapshotUiEnv());
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
