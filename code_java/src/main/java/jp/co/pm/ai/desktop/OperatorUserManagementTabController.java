package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
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
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.PasswordField;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.io.FactoryOperatorUserBackupStore;
import jp.co.pm.ai.desktop.print.FactoryOperatorUserPdfExporter;
import jp.co.pm.ai.desktop.io.FactoryOperatorUserBackupStore.FactoryOperatorUserBackupEntry;

/** 工場別の配台システム操作者名と PIN（4～10 桁）の管理タブ（管理者パスワードで開く）。 */
public final class OperatorUserManagementTabController {

    private static final DateTimeFormatter BACKUP_TS =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm").withZone(ZoneId.systemDefault());

    static final class OperatorRow {
        private final SimpleStringProperty name = new SimpleStringProperty();
        private final SimpleStringProperty pinStatus = new SimpleStringProperty();
        private final SimpleStringProperty adminPin = new SimpleStringProperty();

        OperatorRow(String name, String pinStatus, String adminPin) {
            this.name.set(name);
            this.pinStatus.set(pinStatus);
            this.adminPin.set(adminPin);
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

        String getAdminPin() {
            return adminPin.get();
        }

        SimpleStringProperty adminPinProperty() {
            return adminPin;
        }
    }

    private MainShellController shell;
    private boolean suppressManagedFactoryListener;

    @FXML
    private ComboBox<FactorySite> managedFactoryCombo;

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
    private Button assignPinButton;

    @FXML
    private Button unlockPinButton;

    @FXML
    private Button viewPinButton;

    @FXML
    private Button refreshButton;

    @FXML
    private Button exportUsersPdfButton;

    @FXML
    private Button openUsersPdfButton;

    @FXML
    private Label usersPdfPathLabel;

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
            nameCol.setCellValueFactory(row -> row.getValue().nameProperty());
            nameCol.setPrefWidth(180);
            TableColumn<OperatorRow, String> pinCol = new TableColumn<>("状態");
            pinCol.setCellValueFactory(row -> row.getValue().pinStatusProperty());
            pinCol.setPrefWidth(100);
            TableColumn<OperatorRow, String> adminPinCol = new TableColumn<>("PIN（管理者閲覧）");
            adminPinCol.setCellValueFactory(row -> row.getValue().adminPinProperty());
            adminPinCol.setPrefWidth(140);
            operatorTableView.getColumns().setAll(nameCol, pinCol, adminPinCol);
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
        wireManagedFactoryCombo();
        refreshPresentation();
    }

    private void wireManagedFactoryCombo() {
        if (managedFactoryCombo == null) {
            return;
        }
        managedFactoryCombo.setDisable(true);
        syncManagedFactoryComboToAppFactory();
    }

    private FactorySite effectiveAppFactory() {
        return shell != null
                ? GlobalInitSettingTarget.loadEffective(shell.snapshotUiEnv())
                : GlobalInitSettingTarget.load();
    }

    private void syncManagedFactoryComboToAppFactory() {
        if (managedFactoryCombo == null) {
            return;
        }
        FactorySite app = effectiveAppFactory();
        suppressManagedFactoryListener = true;
        try {
            managedFactoryCombo.getItems().setAll(app);
            managedFactoryCombo.setValue(app);
        } finally {
            suppressManagedFactoryListener = false;
        }
    }

    private FactorySite managedFactory() {
        return effectiveAppFactory();
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
            FactorySite site = managedFactory();
            String pin = FactoryOperatorUserStore.addName(site, name);
            if (newNameField != null) {
                newNameField.clear();
            }
            refreshPresentation();
            shell.appendLog("[operator-user] 名前を追加: " + name + " （" + site.displayLabelJa() + "）");
            Alert info = new Alert(AlertType.INFORMATION);
            info.setTitle("ユーザー追加完了");
            info.setHeaderText(null);
            info.setContentText(
                    "「"
                            + name
                            + "」を追加しました。\n"
                            + "初期 PIN（ランダム）は "
                            + pin
                            + " です。\n"
                            + "初回ログイン時に PIN 変更が必要です。\n"
                            + "この画面を閉じると再表示できません。必ず控えてください。");
            if (shell.primaryStageForDialogs() != null) {
                info.initOwner(shell.primaryStageForDialogs());
            }
            info.showAndWait();
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
            FactorySite site = managedFactory();
            FactoryOperatorUserStore.removeName(site, name);
            refreshPresentation();
            shell.appendLog("[operator-user] 名前を削除: " + name + " （" + site.displayLabelJa() + "）");
            if (site == GlobalInitSettingTarget.load()
                    && FactoryOperatorUserStore.sessionOperatorName().isBlank()) {
                shell.requireOperatorSelectionForFactory(GlobalInitSettingTarget.load(), false);
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
            FactorySite site = managedFactory();
            FactoryOperatorUserStore.resetNamesToDefaults(site);
            refreshPresentation();
            shell.appendLog("[operator-user] 名前一覧を既定に戻しました（" + site.displayLabelJa() + "）");
            if (site == GlobalInitSettingTarget.load()
                    && FactoryOperatorUserStore.sessionOperatorName().isBlank()) {
                shell.requireOperatorSelectionForFactory(GlobalInitSettingTarget.load(), false);
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
        FactorySite site = managedFactory();
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
                        ? "「"
                                + name
                                + "」の PIN を再発行します。旧 PIN は使えなくなります。\n"
                                + "初回ログイン時に PIN 変更が必要です。よろしいですか？"
                        : "「"
                                + name
                                + "」にランダム PIN（"
                                + FactoryOperatorUserStore.pinLengthRangeDescriptionJa()
                                + "）を新規発行します。\n"
                                + "初回ログイン時に PIN 変更が必要です。よろしいですか？");
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
                            + "初回ログイン時に PIN 変更が必要です。\n"
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
    private void onAssignPinManuallyAction() {
        if (shell == null || operatorTableView == null) {
            return;
        }
        OperatorRow sel = operatorTableView.getSelectionModel().getSelectedItem();
        if (sel == null || sel.getName().isBlank()) {
            warn("PIN 手動設定", "PIN を設定するユーザーを一覧から選んでください。");
            return;
        }
        String name = sel.getName();
        FactorySite site = managedFactory();
        Dialog<ButtonType> dialog = new Dialog<>();
        if (shell.primaryStageForDialogs() != null) {
            dialog.initOwner(shell.primaryStageForDialogs());
        }
        dialog.setTitle("PIN 手動設定");
        dialog.setHeaderText(null);
        Label hint =
                new Label(
                        "操作者「"
                                + name
                                + "」（"
                                + site.displayLabelJa()
                                + "）の PIN（"
                                + FactoryOperatorUserStore.pinLengthRangeDescriptionJa()
                                + "）を入力してください。");
        hint.setWrapText(true);
        PasswordField pinField = new PasswordField();
        pinField.setPromptText("PIN");
        PasswordField confirmField = new PasswordField();
        confirmField.setPromptText("PIN（確認）");
        dialog.getDialogPane().setContent(new VBox(8, hint, new Label("PIN:"), pinField, new Label("確認:"), confirmField));
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        Optional<ButtonType> ans = dialog.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        String pin = pinField.getText() != null ? pinField.getText().strip() : "";
        String confirm = confirmField.getText() != null ? confirmField.getText().strip() : "";
        if (FactoryOperatorUserStore.normalizePin(pin) == null) {
            warn("PIN 手動設定", FactoryOperatorUserStore.pinLengthRangeDescriptionJa() + "を入力してください。");
            return;
        }
        if (!pin.equals(confirm)) {
            warn("PIN 手動設定", "PIN と確認入力が一致しません。");
            return;
        }
        try {
            FactoryOperatorUserStore.assignPinByAdmin(site, name, pin);
            refreshPresentation();
            shell.appendLog(
                    "[operator-user] PIN を手動設定: "
                            + name
                            + " （"
                            + site.displayLabelJa()
                            + "）");
            Alert done = new Alert(AlertType.INFORMATION);
            done.setTitle("PIN 手動設定");
            done.setHeaderText(null);
            done.setContentText("操作者「" + name + "」の PIN を設定しました。");
            if (shell.primaryStageForDialogs() != null) {
                done.initOwner(shell.primaryStageForDialogs());
            }
            done.showAndWait();
        } catch (Exception ex) {
            warn("PIN 手動設定", ex.getMessage() != null ? ex.getMessage() : ex.toString());
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
        FactorySite site = managedFactory();
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
    private void onViewPinAction() {
        if (shell == null || operatorTableView == null) {
            return;
        }
        OperatorRow sel = operatorTableView.getSelectionModel().getSelectedItem();
        if (sel == null || sel.getName().isBlank()) {
            warn("PIN 閲覧", "PIN を表示するユーザーを一覧から選んでください。");
            return;
        }
        FactorySite site = managedFactory();
        String name = sel.getName();
        try {
            if (!FactoryOperatorUserStore.hasPin(site, name)) {
                warn("PIN 閲覧", "「" + name + "」には PIN が設定されていません。");
                return;
            }
            Optional<String> pin = FactoryOperatorUserStore.adminViewablePin(site, name);
            Alert info = new Alert(AlertType.INFORMATION);
            info.setTitle("PIN 閲覧（管理者）");
            info.setHeaderText(null);
            info.setContentText(
                    "操作者「"
                            + name
                            + "」（"
                            + site.displayLabelJa()
                            + "）の PIN は "
                            + pin.orElse("（記録なし。PIN 再発行で新しい PIN を確認してください）")
                            + " です。");
            if (shell.primaryStageForDialogs() != null) {
                info.initOwner(shell.primaryStageForDialogs());
            }
            info.showAndWait();
        } catch (Exception ex) {
            warn("PIN 閲覧", ex.getMessage() != null ? ex.getMessage() : ex.toString());
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
    private void onExportUsersPdfAction() {
        if (shell == null) {
            return;
        }
        FactorySite site = managedFactory();
        try {
            var ui = shell.snapshotUiEnv();
            Path outputPath = FactoryOperatorUserPdfExporter.resolveOutputPath(ui, site);
            FactoryOperatorUserPdfExporter.export(
                    outputPath,
                    site,
                    buildPdfRows(site),
                    FactoryOperatorUserStore.sessionOperatorName(),
                    Instant.now(),
                    FactoryOperatorUserStore.storePath().toString());
            lastExportedUsersPdfPath = outputPath;
            refreshUsersPdfControls(ui, site);
            shell.appendLog("[operator-user-pdf] 出力: " + outputPath);
            try {
                DesktopFileOpener.openFile(outputPath);
            } catch (Exception openEx) {
                shell.appendLog(
                        "[operator-user-pdf] open: "
                                + (openEx.getMessage() != null ? openEx.getMessage() : openEx.toString()));
            }
            Alert done = new Alert(AlertType.INFORMATION);
            done.setTitle("PDF 出力");
            done.setHeaderText(null);
            done.setContentText("ユーザー管理情報を PDF 化しました。\n" + outputPath);
            if (shell.primaryStageForDialogs() != null) {
                done.initOwner(shell.primaryStageForDialogs());
            }
            done.showAndWait();
        } catch (Exception ex) {
            warn("PDF 出力", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onOpenUsersPdfAction() {
        if (shell == null) {
            return;
        }
        FactorySite site = managedFactory();
        Path path = lastExportedUsersPdfPath;
        if (path == null || !Files.isRegularFile(path)) {
            path = FactoryOperatorUserPdfExporter.resolveOutputPath(shell.snapshotUiEnv(), site);
        }
        if (!Files.isRegularFile(path)) {
            warn("PDF を開く", "PDF が未作成です。先に PDF 出力を実行してください。");
            refreshUsersPdfControls(shell.snapshotUiEnv(), site);
            return;
        }
        try {
            DesktopFileOpener.openFile(path);
            shell.appendLog("[operator-user-pdf] 開く: " + path);
        } catch (Exception ex) {
            warn("PDF を開く", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
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

    private Path lastExportedUsersPdfPath;

    private List<FactoryOperatorUserPdfExporter.Row> buildPdfRows(FactorySite site) throws IOException {
        List<FactoryOperatorUserPdfExporter.Row> rows = new ArrayList<>();
        for (String name : FactoryOperatorUserStore.namesForFactory(site)) {
            rows.add(
                    new FactoryOperatorUserPdfExporter.Row(
                            name,
                            FactoryOperatorUserStore.pinStatusLabel(site, name),
                            FactoryOperatorUserStore.adminPinDisplayLabel(site, name)));
        }
        return rows;
    }

    private void refreshUsersPdfControls(java.util.Map<String, String> ui, FactorySite site) {
        Path expected = FactoryOperatorUserPdfExporter.resolveOutputPath(ui, site);
        if (Files.isRegularFile(expected)) {
            lastExportedUsersPdfPath = expected;
        }
        if (openUsersPdfButton != null) {
            openUsersPdfButton.setDisable(
                    lastExportedUsersPdfPath == null
                            || !Files.isRegularFile(lastExportedUsersPdfPath));
        }
        if (usersPdfPathLabel != null) {
            usersPdfPathLabel.setText("PDF 出力先: " + expected);
        }
    }

    private void refreshPresentation() {
        syncManagedFactoryComboToAppFactory();
        FactorySite site = effectiveAppFactory();
        if (factoryLabel != null) {
            factoryLabel.setText(
                    "環境変数の利用工場（"
                            + site.displayLabelJa()
                            + "）のユーザー一覧・PIN のみ編集できます。");
        }
        if (sessionOperatorLabel != null) {
            String op = FactoryOperatorUserStore.sessionOperatorName();
            sessionOperatorLabel.setText(
                    op.isBlank()
                            ? "現在の操作者: （未選択）"
                            : "現在の操作者: " + op + " （" + site.displayLabelJa() + "）");
        }
        if (changeSessionOperatorButton != null) {
            changeSessionOperatorButton.setDisable(false);
        }
        if (operatorTableView != null) {
            try {
                FactoryOperatorUserStore.ensureStoreFileOnDisk();
                List<String> names = FactoryOperatorUserStore.namesForFactory(site);
                List<OperatorRow> rows = new ArrayList<>();
                for (String name : names) {
                    rows.add(
                            new OperatorRow(
                                    name,
                                    FactoryOperatorUserStore.pinStatusLabel(site, name),
                                    FactoryOperatorUserStore.adminPinDisplayLabel(site, name)));
                }
                operatorTableView.setItems(FXCollections.observableArrayList(rows));
            } catch (IOException ex) {
                operatorTableView.setItems(FXCollections.observableArrayList());
                String msg = ex.getMessage() != null ? ex.getMessage() : ex.toString();
                if (shell != null) {
                    shell.appendLog("[operator-user] 一覧読込失敗: " + msg);
                }
                warn("ユーザー一覧", "操作者名設定の読込に失敗しました。\n" + msg);
            }
        }
        if (shell != null) {
            refreshBackupList(shell.snapshotUiEnv());
            refreshUsersPdfControls(shell.snapshotUiEnv(), site);
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
