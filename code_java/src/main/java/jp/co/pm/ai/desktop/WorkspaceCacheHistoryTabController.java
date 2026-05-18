package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Optional;

import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;
import javafx.scene.control.TextInputDialog;

import jp.co.pm.ai.desktop.config.Stage1AiCacheClearer;
import jp.co.pm.ai.desktop.config.WorkspaceCacheArchiveStore;
import jp.co.pm.ai.desktop.config.WorkspaceCacheArchiveStore.WorkspaceCacheArchiveEntry;

/**
 * ワークスペースキャッシュ（AI 備考・配台 shaped JSON 等）の退避履歴タブ。
 */
public final class WorkspaceCacheHistoryTabController {

    private static final DateTimeFormatter TS =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm").withZone(ZoneId.systemDefault());

    private MainShellController shell;

    @FXML
    private TextField archiveLabelField;

    @FXML
    private Button archiveOnlyButton;

    @FXML
    private Button archiveAndClearButton;

    @FXML
    private Button restoreButton;

    @FXML
    private Button renameButton;

    @FXML
    private Button deleteButton;

    @FXML
    private Button refreshButton;

    @FXML
    private ListView<WorkspaceCacheArchiveEntry> archiveListView;

    @FXML
    private Label storePathLabel;

    @FXML
    private void initialize() {
        if (archiveListView != null) {
            archiveListView.setCellFactory(
                    lv ->
                            new ListCell<>() {
                                @Override
                                protected void updateItem(
                                        WorkspaceCacheArchiveEntry item, boolean empty) {
                                    super.updateItem(item, empty);
                                    if (empty || item == null) {
                                        setText(null);
                                        return;
                                    }
                                    String ts = TS.format(Instant.ofEpochMilli(item.createdAtMillis()));
                                    String lb =
                                            item.label() != null && !item.label().isBlank()
                                                    ? item.label()
                                                    : "（無題）";
                                    String reason =
                                            reasonLabelJa(item.reason() != null ? item.reason() : "");
                                    setText(
                                            ts
                                                    + "  —  "
                                                    + lb
                                                    + "  ["
                                                    + reason
                                                    + ", "
                                                    + item.fileCount()
                                                    + " ファイル]");
                                }
                            });
        }
        refreshList();
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        if (storePathLabel != null) {
            storePathLabel.setText(
                    "保存先: " + WorkspaceCacheArchiveStore.rootDirectory().toAbsolutePath());
        }
    }

    void refreshListQuietly() {
        refreshList();
    }

    @FXML
    private void onRefreshListAction() {
        refreshList();
    }

    private void refreshList() {
        if (archiveListView == null) {
            return;
        }
        archiveListView.getItems().setAll(WorkspaceCacheArchiveStore.loadIndex());
    }

    @FXML
    private void onArchiveOnlyAction() {
        runArchive(false);
    }

    @FXML
    private void onArchiveAndClearAction() {
        runArchive(true);
    }

    private void runArchive(boolean clearAfter) {
        if (shell == null) {
            return;
        }
        String label =
                archiveLabelField != null && archiveLabelField.getText() != null
                        ? archiveLabelField.getText().strip()
                        : "";
        if (label.isBlank()) {
            label = clearAfter ? "手動・退避してクリア" : "手動・退避のみ";
        }
        try {
            if (!Stage1AiCacheClearer.hasAnyExistingDiskCache(shell.snapshotUiEnv())) {
                showInfo("退避", "退避対象のキャッシュファイルがありません。");
                return;
            }
            WorkspaceCacheArchiveEntry created =
                    WorkspaceCacheArchiveStore.archiveDiskCaches(
                            shell.snapshotUiEnv(), label, clearAfter ? "manual-clear" : "manual-archive");
            if (created == null) {
                showInfo("退避", "退避対象のキャッシュファイルがありません。");
                return;
            }
            shell.appendLog("[cache-archive] キャッシュを退避しました（ID: " + created.id() + "）。");
            if (clearAfter) {
                Stage1AiCacheClearer.ClearResult cleared =
                        Stage1AiCacheClearer.clearBeforeStage1Run(shell.snapshotUiEnv());
                for (String line : cleared.detailLines()) {
                    shell.appendLog(line);
                }
                shell.appendLog("[cache-archive] キャッシュをクリアしました。");
            }
            if (archiveLabelField != null) {
                archiveLabelField.clear();
            }
            refreshList();
        } catch (Exception ex) {
            showError("退避", "退避に失敗しました", ex);
        }
    }

    @FXML
    private void onRestoreArchiveAction() {
        if (shell == null) {
            return;
        }
        WorkspaceCacheArchiveEntry sel =
                archiveListView != null
                        ? archiveListView.getSelectionModel().getSelectedItem()
                        : null;
        if (sel == null) {
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("復元の確認");
        confirm.setHeaderText(null);
        confirm.setContentText(
                "選択した退避（"
                        + (sel.label() != null && !sel.label().isBlank() ? sel.label() : sel.id())
                        + "）の内容で、現在のワークスペース上のキャッシュファイルを上書きします。\n"
                        + "配台手動修正・納期管理ビューは再読み込みが必要になる場合があります。\n"
                        + "続行しますか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            shell.restoreWorkspaceCacheArchive(sel);
            refreshList();
        } catch (Exception ex) {
            showError("復元", "復元に失敗しました", ex);
        }
    }

    @FXML
    private void onRenameArchiveAction() {
        if (shell == null) {
            return;
        }
        WorkspaceCacheArchiveEntry sel =
                archiveListView != null
                        ? archiveListView.getSelectionModel().getSelectedItem()
                        : null;
        if (sel == null) {
            return;
        }
        TextInputDialog d = new TextInputDialog(sel.label() != null ? sel.label() : "");
        d.setTitle("ラベル変更");
        d.setHeaderText(null);
        d.setContentText("表示ラベル");
        if (shell.primaryStageForDialogs() != null) {
            d.initOwner(shell.primaryStageForDialogs());
        }
        Optional<String> r = d.showAndWait();
        if (r.isEmpty()) {
            return;
        }
        try {
            WorkspaceCacheArchiveStore.updateEntryLabel(sel, r.get());
            refreshList();
        } catch (IOException ex) {
            showError("ラベル変更", ex.getMessage(), ex);
        }
    }

    @FXML
    private void onDeleteArchiveAction() {
        if (shell == null) {
            return;
        }
        WorkspaceCacheArchiveEntry sel =
                archiveListView != null
                        ? archiveListView.getSelectionModel().getSelectedItem()
                        : null;
        if (sel == null) {
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setContentText("この退避履歴を削除しますか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            WorkspaceCacheArchiveStore.deleteEntry(sel);
            refreshList();
        } catch (IOException ex) {
            showError("削除", ex.getMessage(), ex);
        }
    }

    private static String reasonLabelJa(String reason) {
        if (reason == null || reason.isBlank()) {
            return "不明";
        }
        return switch (reason) {
            case "stage1-clear" -> "段階1クリア前";
            case "manual-clear" -> "手動・退避してクリア";
            case "manual-archive" -> "手動・退避のみ";
            default -> reason;
        };
    }

    private void showInfo(String title, String message) {
        Alert a = new Alert(AlertType.INFORMATION);
        a.setTitle(title);
        a.setHeaderText(null);
        a.setContentText(message);
        if (shell != null && shell.primaryStageForDialogs() != null) {
            a.initOwner(shell.primaryStageForDialogs());
        }
        a.showAndWait();
    }

    private void showError(String title, String header, Exception ex) {
        Alert a = new Alert(AlertType.ERROR);
        a.setTitle(title);
        a.setHeaderText(header);
        a.setContentText(ex.getMessage() != null ? ex.getMessage() : ex.getClass().getSimpleName());
        if (shell != null && shell.primaryStageForDialogs() != null) {
            a.initOwner(shell.primaryStageForDialogs());
        }
        a.showAndWait();
    }
}
