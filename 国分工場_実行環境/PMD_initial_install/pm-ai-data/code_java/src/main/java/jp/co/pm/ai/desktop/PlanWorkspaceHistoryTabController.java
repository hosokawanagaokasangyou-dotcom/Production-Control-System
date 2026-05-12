package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
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

import com.fasterxml.jackson.databind.JsonNode;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.PlanWorkspaceSessionFragment;
import jp.co.pm.ai.desktop.config.PlanWorkspaceSnapshotStore;
import jp.co.pm.ai.desktop.config.PlanWorkspaceSnapshotStore.PlanWorkspaceSnapshotEntry;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchDocument;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchJsonIo;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * 配台ワークスペース（結果 JSON・ガント表示・列順）のスナップショット履歴タブ。
 */
public final class PlanWorkspaceHistoryTabController {

    private static final DateTimeFormatter TS =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm").withZone(ZoneId.systemDefault());

    private MainShellController shell;

    @FXML
    private TextField newSnapshotLabelField;

    @FXML
    private Button captureButton;

    @FXML
    private Button restoreButton;

    @FXML
    private Button renameButton;

    @FXML
    private Button deleteButton;

    @FXML
    private Button refreshButton;

    @FXML
    private ListView<PlanWorkspaceSnapshotEntry> snapshotListView;

    @FXML
    private Label storePathLabel;

    @FXML
    private void initialize() {
        if (snapshotListView != null) {
            snapshotListView.setCellFactory(
                    lv ->
                            new ListCell<>() {
                                @Override
                                protected void updateItem(PlanWorkspaceSnapshotEntry item, boolean empty) {
                                    super.updateItem(item, empty);
                                    if (empty || item == null) {
                                        setText(null);
                                        return;
                                    }
                                    String ts = TS.format(Instant.ofEpochMilli(item.createdAtMillis()));
                                    String lb = item.label() != null && !item.label().isBlank() ? item.label() : "（無題）";
                                    setText(ts + "  —  " + lb);
                                }
                            });
        }
        refreshList();
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        if (storePathLabel != null) {
            storePathLabel.setText(
                    "保存先: " + PlanWorkspaceSnapshotStore.rootDirectory().toAbsolutePath());
        }
    }

    @FXML
    private void onRefreshListAction() {
        refreshList();
    }

    private void refreshList() {
        if (snapshotListView == null) {
            return;
        }
        snapshotListView.getItems().setAll(PlanWorkspaceSnapshotStore.loadIndex());
    }

    @FXML
    private void onCaptureSnapshotAction() {
        if (shell == null) {
            return;
        }
        String label = newSnapshotLabelField != null && newSnapshotLabelField.getText() != null
                ? newSnapshotLabelField.getText().strip()
                : "";
        try {
            shell.persistDesktopSessionNow();
            PlanWorkspaceSessionFragment frag =
                    PlanWorkspaceSessionFragment.fromDesktopSession(shell.collectDesktopSessionSnapshot());
            JsonNode colPart = TableColumnOrderPersistence.capturePlanWorkspaceColumnOrderPartial();
            ResultDispatchDocument doc = shell.snapshotDispatchDocumentForPlanWorkspace();
            if (doc == null) {
                throw new IllegalStateException("配台計画手動修正タブが未初期化です");
            }
            Path tmp = Files.createTempFile("pm-plan-workspace-snap-", ".json");
            try {
                ResultDispatchJsonIo.write(tmp, doc);
                PlanWorkspaceSnapshotEntry created =
                        PlanWorkspaceSnapshotStore.appendSnapshot(label, frag, colPart, tmp);
                shell.tryExportResultDispatchTableXlsxNearJson(
                        PlanWorkspaceSnapshotStore.resultDispatchJsonPath(created));
            } finally {
                Files.deleteIfExists(tmp);
            }
            if (newSnapshotLabelField != null) {
                newSnapshotLabelField.clear();
            }
            shell.appendLog("[plan-workspace-snapshot] 保存しました");
            refreshList();
        } catch (Exception ex) {
            Alert a = new Alert(AlertType.ERROR);
            a.setTitle("スナップショット");
            a.setHeaderText("保存に失敗しました");
            a.setContentText(ex.getMessage() != null ? ex.getMessage() : ex.getClass().getSimpleName());
            if (shell.primaryStageForDialogs() != null) {
                a.initOwner(shell.primaryStageForDialogs());
            }
            a.showAndWait();
        }
    }

    @FXML
    private void onRestoreSnapshotAction() {
        if (shell == null) {
            return;
        }
        PlanWorkspaceSnapshotEntry sel = snapshotListView != null ? snapshotListView.getSelectionModel().getSelectedItem() : null;
        if (sel == null) {
            return;
        }
        Path canonical = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("復元の確認");
        confirm.setHeaderText(null);
        confirm.setContentText(
                "選択したスナップショットの内容で次を上書きします。\n"
                        + "・正規の結果_配台表 JSON（\n"
                        + canonical
                        + "）\n"
                        + "・同階層の 結果_配台表.xlsx（段階2と同一の export_result_dispatch_from_json 経路で再生成を試みます）\n"
                        + "・配台計画入力パス・段階1プレビュー・設備ガント表示・担当バッジ位置・列順（該当キーのみ）\n"
                        + "実行タブのブックパス等は維持されます。\n"
                        + "続行しますか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            shell.restorePlanWorkspaceSnapshot(sel);
            shell.appendLog("[plan-workspace-snapshot] 復元しました: " + sel.id());
            refreshList();
        } catch (Exception ex) {
            Alert a = new Alert(AlertType.ERROR);
            a.setTitle("復元");
            a.setHeaderText("復元に失敗しました");
            a.setContentText(ex.getMessage() != null ? ex.getMessage() : ex.getClass().getSimpleName());
            if (shell.primaryStageForDialogs() != null) {
                a.initOwner(shell.primaryStageForDialogs());
            }
            a.showAndWait();
        }
    }

    @FXML
    private void onRenameSnapshotAction() {
        if (shell == null) {
            return;
        }
        PlanWorkspaceSnapshotEntry sel = snapshotListView != null ? snapshotListView.getSelectionModel().getSelectedItem() : null;
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
            PlanWorkspaceSnapshotStore.updateEntryLabel(sel, r.get());
            refreshList();
        } catch (IOException ex) {
            Alert a = new Alert(AlertType.ERROR);
            a.setContentText(ex.getMessage());
            a.showAndWait();
        }
    }

    @FXML
    private void onDeleteSnapshotAction() {
        if (shell == null) {
            return;
        }
        PlanWorkspaceSnapshotEntry sel = snapshotListView != null ? snapshotListView.getSelectionModel().getSelectedItem() : null;
        if (sel == null) {
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setContentText("このスナップショットを削除しますか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            PlanWorkspaceSnapshotStore.deleteEntry(sel);
            refreshList();
        } catch (IOException ex) {
            Alert a = new Alert(AlertType.ERROR);
            a.setContentText(ex.getMessage());
            a.showAndWait();
        }
    }
}
