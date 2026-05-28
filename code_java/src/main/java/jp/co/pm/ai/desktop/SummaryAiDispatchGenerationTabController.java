package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
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

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.io.SummaryAiDispatchGenerationStore;
import jp.co.pm.ai.desktop.io.SummaryAiDispatchGenerationStore.SummaryAiDispatchGenerationEntry;

/**
 * サマリ {@link AppPaths#SUMMARY_AI_DISPATCH_XLSX} の世代退避・復元タブ。
 */
public final class SummaryAiDispatchGenerationTabController {

    private static final DateTimeFormatter TS =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm").withZone(ZoneId.systemDefault());

    private MainShellController shell;

    @FXML
    private TextField manualArchiveLabelField;

    @FXML
    private Button manualArchiveButton;

    @FXML
    private Button openButton;

    @FXML
    private Button restoreButton;

    @FXML
    private Button renameButton;

    @FXML
    private Button deleteButton;

    @FXML
    private Button refreshButton;

    @FXML
    private ListView<SummaryAiDispatchGenerationEntry> generationListView;

    @FXML
    private Label storePathLabel;

    @FXML
    private Label operatorScopeLabel;

    @FXML
    private Label historyHeadingLabel;

    @FXML
    private void initialize() {
        if (generationListView != null) {
            generationListView.setCellFactory(
                    lv ->
                            new ListCell<>() {
                                @Override
                                protected void updateItem(
                                        SummaryAiDispatchGenerationEntry item, boolean empty) {
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
                                            SummaryAiDispatchGenerationStore.reasonLabelJa(item.reason());
                                    String owner =
                                            item.operatorUser() != null && !item.operatorUser().isBlank()
                                                    ? item.operatorUser()
                                                    : "（不明）";
                                    setText(
                                            "["
                                                    + owner
                                                    + "] "
                                                    + ts
                                                    + "  —  "
                                                    + lb
                                                    + "  ["
                                                    + reason
                                                    + "]");
                                }
                            });
            generationListView
                    .getSelectionModel()
                    .selectedItemProperty()
                    .addListener((obs, old, sel) -> updateEntryActionButtons(sel));
        }
        refreshList();
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        refreshStorePathLabel();
    }

    void refreshListQuietly() {
        refreshList();
    }

    @FXML
    private void onRefreshListAction() {
        refreshList();
    }

    private void refreshList() {
        if (generationListView == null || shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        List<SummaryAiDispatchGenerationEntry> entries =
                SummaryAiDispatchGenerationStore.loadAllGenerations(ui);
        generationListView.getItems().setAll(entries);
        refreshScopeLabels(ui, entries);
        updateEntryActionButtons(generationListView.getSelectionModel().getSelectedItem());
    }

    private void refreshScopeLabels(
            Map<String, String> ui, List<SummaryAiDispatchGenerationEntry> allEntries) {
        String operator = SummaryAiDispatchGenerationStore.resolveOperatorUser(ui);
        long ownCount =
                allEntries.stream()
                        .filter(e -> SummaryAiDispatchGenerationStore.isCreatedByCurrentUser(e, ui))
                        .count();
        if (operatorScopeLabel != null) {
            operatorScopeLabel.setText(
                    "操作者: "
                            + operator
                            + "　自分の退避 "
                            + ownCount
                            + " / "
                            + SummaryAiDispatchGenerationStore.MAX_GENERATIONS_PER_USER
                            + " 件　一覧 "
                            + allEntries.size()
                            + " 件（全操作者）");
        }
        if (historyHeadingLabel != null) {
            historyHeadingLabel.setText("退避履歴（全操作者・新しい順）");
        }
    }

    private void updateEntryActionButtons(SummaryAiDispatchGenerationEntry sel) {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        boolean own = SummaryAiDispatchGenerationStore.isCreatedByCurrentUser(sel, ui);
        if (deleteButton != null) {
            deleteButton.setDisable(sel == null || !own);
        }
        if (renameButton != null) {
            renameButton.setDisable(sel == null || !own);
        }
    }

    @FXML
    private void onOpenCurrentSummaryAction() {
        if (shell == null) {
            return;
        }
        if (shell.isSummaryAiDispatchExportLocked()) {
            showInfo("現行を開く", "サマリ xlsx を作成中のため開けません。");
            return;
        }
        Path current = AppPaths.summaryAiDispatchXlsxPath(shell.snapshotUiEnv());
        if (!Files.isRegularFile(current)) {
            showInfo("現行を開く", "現行サマリ Excel が見つかりません:\n" + current);
            return;
        }
        try {
            DesktopFileOpener.openFileReadOnly(current);
            shell.appendLog("[summary-generation] opened current (read-only): " + current);
        } catch (Exception ex) {
            showError("現行を開く", "ファイルを開けませんでした", ex);
        }
    }

    private void refreshStorePathLabel() {
        if (storePathLabel == null || shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        storePathLabel.setText(
                "退避先: "
                        + SummaryAiDispatchGenerationStore.resolveUserGenerationsRoot(ui)
                        + "　現行（上書き先）: "
                        + AppPaths.summaryAiDispatchXlsxPath(ui));
    }

    @FXML
    private void onManualArchiveAction() {
        if (shell == null) {
            return;
        }
        if (shell.isSummaryAiDispatchExportLocked()) {
            showInfo("手動退避", "サマリ xlsx を作成中のため退避できません。");
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path current = AppPaths.summaryAiDispatchXlsxPath(ui);
        if (!Files.isRegularFile(current)) {
            showInfo("手動退避", "現行サマリ Excel が見つかりません:\n" + current);
            return;
        }
        String label =
                manualArchiveLabelField != null && manualArchiveLabelField.getText() != null
                        ? manualArchiveLabelField.getText().strip()
                        : "";
        try {
            Optional<SummaryAiDispatchGenerationEntry> created =
                    SummaryAiDispatchGenerationStore.archiveCurrent(ui, label);
            if (created.isEmpty()) {
                showInfo("手動退避", "退避対象がありません。");
                return;
            }
            shell.appendLog(
                    "[summary-generation] 手動退避: "
                            + created.get().displayLabel()
                            + " ("
                            + created.get().id()
                            + ")");
            if (manualArchiveLabelField != null) {
                manualArchiveLabelField.clear();
            }
            refreshList();
        } catch (Exception ex) {
            showError("手動退避", "退避に失敗しました", ex);
        }
    }

    @FXML
    private void onOpenGenerationAction() {
        if (shell == null) {
            return;
        }
        SummaryAiDispatchGenerationEntry sel = selectedEntry();
        if (sel == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path workbook = sel.resolveWorkbookPathForOperator(ui);
        if (!Files.isRegularFile(workbook)) {
            showInfo("開く", "退避ファイルが見つかりません:\n" + workbook);
            return;
        }
        try {
            DesktopFileOpener.openFileReadOnly(workbook);
            shell.appendLog("[summary-generation] opened (read-only): " + workbook);
        } catch (Exception ex) {
            showError("開く", "ファイルを開けませんでした", ex);
        }
    }

    @FXML
    private void onRestoreGenerationAction() {
        if (shell == null) {
            return;
        }
        SummaryAiDispatchGenerationEntry sel = selectedEntry();
        if (sel == null) {
            return;
        }
        if (shell.isSummaryAiDispatchExportLocked()) {
            showInfo("復元", "サマリ xlsx を作成中のため復元できません。");
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("復元の確認");
        confirm.setHeaderText(null);
        confirm.setContentText(
                "選択した世代（"
                        + sel.displayLabel()
                        + "）で現行サマリ Excel を上書きします。\n"
                        + "復元前に現行ブックは自動退避されます。\n\n続行しますか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            Map<String, String> ui = shell.snapshotUiEnv();
            SummaryAiDispatchGenerationStore.restoreToCurrentWorkbook(sel, ui);
            shell.appendLog(
                    "[summary-generation] 復元しました: "
                            + AppPaths.summaryAiDispatchXlsxPath(ui)
                            + " ← "
                            + sel.displayLabel());
            refreshList();
        } catch (Exception ex) {
            showError("復元", "復元に失敗しました", ex);
        }
    }

    @FXML
    private void onRenameGenerationAction() {
        if (shell == null) {
            return;
        }
        SummaryAiDispatchGenerationEntry sel = selectedEntry();
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
            SummaryAiDispatchGenerationStore.updateEntryLabel(sel, shell.snapshotUiEnv(), r.get());
            refreshList();
        } catch (Exception ex) {
            showError("ラベル変更", ex.getMessage(), ex);
        }
    }

    @FXML
    private void onDeleteGenerationAction() {
        if (shell == null) {
            return;
        }
        SummaryAiDispatchGenerationEntry sel = selectedEntry();
        if (sel == null) {
            return;
        }
        if (!SummaryAiDispatchGenerationStore.isCreatedByCurrentUser(sel, shell.snapshotUiEnv())) {
            showInfo("削除", "自分が作成した退避のみ削除できます。");
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("削除の確認");
        confirm.setHeaderText(null);
        confirm.setContentText("自分が作成したこの世代を削除しますか？");
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        try {
            SummaryAiDispatchGenerationStore.deleteEntry(sel, shell.snapshotUiEnv());
            refreshList();
        } catch (Exception ex) {
            showError("削除", ex.getMessage(), ex);
        }
    }

    private SummaryAiDispatchGenerationEntry selectedEntry() {
        return generationListView != null
                ? generationListView.getSelectionModel().getSelectedItem()
                : null;
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
