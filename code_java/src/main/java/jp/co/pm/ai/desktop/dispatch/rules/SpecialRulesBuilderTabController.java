package jp.co.pm.ai.desktop.dispatch.rules;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.ObjectNode;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.CheckBox;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.dispatch.rules.history.DispatchRuleHistoryStore;
import jp.co.pm.ai.desktop.dispatch.rules.migration.DispatchRuleMigrationService;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleDocument;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleEntry;
import jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths;
import jp.co.pm.ai.desktop.dispatch.rules.stage.DispatchRuleBuilderRunContext;
import jp.co.pm.ai.desktop.dispatch.rules.ui.editor.DispatchRuleGraphEditorPane;
import jp.co.pm.ai.desktop.dispatch.rules.validation.DispatchRuleConflictChecker;

/** Rule builder child tab. */
public final class SpecialRulesBuilderTabController {

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private MainShellController shell;
    private DispatchRuleDocument document = new DispatchRuleDocument();
    private Path workPath;
    private boolean suppressDirty;

    @FXML private Label schemaBadge;
    @FXML private Label runStatusBanner;
    @FXML private TextField pathField;
    @FXML private ComboBox<DispatchRuleEntry> ruleCombo;
    @FXML private DispatchRuleGraphEditorPane graphEditor;
    @FXML private TextArea inspectorArea;
    @FXML private TextArea conflictArea;
    @FXML private ListView<String> historyList;
    @FXML private Label applyOrderSummary;
    @FXML private CheckBox ruleEnabledCheck;

    private final ObservableList<DispatchRuleEntry> ruleItems = FXCollections.observableArrayList();

    @FXML
    private void initialize() {
        ruleCombo.setItems(ruleItems);
        ruleCombo.setCellFactory(cb -> new ListCell<>() {
            @Override
            protected void updateItem(DispatchRuleEntry item, boolean empty) {
                super.updateItem(item, empty);
                setText(empty || item == null ? null : item.id + " " + item.name);
            }
        });
        ruleCombo.setButtonCell(ruleCombo.getCellFactory().call(null));
        ruleCombo.getSelectionModel().selectedItemProperty().addListener((o, a, b) -> showSelectedRule());
        DispatchRuleBuilderRunContext.get().setBannerConsumer(text -> {
            if (runStatusBanner != null) {
                runStatusBanner.setText(text);
            }
        });
        if (historyList != null) {
            historyList.setOnMouseClicked(
                    e -> {
                        if (e.getClickCount() >= 2) {
                            onRestoreHistory();
                        }
                    });
        }
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        reloadFromDisk(false);
    }

    DispatchRuleDocument snapshotDocument() {
        return document;
    }

    @FXML
    private void onReloadAction() {
        reloadFromDisk(true);
    }

    @FXML
    private void onSaveAction() {
        if (shell == null || workPath == null) {
            return;
        }
        try {
            document.schemaVersion = DispatchRuleMigrationService.CURRENT_SCHEMA_VERSION;
            document.savedAt = Instant.now().toString();
            syncSelectedRuleFromUi();
            Files.createDirectories(workPath.getParent());
            JSON.writeValue(workPath.toFile(), document);
            DispatchRuleHistoryStore.appendAutoSave(shell.dispatchRulesUiEnv(), workPath);
            DispatchRuleBuilderRunContext.get().setDirty(false);
            refreshHistory();
            shell.dispatchRulesAppendLog("[dispatch-rules] saved: " + workPath);
            refreshConflicts();
        } catch (IOException ex) {
            shell.showErrorDialog("保存エラー", ex.getMessage());
        }
    }

    @FXML
    private void onValidateAction() {
        refreshConflicts();
        shell.showInformationDialog("検証", conflictArea.getText());
    }

    @FXML
    private void onConflictAction() {
        refreshConflicts();
    }

    @FXML
    private void onRestoreHistory() {
        if (shell == null || historyList == null || workPath == null) {
            return;
        }
        String selected = historyList.getSelectionModel().getSelectedItem();
        if (selected == null || selected.isBlank()) {
            return;
        }
        String id = selected.split(" ", 2)[0];
        try {
            DispatchRuleHistoryStore.restore(shell.dispatchRulesUiEnv(), workPath, id);
            reloadFromDisk(false);
            shell.showInformationDialog("復元", "履歴から復元しました（次回実行から反映）");
        } catch (IOException ex) {
            shell.showErrorDialog("復元エラー", ex.getMessage());
        }
    }

    @FXML
    private void onToggleEnabledAction() {
        DispatchRuleEntry sel = ruleCombo.getSelectionModel().getSelectedItem();
        if (sel != null && ruleEnabledCheck != null) {
            sel.enabled = ruleEnabledCheck.isSelected();
            markDirty();
            refreshApplyOrderSummary();
        }
    }

    private void reloadFromDisk(boolean dialog) {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.dispatchRulesUiEnv();
        DispatchRulePaths.ensureWorkJsonFromRepoIfMissing(ui);
        workPath = DispatchRulePaths.resolveWorkJson(ui);
        pathField.setText(workPath.toString());
        try {
            if (!Files.isRegularFile(workPath)) {
                document = new DispatchRuleDocument();
            } else {
                ObjectNode raw = (ObjectNode) JSON.readTree(Files.readString(workPath, StandardCharsets.UTF_8));
                ObjectNode migrated = DispatchRuleMigrationService.migrate(raw);
                document = JSON.treeToValue(migrated, DispatchRuleDocument.class);
            }
            suppressDirty = true;
            ruleItems.setAll(document.rules);
            if (!ruleItems.isEmpty()) {
                ruleCombo.getSelectionModel().select(0);
            }
            schemaBadge.setText("schema v" + document.schemaVersion);
            refreshApplyOrderSummary();
            refreshConflicts();
            refreshHistory();
            DispatchRuleBuilderRunContext.get().setDirty(false);
            suppressDirty = false;
            if (dialog) {
                shell.showInformationDialog("再読込", "特別ルール JSON を読み込みました。");
            }
        } catch (IOException ex) {
            shell.showErrorDialog("読込エラー", ex.getMessage());
        }
    }

    private void showSelectedRule() {
        DispatchRuleEntry sel = ruleCombo.getSelectionModel().getSelectedItem();
        if (sel == null) {
            graphEditor.setGraph(null);
            inspectorArea.setText("");
            if (ruleEnabledCheck != null) {
                ruleEnabledCheck.setSelected(false);
            }
            return;
        }
        graphEditor.setGraph(sel.graph);
        if (ruleEnabledCheck != null) {
            ruleEnabledCheck.setSelected(sel.enabled);
        }
        inspectorArea.setText(
                "id: "
                        + sel.id
                        + "\nname: "
                        + sel.name
                        + "\napplyOrder: "
                        + sel.applyOrder
                        + "\nexecutionMode: "
                        + sel.executionMode
                        + "\nenabled: "
                        + sel.enabled);
        graphEditor.setOnNodeSelected(
                nodeId -> {
                    sel.graph.nodes.stream()
                            .filter(n -> nodeId.equals(n.id))
                            .findFirst()
                            .ifPresent(
                                    n ->
                                            inspectorArea.appendText(
                                                    "\n\nnode: "
                                                            + n.id
                                                            + " type="
                                                            + n.type
                                                            + "\nparams="
                                                            + n.params));
                    graphEditor.setHighlightedNodeId(nodeId);
                });
    }

    private void syncSelectedRuleFromUi() {
        DispatchRuleEntry sel = ruleCombo.getSelectionModel().getSelectedItem();
        if (sel != null && ruleEnabledCheck != null) {
            sel.enabled = ruleEnabledCheck.isSelected();
        }
    }

    private void refreshConflicts() {
        var report = DispatchRuleConflictChecker.check(document);
        StringBuilder sb = new StringBuilder();
        sb.append("errors=").append(report.errorCount()).append(" warnings=").append(report.warningCount()).append('\n');
        report.conflicts().forEach(c -> sb.append(c.severity()).append(' ').append(c.message()).append('\n'));
        conflictArea.setText(sb.toString());
    }

    private void refreshApplyOrderSummary() {
        List<String> enabled =
                document.rules.stream().filter(r -> r.enabled).sorted((a, b) -> Integer.compare(a.applyOrder, b.applyOrder)).map(r -> r.id).toList();
        applyOrderSummary.setText("有効 " + enabled.size() + " 件: " + String.join(" → ", enabled));
    }

    private void refreshHistory() {
        if (historyList == null || shell == null) {
            return;
        }
        try {
            List<DispatchRuleHistoryStore.HistoryEntry> entries =
                    DispatchRuleHistoryStore.listEntries(shell.dispatchRulesUiEnv());
            historyList.getItems().setAll(
                    entries.stream().map(e -> e.id() + " " + e.label() + " " + e.summary()).toList());
        } catch (IOException ex) {
            historyList.getItems().clear();
        }
    }

    private void markDirty() {
        if (!suppressDirty) {
            DispatchRuleBuilderRunContext.get().setDirty(true);
        }
        refreshApplyOrderSummary();
        refreshConflicts();
    }
}
