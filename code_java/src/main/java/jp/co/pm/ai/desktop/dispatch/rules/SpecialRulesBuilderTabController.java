package jp.co.pm.ai.desktop.dispatch.rules;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.ObjectNode;

import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.CheckBoxTableCell;
import javafx.scene.control.cell.PropertyValueFactory;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.dispatch.rules.execution.DispatchRuleExecutionPlanner;
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

    private static final List<String> EXECUTION_MODES = List.of("auto", "dsl", "legacy");

    private MainShellController shell;
    private Runnable afterSave;
    private DispatchRuleDocument document = new DispatchRuleDocument();
    private Path workPath;
    private boolean suppressDirty;

    @FXML private Label schemaBadge;
    @FXML private Label runStatusBanner;
    @FXML private TextField pathField;
    @FXML private ComboBox<DispatchRuleEntry> ruleCombo;
    @FXML private ComboBox<String> executionModeCombo;
    @FXML private DispatchRuleGraphEditorPane graphEditor;
    @FXML private TextArea inspectorArea;
    @FXML private TextArea conflictArea;
    @FXML private ListView<String> historyList;
    @FXML private Label applyOrderSummary;
    @FXML private CheckBox ruleEnabledCheck;
    @FXML private TableView<RuleRow> ruleTable;
    @FXML private TableColumn<RuleRow, Number> orderCol;
    @FXML private TableColumn<RuleRow, Boolean> enabledCol;
    @FXML private TableColumn<RuleRow, String> idCol;
    @FXML private TableColumn<RuleRow, String> nameCol;
    @FXML private TableColumn<RuleRow, String> modeCol;
    @FXML private TableColumn<RuleRow, Number> applyOrderCol;

    private final ObservableList<DispatchRuleEntry> ruleItems = FXCollections.observableArrayList();
    private final ObservableList<RuleRow> ruleRows = FXCollections.observableArrayList();

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
        if (executionModeCombo != null) {
            executionModeCombo.getItems().setAll(EXECUTION_MODES);
            executionModeCombo.valueProperty().addListener((o, a, b) -> onExecutionModeChanged());
        }
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
        setupRuleTable();
    }

    private void setupRuleTable() {
        if (ruleTable == null) {
            return;
        }
        ruleTable.setItems(ruleRows);
        orderCol.setCellValueFactory(new PropertyValueFactory<>("displayOrder"));
        enabledCol.setCellValueFactory(new PropertyValueFactory<>("enabled"));
        enabledCol.setCellFactory(CheckBoxTableCell.forTableColumn(enabledCol));
        idCol.setCellValueFactory(new PropertyValueFactory<>("id"));
        nameCol.setCellValueFactory(new PropertyValueFactory<>("name"));
        modeCol.setCellValueFactory(new PropertyValueFactory<>("executionMode"));
        applyOrderCol.setCellValueFactory(new PropertyValueFactory<>("applyOrder"));
        modeCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final ComboBox<String> combo = new ComboBox<>(FXCollections.observableArrayList(EXECUTION_MODES));

                            {
                                combo.valueProperty()
                                        .addListener(
                                                (o, a, b) -> {
                                                    RuleRow row = getTableRow() != null ? getTableRow().getItem() : null;
                                                    if (row != null && b != null && row.entry != null) {
                                                        row.entry.executionMode = b;
                                                        row.executionMode.set(b);
                                                        markDirty();
                                                    }
                                                });
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || getTableRow() == null || getTableRow().getItem() == null) {
                                    setGraphic(null);
                                    return;
                                }
                                RuleRow row = getTableRow().getItem();
                                combo.setValue(row.entry.executionMode);
                                setGraphic(combo);
                            }
                        });
        ruleTable.getSelectionModel().selectedItemProperty().addListener((o, a, b) -> selectRuleEntry(b));
        enabledCol.setEditable(true);
        ruleTable.setEditable(true);
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        reloadFromDisk(false);
    }

    public void setAfterSave(Runnable afterSave) {
        this.afterSave = afterSave;
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
            syncTableToDocument();
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
            if (afterSave != null) {
                afterSave.run();
            }
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
            refreshRuleTable();
        }
    }

    @FXML
    private void onMoveRuleUpAction() {
        moveSelectedRule(-1);
    }

    @FXML
    private void onMoveRuleDownAction() {
        moveSelectedRule(1);
    }

    @FXML
    private void onEnableAllRulesAction() {
        document.rules.forEach(r -> r.enabled = true);
        markDirty();
        refreshRuleTable();
    }

    @FXML
    private void onDisableAllRulesAction() {
        document.rules.forEach(r -> r.enabled = false);
        markDirty();
        refreshRuleTable();
    }

    private void onExecutionModeChanged() {
        DispatchRuleEntry sel = ruleCombo.getSelectionModel().getSelectedItem();
        if (sel != null && executionModeCombo != null && executionModeCombo.getValue() != null) {
            sel.executionMode = executionModeCombo.getValue();
            markDirty();
            refreshRuleTable();
        }
    }

    private void moveSelectedRule(int delta) {
        RuleRow row = ruleTable != null ? ruleTable.getSelectionModel().getSelectedItem() : null;
        if (row == null) {
            return;
        }
        List<DispatchRuleEntry> sorted = sortedRulesMutable();
        int idx = sorted.indexOf(row.entry);
        if (idx < 0) {
            return;
        }
        int next = idx + delta;
        if (next < 0 || next >= sorted.size()) {
            return;
        }
        DispatchRuleEntry swap = sorted.get(next);
        int tmp = row.entry.applyOrder;
        row.entry.applyOrder = swap.applyOrder;
        swap.applyOrder = tmp;
        markDirty();
        refreshRuleTable();
        ruleTable.getSelectionModel().select(row);
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
            refreshRuleTable();
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
            if (executionModeCombo != null) {
                executionModeCombo.setValue("auto");
            }
            return;
        }
        graphEditor.setGraph(sel.graph);
        if (ruleEnabledCheck != null) {
            ruleEnabledCheck.setSelected(sel.enabled);
        }
        if (executionModeCombo != null) {
            executionModeCombo.setValue(
                    sel.executionMode != null && !sel.executionMode.isBlank() ? sel.executionMode : "auto");
        }
        boolean engineOn =
                "1".equals(Optional.ofNullable(shell).map(s -> s.dispatchRulesUiEnv().get("PM_AI_DISPATCH_RULE_ENGINE")).orElse("0"));
        var planned = DispatchRuleExecutionPlanner.resolveSource(sel, engineOn);
        inspectorArea.setText(
                "id: "
                        + sel.id
                        + "\nname: "
                        + sel.name
                        + "\napplyOrder: "
                        + sel.applyOrder
                        + "\nexecutionMode: "
                        + sel.executionMode
                        + "\nplanned: "
                        + planned
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

    private void selectRuleEntry(RuleRow row) {
        if (row == null || row.entry == null) {
            return;
        }
        ruleCombo.getSelectionModel().select(row.entry);
    }

    private void syncSelectedRuleFromUi() {
        DispatchRuleEntry sel = ruleCombo.getSelectionModel().getSelectedItem();
        if (sel != null && ruleEnabledCheck != null) {
            sel.enabled = ruleEnabledCheck.isSelected();
        }
    }

    private void syncTableToDocument() {
        for (RuleRow row : ruleRows) {
            if (row.entry != null) {
                row.entry.enabled = row.enabled.get();
                row.entry.executionMode = row.executionMode.get();
            }
        }
    }

    private void refreshRuleTable() {
        if (ruleTable == null) {
            return;
        }
        ruleRows.clear();
        List<DispatchRuleEntry> sorted = sortedRulesMutable();
        int display = 1;
        for (DispatchRuleEntry r : sorted) {
            RuleRow row = new RuleRow(r);
            if (r.enabled) {
                row.displayOrder.set(display++);
            } else {
                row.displayOrder.set(0);
            }
            ruleRows.add(row);
        }
        refreshApplyOrderSummary();
    }

    private List<DispatchRuleEntry> sortedRulesMutable() {
        List<DispatchRuleEntry> sorted = new ArrayList<>(document.rules);
        sorted.sort(Comparator.comparingInt(r -> r.applyOrder));
        return sorted;
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
                document.rules.stream()
                        .filter(r -> r.enabled)
                        .sorted(Comparator.comparingInt(r -> r.applyOrder))
                        .map(r -> r.id)
                        .toList();
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

    public static final class RuleRow {
        final DispatchRuleEntry entry;
        final SimpleIntegerProperty displayOrder = new SimpleIntegerProperty();
        final SimpleBooleanProperty enabled;
        final SimpleStringProperty id;
        final SimpleStringProperty name;
        final SimpleStringProperty executionMode;
        final SimpleIntegerProperty applyOrder;

        RuleRow(DispatchRuleEntry entry) {
            this.entry = entry;
            this.enabled = new SimpleBooleanProperty(entry.enabled);
            this.id = new SimpleStringProperty(entry.id);
            this.name = new SimpleStringProperty(entry.name);
            this.executionMode =
                    new SimpleStringProperty(
                            entry.executionMode != null && !entry.executionMode.isBlank()
                                    ? entry.executionMode
                                    : "auto");
            this.applyOrder = new SimpleIntegerProperty(entry.applyOrder);
            this.enabled.addListener(
                    (o, a, b) -> {
                        entry.enabled = b;
                    });
        }

        public int getDisplayOrder() {
            return displayOrder.get();
        }

        public SimpleIntegerProperty displayOrderProperty() {
            return displayOrder;
        }

        public boolean isEnabled() {
            return enabled.get();
        }

        public SimpleBooleanProperty enabledProperty() {
            return enabled;
        }

        public String getId() {
            return id.get();
        }

        public String getName() {
            return name.get();
        }

        public String getExecutionMode() {
            return executionMode.get();
        }

        public int getApplyOrder() {
            return applyOrder.get();
        }
    }
}
