package jp.co.pm.ai.desktop.dispatch.rules;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.CheckBoxTableCell;
import javafx.scene.control.cell.ComboBoxTableCell;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.dispatch.rules.migration.DispatchRuleMigrationService;
import jp.co.pm.ai.desktop.dispatch.rules.model.ProcessMachinePriorityEntry;
import jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths;

/**
 * 工程名+機械名の配台優先。同一実機械上の連続選好。グローバル試行順は書き換えない。
 */
public final class ProcessMachinePriorityTabController {

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private static final ObservableList<String> PRIORITY_CHOICES =
            FXCollections.observableArrayList(
                    ProcessMachinePriorityEntry.PRIORITY_HIGHEST,
                    ProcessMachinePriorityEntry.PRIORITY_HIGH,
                    ProcessMachinePriorityEntry.PRIORITY_NORMAL,
                    ProcessMachinePriorityEntry.PRIORITY_LOW);

    private MainShellController shell;
    private final ObservableList<Row> rows = FXCollections.observableArrayList();
    private Runnable afterSave;

    @FXML private VBox root;
    private TableView<Row> table;
    private Label statusLabel;

    @FXML
    private void initialize() {
        if (root == null) {
            return;
        }
        table = new TableView<>();
        table.setEditable(true);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
        Label hint =
                new Label(
                        "工程名+機械名ごとに優先度を選びます。機械名が空ならその工程の全機械。"
                                + "既定はエンボス・通常・連続。同一実機械上で連続配台し、配台試行順番の連番塊は作りません。");
        hint.setWrapText(true);
        Button add = new Button("行を追加");
        add.setOnAction(e -> onAddRow());
        Button del = new Button("行を削除");
        del.setOnAction(e -> onRemoveRow());
        Button reload = new Button("再読込");
        reload.setOnAction(e -> reloadFromDisk());
        Button save = new Button("保存");
        save.setOnAction(e -> saveToDisk());
        HBox bar = new HBox(8, add, del, reload, save);
        statusLabel = new Label();
        statusLabel.setWrapText(true);
        root.setSpacing(8);
        root.setPadding(new Insets(12));
        VBox.setVgrow(table, Priority.ALWAYS);

        TableColumn<Row, String> processCol = new TableColumn<>("工程名");
        processCol.setCellValueFactory(c -> c.getValue().processName);
        processCol.setCellFactory(TextFieldTableCell.forTableColumn());
        processCol.setOnEditCommit(ev -> ev.getRowValue().processName.set(ev.getNewValue()));
        TableColumn<Row, String> machineCol = new TableColumn<>("機械名");
        machineCol.setCellValueFactory(c -> c.getValue().machineName);
        machineCol.setCellFactory(TextFieldTableCell.forTableColumn());
        machineCol.setOnEditCommit(ev -> ev.getRowValue().machineName.set(ev.getNewValue()));
        TableColumn<Row, String> prioCol = new TableColumn<>("優先度");
        prioCol.setCellValueFactory(c -> c.getValue().priority);
        prioCol.setCellFactory(ComboBoxTableCell.forTableColumn(PRIORITY_CHOICES));
        prioCol.setOnEditCommit(ev -> ev.getRowValue().priority.set(ev.getNewValue()));
        TableColumn<Row, Boolean> consCol = new TableColumn<>("連続配置");
        consCol.setCellValueFactory(c -> c.getValue().consecutive);
        consCol.setCellFactory(CheckBoxTableCell.forTableColumn(consCol));
        TableColumn<Row, Boolean> enCol = new TableColumn<>("有効");
        enCol.setCellValueFactory(c -> c.getValue().enabled);
        enCol.setCellFactory(CheckBoxTableCell.forTableColumn(enCol));
        table.getColumns().setAll(List.of(processCol, machineCol, prioCol, consCol, enCol));
        table.setItems(rows);
        table.setPlaceholder(new Label("行がありません"));
        root.getChildren().setAll(hint, bar, table, statusLabel);
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        reloadFromDisk();
    }

    public void setAfterSave(Runnable afterSave) {
        this.afterSave = afterSave;
    }

    private void onAddRow() {
        rows.add(Row.fromEntry(ProcessMachinePriorityEntry.defaultEmboss()));
    }

    private void onRemoveRow() {
        if (table == null) {
            return;
        }
        Row sel = table.getSelectionModel().getSelectedItem();
        if (sel != null) {
            rows.remove(sel);
        }
    }

    private void reloadFromDisk() {
        if (shell == null) {
            return;
        }
        try {
            DispatchRulePaths.ensureWorkJsonFromRepoIfMissing(shell.dispatchRulesUiEnv());
            Path work = DispatchRulePaths.resolveWorkJson(shell.dispatchRulesUiEnv());
            rows.setAll(loadRows(work));
            if (statusLabel != null) {
                statusLabel.setText(work.toString());
            }
        } catch (IOException ex) {
            if (statusLabel != null) {
                statusLabel.setText("読込失敗: " + ex.getMessage());
            }
        }
    }

    private void saveToDisk() {
        if (shell == null) {
            return;
        }
        try {
            Path work = DispatchRulePaths.resolveWorkJson(shell.dispatchRulesUiEnv());
            Files.createDirectories(work.getParent());
            ObjectNode rootNode;
            if (Files.isRegularFile(work)) {
                JsonNode parsed = JSON.readTree(Files.readString(work, StandardCharsets.UTF_8));
                rootNode = parsed != null && parsed.isObject() ? (ObjectNode) parsed : JSON.createObjectNode();
            } else {
                rootNode = JSON.createObjectNode();
                rootNode.put("schemaVersion", DispatchRuleMigrationService.CURRENT_SCHEMA_VERSION);
            }
            ArrayNode arr = JSON.createArrayNode();
            for (Row row : rows) {
                ProcessMachinePriorityEntry entry = row.toEntry();
                if (entry.processName == null || entry.processName.isBlank()) {
                    continue;
                }
                arr.add(JSON.valueToTree(entry));
            }
            rootNode.set("processMachinePriorities", arr);
            Files.writeString(
                    work,
                    JSON.writerWithDefaultPrettyPrinter().writeValueAsString(rootNode),
                    StandardCharsets.UTF_8);
            if (statusLabel != null) {
                statusLabel.setText("保存しました: " + work);
            }
            if (afterSave != null) {
                afterSave.run();
            }
        } catch (IOException ex) {
            shell.showErrorDialog("保存エラー", ex.getMessage());
        }
    }

    private static List<Row> loadRows(Path work) throws IOException {
        List<Row> out = new ArrayList<>();
        if (!Files.isRegularFile(work)) {
            out.add(Row.fromEntry(ProcessMachinePriorityEntry.defaultEmboss()));
            return out;
        }
        JsonNode rootNode = JSON.readTree(Files.readString(work, StandardCharsets.UTF_8));
        JsonNode arr = rootNode.get("processMachinePriorities");
        if (arr == null || !arr.isArray()) {
            out.add(Row.fromEntry(ProcessMachinePriorityEntry.defaultEmboss()));
            return out;
        }
        for (JsonNode n : arr) {
            ProcessMachinePriorityEntry e = JSON.treeToValue(n, ProcessMachinePriorityEntry.class);
            if (e != null) {
                out.add(Row.fromEntry(e));
            }
        }
        return out;
    }

    public static final class Row {
        final SimpleStringProperty processName = new SimpleStringProperty("");
        final SimpleStringProperty machineName = new SimpleStringProperty("");
        final SimpleStringProperty priority =
                new SimpleStringProperty(ProcessMachinePriorityEntry.PRIORITY_NORMAL);
        final SimpleBooleanProperty consecutive = new SimpleBooleanProperty(true);
        final SimpleBooleanProperty enabled = new SimpleBooleanProperty(true);

        static Row fromEntry(ProcessMachinePriorityEntry e) {
            Row r = new Row();
            r.processName.set(e.processName == null ? "" : e.processName);
            r.machineName.set(e.machineName == null ? "" : e.machineName);
            r.priority.set(
                    e.priority == null || e.priority.isBlank()
                            ? ProcessMachinePriorityEntry.PRIORITY_NORMAL
                            : e.priority);
            r.consecutive.set(e.consecutive);
            r.enabled.set(e.enabled);
            return r;
        }

        ProcessMachinePriorityEntry toEntry() {
            ProcessMachinePriorityEntry e = new ProcessMachinePriorityEntry();
            e.processName = processName.get() == null ? "" : processName.get().strip();
            e.machineName = machineName.get() == null ? "" : machineName.get().strip();
            e.priority =
                    priority.get() == null || priority.get().isBlank()
                            ? ProcessMachinePriorityEntry.PRIORITY_NORMAL
                            : priority.get();
            e.consecutive = consecutive.get();
            e.enabled = enabled.get();
            return e;
        }
    }
}
