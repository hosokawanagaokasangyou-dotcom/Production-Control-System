package jp.co.pm.ai.desktop;

import java.io.File;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.TabularCellHighlight;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;

/**
 * \u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b tab; layout {@code PlanInputTab.fxml}.
 */
public final class PlanInputTabController {

    public static final String ENV_PM_AI_PLAN_INPUT_PATH = "PM_AI_PLAN_INPUT_PATH";
    public static final String ENV_TASK_PLAN_SHEET = "TASK_PLAN_SHEET";

    public static final String DEFAULT_PLAN_INPUT_SHEET_NAME =
            "\u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b";

    private static final String HINT_TEXT =
            "PM_AI_PLAN_INPUT_PATH \u3068\u540c\u3058\u30d5\u30a1\u30a4\u30eb\u3092\u7de8\u96c6\u3057\u307e\u3059"
                    + " (\u6bb5\u968e2 load_planning_tasks_df)\u3002Excel \u306f"
                    + " \u30b7\u30fc\u30c8\u540d\u3092\u6307\u5b9a\u3002\u4fdd\u5b58\u6642"
                    + " .xlsx \u306f\u30c7\u30fc\u30bf\u306e\u307f\uff08\u30de\u30af\u30ed\u306f\u524a\u9664\u3055\u308c\u307e\u3059\uff09\u3002";

    private Stage ownerStage;

    private MainShellController shell;

    @FXML
    private TextField pathField;

    @FXML
    private TextField sheetField;

    @FXML
    private Button browseButton;

    @FXML
    private Button loadButton;

    @FXML
    private Button saveButton;

    @FXML
    private Button addRowButton;

    @FXML
    private Button removeRowsButton;

    @FXML
    private HBox columnStripHost;

    @FXML
    private TextField colWidthField;

    @FXML
    private Label hintLabel;

    @FXML
    private TableView<ObservableList<String>> table;

    private final List<String> headersRef = new ArrayList<>();
    private ObservableList<ObservableList<String>> rows;
    private final AtomicBoolean suppressColumnOrderPersistence = new AtomicBoolean(false);
    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());

    @FXML
    private void initialize() {
        pathField.setPromptText("PM_AI_PLAN_INPUT_PATH ? .csv / .xlsx / .xlsm");
        sheetField.setText(DEFAULT_PLAN_INPUT_SHEET_NAME);
        sheetField.setPromptText("Excel sheet name (TASK_PLAN_SHEET / TASK_PLAN_SHEET)");
        colWidthField.setText("112");
        hintLabel.setText(HINT_TEXT);

        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.setEditable(true);
        table.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        rows = FXCollections.observableArrayList();
        table.setItems(rows);
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        ownerStage = shell.getPrimaryStage();

        columnStripHost.getChildren().setAll(TableViewColumnSettingsStrip.create(table, this::applyDynamicColumnWidths, false));

        shell.acceptReloadAfterStage1PlanInput(
                () -> {
                    Map<String, String> env = shell.snapshotUiEnv();
                    if (env != null) {
                        pathField.setText(AppPaths.defaultStage1PlanTasksPath(env).toString());
                    }
                    sheetField.setText(AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
                    loadFromCurrentPath();
                });

        TableColumnOrderPersistence.installColumnLayoutWatcher(
                table,
                TableColumnOrderPersistence.TableId.PLAN_INPUT,
                suppressColumnOrderPersistence::get);

        javafx.application.Platform.runLater(
                () -> {
                    syncFromEnv();
                    if (!pathField.getText().isBlank()) {
                        loadFromCurrentPath();
                    }
                });
    }

    @FXML
    private void onBrowseButtonAction() {
        FileChooser ch = new FileChooser();
        ch.setTitle("\u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b ? \u30d5\u30a1\u30a4\u30eb");
        ch.getExtensionFilters()
                .addAll(
                        new FileChooser.ExtensionFilter("Tabular", "*.csv", "*.xlsx", "*.xlsm"),
                        new FileChooser.ExtensionFilter("All", "*.*"));
        File f = ch.showOpenDialog(ownerStage);
        if (f != null) {
            pathField.setText(f.getAbsolutePath());
        }
    }

    @FXML
    private void onLoadButtonAction() {
        syncFromEnv();
        loadFromCurrentPath();
    }

    @FXML
    private void onSaveButtonAction() {
        Path path = Path.of(pathField.getText().trim());
        if (pathField.getText().isBlank()) {
            shell.appendLog("[plan-input] save: path is empty");
            return;
        }
        try {
            List<List<String>> dataRows = new ArrayList<>();
            for (ObservableList<String> r : rows) {
                List<String> copy = new ArrayList<>(r);
                while (copy.size() < headersRef.size()) {
                    copy.add("");
                }
                while (copy.size() > headersRef.size()) {
                    copy.remove(copy.size() - 1);
                }
                dataRows.add(copy);
            }
            PlanInputTabularIo.write(
                    path,
                    sheetField.getText().trim().isEmpty()
                            ? DEFAULT_PLAN_INPUT_SHEET_NAME
                            : sheetField.getText().trim(),
                    new PlanInputTabularIo.TabularSheet(headersRef, dataRows));
            shell.appendLog("[plan-input] saved " + path);
        } catch (Exception ex) {
            shell.appendLog("[plan-input] save error: " + ex.getMessage());
        }
    }

    @FXML
    private void onAddRowButtonAction() {
        if (headersRef.isEmpty()) {
            shell.appendLog("[plan-input] load a file first (headers required)");
            return;
        }
        ObservableList<String> r = FXCollections.observableArrayList();
        for (int i = 0; i < headersRef.size(); i++) {
            r.add("");
        }
        rows.add(r);
    }

    @FXML
    private void onRemoveRowsButtonAction() {
        var sel = table.getSelectionModel().getSelectedItems();
        if (sel.isEmpty()) {
            return;
        }
        rows.removeAll(sel);
        shell.appendLog("[plan-input] removed " + sel.size() + " row(s)");
    }

    private void applyDynamicColumnWidths() {
        double w = 112;
        try {
            w = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
        } catch (NumberFormatException ignored) {
        }
        for (TableColumn<ObservableList<String>, ?> c : table.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private void rebuildColumns() {
        suppressColumnOrderPersistence.set(true);
        try {
            double colW = 112;
            try {
                colW = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
            } catch (NumberFormatException ignored) {
            }
            List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            headersRef, persistedLayout.get(), colW);
            table.getColumns().clear();
            for (int i = 0; i < headersRef.size(); i++) {
                final int idx = i;
                double prefW = i < widths.size() ? widths.get(i) : colW;
                String title =
                        !headersRef.get(i).isBlank()
                                ? headersRef.get(i)
                                : ("\u5217" + (i + 1));
                TableColumn<ObservableList<String>, String> col = new TableColumn<>(title);
                col.setCellValueFactory(
                        cd -> {
                            ObservableList<String> row = cd.getValue();
                            String v =
                                    row != null && idx < row.size()
                                            ? row.get(idx)
                                            : "";
                            return new javafx.beans.property.SimpleStringProperty(v);
                        });
                col.setCellFactory(TabularCellHighlight.planInputUnprocessedHighlightCellFactory(title));
                col.setOnEditCommit(
                        ev -> {
                            ObservableList<String> row = ev.getRowValue();
                            if (row == null) {
                                return;
                            }
                            while (row.size() <= idx) {
                                row.add("");
                            }
                            row.set(
                                    idx,
                                    ev.getNewValue() != null ? ev.getNewValue() : "");
                            table.refresh();
                        });
                col.setPrefWidth(prefW);
                table.getColumns().add(col);
            }
        } finally {
            suppressColumnOrderPersistence.set(false);
        }
    }

    private void applyLoaded() {
        rebuildColumns();
        table.refresh();
    }

    private void syncFromEnv() {
        Map<String, String> env = shell.snapshotUiEnv();
        if (env != null) {
            String p = trim(env.get(ENV_PM_AI_PLAN_INPUT_PATH));
            if (!p.isEmpty() && pathField.getText().isBlank()) {
                pathField.setText(p);
            }
            String sh = trim(env.get(ENV_TASK_PLAN_SHEET));
            if (!sh.isEmpty()) {
                sheetField.setText(sh);
            }
        }
    }

    private void loadFromCurrentPath() {
        Path path = Path.of(pathField.getText().trim());
        if (!java.nio.file.Files.isRegularFile(path)) {
            shell.appendLog("[plan-input] file not found: " + path);
            return;
        }
        String sheetName = sheetField.getText().trim();
        if (sheetName.isEmpty()) {
            sheetName = DEFAULT_PLAN_INPUT_SHEET_NAME;
        }
        try {
            PlanInputTabularIo.TabularSheet sheet = PlanInputTabularIo.read(path, sheetName);
            headersRef.clear();
            headersRef.addAll(sheet.headers());
            rows.clear();
            for (List<String> line : sheet.rows()) {
                ObservableList<String> r = FXCollections.observableArrayList(line);
                while (r.size() < headersRef.size()) {
                    r.add("");
                }
                while (r.size() > headersRef.size()) {
                    r.remove(r.size() - 1);
                }
                rows.add(r);
            }
            List<TableColumnOrderPersistence.ColumnSpec> lay =
                    TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.PLAN_INPUT);
            persistedLayout.set(lay);
            TableColumnOrderPersistence.applyLogicalColumnOrder(
                    headersRef,
                    rows,
                    lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList());
            applyLoaded();
            shell.appendLog(
                    "[plan-input] loaded rows="
                            + rows.size()
                            + " cols="
                            + headersRef.size()
                            + " path="
                            + path);
        } catch (Exception ex) {
            shell.appendLog("[plan-input] load error: " + ex.getMessage());
        }
    }

    private static String trim(String s) {
        return s != null ? s.trim() : "";
    }
}
