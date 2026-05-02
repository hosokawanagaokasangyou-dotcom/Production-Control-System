package jp.co.pm.ai.desktop;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.text.Text;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.EnvVarDocs;
import jp.co.pm.ai.desktop.config.UiRefEnvDefaults;
import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;
import jp.co.pm.ai.desktop.io.ExcelSheetTitlesProbe;
import jp.co.pm.ai.desktop.ipc.IpcStdoutTap;
import jp.co.pm.ai.desktop.ui.ActualsDataStatusPane;
import jp.co.pm.ai.desktop.ui.ExcludeRulesEditorPane;
import jp.co.pm.ai.desktop.ui.FileChooserForEnvKey;
import jp.co.pm.ai.desktop.ui.PlanInputEditorPane;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.embed.swing.SwingNode;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.awt.GraphicsEnvironment;

import javax.swing.SwingUtilities;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

/**
 * MVP: ProcessBuilder \u3067\u6bb5\u968e1/2 \u8d77\u52d5\u3001\u30ed\u30b0\u8868\u793a\u3001JFreeChart \u57cb\u3081\u8fbc\u307f\u30b5\u30f3\u30d7\u30eb\u3002
 * bootstrap-javafx / mvp-python-bridge
 */
public class PmAiFxApp extends Application {

    private static final String STAGE1 = "task_extract_stage1.py";
    private static final String STAGE2 = "plan_simulation_stage2.py";
    private static final String PREFIX_CHILD = "[child] ";
    private static final String NDJSON_START = PREFIX_CHILD + "{";

    private static final List<String> BOOTSTRAP_ORDER = List.of(
            AppPaths.KEY_PM_AI_PYTHON,
            AppPaths.KEY_PM_AI_REPO_ROOT,
            AppPaths.KEY_PM_AI_CODE_PYTHON_DIR,
            AppPaths.KEY_PM_AI_WORKSPACE,
            AppPaths.KEY_GEMINI_CREDENTIALS_JSON,
            AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
            AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
            AppPaths.KEY_PM_AI_COLUMN_CONFIG_WORKBOOK,
            AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK,
            AppPaths.KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV,
            AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET,
            AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
            AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
            AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR);

    private TextArea logArea;
    private TextField workbookField;
    private TextField pythonExeField;
    private TextField scriptDirField;
    private Label statusLabel;
    private ObservableList<EnvVarRow> envRows;
    private final AtomicBoolean runLock = new AtomicBoolean(false);

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("\u5de5\u7a0b\u7ba1\u7406 AI \u914d\u53f0 \u2014 JavaFX MVP");

        envRows = FXCollections.observableArrayList();
        populateEnvRows(envRows);
        Map<String, String> ui0 = collectUiEnv();

        workbookField = new TextField();
        workbookField.setPromptText(
                "\u4efb\u610f \u2014 \u7a7a\u3067\u3088\u3044\uff08\u6bb5\u968e1\u306f PM_AI_PROCESSING_PLAN_PATH \u306a\u3069\u3067\u53ef\uff09\u3002"
                        + " \u6307\u5b9a\u6642\u306f ProcessBuilder \u304c TASK_INPUT_WORKBOOK \u3092\u4ed8\u4e0e");
        workbookField.setText(AppPaths.resolveTaskInputWorkbook(ui0).map(Path::toString).orElse(""));

        Button browseWb = new Button("\u53c2\u7167\u2026");
        browseWb.setOnAction(e -> pickWorkbook(primaryStage));
        Button detectWb = new Button("\u81ea\u52d5\u691c\u51fa");
        detectWb.setOnAction(
                e -> workbookField.setText(
                        AppPaths.resolveTaskInputWorkbook(collectUiEnv())
                                .map(Path::toString)
                                .orElse("")));

        pythonExeField = new TextField(firstNonBlank(ui0.get(AppPaths.KEY_PM_AI_PYTHON), defaultOsPython()));
        pythonExeField.setPromptText("Python executable (\u74b0\u5883\u5909\u6570 PM_AI_PYTHON)");

        scriptDirField = new TextField(
                firstNonBlank(ui0.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR), AppPaths.resolvePythonScriptDir(ui0).toString()));
        scriptDirField.setPromptText("code/python (\u74b0\u5883\u5909\u6570 PM_AI_CODE_PYTHON_DIR)");

        Button refreshDir = new Button("\u81ea\u52d5\u691c\u51fa");
        refreshDir.setOnAction(
                e -> scriptDirField.setText(AppPaths.resolvePythonScriptDir(collectUiEnv()).toString()));

        Button peekSheets = new Button("\u30b7\u30fc\u30c8\u4e00\u89a7 (POI)");
        peekSheets.setOnAction(e -> peekSheetsAction());

        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(12));
        int r = 0;
        grid.add(new Label("\u30de\u30af\u30ed\u5b9f\u884c\u30d6\u30c3\u30af\uff08\u4efb\u610f\uff09"), 0, r);
        grid.add(workbookField, 1, r);
        grid.add(new HBox(4, browseWb, detectWb), 2, r);
        r++;
        grid.add(new Label("Python"), 0, r);
        grid.add(pythonExeField, 1, r);
        r++;
        grid.add(new Label("\u30b9\u30af\u30ea\u30d7\u30c8\u30c7\u30a3\u30ec\u30af\u30c8\u30ea"), 0, r);
        grid.add(scriptDirField, 1, r);
        grid.add(refreshDir, 2, r);
        r++;
        HBox actions = new HBox(8,
                buttonStage(STAGE1, "\u6bb5\u968e1 \u5b9f\u884c"),
                buttonStage(STAGE2, "\u6bb5\u968e2 \u5b9f\u884c"),
                peekSheets);
        grid.add(actions, 1, r);

        logArea = new TextArea();
        logArea.setEditable(false);
        logArea.setWrapText(true);
        logArea.setMinHeight(160);
        logArea.setPrefRowCount(12);

        statusLabel = new Label(
                "exit: n/a \u2014 0=OK / 1=error / 2=fatal / 3=PlanningValidationError / 9=cancel");

        Label logCaption = new Label("log (stdout+stderr merged)");
        VBox topBox = new VBox(8, grid, logCaption);
        topBox.setFillWidth(true);

        BorderPane mainPane = new BorderPane();
        mainPane.setTop(topBox);
        BorderPane.setMargin(topBox, new Insets(0, 0, 8, 0));
        mainPane.setCenter(logArea);
        mainPane.setBottom(statusLabel);
        BorderPane.setMargin(statusLabel, new Insets(8, 0, 0, 0));
        mainPane.setPadding(new Insets(0, 12, 12, 12));

        Tab tabMain = new Tab("\u5b9f\u884c\u30fb\u30ed\u30b0", mainPane);
        tabMain.setClosable(false);

        Tab tabChart = new Tab("\u30b0\u30e9\u30d5 (JFreeChart)", buildChartPane());
        tabChart.setClosable(false);

        Tab tabPlanInput =
                new Tab(
                        "\u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b",
                        PlanInputEditorPane.create(
                                primaryStage, this::collectUiEnv, this::appendLog));
        tabPlanInput.setClosable(false);

        Tab tabExcludeRules =
                new Tab(
                        "\u914d\u53f0\u4e0d\u8981\u30eb\u30fc\u30eb (JSON)",
                        ExcludeRulesEditorPane.create(primaryStage, this::collectUiEnv, this::appendLog));
        tabExcludeRules.setClosable(false);

        Tab tabEnv = new Tab("\u74b0\u5883\u5909\u6570", buildEnvPane(primaryStage));
        tabEnv.setClosable(false);

        Tab tabActuals =
                new Tab(
                        "\u5b9f\u7e3eDATA\u30fb\u660e\u7d30\u306e\u53d6\u5f97\u72b6\u614b",
                        ActualsDataStatusPane.create(this::buildActualsStatusRequest, this::appendLog));
        tabActuals.setClosable(false);

        TabPane tabs =
                new TabPane(tabMain, tabEnv, tabChart, tabPlanInput, tabExcludeRules, tabActuals);
        BorderPane root = new BorderPane(tabs);
        Scene scene = new Scene(root, 960, 640);
        primaryStage.setScene(scene);
        primaryStage.setMinWidth(640);
        primaryStage.setMinHeight(480);
        primaryStage.setOnShown(e -> {
            primaryStage.toFront();
            tabs.getSelectionModel().select(tabMain);
        });
        primaryStage.show();

        appendLog("[boot] PYTHONUTF8=1 PYTHONIOENCODING=utf-8 for child process.");
    }

    private Parent buildEnvPane(Stage ownerStage) {
        Label hint = new Label(
                "OS \u74b0\u5883\u5909\u6570\u306f\u53c2\u7167\u3057\u307e\u305b\u3093\u3002\u3053\u306e\u30bf\u30d6\u3067\u96c6\u7d04\u3002"
                        + " \u521d\u671f\u5024: ui_ref_env_defaults.json + \u30ed\u30b8\u30c3\u30af\u8aac\u660e\u3002"
                        + " \u5b50\u30d7\u30ed\u30bb\u30b9: \u3053\u306e\u8868 + \u30e9\u30f3\u30c1\u30e3\u30fc\u306e TASK_INPUT_WORKBOOK"
                        + "\uff08\u30de\u30af\u30ed\u30d6\u30c3\u30af\u306f\u4efb\u610f\uff09\u2192 PYTHONUTF8 \u6700\u7d42\u56fa\u5b9a\u3002"
                        + " PM_AI_SKIP_WORKBOOK_ENV_SHEET \u304c\u7a7a\u306e\u3068\u304d\u306f 1 \u3068\u3057\u3066"
                        + "\u30de\u30af\u30ed\u300c\u8a2d\u5b9a_\u74b0\u5883\u5909\u6570\u300d\u30b7\u30fc\u30c8\u3092\u8aad\u307e\u306a\u3044\u3002"
                        + " \u30d5\u30a9\u30eb\u30c0\u578b\u306f\u300c\u30d5\u30a9\u30eb\u30c0...\u300d\u3001\u5404\u30d5\u30a1\u30a4\u30eb\u578b\u306f"
                        + "\u5909\u6570\u540d\u306b\u5fdc\u3058\u3066 JSON / Excel / CSV \u306e\u62e1\u5f35\u5b50\u3092\u8868\u793a\u3002");
        hint.setWrapText(true);

        TableView<EnvVarRow> table = new TableView<>(envRows);
        table.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        table.setEditable(true);
        table.setMinHeight(180);
        table.setPrefHeight(260);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

        TableColumn<EnvVarRow, String> nameCol = new TableColumn<>("\u5909\u6570\u540d");
        nameCol.setCellValueFactory(cdf -> cdf.getValue().nameProperty());
        nameCol.setCellFactory(TextFieldTableCell.forTableColumn());
        nameCol.setOnEditCommit(
                e -> {
                    e.getRowValue().setName(e.getNewValue());
                    table.refresh();
                });
        nameCol.setPrefWidth(220);

        TableColumn<EnvVarRow, String> valueCol = new TableColumn<>("\u5024");
        valueCol.setCellValueFactory(cdf -> cdf.getValue().valueProperty());
        valueCol.setCellFactory(TextFieldTableCell.forTableColumn());
        valueCol.setOnEditCommit(
                e -> {
                    e.getRowValue().setValue(e.getNewValue());
                    table.refresh();
                });

        TableColumn<EnvVarRow, Void> folderCol = new TableColumn<>("\u9078\u629e");
        folderCol.setPrefWidth(120);
        folderCol.setSortable(false);
        folderCol.setCellFactory(
                col -> new TableCell<EnvVarRow, Void>() {
                    private final Button pickFolder = new Button("\u30d5\u30a9\u30eb\u30c0...");
                    private final Button pickFile = new Button("\u30d5\u30a1\u30a4\u30eb...");

                    {
                        pickFolder.setOnAction(
                                ev -> {
                                    EnvVarRow row = getTableRow() != null ? getTableRow().getItem() : null;
                                    if (row == null) {
                                        return;
                                    }
                                    DirectoryChooser dc = new DirectoryChooser();
                                    dc.setTitle(
                                            "\u30d5\u30a9\u30eb\u30c0\u3092\u9078\u629e: "
                                                    + row.getName());
                                    String cur = row.getValue();
                                    if (cur != null && !cur.isBlank()) {
                                        try {
                                            Path p = Path.of(cur.trim());
                                            if (Files.isDirectory(p)) {
                                                dc.setInitialDirectory(p.toFile());
                                            } else {
                                                Path par = p.getParent();
                                                if (par != null && Files.isDirectory(par)) {
                                                    dc.setInitialDirectory(par.toFile());
                                                }
                                            }
                                        } catch (Exception ignored) {
                                            // keep default initial directory
                                        }
                                    }
                                    File f = dc.showDialog(ownerStage);
                                    if (f != null) {
                                        row.setValue(f.getAbsolutePath());
                                        table.refresh();
                                    }
                                });
                        pickFile.setOnAction(
                                ev -> {
                                    EnvVarRow row = getTableRow() != null ? getTableRow().getItem() : null;
                                    if (row == null) {
                                        return;
                                    }
                                    FileChooser fc = new FileChooser();
                                    fc.setTitle(
                                            "\u30d5\u30a1\u30a4\u30eb\u3092\u9078\u629e: "
                                                    + row.getName());
                                    FileChooserForEnvKey.apply(fc, row.getName());
                                    String cur = row.getValue();
                                    if (cur != null && !cur.isBlank()) {
                                        try {
                                            Path p = Path.of(cur.trim());
                                            if (Files.isRegularFile(p)) {
                                                fc.setInitialDirectory(
                                                        p.getParent() != null
                                                                ? p.getParent().toFile()
                                                                : null);
                                                fc.setInitialFileName(p.getFileName().toString());
                                            } else if (Files.isDirectory(p)) {
                                                fc.setInitialDirectory(p.toFile());
                                            } else {
                                                Path par = p.getParent();
                                                if (par != null && Files.isDirectory(par)) {
                                                    fc.setInitialDirectory(par.toFile());
                                                }
                                            }
                                        } catch (Exception ignored) {
                                            // keep defaults
                                        }
                                    }
                                    File f = fc.showOpenDialog(ownerStage);
                                    if (f != null) {
                                        row.setValue(f.getAbsolutePath());
                                        table.refresh();
                                    }
                                });
                    }

                    @Override
                    protected void updateItem(Void item, boolean empty) {
                        super.updateItem(item, empty);
                        if (empty) {
                            setGraphic(null);
                            return;
                        }
                        TableRow<EnvVarRow> tr = getTableRow();
                        EnvVarRow row = tr != null ? tr.getItem() : null;
                        if (row == null) {
                            int i = getIndex();
                            if (i >= 0 && i < getTableView().getItems().size()) {
                                row = getTableView().getItems().get(i);
                            }
                        }
                        String key = row != null && row.getName() != null ? row.getName() : "";
                        if (row != null && AppPaths.isFolderPathEnvKey(key)) {
                            setGraphic(pickFolder);
                        } else if (row != null && AppPaths.isFilePathEnvKey(key)) {
                            setGraphic(pickFile);
                        } else {
                            setGraphic(null);
                        }
                    }
                });

        TableColumn<EnvVarRow, String> descCol = new TableColumn<>("\u8aac\u660e\uff08\u30b7\u30fc\u30c8+\u30ed\u30b8\u30c3\u30af\uff09");
        descCol.setCellValueFactory(cdf -> cdf.getValue().descriptionProperty());
        descCol.setPrefWidth(420);
        descCol.setCellFactory(
                col -> new TableCell<EnvVarRow, String>() {
                    private final Text text = new Text();

                    {
                        text.wrappingWidthProperty().bind(col.widthProperty().subtract(16));
                    }

                    @Override
                    protected void updateItem(String item, boolean empty) {
                        super.updateItem(item, empty);
                        if (empty || item == null) {
                            setGraphic(null);
                        } else {
                            text.setText(item);
                            setGraphic(text);
                        }
                    }
                });

        table.getColumns().setAll(nameCol, valueCol, folderCol, descCol);

        Runnable resetEnvColumns =
                () -> {
                    nameCol.setPrefWidth(220);
                    valueCol.setPrefWidth(280);
                    folderCol.setPrefWidth(120);
                    descCol.setPrefWidth(420);
                };

        HBox envColStrip = TableViewColumnSettingsStrip.create(table, resetEnvColumns, true);

        Button addRow = new Button("\u884c\u3092\u8ffd\u52a0");
        addRow.setOnAction(
                e -> {
                    EnvVarRow r = new EnvVarRow();
                    r.setDescription("");
                    envRows.add(r);
                });
        Button delRow = new Button("\u884c\u3092\u524a\u9664");
        delRow.setOnAction(e -> {
            var sel = table.getSelectionModel().getSelectedItems();
            if (!sel.isEmpty()) {
                envRows.removeAll(sel);
            } else if (!envRows.isEmpty()) {
                envRows.remove(envRows.size() - 1);
            }
            if (envRows.isEmpty()) {
                envRows.add(new EnvVarRow());
            }
        });

        HBox btnRow = new HBox(8, addRow, delRow);
        VBox box = new VBox(8, hint, envColStrip, table, btnRow);
        box.setFillWidth(true);
        VBox.setVgrow(table, Priority.ALWAYS);
        BorderPane wrap = new BorderPane(box);
        wrap.setPadding(new Insets(12));
        return wrap;
    }

    private Parent buildChartPane() {
        SwingNode swingNode = new SwingNode();
        DefaultCategoryDataset ds = new DefaultCategoryDataset();
        ds.addValue(12, "actual", "M-A");
        ds.addValue(8, "actual", "M-B");
        ds.addValue(15, "plan", "M-A");
        SwingUtilities.invokeLater(() -> {
            JFreeChart chart = ChartFactory.createBarChart(
                    "sample by equipment (JFreeChart)",
                    "equipment",
                    "qty",
                    ds,
                    PlotOrientation.VERTICAL,
                    true,
                    true,
                    false);
            ChartPanel panel = new ChartPanel(chart);
            panel.setFillZoomRectangle(true);
            javafx.application.Platform.runLater(() -> swingNode.setContent(panel));
        });
        BorderPane bp = new BorderPane(swingNode);
        bp.setPadding(new Insets(12));
        bp.setMinHeight(400);
        BorderPane.setMargin(swingNode, new Insets(8));
        return bp;
    }

    private Button buttonStage(String script, String label) {
        Button b = new Button(label);
        b.setOnAction(e -> runStage(script));
        return b;
    }

    private void pickWorkbook(Stage stage) {
        FileChooser ch = new FileChooser();
        ch.setTitle("\u30de\u30af\u30ed\u30d6\u30c3\u30af\u9078\u629e");
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel", "*.xlsm", "*.xlsx"));
        var f = ch.showOpenDialog(stage);
        if (f != null) {
            workbookField.setText(f.getAbsolutePath());
        }
    }

    private void peekSheetsAction() {
        String p = effectiveTaskInputWorkbookPath();
        if (p.isEmpty()) {
            appendLog(
                    "[POI] \u30de\u30af\u30ed\u5b9f\u884c\u30d6\u30c3\u30af\u304c\u7a7a\uff08\u4efb\u610f\uff09\u3002"
                            + " \u30b7\u30fc\u30c8\u4e00\u89a7\u306b\u306f\u30d1\u30b9\u304c\u5fc5\u8981\u3067\u3059\u3002");
            return;
        }
        try {
            var names = ExcelSheetTitlesProbe.sheetNames(Path.of(p));
            appendLog("[POI] sheets=" + names.size() + " " + String.join(", ", names));
        } catch (Exception ex) {
            appendLog("[POI] error: " + ex.getMessage());
        }
    }

    private void runStage(String script) {
        if (!runLock.compareAndSet(false, true)) {
            appendLog("[busy] already running (single flight).");
            return;
        }
        Map<String, String> uiRun = collectUiEnv();
        Path py =
                Path.of(
                        firstNonBlank(
                                uiRun.get(AppPaths.KEY_PM_AI_PYTHON),
                                pythonExeField.getText().trim()));
        Path dir =
                Path.of(
                        firstNonBlank(
                                uiRun.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                scriptDirField.getText().trim()));
        String wb = effectiveTaskInputWorkbookPath();
        appendLog("--- start: " + script + " ---");
        RunRequest req = new RunRequest(py, dir, script, wb, childEnvForPython(uiRun));
        statusLabel.setText("running\u2026");

        PythonProcessRunner.runAsync(
                        req,
                        line -> {
                            if (line.startsWith(NDJSON_START)) {
                                String payload = line.substring(PREFIX_CHILD.length());
                                IpcStdoutTap.handleLine(payload, this::appendLog);
                            } else {
                                appendLog(line);
                            }
                        },
                        ex -> appendLog("[error] " + ex.getMessage()))
                .whenComplete((code, err) -> {
                    runLock.set(false);
                    javafx.application.Platform.runLater(() -> {
                        if (err != null) {
                            statusLabel.setText("failed: " + err.getMessage());
                            appendLog("[end] exceptional exit");
                        } else {
                            int c = code != null ? code : -1;
                            statusLabel.setText(exitCodeLegend(c));
                            appendLog("[end] exitCode=" + c + " " + exitHint(c));
                        }
                    });
                });
    }

    private static String exitCodeLegend(int code) {
        return "exit="
                + code
                + " \u2014 0=OK / 1=error / 2=fatal / 3=PlanningValidationError / 9=cancel";
    }

    private static String exitHint(int code) {
        return switch (code) {
            case 0 -> "(success)";
            case 1 -> "(general failure)";
            case 2 -> "(fatal / missing TASK_INPUT / file)";
            case 3 -> "(PlanningValidationError)";
            case 9 -> "(user cancel)";
            default -> "";
        };
    }

    /**
     * Main macro-book field, else {@link AppPaths#resolveTaskInputWorkbook(Map)} (workspace / repo .xlsm).
     * Python still receives {@code TASK_INPUT_WORKBOOK} only via {@link PythonProcessRunner}, not the env tab.
     */
    private String effectiveTaskInputWorkbookPath() {
        String t = workbookField.getText() != null ? workbookField.getText().trim() : "";
        if (!t.isEmpty()) {
            return t;
        }
        return AppPaths.resolveTaskInputWorkbook(collectUiEnv()).map(Path::toString).orElse("");
    }

    /** Probe script {@code pm_ai_actuals_status.py}: same env merge as stage1/2. */
    private RunRequest buildActualsStatusRequest() {
        Map<String, String> uiRun = collectUiEnv();
        Path py =
                Path.of(
                        firstNonBlank(
                                uiRun.get(AppPaths.KEY_PM_AI_PYTHON),
                                pythonExeField.getText().trim()));
        Path dir =
                Path.of(
                        firstNonBlank(
                                uiRun.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                scriptDirField.getText().trim()));
        String wb = effectiveTaskInputWorkbookPath();
        return new RunRequest(py, dir, "pm_ai_actuals_status.py", wb, childEnvForPython(uiRun));
    }

    /**
     * Env tab keys passed to Python; omits VBA-era workbook keys (launcher sets {@code TASK_INPUT_WORKBOOK}).
     */
    private static Map<String, String> childEnvForPython(Map<String, String> ui) {
        Map<String, String> m = new HashMap<>(ui);
        m.remove("TASK_INPUT_WORKBOOK");
        m.remove("PM_AI_TASK_INPUT_WORKBOOK");
        String skip = m.get(AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET);
        if (skip == null || skip.isBlank()) {
            m.put(AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET, "1");
        }
        return m;
    }

    /**
     * Child-process env from the \u74b0\u5883\u5909\u6570 tab (same skip rules as workbook sheet: empty name, #).
     */
    private Map<String, String> collectUiEnv() {
        Map<String, String> m = new HashMap<>();
        if (envRows == null) {
            return m;
        }
        for (EnvVarRow row : envRows) {
            String k = row.getName() != null ? row.getName().trim() : "";
            if (k.isEmpty() || k.startsWith("#")) {
                continue;
            }
            m.put(k, row.getValue() != null ? row.getValue() : "");
        }
        return m;
    }

    private void appendLog(String line) {
        logArea.appendText(line + "\n");
    }

    private static String defaultOsPython() {
        return System.getProperty("os.name", "").toLowerCase().contains("win") ? "python" : "python3";
    }

    private static String firstNonBlank(String... parts) {
        if (parts == null) {
            return "";
        }
        for (String p : parts) {
            if (p != null && !p.isBlank()) {
                return p.trim();
            }
        }
        return "";
    }

    private void populateEnvRows(ObservableList<EnvVarRow> rows) {
        LinkedHashMap<String, EnvVarRow> sheet = new LinkedHashMap<>();
        for (WorkbookEnvSheetReader.RowEntry e : UiRefEnvDefaults.loadOrEmpty()) {
            EnvVarRow row = new EnvVarRow();
            row.setName(e.key());
            row.setValue(e.value() != null ? e.value() : "");
            row.setDescription(EnvVarDocs.mergeDescriptions(e.description(), e.key()));
            sheet.put(e.key(), row);
        }
        Map<String, String> empty = Map.of();
        LinkedHashMap<String, EnvVarRow> ordered = new LinkedHashMap<>();
        for (String k : BOOTSTRAP_ORDER) {
            EnvVarRow existing = sheet.remove(k);
            if (existing != null) {
                maybeFillEmptyBootstrap(existing, k, empty);
                ordered.put(k, existing);
            } else {
                ordered.put(k, newBootstrapRow(k, empty));
            }
        }
        ordered.putAll(sheet);
        rows.setAll(new ArrayList<>(ordered.values()));
        if (rows.isEmpty()) {
            rows.add(new EnvVarRow());
        }
    }

    private static void maybeFillEmptyBootstrap(EnvVarRow r, String k, Map<String, String> ui) {
        if (r.getValue() != null && !r.getValue().isBlank()) {
            return;
        }
        switch (k) {
            case AppPaths.KEY_PM_AI_PYTHON -> r.setValue(defaultOsPython());
            case AppPaths.KEY_PM_AI_REPO_ROOT -> r.setValue(AppPaths.resolveRepoRoot(ui).toString());
            case AppPaths.KEY_PM_AI_CODE_PYTHON_DIR -> r.setValue(AppPaths.resolvePythonScriptDir(ui).toString());
            case AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR ->
                    r.setValue(AppPaths.resolveTaskInputSourceDir(ui).toString());
            case AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR ->
                    r.setValue(AppPaths.resolveActualDetailSourceDir(ui).toString());
            case AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR ->
                    r.setValue(AppPaths.resolveResultDispatchTableDir(ui).toString());
            case AppPaths.KEY_GEMINI_CREDENTIALS_JSON -> {
                Path cand =
                        AppPaths.resolveRepoRoot(ui)
                                .resolve("code")
                                .resolve("gemini_credentials.encrypted.json");
                if (Files.isRegularFile(cand)) {
                    r.setValue(cand.toAbsolutePath().normalize().toString());
                }
            }
            case AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON -> {
                Path cand =
                        AppPaths.resolveRepoRoot(ui).resolve("code").resolve("exclude_rules.json");
                if (Files.isRegularFile(cand)) {
                    r.setValue(cand.toAbsolutePath().normalize().toString());
                }
            }
            case AppPaths.KEY_PM_AI_MASTER_WORKBOOK ->
                    AppPaths.resolveMasterWorkbookCandidate(ui)
                            .ifPresent(p -> r.setValue(p.toString()));
            case AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET -> r.setValue("1");
            default -> {
                /* PM_AI_WORKSPACE stays empty */
            }
        }
    }

    private static EnvVarRow newBootstrapRow(String k, Map<String, String> ui) {
        EnvVarRow r = new EnvVarRow();
        r.setName(k);
        r.setDescription(EnvVarDocs.mergeDescriptions("", k));
        switch (k) {
            case AppPaths.KEY_PM_AI_PYTHON -> r.setValue(defaultOsPython());
            case AppPaths.KEY_PM_AI_REPO_ROOT -> r.setValue(AppPaths.resolveRepoRoot(ui).toString());
            case AppPaths.KEY_PM_AI_CODE_PYTHON_DIR -> r.setValue(AppPaths.resolvePythonScriptDir(ui).toString());
            case AppPaths.KEY_PM_AI_WORKSPACE -> r.setValue("");
            case AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR ->
                    r.setValue(AppPaths.resolveTaskInputSourceDir(ui).toString());
            case AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR ->
                    r.setValue(AppPaths.resolveActualDetailSourceDir(ui).toString());
            case AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR ->
                    r.setValue(AppPaths.resolveResultDispatchTableDir(ui).toString());
            case AppPaths.KEY_GEMINI_CREDENTIALS_JSON -> {
                Path cand =
                        AppPaths.resolveRepoRoot(ui)
                                .resolve("code")
                                .resolve("gemini_credentials.encrypted.json");
                r.setValue(
                        Files.isRegularFile(cand)
                                ? cand.toAbsolutePath().normalize().toString()
                                : "");
            }
            case AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON -> {
                Path cand =
                        AppPaths.resolveRepoRoot(ui).resolve("code").resolve("exclude_rules.json");
                r.setValue(
                        Files.isRegularFile(cand)
                                ? cand.toAbsolutePath().normalize().toString()
                                : "");
            }
            case AppPaths.KEY_PM_AI_MASTER_WORKBOOK ->
                    r.setValue(
                            AppPaths.resolveMasterWorkbookCandidate(ui)
                                    .map(Path::toString)
                                    .orElse(""));
            case AppPaths.KEY_PM_AI_COLUMN_CONFIG_WORKBOOK,
                    AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK,
                    AppPaths.KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV -> r.setValue("");
            case AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET -> r.setValue("1");
            default -> r.setValue("");
        }
        return r;
    }

    public static void main(String[] args) {
        System.setProperty("file.encoding", "UTF-8");
        if (GraphicsEnvironment.isHeadless()) {
            System.err.println(
                    "[PmAiFxApp] No graphical display (headless). "
                            + "Run on Windows desktop, or on WSL set DISPLAY for JavaFX (e.g. WSLg / VcXsrv). "
                            + "Do not run javafx:run from SSH without X forwarding.");
            System.exit(2);
        }
        launch(args);
    }

    /** Editable row for {@link TableView} (setting_ env sheet parity). */
    static final class EnvVarRow {
        private final StringProperty name = new SimpleStringProperty("");
        private final StringProperty value = new SimpleStringProperty("");
        private final StringProperty description = new SimpleStringProperty("");

        String getName() {
            return name.get();
        }

        void setName(String v) {
            name.set(v != null ? v : "");
        }

        StringProperty nameProperty() {
            return name;
        }

        String getValue() {
            return value.get();
        }

        void setValue(String v) {
            value.set(v != null ? v : "");
        }

        StringProperty valueProperty() {
            return value;
        }

        String getDescription() {
            return description.get();
        }

        void setDescription(String v) {
            description.set(v != null ? v : "");
        }

        StringProperty descriptionProperty() {
            return description;
        }
    }
}
