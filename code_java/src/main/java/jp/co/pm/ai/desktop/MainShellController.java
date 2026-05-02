package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.TabPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.EnvVarDocs;
import jp.co.pm.ai.desktop.config.UiRefEnvDefaults;
import jp.co.pm.ai.desktop.io.ExcelSheetTitlesProbe;
import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;
import jp.co.pm.ai.desktop.ipc.IpcStdoutTap;
import jp.co.pm.ai.desktop.ui.ActualsDataStatusPane;
import jp.co.pm.ai.desktop.ui.ExcludeRulesEditorPane;
import jp.co.pm.ai.desktop.ui.PlanInputEditorPane;
import jp.co.pm.ai.desktop.ui.Stage1ShapedOutputPreviewPane;

/**
 * Main window controller \u2014 business logic moved from legacy inline {@link PmAiFxApp}.
 * Layout: {@code MainShell.fxml} and tab FXML files.
 */
public final class MainShellController {

    private static final String STAGE1 = "task_extract_stage1.py";
    private static final String STAGE2 = "plan_simulation_stage2.py";
    private static final String PREFIX_CHILD = "[child] ";
    private static final String NDJSON_START = PREFIX_CHILD + "{";

    private static final List<String> BOOTSTRAP_ORDER =
            List.of(
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

    private final Stage primaryStage;

    @FXML
    private TabPane tabPane;

    @FXML
    private MainRunTabController mainRunTabController;

    @FXML
    private EnvTabController envTabController;

    @FXML
    private BorderPane planInputHost;

    @FXML
    private BorderPane stage1Host;

    @FXML
    private BorderPane excludeRulesHost;

    @FXML
    private BorderPane actualsHost;

    private ObservableList<EnvVarRow> envRows;
    private final AtomicBoolean runLock = new AtomicBoolean(false);

    /** Set by {@link Stage1ShapedOutputPreviewPane#create}; runs after stage 1 exits 0. */
    private Runnable reloadAfterStage1Preview;

    /** Set by {@link PlanInputEditorPane#create}; loads {@code plan_input_tasks.xlsx}. */
    private Runnable reloadAfterStage1PlanInput;

    MainShellController(Stage primaryStage) {
        this.primaryStage = primaryStage;
    }

    @FXML
    private void initialize() {
        envRows = FXCollections.observableArrayList();
        populateEnvRows(envRows);
        Map<String, String> ui0 = collectUiEnv();

        mainRunTabController.bindShell(this);
        envTabController.bindShell(this);

        mainRunTabController
                .getWorkbookField()
                .setPromptText(
                        "\u4efb\u610f \u2014 \u7a7a\u3067\u3088\u3044\uff08\u6bb5\u968e1\u306f PM_AI_PROCESSING_PLAN_PATH \u306a\u3069\u3067\u53ef\uff09\u3002"
                                + " \u6307\u5b9a\u6642\u306f ProcessBuilder \u304c TASK_INPUT_WORKBOOK \u3092\u4ed8\u4e0e");
        mainRunTabController
                .getWorkbookField()
                .setText(AppPaths.resolveTaskInputWorkbook(ui0).map(Path::toString).orElse(""));
        mainRunTabController
                .getPythonExeField()
                .setText(firstNonBlank(ui0.get(AppPaths.KEY_PM_AI_PYTHON), defaultOsPython()));
        mainRunTabController
                .getPythonExeField()
                .setPromptText("Python executable (\u74b0\u5883\u5909\u6570 PM_AI_PYTHON)");
        mainRunTabController
                .getScriptDirField()
                .setText(
                        firstNonBlank(
                                ui0.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                AppPaths.resolvePythonScriptDir(ui0).toString()));
        mainRunTabController
                .getScriptDirField()
                .setPromptText("code/python (\u74b0\u5883\u5909\u6570 PM_AI_CODE_PYTHON_DIR)");

        planInputHost.setCenter(
                PlanInputEditorPane.create(
                        primaryStage,
                        this::collectUiEnv,
                        this::appendLog,
                        rr -> reloadAfterStage1PlanInput = rr));
        stage1Host.setCenter(
                Stage1ShapedOutputPreviewPane.create(
                        primaryStage,
                        this::collectUiEnv,
                        this::appendLog,
                        rr -> reloadAfterStage1Preview = rr));
        excludeRulesHost.setCenter(
                ExcludeRulesEditorPane.create(primaryStage, this::collectUiEnv, this::appendLog));
        actualsHost.setCenter(
                ActualsDataStatusPane.create(this::buildActualsStatusRequest, this::appendLog));

        primaryStage.setMinWidth(640);
        primaryStage.setMinHeight(480);
        primaryStage.setOnShown(
                e -> {
                    primaryStage.toFront();
                    tabPane.getSelectionModel().selectFirst();
                });
    }

    Stage getPrimaryStage() {
        return primaryStage;
    }

    ObservableList<EnvVarRow> getEnvRows() {
        return envRows;
    }

    void appendBootMessage() {
        appendLog("[boot] PYTHONUTF8=1 PYTHONIOENCODING=utf-8 for child process.");
    }

    void attachStageButtons(HBox box) {
        box.getChildren()
                .setAll(
                        buttonStage(STAGE1, "\u6bb5\u968e1 \u5b9f\u884c"),
                        buttonStage(STAGE2, "\u6bb5\u968e2 \u5b9f\u884c"),
                        peekSheetsButton());
    }

    private Button peekSheetsButton() {
        Button peekSheets = new Button("\u30b7\u30fc\u30c8\u4e00\u89a7 (POI)");
        peekSheets.setOnAction(e -> peekSheetsAction());
        return peekSheets;
    }

    void pickWorkbook() {
        FileChooser ch = new FileChooser();
        ch.setTitle("\u30de\u30af\u30ed\u30d6\u30c3\u30af\u9078\u629e");
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel", "*.xlsm", "*.xlsx"));
        var f = ch.showOpenDialog(primaryStage);
        if (f != null) {
            mainRunTabController.getWorkbookField().setText(f.getAbsolutePath());
        }
    }

    String resolveTaskInputWorkbookFromEnv() {
        return AppPaths.resolveTaskInputWorkbook(collectUiEnv()).map(Path::toString).orElse("");
    }

    String resolvePythonScriptDirFromEnv() {
        return AppPaths.resolvePythonScriptDir(collectUiEnv()).toString();
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

    private Button buttonStage(String script, String label) {
        Button b = new Button(label);
        b.setOnAction(e -> runStage(script));
        return b;
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
                                mainRunTabController.getPythonExeField().getText().trim()));
        Path dir =
                Path.of(
                        firstNonBlank(
                                uiRun.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                mainRunTabController.getScriptDirField().getText().trim()));
        String wb = effectiveTaskInputWorkbookPath();
        appendLog("--- start: " + script + " ---");
        RunRequest req = new RunRequest(py, dir, script, wb, childEnvForPython(uiRun));
        mainRunTabController.getStatusLabel().setText("running\u2026");

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
                .whenComplete(
                        (code, err) -> {
                            runLock.set(false);
                            javafx.application.Platform.runLater(
                                    () -> {
                                        if (err != null) {
                                            mainRunTabController
                                                    .getStatusLabel()
                                                    .setText("failed: " + err.getMessage());
                                            appendLog("[end] exceptional exit");
                                        } else {
                                            int c = code != null ? code : -1;
                                            mainRunTabController.getStatusLabel().setText(exitCodeLegend(c));
                                            appendLog("[end] exitCode=" + c + " " + exitHint(c));
                                            if (STAGE1.equals(script) && c == 0) {
                                                if (reloadAfterStage1Preview != null) {
                                                    reloadAfterStage1Preview.run();
                                                }
                                                if (reloadAfterStage1PlanInput != null) {
                                                    reloadAfterStage1PlanInput.run();
                                                }
                                            }
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
        String t =
                mainRunTabController.getWorkbookField().getText() != null
                        ? mainRunTabController.getWorkbookField().getText().trim()
                        : "";
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
                                mainRunTabController.getPythonExeField().getText().trim()));
        Path dir =
                Path.of(
                        firstNonBlank(
                                uiRun.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                mainRunTabController.getScriptDirField().getText().trim()));
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
        mainRunTabController.appendLog(line);
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
                    AppPaths.resolveMasterWorkbookCandidate(ui).ifPresent(p -> r.setValue(p.toString()));
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
                            AppPaths.resolveMasterWorkbookCandidate(ui).map(Path::toString).orElse(""));
            case AppPaths.KEY_PM_AI_COLUMN_CONFIG_WORKBOOK,
                    AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK,
                    AppPaths.KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV -> r.setValue("");
            case AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET -> r.setValue("1");
            default -> r.setValue("");
        }
        return r;
    }
}
