package jp.co.pm.ai.desktop;

import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Rectangle2D;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TabPane;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.util.Duration;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.DesktopTheme;
import jp.co.pm.ai.desktop.config.EnvVarDocs;
import jp.co.pm.ai.desktop.config.UiEnvRowSnapshot;
import jp.co.pm.ai.desktop.config.UiRefEnvDefaults;
import jp.co.pm.ai.desktop.io.ExcelSheetTitlesProbe;
import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;
import jp.co.pm.ai.desktop.ipc.IpcStdoutTap;

/**
 * Main window controller \u2014 business logic moved from legacy inline {@link PmAiFxApp}.
 * Layout: {@code MainShell.fxml} and tab FXML files.
 */
public final class MainShellController {

    private static final String STAGE1 = "task_extract_stage1.py";
    private static final String STAGE2 = "plan_simulation_stage2.py";
    private static final String PREFIX_CHILD = "[child] ";
    private static final String NDJSON_START = PREFIX_CHILD + "{";

    /** Legacy keys removed from the env tab and never passed to Python children. */
    private static final Set<String> REMOVED_ENV_VAR_KEYS =
            Set.of("TASK_INPUT_WORKBOOK", "PM_AI_TASK_INPUT_WORKBOOK");

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
    private Button clearTabFiltersButton;

    @FXML
    private ComboBox<DesktopTheme> themeCombo;

    @FXML
    private MainRunTabController mainRunTabController;

    @FXML
    private EnvTabController envTabController;

    @FXML
    private PlanInputTabController planInputTabController;

    @FXML
    private Stage1PreviewTabController stage1PreviewTabController;

    @FXML
    private ExcludeRulesTabController excludeRulesTabController;

    @FXML
    private ActualsStatusTabController actualsStatusTabController;

    @FXML
    private MasterReadSummaryTabController masterReadSummaryTabController;

    private ObservableList<EnvVarRow> envRows;
    private final AtomicBoolean runLock = new AtomicBoolean(false);
    private final AtomicBoolean suppressEnvSessionPersistence = new AtomicBoolean(false);
    private final PauseTransition uiEnvSaveDebounce = new PauseTransition(Duration.millis(400));

    private DesktopTheme pendingTheme = DesktopTheme.LIGHT;

    /** Set by {@link Stage1PreviewTabController}; runs after stage 1 exits 0. */
    private Runnable reloadAfterStage1Preview;

    /** Set by {@link PlanInputTabController}; loads {@code plan_input_tasks.xlsx}. */
    private Runnable reloadAfterStage1PlanInput;

    MainShellController(Stage primaryStage) {
        this.primaryStage = primaryStage;
    }

    @FXML
    private void initialize() {
        suppressEnvSessionPersistence.set(true);
        try {
            envRows = FXCollections.observableArrayList();
            populateEnvRows(envRows);
            Map<String, String> ui0 = collectUiEnv();

            mainRunTabController.bindShell(this);
            envTabController.bindShell(this);
            masterReadSummaryTabController.bindShell(this);

        mainRunTabController
                .getWorkbookField()
                .setPromptText(
                        "\u4efb\u610f\u3002\u7a7a\u3067\u3088\u3044\uff08\u6bb5\u968e1/2\u306f\u74b0\u5883\u5909\u6570\u30bf\u30d6\u306e PM_AI_* \u304c\u672c\u7dda\uff09\u3002"
                                + "\u30de\u30b9\u30bf\u8aad\u8fbc\u30b5\u30de\u30ea\u7b49\u306eUI\u63a8\u5b9a\u7528\u3002");
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

        planInputTabController.bindShell(this);
        stage1PreviewTabController.bindShell(this);
        excludeRulesTabController.bindShell(this);
        actualsStatusTabController.bindShell(this);

        primaryStage.setMinWidth(640);
        primaryStage.setMinHeight(480);

            applyDesktopSession(DesktopSessionStateStore.load());
        } finally {
            suppressEnvSessionPersistence.set(false);
        }

        installUiEnvAutoSave();

        primaryStage.setOnCloseRequest(e -> DesktopSessionStateStore.save(collectDesktopSession()));

        primaryStage.setOnShown(
                e -> {
                    primaryStage.toFront();
                    tabPane.getSelectionModel().selectFirst();
                });

        tabPane
                .getSelectionModel()
                .selectedIndexProperty()
                .addListener(
                        (obs, prev, idx) ->
                                updateClearTabFiltersButton(idx != null ? idx.intValue() : -1));
        updateClearTabFiltersButton(tabPane.getSelectionModel().getSelectedIndex());
    }

    /**
     * Invoked from {@link PmAiFxApp} after {@link Scene} creation so theme stylesheets can target the scene.
     */
    public void finishStartup(Scene scene) {
        if (themeCombo == null) {
            return;
        }
        themeCombo.getItems().setAll(DesktopTheme.values());
        themeCombo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(DesktopTheme t) {
                        return t != null ? t.displayLabel() : "";
                    }

                    @Override
                    public DesktopTheme fromString(String s) {
                        return DesktopTheme.fromDisplayLabel(s);
                    }
                });
        DesktopTheme initial = pendingTheme != null ? pendingTheme : DesktopTheme.LIGHT;
        initial.applyTo(scene);
        themeCombo.setValue(initial);
        themeCombo
                .valueProperty()
                .addListener(
                        (obs, oldV, newV) -> {
                            if (newV != null) {
                                newV.applyTo(scene);
                            }
                            mainRunTabController.refreshLogThemeCells();
                        });
        Platform.runLater(mainRunTabController::refreshLogThemeCells);
    }

    /** Theme shown in toolbar (for components that need dark/light tint hints). */
    DesktopTheme currentDesktopTheme() {
        if (themeCombo != null && themeCombo.getValue() != null) {
            return themeCombo.getValue();
        }
        return pendingTheme != null ? pendingTheme : DesktopTheme.LIGHT;
    }

    private static boolean tabSupportsClearFilters(int idx) {
        return idx == 1 || idx == 2 || idx == 3 || idx == 5;
    }

    private void updateClearTabFiltersButton(int idx) {
        if (clearTabFiltersButton == null) {
            return;
        }
        clearTabFiltersButton.setDisable(!tabSupportsClearFilters(idx));
    }

    @FXML
    private void onClearTabFiltersAction() {
        switch (tabPane.getSelectionModel().getSelectedIndex()) {
            case 1 -> envTabController.clearColumnFiltersAndSort();
            case 2 -> planInputTabController.clearColumnFiltersAndSort();
            case 3 -> stage1PreviewTabController.clearColumnFiltersAndSort();
            case 5 -> actualsStatusTabController.clearColumnFiltersAndSort();
            default -> {
                /* tabs without TableFilter / Spreadsheet filter row */
            }
        }
    }

    private void applyDesktopSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        applyUiEnvRowsFromSession(s);
        planInputTabController.restoreDesktopSessionPaths(s.planInputPath(), s.planInputSheet());
        stage1PreviewTabController.restoreDesktopSessionPaths(s.stage1PreviewPath(), s.stage1PreviewSheet());
        excludeRulesTabController.restoreDesktopSessionPath(s.excludeRulesPath());
        if (nonBlank(s.mainRunWorkbook())) {
            mainRunTabController.getWorkbookField().setText(s.mainRunWorkbook());
        }
        if (nonBlank(s.mainRunPythonExe())) {
            mainRunTabController.getPythonExeField().setText(s.mainRunPythonExe());
        }
        if (nonBlank(s.mainRunScriptDir())) {
            mainRunTabController.getScriptDirField().setText(s.mainRunScriptDir());
        }
        mainRunTabController.applyLogFontFromSession(s.logFontFamily(), s.logFontSize());
        applyWindowGeometry(s);
        pendingTheme = DesktopTheme.fromStored(s.uiTheme());
        Platform.runLater(() -> excludeRulesTabController.tryStartupLoadFromPathField());
    }

    private void applyWindowGeometry(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        double w = s.windowWidth();
        double h = s.windowHeight();
        double minW = primaryStage.getMinWidth();
        double minH = primaryStage.getMinHeight();
        if (Double.isFinite(w)
                && Double.isFinite(h)
                && w >= minW
                && h >= minH) {
            primaryStage.setWidth(w);
            primaryStage.setHeight(h);
        }
        double x = s.windowX();
        double y = s.windowY();
        if (Double.isFinite(x) && Double.isFinite(y)) {
            Rectangle2D screen = Screen.getPrimary().getVisualBounds();
            double ww = primaryStage.getWidth();
            double hh = primaryStage.getHeight();
            double maxX = Math.max(screen.getMinX(), screen.getMaxX() - ww);
            double maxY = Math.max(screen.getMinY(), screen.getMaxY() - hh);
            primaryStage.setX(clamp(x, screen.getMinX(), maxX));
            primaryStage.setY(clamp(y, screen.getMinY(), maxY));
        }
    }

    private static double clamp(double v, double lo, double hi) {
        if (hi < lo) {
            return lo;
        }
        return Math.max(lo, Math.min(hi, v));
    }

    private DesktopSessionState collectDesktopSession() {
        return new DesktopSessionState(
                planInputTabController.snapshotPlanInputPath(),
                planInputTabController.snapshotPlanInputSheet(),
                stage1PreviewTabController.snapshotStage1PreviewPath(),
                stage1PreviewTabController.snapshotStage1PreviewSheet(),
                excludeRulesTabController.snapshotExcludeRulesPath(),
                nz(mainRunTabController.getWorkbookField().getText()),
                nz(mainRunTabController.getPythonExeField().getText()),
                nz(mainRunTabController.getScriptDirField().getText()),
                primaryStage.getWidth(),
                primaryStage.getHeight(),
                primaryStage.getX(),
                primaryStage.getY(),
                themeCombo != null && themeCombo.getValue() != null
                        ? themeCombo.getValue().storedId()
                        : DesktopTheme.LIGHT.storedId(),
                mainRunTabController.snapshotLogFontFamily(),
                mainRunTabController.snapshotLogFontSize(),
                snapshotUiEnvRows());
    }

    private List<UiEnvRowSnapshot> snapshotUiEnvRows() {
        if (envRows == null) {
            return List.of();
        }
        List<UiEnvRowSnapshot> out = new ArrayList<>(envRows.size());
        for (EnvVarRow r : envRows) {
            String key = nz(r.getName());
            if (REMOVED_ENV_VAR_KEYS.contains(key)) {
                continue;
            }
            out.add(
                    new UiEnvRowSnapshot(
                            key,
                            r.getValue() != null ? r.getValue() : "",
                            r.getDescription() != null ? r.getDescription() : ""));
        }
        return List.copyOf(out);
    }

    private void applyUiEnvRowsFromSession(DesktopSessionState s) {
        if (s == null || s.uiEnvRows() == null || s.uiEnvRows().isEmpty() || envRows == null) {
            return;
        }
        List<EnvVarRow> restored = new ArrayList<>(s.uiEnvRows().size());
        for (UiEnvRowSnapshot snap : s.uiEnvRows()) {
            String nm = snap.name() != null ? snap.name().trim() : "";
            if (REMOVED_ENV_VAR_KEYS.contains(nm)) {
                continue;
            }
            EnvVarRow row = new EnvVarRow();
            String name = snap.name() != null ? snap.name() : "";
            row.setName(name);
            row.setValue(snap.value() != null ? snap.value() : "");
            String desc = snap.description() != null ? snap.description() : "";
            if (desc.isBlank() && !name.trim().isEmpty()) {
                desc = EnvVarDocs.mergeDescriptions("", name.trim());
            }
            row.setDescription(desc);
            restored.add(row);
        }
        envRows.setAll(restored);
        stripRemovedEnvVarRows(envRows);
    }

    /** Debounced session flush when run-tab log font changes. */
    void scheduleDesktopSessionSave() {
        if (!suppressEnvSessionPersistence.get()) {
            uiEnvSaveDebounce.playFromStart();
        }
    }

    private void installUiEnvAutoSave() {
        uiEnvSaveDebounce.setOnFinished(
                e -> {
                    if (!suppressEnvSessionPersistence.get()) {
                        DesktopSessionStateStore.save(collectDesktopSession());
                    }
                });
        Runnable schedule = () -> uiEnvSaveDebounce.playFromStart();
        envRows.addListener(
                (ListChangeListener<EnvVarRow>)
                        c -> {
                            while (c.next()) {
                                if (c.wasAdded()) {
                                    for (EnvVarRow row : c.getAddedSubList()) {
                                        hookEnvRowForAutoSave(row, schedule);
                                    }
                                }
                            }
                            schedule.run();
                        });
        for (EnvVarRow row : envRows) {
            hookEnvRowForAutoSave(row, schedule);
        }
    }

    private static void hookEnvRowForAutoSave(EnvVarRow row, Runnable schedule) {
        row.nameProperty().addListener((o, a, b) -> schedule.run());
        row.valueProperty().addListener((o, a, b) -> schedule.run());
        row.descriptionProperty().addListener((o, a, b) -> schedule.run());
    }

    private static boolean nonBlank(String v) {
        return v != null && !v.isBlank();
    }

    private static String nz(String s) {
        return s != null ? s.trim() : "";
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

    /**
     * After stage 1 writes {@code json/stage1_exclude_rules.json}, mirror the path into the env tab so
     * {@code PM_AI_EXCLUDE_RULES_JSON} matches the next child-process run.
     */
    private void applyStage1ExcludeRulesJsonToEnvTab() {
        if (envRows == null) {
            return;
        }
        try {
            Map<String, String> ui = collectUiEnv();
            Path p = AppPaths.stage1ExcludeRulesJsonPath(ui);
            if (!Files.isRegularFile(p)) {
                Path legacy = AppPaths.stage1ExcludeRulesJsonPathLegacyUnderPython(ui);
                if (Files.isRegularFile(legacy)) {
                    p = legacy;
                }
            }
            if (!Files.isRegularFile(p)) {
                appendLog(
                        "[env] PM_AI_EXCLUDE_RULES_JSON: "
                                + p
                                + " \u304c\u7121\u3044\u305f\u3081\u74b0\u5883\u5909\u6570\u30bf\u30d6\u306f\u672a\u66f4\u65b0"
                                + "\uff08\u914d\u53f0\u4e0d\u8981 JSON \u672a\u751f\u6210\u307e\u305f\u306f cwd/json \u3068\u4e00\u81f4\u3057\u306a\u3044\uff09\u3002");
                return;
            }
            String pathStr = p.toString();
            for (EnvVarRow row : envRows) {
                String k = row.getName() != null ? row.getName().trim() : "";
                if (AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON.equals(k)) {
                    row.setValue(pathStr);
                    appendLog("[env] PM_AI_EXCLUDE_RULES_JSON=" + pathStr);
                    return;
                }
            }
            appendLog(
                    "[env] PM_AI_EXCLUDE_RULES_JSON \u884c\u304c\u898b\u3064\u304b\u3089\u306a\u3044\u305f\u3081\u672a\u66f4\u65b0\u3002");
        } catch (Exception ex) {
            appendLog("[env] PM_AI_EXCLUDE_RULES_JSON \u66f4\u65b0\u5931\u6557: " + ex.getMessage());
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
                                                applyStage1ExcludeRulesJsonToEnvTab();
                                                if (reloadAfterStage1Preview != null) {
                                                    reloadAfterStage1Preview.run();
                                                }
                                                if (reloadAfterStage1PlanInput != null) {
                                                    reloadAfterStage1PlanInput.run();
                                                }
                                            }
                                            if (STAGE2.equals(script) && c == 0) {
                                                refreshStage2OutputArtifacts();
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
            case 2 -> "(fatal / missing master or processing-plan file)";
            case 3 -> "(PlanningValidationError)";
            case 9 -> "(user cancel)";
            default -> "";
        };
    }

    /**
     * Optional macro-book path from the main-run tab (sheet probe, master path resolution in Java UI).
     * Stage 1/2 child processes do not receive legacy {@link #REMOVED_ENV_VAR_KEYS}; use
     * {@code PM_AI_PLAN_INPUT_PATH} and related keys from the env tab. {@link PythonProcessRunner} ignores the
     * workbook component of {@link PythonProcessRunner.RunRequest} for environment injection.
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

    /** Same as {@link #effectiveTaskInputWorkbookPath()} for Java UI helpers (e.g. master workbook open); not Python env. */
    String effectiveTaskInputWorkbookPathForShell() {
        return effectiveTaskInputWorkbookPath();
    }

    /** Probe script {@code master_read_summary.py}: same env merge as stage1/2. */
    RunRequest buildMasterReadSummaryRequest() {
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
        return new RunRequest(
                py,
                dir,
                MasterReadSummaryTabController.scriptName(),
                wb,
                childEnvForPython(uiRun));
    }

    /** Probe script {@code pm_ai_actuals_status.py}: same env merge as stage1/2. */
    RunRequest buildActualsStatusRequest() {
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
     * Env tab keys passed to Python; strips legacy workbook keys ({@link #REMOVED_ENV_VAR_KEYS}).
     * If {@code PM_AI_PLAN_INPUT_PATH} / {@code TASK_PLAN_SHEET} are unset in the env tab, values from
     * the \u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b tab are applied so \u6bb5\u968e2 matches the
     * file the user is editing there.
     */
    private Map<String, String> childEnvForPython(Map<String, String> ui) {
        Map<String, String> m = new HashMap<>(ui);
        for (String k : REMOVED_ENV_VAR_KEYS) {
            m.remove(k);
        }
        String skip = m.get(AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET);
        if (skip == null || skip.isBlank()) {
            m.put(AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET, "1");
        }
        overlayPlanInputTabPathsIfEnvBlank(m);
        return m;
    }

    /**
     * Fills {@link PlanInputTabController#ENV_PM_AI_PLAN_INPUT_PATH} and {@link
     * PlanInputTabController#ENV_TASK_PLAN_SHEET} from the dedicated plan-input tab when the env tab
     * leaves them blank.
     */
    private void overlayPlanInputTabPathsIfEnvBlank(Map<String, String> m) {
        String pipKey = PlanInputTabController.ENV_PM_AI_PLAN_INPUT_PATH;
        String pip = m.get(pipKey);
        if (pip == null || pip.isBlank()) {
            String tab = planInputTabController.snapshotPlanInputPath();
            if (tab != null && !tab.isBlank()) {
                m.put(pipKey, tab.trim());
            }
        }
        String tpsKey = PlanInputTabController.ENV_TASK_PLAN_SHEET;
        String tps = m.get(tpsKey);
        if (tps == null || tps.isBlank()) {
            String tabSheet = planInputTabController.snapshotPlanInputSheet();
            if (tabSheet != null && !tabSheet.isBlank()) {
                m.put(tpsKey, tabSheet.trim());
            }
        }
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
            if (k.isEmpty() || k.startsWith("#") || REMOVED_ENV_VAR_KEYS.contains(k)) {
                continue;
            }
            m.put(k, row.getValue() != null ? row.getValue() : "");
        }
        return m;
    }

    /** Same-package tab controllers append run-tab log lines here. */
    void appendLog(String line) {
        mainRunTabController.appendLog(line);
    }

    Map<String, String> snapshotUiEnv() {
        return collectUiEnv();
    }

    void acceptReloadAfterStage1PlanInput(Runnable r) {
        this.reloadAfterStage1PlanInput = r;
    }

    void acceptReloadAfterStage1Preview(Runnable r) {
        this.reloadAfterStage1Preview = r;
    }

    void triggerStage1() {
        runStage(STAGE1);
    }

    void triggerStage2() {
        runStage(STAGE2);
    }

    /**
     * After stage-2 success, show newest {@code production_plan_multi_day_*.xlsx} and {@code member_schedule_*.xlsx}
     * under {@link AppPaths#defaultPlanningOutputDir} in the run tab (same folder as plan-input export).
     */
    private void refreshStage2OutputArtifacts() {
        try {
            Map<String, String> ui = collectUiEnv();
            Path dir = AppPaths.defaultPlanningOutputDir(ui);
            if (!Files.isDirectory(dir)) {
                mainRunTabController.setStage2ArtifactPaths("", "");
                appendLog(
                        "[stage2-ui] "
                                + "\u51fa\u529b\u30d5\u30a9\u30eb\u30c0\u304c\u3042\u308a\u307e\u305b\u3093: "
                                + dir);
                return;
            }
            Path newestPlan = null;
            long planTime = Long.MIN_VALUE;
            Path newestMember = null;
            long memberTime = Long.MIN_VALUE;
            try (DirectoryStream<Path> stream =
                    Files.newDirectoryStream(dir, "production_plan_multi_day_*.xlsx")) {
                for (Path p : stream) {
                    long t = Files.getLastModifiedTime(p).toMillis();
                    if (t >= planTime) {
                        planTime = t;
                        newestPlan = p;
                    }
                }
            }
            try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir, "member_schedule_*.xlsx")) {
                for (Path p : stream) {
                    long t = Files.getLastModifiedTime(p).toMillis();
                    if (t >= memberTime) {
                        memberTime = t;
                        newestMember = p;
                    }
                }
            }
            String planStr = newestPlan != null ? newestPlan.toString() : "";
            String memStr = newestMember != null ? newestMember.toString() : "";
            mainRunTabController.setStage2ArtifactPaths(planStr, memStr);
            if (!planStr.isEmpty() || !memStr.isEmpty()) {
                appendLog(
                        "[stage2-ui] "
                                + "\u6700\u65b0\u6210\u679c\u7269: production_plan="
                                + planStr
                                + " | member_schedule="
                                + memStr);
            }
        } catch (Exception ex) {
            appendLog(
                    "[stage2-ui] "
                            + "\u6210\u679c\u7269\u30d1\u30b9\u66f4\u65b0\u30a8\u30e9\u30fc: "
                            + ex.getMessage());
        }
    }

    void triggerPeekSheets() {
        peekSheetsAction();
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
        stripRemovedEnvVarRows(rows);
        if (rows.isEmpty()) {
            rows.add(new EnvVarRow());
        }
    }

    private static void stripRemovedEnvVarRows(ObservableList<EnvVarRow> rows) {
        if (rows == null) {
            return;
        }
        rows.removeIf(
                r -> {
                    String n = r.getName() != null ? r.getName().trim() : "";
                    return REMOVED_ENV_VAR_KEYS.contains(n);
                });
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
