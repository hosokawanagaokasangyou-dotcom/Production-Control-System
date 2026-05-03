package jp.co.pm.ai.desktop;

import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
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
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Tab;
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
import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;
import jp.co.pm.ai.desktop.ipc.IpcStdoutTap;

/**
 * Main window controller\uff08\u5f93\u6765\u306f {@link PmAiFxApp} \u5185\u8535\u3060\u3063\u305f\u696d\u52d9\u30ed\u30b8\u30c3\u30af\u3092\u5206\u96e2\uff09\u3002
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

    /**
     * Dropped from the env tab (defaults and session); not used in normal operation. Python still accepts
     * these if set in the real OS environment.
     */
    private static final Set<String> DROPPED_ENV_TAB_ROW_KEYS =
            Set.of(
                    "DEBUG_TASK_ID",
                    "TRACE_TEAM_ASSIGN_TASK_ID",
                    "EXCLUDE_RULES_TEST_E1234",
                    "EXCLUDE_RULES_TEST_E1234_ROW");

    private static final List<String> BOOTSTRAP_ORDER =
            List.of(
                    AppPaths.KEY_PM_AI_PYTHON,
                    AppPaths.KEY_PM_AI_REPO_ROOT,
                    AppPaths.KEY_PM_AI_OUTPUT_DIR,
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

    /** Keys in {@link #BOOTSTRAP_ORDER} for quick membership checks. */
    private static final Set<String> BOOTSTRAP_KEY_SET = Set.copyOf(BOOTSTRAP_ORDER);

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

    @FXML
    private ResultDispatchTableTabController resultDispatchTableTabController;

    @FXML
    private MachineCalendarTabController machineCalendarTabController;

    @FXML
    private DispatchInteractiveTabController dispatchInteractiveTabController;

    @FXML
    private PlanResultViewerTabController planResultViewerTabController;

    @FXML
    private OperatorCardTabController operatorCardTabController;

    @FXML
    private Tab mainShellTabRun;

    @FXML
    private Tab mainShellTabEnv;

    @FXML
    private Tab mainShellTabMasterSummary;

    @FXML
    private Tab mainShellTabPlanInput;

    @FXML
    private Tab mainShellTabStage1Preview;

    @FXML
    private Tab mainShellTabExcludeRules;

    @FXML
    private Tab mainShellTabActualsStatus;

    @FXML
    private Tab mainShellTabResultDispatch;

    @FXML
    private Tab mainShellTabMachineCalendarJson;

    @FXML
    private Tab mainShellTabDispatchInteractive;

    @FXML
    private Tab mainShellTabPlanResultViewer;

    @FXML
    private Tab mainShellTabOperatorCard;

    private ObservableList<EnvVarRow> envRows;
    private final AtomicBoolean runLock = new AtomicBoolean(false);

    /** Non-null while a stage script is running; equals {@link #STAGE1} or {@link #STAGE2}. */
    private volatile String activeRunStageScript;
    private final AtomicBoolean suppressEnvSessionPersistence = new AtomicBoolean(false);
    private final PauseTransition uiEnvSaveDebounce = new PauseTransition(Duration.millis(400));
    /** Assigned in {@link #installUiEnvAutoSave()} for {@link #resetEnvRowsToDefaults()}. */
    private Runnable uiEnvPersistSchedule;
    private final AtomicBoolean envResetInProgress = new AtomicBoolean(false);

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
            planResultViewerTabController.bindShell(this);

            operatorCardTabController.bindShell(this);

        mainRunTabController
                .getWorkbookField()
                .setPromptText(
                        "\u4efb\u610f\u3002\u7a7a\u6b04\u306e\u3068\u304d\u306f\u6bb5\u968e1/2\u5b9f\u884c\u6642\u306b\u3001\u74b0\u5883\u5909\u6570\u3068\u65e2\u5b9a\u306e\u30d6\u30fc\u30c8\u30b9\u30c8\u30e9\u30c3\u30d7\u3067\u30bf\u30b9\u30af\u5165\u529b\u30d6\u30c3\u30af\u306e\u30d1\u30b9\u304c\u6c7a\u307e\u308a\u307e\u3059\u3002"
                                + " PM_AI_* \u304c\u901a\u5e38\u904b\u7528\u306e\u8ef8\u3067\u3059\uff08\u30de\u30b9\u30bf\u8aad\u8fbc\u3084\u30d1\u30b9\u6307\u5b9a\u306e\u88dc\u52a9\u30d2\u30f3\u30c8\uff09\u3002");
        mainRunTabController
                .getWorkbookField()
                .setText(AppPaths.resolveTaskInputWorkbook(ui0).map(Path::toString).orElse(""));
        mainRunTabController
                .getPythonExeField()
                .setText(firstNonBlank(ui0.get(AppPaths.KEY_PM_AI_PYTHON), defaultOsPython()));
        mainRunTabController
                .getPythonExeField()
                .setPromptText("Python executable (\u672a\u8a2d\u5b9a\u6642\u306f\u74b0\u5883\u5909\u6570 PM_AI_PYTHON)");
        mainRunTabController
                .getScriptDirField()
                .setText(
                        firstNonBlank(
                                ui0.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                AppPaths.resolvePythonScriptDir(ui0).toString()));
        mainRunTabController
                .getScriptDirField()
                .setPromptText("code/python (\u672a\u8a2d\u5b9a\u6642\u306f\u74b0\u5883\u5909\u6570 PM_AI_CODE_PYTHON_DIR)");

        planInputTabController.bindShell(this);
        stage1PreviewTabController.bindShell(this);
        excludeRulesTabController.bindShell(this);
        actualsStatusTabController.bindShell(this);
        resultDispatchTableTabController.bindShell(this);
        machineCalendarTabController.bindShell(this);
        dispatchInteractiveTabController.bindShell(this);

        primaryStage.setMinWidth(640);
        primaryStage.setMinHeight(480);

            applyDesktopSession(DesktopSessionStateStore.load());
        } finally {
            suppressEnvSessionPersistence.set(false);
        }

        installUiEnvAutoSave();

        applyRepoFolderPathNormalization();

        primaryStage.setOnCloseRequest(e -> DesktopSessionStateStore.save(collectDesktopSession()));

        primaryStage.setOnShown(
                e -> {
                    primaryStage.toFront();
                    tabPane.getSelectionModel().selectFirst();
                });

        tabPane
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, prevTab, newTab) -> {
                            if (prevTab == mainShellTabRun && newTab != mainShellTabRun) {
                                DesktopSessionStateStore.save(collectDesktopSession());
                            }
                            updateClearTabFiltersButton(newTab);
                        });
        tabPane
                .getTabs()
                .addListener(
                        (ListChangeListener<Tab>)
                                c -> {
                                    if (!suppressEnvSessionPersistence.get()) {
                                        DesktopSessionStateStore.save(collectDesktopSession());
                                    }
                                });
        updateClearTabFiltersButton(tabPane.getSelectionModel().getSelectedItem());
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

    private boolean tabSupportsClearFilters(Tab t) {
        return t == mainShellTabEnv
                || t == mainShellTabPlanInput
                || t == mainShellTabStage1Preview
                || t == mainShellTabActualsStatus
                || t == mainShellTabResultDispatch
                || t == mainShellTabPlanResultViewer;
    }

    private void updateClearTabFiltersButton(Tab selected) {
        if (clearTabFiltersButton == null) {
            return;
        }
        clearTabFiltersButton.setDisable(!tabSupportsClearFilters(selected));
    }

    @FXML
    private void onClearTabFiltersAction() {
        Tab sel = tabPane.getSelectionModel().getSelectedItem();
        if (sel == mainShellTabEnv) {
            envTabController.clearColumnFiltersAndSort();
        } else if (sel == mainShellTabPlanInput) {
            planInputTabController.clearColumnFiltersAndSort();
        } else if (sel == mainShellTabStage1Preview) {
            stage1PreviewTabController.clearColumnFiltersAndSort();
        } else if (sel == mainShellTabActualsStatus) {
            actualsStatusTabController.clearColumnFiltersAndSort();
        } else if (sel == mainShellTabResultDispatch) {
            resultDispatchTableTabController.clearColumnFiltersAndSort();
        } else if (sel == mainShellTabPlanResultViewer) {
            planResultViewerTabController.clearColumnFiltersAndSort();
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
        mainRunTabController.restoreRunLogUiFromSession(
                s.mainRunLogFilter(), s.mainRunLogLines(), s.mainRunLogScroll());
        if (nonBlank(s.mainRunStage2ProductionPlan())
                || nonBlank(s.mainRunStage2MemberSchedule())) {
            mainRunTabController.setStage2ArtifactPaths(
                    nz(s.mainRunStage2ProductionPlan()),
                    nz(s.mainRunStage2MemberSchedule()));
        }
        mainRunTabController.applyStage2WriteExcelFromSession(s.mainRunStage2WriteExcel());
        mainRunTabController.applyStage2ResultBookFontFromSession(s.mainRunStage2ResultBookFont());
        applyWindowGeometry(s);
        applyMainShellTabOrder(s.mainShellTabOrder());
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
                mainRunTabController.snapshotLogFilterName(),
                mainRunTabController.snapshotPersistedLogLines(),
                mainRunTabController.snapshotLogScrollProportion(),
                mainRunTabController.snapshotStage2ProductionPlanPath(),
                mainRunTabController.snapshotStage2MemberSchedulePath(),
                mainRunTabController.snapshotStage2WriteExcel(),
                mainRunTabController.snapshotStage2ResultBookFont(),
                snapshotUiEnvRows(),
                snapshotMainShellTabOrder());
    }

    private MainShellTabId mainShellTabId(Tab t) {
        if (t == null) {
            return null;
        }
        if (t == mainShellTabRun) {
            return MainShellTabId.RUN;
        }
        if (t == mainShellTabEnv) {
            return MainShellTabId.ENV;
        }
        if (t == mainShellTabMasterSummary) {
            return MainShellTabId.MASTER_SUMMARY;
        }
        if (t == mainShellTabPlanInput) {
            return MainShellTabId.PLAN_INPUT;
        }
        if (t == mainShellTabStage1Preview) {
            return MainShellTabId.STAGE1_PREVIEW;
        }
        if (t == mainShellTabExcludeRules) {
            return MainShellTabId.EXCLUDE_RULES;
        }
        if (t == mainShellTabActualsStatus) {
            return MainShellTabId.ACTUALS_STATUS;
        }
        if (t == mainShellTabResultDispatch) {
            return MainShellTabId.RESULT_DISPATCH;
        }
        if (t == mainShellTabMachineCalendarJson) {
            return MainShellTabId.MACHINE_CALENDAR_JSON;
        }
        if (t == mainShellTabDispatchInteractive) {
            return MainShellTabId.DISPATCH_INTERACTIVE;
        }
        if (t == mainShellTabPlanResultViewer) {
            return MainShellTabId.PLAN_RESULT_VIEWER;
        }
        if (t == mainShellTabOperatorCard) {
            return MainShellTabId.OPERATOR_CARD;
        }
        return null;
    }

    private Tab mainShellTabFor(MainShellTabId id) {
        if (id == null) {
            return null;
        }
        return switch (id) {
            case RUN -> mainShellTabRun;
            case ENV -> mainShellTabEnv;
            case MASTER_SUMMARY -> mainShellTabMasterSummary;
            case PLAN_INPUT -> mainShellTabPlanInput;
            case STAGE1_PREVIEW -> mainShellTabStage1Preview;
            case EXCLUDE_RULES -> mainShellTabExcludeRules;
            case ACTUALS_STATUS -> mainShellTabActualsStatus;
            case RESULT_DISPATCH -> mainShellTabResultDispatch;
            case MACHINE_CALENDAR_JSON -> mainShellTabMachineCalendarJson;
            case DISPATCH_INTERACTIVE -> mainShellTabDispatchInteractive;
            case PLAN_RESULT_VIEWER -> mainShellTabPlanResultViewer;
            case OPERATOR_CARD -> mainShellTabOperatorCard;
        };
    }

    private List<String> snapshotMainShellTabOrder() {
        if (tabPane == null) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (Tab t : tabPane.getTabs()) {
            MainShellTabId id = mainShellTabId(t);
            if (id != null) {
                out.add(id.key());
            }
        }
        return List.copyOf(out);
    }

    private void applyMainShellTabOrder(List<String> orderKeys) {
        if (tabPane == null || orderKeys == null || orderKeys.isEmpty()) {
            return;
        }
        ObservableList<Tab> tabs = tabPane.getTabs();
        if (tabs.isEmpty()) {
            return;
        }
        ArrayList<Tab> newOrder = new ArrayList<>();
        HashSet<Tab> seen = new HashSet<>();
        for (String key : orderKeys) {
            MainShellTabId id = MainShellTabId.fromKey(key);
            if (id == null) {
                continue;
            }
            Tab t = mainShellTabFor(id);
            if (t != null && seen.add(t)) {
                newOrder.add(t);
            }
        }
        for (MainShellTabId id : MainShellTabId.values()) {
            Tab t = mainShellTabFor(id);
            if (t != null && seen.add(t)) {
                newOrder.add(t);
            }
        }
        if (newOrder.size() != tabs.size()) {
            return;
        }
        tabs.setAll(newOrder);
    }

    private static boolean omitEnvRowKey(String name) {
        String k = name != null ? name.trim() : "";
        return REMOVED_ENV_VAR_KEYS.contains(k) || DROPPED_ENV_TAB_ROW_KEYS.contains(k);
    }

    private List<UiEnvRowSnapshot> snapshotUiEnvRows() {
        if (envRows == null) {
            return List.of();
        }
        List<UiEnvRowSnapshot> out = new ArrayList<>(envRows.size());
        for (EnvVarRow r : envRows) {
            String key = nz(r.getName());
            if (omitEnvRowKey(key)) {
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
            if (omitEnvRowKey(nm)) {
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
        mergeMissingBootstrapEnvRows();
    }

    /**
     * Session snapshots may omit rows that were added in a later app version. Rebuild env rows so
     * {@link #BOOTSTRAP_ORDER} keys still appear (same order as {@link #populateEnvRows}).
     */
    private void mergeMissingBootstrapEnvRows() {
        if (envRows == null) {
            return;
        }
        Map<String, String> ui = collectUiEnv();
        LinkedHashMap<String, EnvVarRow> byKey = new LinkedHashMap<>();
        for (EnvVarRow r : envRows) {
            String name = r.getName() != null ? r.getName().trim() : "";
            if (name.isEmpty() || omitEnvRowKey(name)) {
                continue;
            }
            byKey.putIfAbsent(name, r);
        }
        ArrayList<EnvVarRow> out = new ArrayList<>(envRows.size() + BOOTSTRAP_ORDER.size());
        for (String k : BOOTSTRAP_ORDER) {
            EnvVarRow row = byKey.get(k);
            if (row != null) {
                maybeFillEmptyBootstrap(row, k, ui);
                out.add(row);
            } else {
                out.add(newBootstrapRow(k, ui));
            }
        }
        HashSet<String> seen = new HashSet<>(BOOTSTRAP_KEY_SET);
        List<EnvVarRow> trailingEmpty = new ArrayList<>();
        for (EnvVarRow r : envRows) {
            String name = r.getName() != null ? r.getName().trim() : "";
            if (omitEnvRowKey(name)) {
                continue;
            }
            if (name.isEmpty()) {
                trailingEmpty.add(r);
                continue;
            }
            if (BOOTSTRAP_KEY_SET.contains(name)) {
                continue;
            }
            if (seen.contains(name)) {
                continue;
            }
            out.add(r);
            seen.add(name);
        }
        out.addAll(trailingEmpty);
        envRows.setAll(out);
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
        this.uiEnvPersistSchedule = schedule;
        envRows.addListener(
                (ListChangeListener<EnvVarRow>)
                        c -> {
                            while (c.next()) {
                                if (envResetInProgress.get()) {
                                    continue;
                                }
                                if (c.wasAdded()) {
                                    for (EnvVarRow row : c.getAddedSubList()) {
                                        hookEnvRowForAutoSave(row, schedule);
                                    }
                                }
                            }
                            if (!envResetInProgress.get()) {
                                schedule.run();
                            }
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

    /**
     * Resets the env-var table to bundled defaults ({@link UiRefEnvDefaults}) and reapplies bootstrap fills.
     * Shows a confirmation dialog first.
     */
    void confirmAndResetEnvRowsToDefaults() {
        Alert alert = new Alert(AlertType.CONFIRMATION);
        alert.initOwner(primaryStage);
        alert.setTitle("\u74b0\u5883\u5909\u6570\u3092\u521d\u671f\u5024\u306b\u623b\u3059");
        alert.setHeaderText(null);
        alert.setContentText(
                "ui_ref_env_defaults.json \u306e\u65e2\u5b9a\u884c\u306b\u623b\u3057\u307e\u3059\u3002"
                        + "\u672a\u4fdd\u5b58\u306e\u7de8\u96c6\u3068\u3001\u30bb\u30c3\u30b7\u30e7\u30f3\u306b\u4fdd\u5b58\u3057\u3066\u3044\u305f\u5404\u30bf\u30d6\u306e\u5024\uff08Python \u30d1\u30b9\u7b49\uff09\u3082\u5931\u308f\u308c\u307e\u3059\u3002"
                        + "\u7d9a\u884c\u3057\u307e\u3059\u304b\uff1f");
        Optional<ButtonType> ans = alert.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        resetEnvRowsToDefaults();
    }

    private void resetEnvRowsToDefaults() {
        if (envRows == null) {
            return;
        }
        suppressEnvSessionPersistence.set(true);
        envResetInProgress.set(true);
        try {
            populateEnvRows(envRows);
            Runnable sched = uiEnvPersistSchedule;
            if (sched != null) {
                for (EnvVarRow row : envRows) {
                    hookEnvRowForAutoSave(row, sched);
                }
            }
            Map<String, String> ui = collectUiEnv();
            mainRunTabController
                    .getPythonExeField()
                    .setText(
                            firstNonBlank(
                                    ui.get(AppPaths.KEY_PM_AI_PYTHON), defaultOsPython()));
            mainRunTabController
                    .getScriptDirField()
                    .setText(
                            firstNonBlank(
                                    ui.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                    AppPaths.resolvePythonScriptDir(ui).toString()));
        } finally {
            envResetInProgress.set(false);
            suppressEnvSessionPersistence.set(false);
        }
        applyRepoFolderPathNormalization();
        DesktopSessionStateStore.save(collectDesktopSession());
        uiEnvSaveDebounce.stop();
    }

    void appendBootMessage() {
        mainRunTabController.appendLog(
                "[boot] PYTHONUTF8=1 PYTHONIOENCODING=utf-8 for child process.", false);
        Platform.runLater(
                () -> {
                    mainRunTabController.flushPendingSessionLogScroll();
                    Platform.runLater(mainRunTabController::flushPendingSessionLogScroll);
                });
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
                                + " \u304c\u7121\u3044\u305f\u3081\u3001\u74b0\u5883\u5909\u6570\u30bf\u30d6\u306e\u5024\u306f\u672a\u66f4\u65b0\u306e\u307e\u307e\u3067\u3059\u3002"
                                + " \u6bb5\u968e1\u304c\u914d\u53f0\u9664\u5916\u30eb\u30fc\u30eb JSON \u3092\u751f\u6210\u3057\u3066\u3044\u308b\u304b\u3001cwd/json \u306e\u5834\u6240\u304c\u4e00\u81f4\u3057\u3066\u3044\u308b\u304b\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002");
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
                    "[env] PM_AI_EXCLUDE_RULES_JSON \u884c\u304c\u898b\u3064\u304b\u3089\u306a\u3044\u305f\u3081\u672a\u66f4\u65b0\u306e\u307e\u307e\u3067\u3059\u3002");
        } catch (Exception ex) {
            appendLog("[env] PM_AI_EXCLUDE_RULES_JSON \u66f4\u65b0\u306b\u5931\u6557: " + ex.getMessage());
        }
    }

    private void runStage(String script) {
        if (!runLock.compareAndSet(false, true)) {
            appendLog("[busy] already running (single flight).");
            return;
        }
        activeRunStageScript = script;
        applyRunTabGating();
        Map<String, String> uiRun = collectUiEnv();
        if (STAGE2.equals(script)) {
            uiRun.put(
                    AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL,
                    mainRunTabController.snapshotStage2WriteExcel() ? "1" : "0");
            String resultFont = mainRunTabController.snapshotStage2ResultBookFont();
            if (resultFont != null && !resultFont.isBlank()) {
                uiRun.put(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT, resultFont.trim());
            } else {
                uiRun.remove(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT);
            }
        }
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
        mainRunTabController.getStatusLabel().setText("\u5b9f\u884c\u4e2d\u2026");

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
                            activeRunStageScript = null;
                            javafx.application.Platform.runLater(
                                    () -> {
                                        applyRunTabGating();
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
                                                showStageCompletionDialog(
                                                        "\u6bb5\u968e1 \u5b8c\u4e86",
                                                        "\u6bb5\u968e1 \u306e\u51e6\u7406\u304c\u6b63\u5e38\u7d42\u4e86\u3057\u307e\u3057\u305f\u3002");
                                            }
                                            if (STAGE2.equals(script) && c == 0) {
                                                refreshStage2OutputArtifacts();
                                                showStageCompletionDialog(
                                                        "\u6bb5\u968e2 \u5b8c\u4e86",
                                                        "\u6bb5\u968e2 \u306e\u51e6\u7406\u304c\u6b63\u5e38\u7d42\u4e86\u3057\u307e\u3057\u305f\u3002");
                                            }
                                        }
                                    });
                        });
    }

    /**
     * \u6bb5\u968e1\u5b9f\u884c\u4e2d\u306f\u74b0\u5883\u5909\u6570\u30bf\u30d6\u3068\u914d\u53f0\u8a08\u753b\u5165\u529b\u30bf\u30d6\u3092\u7121\u52b9\u5316\u3057\u3001\u6bb5\u968e2\u306b\u6e21\u3059\u524d\u63d0\u304c\u9014\u4e2d\u3067\u5d29\u308c\u308b\u64cd\u4f5c\u3092\u9632\u3050\u3002
     * \u6bb5\u968e2\u5b9f\u884c\u4e2d\u306f\u74b0\u5883\u5909\u6570\u30bf\u30d6\u3068\u6bb5\u968e1\u30d7\u30ec\u30d3\u30e5\u30fc\u30bf\u30d6\u3092\u7121\u52b9\u5316\u3057\u3001\u6bb5\u968e1\u306e\u7d50\u679c\u3068\u98df\u3044\u9055\u3046\u64cd\u4f5c\u3092\u9632\u3050\u3002
     */
    private void applyRunTabGating() {
        if (tabPane == null) {
            return;
        }
        ObservableList<Tab> tabs = tabPane.getTabs();
        if (tabs.isEmpty()) {
            return;
        }
        String script = activeRunStageScript;
        boolean stage1Running = STAGE1.equals(script);
        boolean stage2Running = STAGE2.equals(script);
        for (Tab t : tabs) {
            boolean disable =
                    stage1Running
                            && (t == mainShellTabEnv || t == mainShellTabPlanInput)
                            || stage2Running
                                    && (t == mainShellTabEnv || t == mainShellTabStage1Preview);
            t.setDisable(disable);
        }
        Tab sel = tabPane.getSelectionModel().getSelectedItem();
        if (stage1Running && (sel == mainShellTabEnv || sel == mainShellTabPlanInput)) {
            tabPane.getSelectionModel().select(mainShellTabRun);
        } else if (stage2Running && (sel == mainShellTabEnv || sel == mainShellTabStage1Preview)) {
            tabPane.getSelectionModel().select(mainShellTabRun);
        }
    }

    private void showStageCompletionDialog(String title, String contentText) {
        Alert alert = new Alert(AlertType.INFORMATION);
        alert.initOwner(primaryStage);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(contentText);
        alert.showAndWait();
    }

    private static String exitCodeLegend(int code) {
        return "exit="
                + code
                + " \uff080=OK / 1=error / 2=fatal / 3=PlanningValidationError / 9=cancel\uff09";
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
     * the \u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b tab are applied so that stage-2 uses the
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
            if (k.isEmpty() || k.startsWith("#") || omitEnvRowKey(k)) {
                continue;
            }
            m.put(k, row.getValue() != null ? row.getValue() : "");
        }
        return m;
    }

    /**
     * When folder env values ({@code PM_AI_*}) can be read as under the current repo, rewrites them to
     * canonical absolute paths and saves the session (also syncs main-run script dir for
     * {@code PM_AI_CODE_PYTHON_DIR}).
     */
    private void applyRepoFolderPathNormalization() {
        if (envRows == null) {
            return;
        }
        suppressEnvSessionPersistence.set(true);
        try {
            Map<String, String> ui = collectUiEnv();
            Map<String, String> overrides = AppPaths.normalizedFolderEnvOverrides(ui);
            if (overrides.isEmpty()) {
                return;
            }
            for (EnvVarRow row : envRows) {
                String k = nz(row.getName());
                if (overrides.containsKey(k)) {
                    row.setValue(overrides.get(k));
                }
            }
            String cp = overrides.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR);
            if (cp != null && mainRunTabController != null) {
                mainRunTabController.getScriptDirField().setText(cp);
            }
        } finally {
            suppressEnvSessionPersistence.set(false);
        }
        DesktopSessionStateStore.save(collectDesktopSession());
    }

    /** After {@code machine_calendar_blocks.json} is (re)written, refresh interactive dispatch block overlay. */
    public void notifyMachineCalendarJsonUpdated() {
        if (dispatchInteractiveTabController != null) {
            dispatchInteractiveTabController.refreshCalendarFromSharedJsonFile();
        }
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
            Path newestPlan = newestInOutputDir(dir, "production_plan_multi_day_*.xlsx");
            if (newestPlan == null) {
                newestPlan = newestInOutputDir(dir, "production_plan_multi_day_*.json");
            }
            Path newestMember = newestInOutputDir(dir, "member_schedule_*.xlsx");
            if (newestMember == null) {
                newestMember = newestInOutputDir(dir, "member_schedule_*.json");
            }
            String planStr = newestPlan != null ? newestPlan.toString() : "";
            String memStr = newestMember != null ? newestMember.toString() : "";
            mainRunTabController.setStage2ArtifactPaths(planStr, memStr);
            planResultViewerTabController.tryAutofillJsonFromStage2Xlsx(planStr, memStr);
            operatorCardTabController.tryAutofillMemberJsonFromStage2(memStr);
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
                            + "\u6210\u679c\u30d1\u30b9\u66f4\u65b0\u30a8\u30e9\u30fc: "
                            + ex.getMessage());
        }
    }

    private static Path newestInOutputDir(Path dir, String glob) throws java.io.IOException {
        Path best = null;
        long bestTime = Long.MIN_VALUE;
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir, glob)) {
            for (Path p : stream) {
                long t = Files.getLastModifiedTime(p).toMillis();
                if (t >= bestTime) {
                    bestTime = t;
                    best = p;
                }
            }
        }
        return best;
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
                    return omitEnvRowKey(n);
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
            case AppPaths.KEY_PM_AI_OUTPUT_DIR -> r.setValue(AppPaths.resolveDefaultOutputDir(ui).toString());
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
            case AppPaths.KEY_PM_AI_OUTPUT_DIR -> r.setValue(AppPaths.resolveDefaultOutputDir(ui).toString());
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
