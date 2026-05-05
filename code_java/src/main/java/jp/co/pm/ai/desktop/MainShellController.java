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
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

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
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
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
import jp.co.pm.ai.desktop.config.PushButtonCssEmitter;
import jp.co.pm.ai.desktop.config.PushButtonDesignPrefs;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.config.EnvVarDocs;
import jp.co.pm.ai.desktop.config.UiEnvRowSnapshot;
import jp.co.pm.ai.desktop.config.UiRefEnvDefaults;
import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;
import jp.co.pm.ai.desktop.ipc.IpcStdoutTap;

/**
 * Main window controller（従来は {@link PmAiFxApp} 内蔵だった業務ロジックを分離）。
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
                    AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
                    AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON,
                    AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH);

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
    private HBox shellStageProgressBox;

    @FXML
    private Label shellStageProgressLabel;

    @FXML
    private ProgressBar shellStageProgressBar;

    @FXML
    private ProgressIndicator shellStageBusyIndicator;

    @FXML
    private Button shellStageCancelButton;

    @FXML
    private Region toolbarGrowSpacer;

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
    private SpecialRulesTabController specialRulesTabController;

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
    private EquipmentGanttGraphicTabController equipmentGanttGraphicTabController;

    @FXML
    private GanttPersonBadgeDesignTabController ganttPersonBadgeDesignTabController;

    @FXML
    private UiBadgeDesignTabController uiBadgeDesignTabController;

    @FXML
    private PushButtonDesignTabController pushButtonDesignTabController;

    @FXML
    private OperatorCardTabController operatorCardTabController;

    @FXML
    private Tab mainShellTabRun;

    @FXML
    private Tab mainShellTabUiBadgeDesign;

    @FXML
    private Tab mainShellTabPushButtonDesign;

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
    private Tab mainShellTabSpecialRules;

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
    private Tab mainShellTabEquipmentGanttGraphic;

    @FXML
    private Tab mainShellTabGanttPersonBadgeDesign;

    @FXML
    private Tab mainShellTabOperatorCard;

    private ObservableList<EnvVarRow> envRows;

    private final AtomicBoolean runLock = new AtomicBoolean(false);

    /** Non-null while a stage script is running; equals {@link #STAGE1} or {@link #STAGE2}. */
    private volatile String activeRunStageScript;

    /** Python child process while stage 1/2 is running; cleared on completion or interrupt. */
    private final AtomicReference<Process> activeStageChildProcess = new AtomicReference<>();

    /** {@link #childEnvForPython(Map)} の直近結果（実行タブのキャッシュ表示・ログ用）。 */
    private NetworkSourceDirResolver.Result lastNetworkSourceResolution;

    /**
     * 起動時プローブでソースフォルダが一覧不可なら {@code true}。{@link NetworkSourceDirResolver#resolve(Map, boolean, boolean)}
     * でネットワーク側の一覧を省略する。
     */
    private volatile boolean startupSkipTaskInputSourceDirListing;

    private volatile boolean startupSkipActualDetailSourceDirListing;

    private final AtomicBoolean suppressEnvSessionPersistence = new AtomicBoolean(false);
    private final PauseTransition uiEnvSaveDebounce = new PauseTransition(Duration.millis(400));
    /** Assigned in {@link #installUiEnvAutoSave()} for {@link #resetEnvRowsToDefaults()}. */
    private Runnable uiEnvPersistSchedule;
    private final AtomicBoolean envResetInProgress = new AtomicBoolean(false);

    private DesktopTheme pendingTheme = DesktopTheme.LIGHT;

    private static final String PM_AI_DESKTOP_CSS =
            Objects.requireNonNull(
                            PmAiFxApp.class.getResource("/jp/co/pm/ai/desktop/css/pm-ai-desktop.css"),
                            "pm-ai-desktop.css")
                    .toExternalForm();

    /** Child windows (e.g. dispatch trial log) that mirror the toolbar {@link DesktopTheme}. */
    private final List<Scene> themeTrackedSecondaryScenes = new CopyOnWriteArrayList<>();

    /** Primary shell scene (push-button overridesなどで参照)。 */
    private Scene primaryScene;

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
            equipmentGanttGraphicTabController.bindShell(this);
            if (ganttPersonBadgeDesignTabController != null) {
                ganttPersonBadgeDesignTabController.bindShell(this);
            }
            if (uiBadgeDesignTabController != null) {
                uiBadgeDesignTabController.bindShell(this);
            }
            if (pushButtonDesignTabController != null) {
                pushButtonDesignTabController.bindShell(this);
            }

            operatorCardTabController.bindShell(this);

        mainRunTabController
                .getWorkbookField()
                .setPromptText(
                        "任意。空欄のときは段階1/2実行時に、環境変数と既定のブートストラップでタスク入力ブックのパスが決まります。"
                                + " PM_AI_* が通常運用の軸です（マスタ読込やパス指定の補助ヒント）。");
        mainRunTabController
                .getWorkbookField()
                .setText(AppPaths.resolveTaskInputWorkbook(ui0).map(Path::toString).orElse(""));
        mainRunTabController
                .getPythonExeField()
                .setText(firstNonBlank(ui0.get(AppPaths.KEY_PM_AI_PYTHON), defaultOsPython()));
        mainRunTabController
                .getPythonExeField()
                .setPromptText("Python executable (未設定時は環境変数 PM_AI_PYTHON)");
        mainRunTabController
                .getScriptDirField()
                .setText(
                        firstNonBlank(
                                ui0.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                AppPaths.resolvePythonScriptDir(ui0).toString()));
        mainRunTabController
                .getScriptDirField()
                .setPromptText("code/python (未設定時は環境変数 PM_AI_CODE_PYTHON_DIR)");

        planInputTabController.bindShell(this);
        stage1PreviewTabController.bindShell(this);
        excludeRulesTabController.bindShell(this);
        specialRulesTabController.bindShell(this);
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

        if (toolbarGrowSpacer != null) {
            HBox.setHgrow(toolbarGrowSpacer, Priority.ALWAYS);
        }

        installUiEnvAutoSave();

        applyRepoFolderPathNormalization();

        probeNetworkSourceDirsAtStartup();

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
        this.primaryScene = scene;
        if (themeCombo == null) {
            if (pushButtonDesignTabController != null) {
                pushButtonDesignTabController.installStylesheetWhenSceneReady();
            }
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
                                refreshThemeTrackedSecondaryScenes();
                            }
                            mainRunTabController.refreshLogThemeCells();
                            equipmentGanttGraphicTabController.refreshGraphicForTheme();
                            refreshPushButtonStylesheet();
                        });
        Platform.runLater(mainRunTabController::refreshLogThemeCells);
        if (pushButtonDesignTabController != null) {
            pushButtonDesignTabController.installStylesheetWhenSceneReady();
        }
    }

    /** Theme shown in toolbar (for components that need dark/light tint hints). */
    DesktopTheme currentDesktopTheme() {
        if (themeCombo != null && themeCombo.getValue() != null) {
            return themeCombo.getValue();
        }
        return pendingTheme != null ? pendingTheme : DesktopTheme.LIGHT;
    }

    /**
     * Loads {@code pm-ai-desktop.css} and the current theme overlay onto a secondary {@link Scene},
     * and reapplies the palette when the user changes the theme until {@link #unregisterThemeTrackedScene}.
     */
    public void registerThemeTrackedScene(Scene scene) {
        if (scene == null) {
            return;
        }
        if (!scene.getStylesheets().contains(PM_AI_DESKTOP_CSS)) {
            scene.getStylesheets().add(PM_AI_DESKTOP_CSS);
        }
        currentDesktopTheme().applyTo(scene);
        if (!themeTrackedSecondaryScenes.contains(scene)) {
            themeTrackedSecondaryScenes.add(scene);
        }
    }

    public void unregisterThemeTrackedScene(Scene scene) {
        themeTrackedSecondaryScenes.remove(scene);
    }

    private void refreshThemeTrackedSecondaryScenes() {
        DesktopTheme t = currentDesktopTheme();
        for (Scene s : themeTrackedSecondaryScenes) {
            t.applyTo(s);
        }
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
        equipmentGanttGraphicTabController.applyEquipmentGanttSession(s);
        if (ganttPersonBadgeDesignTabController != null) {
            ganttPersonBadgeDesignTabController.applyPersonBadgeDesignSession(s);
        }
        if (uiBadgeDesignTabController != null) {
            uiBadgeDesignTabController.applyUiBadgeSession(s);
        }
        if (pushButtonDesignTabController != null) {
            pushButtonDesignTabController.applyPushButtonSession(s);
        }
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
        if (ganttPersonBadgeDesignTabController != null) {
            ganttPersonBadgeDesignTabController.flushBadgeEditsBeforeSnapshot();
        }
        if (uiBadgeDesignTabController != null) {
            uiBadgeDesignTabController.flushEditsBeforeSnapshot();
        }
        if (pushButtonDesignTabController != null) {
            pushButtonDesignTabController.flushEditsBeforeSnapshot();
        }
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
                snapshotMainShellTabOrder(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttZoomPercent(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttDateColWidth(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttMachineColWidth(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttProcessColWidth(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttBarFontFamily(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttBarFontPercent(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttRowHeightPercent(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttHeaderHeightPercent(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttSlotWidthPercent(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttShiftWheelHScrollPercent(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeEnabled(),
                snapshotPersonBadgeFontFamily(),
                snapshotPersonBadgeFontPercent(),
                snapshotPersonBadgeFillHex(),
                snapshotPersonBadgeTextHex(),
                snapshotPersonBadgeStrokeHex(),
                snapshotPersonBadgeStrokeWidth(),
                snapshotPersonBadgeCornerRadius(),
                snapshotPersonBadgePill(),
                snapshotPersonBadgeGlowColorHex(),
                snapshotPersonBadgeGlowRadius(),
                snapshotPersonBadgeGlowSpread(),
                snapshotPersonBadgeStylesByLabel(),
                snapshotPersonBadgeStylesByMemberKey(),
                uiBadgeDesignTabController != null
                        ? uiBadgeDesignTabController.snapshotStage1NetworkCacheBadgeLabel()
                        : "",
                uiBadgeDesignTabController != null
                        ? uiBadgeDesignTabController.snapshotStage1NetworkCacheBadgeStyle()
                        : PersonBadgeStyle.networkSourceCacheBadgeDefault(),
                pushButtonDesignTabController != null
                        ? pushButtonDesignTabController.snapshotPrefs()
                        : PushButtonDesignPrefs.inactiveDefaults());
    }

    /** 設備ガントのプレビュー用に、バッジ「既定」スタイルを返す。 */
    public PersonBadgeStyle currentPersonBadgeStyleForGantt() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.previewStyleForGantt()
                : PersonBadgeStyle.defaultStyle();
    }

    /** バッジ表示文字列ごとの見た目（担当者別設定を反映）。 */
    public java.util.function.Function<String, PersonBadgeStyle> personBadgeStyleResolverForGantt() {
        if (ganttPersonBadgeDesignTabController != null) {
            return ganttPersonBadgeDesignTabController::resolveStyleForBadgeLabel;
        }
        return (String __) -> PersonBadgeStyle.defaultStyle();
    }

    /** 設備ガントで検出したバッジキーをデザインタブの候補に追加する。 */
    public void refreshEquipmentGanttObservedBadgeLabels(java.util.Collection<String> labels) {
        if (ganttPersonBadgeDesignTabController != null) {
            ganttPersonBadgeDesignTabController.mergeObservedBadgeLabels(labels);
        }
    }

    /**
     * planning_core と同様に {@code master.xls(x/m)} を解決する。ファイルが無いときは {@code null}。
     */
    public Path resolveMasterWorkbookIfPresent() {
        Path p =
                AppPaths.resolveMasterWorkbookPathResolved(
                        collectUiEnv(), nz(mainRunTabController.getWorkbookField().getText()));
        return Files.isRegularFile(p) ? p.toAbsolutePath().normalize() : null;
    }

    /** バッジデザイン変更後に設備ガント（グラフィック）のみ再描画する。 */
    public void refreshEquipmentGanttGraphicForBadgeChange() {
        if (equipmentGanttGraphicTabController != null) {
            equipmentGanttGraphicTabController.refreshGraphicForPersonBadge();
        }
    }

    private String snapshotPersonBadgeFontFamily() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeFontFamily()
                : "";
    }

    private double snapshotPersonBadgeFontPercent() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeFontPercent()
                : 0d;
    }

    private String snapshotPersonBadgeFillHex() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeFillHex()
                : "";
    }

    private String snapshotPersonBadgeTextHex() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeTextHex()
                : "";
    }

    private String snapshotPersonBadgeStrokeHex() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeStrokeHex()
                : "";
    }

    private double snapshotPersonBadgeStrokeWidth() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeStrokeWidth()
                : -1d;
    }

    private double snapshotPersonBadgeCornerRadius() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeCornerRadius()
                : -1d;
    }

    private boolean snapshotPersonBadgePill() {
        return ganttPersonBadgeDesignTabController != null
                && ganttPersonBadgeDesignTabController.snapshotPersonBadgePill();
    }

    private String snapshotPersonBadgeGlowColorHex() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeGlowColorHex()
                : "";
    }

    private double snapshotPersonBadgeGlowRadius() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeGlowRadius()
                : -1d;
    }

    private double snapshotPersonBadgeGlowSpread() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeGlowSpread()
                : -1d;
    }

    private java.util.Map<String, PersonBadgeStyle> snapshotPersonBadgeStylesByLabel() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeStylesByLabel()
                : java.util.Map.of();
    }

    private java.util.Map<String, PersonBadgeStyle> snapshotPersonBadgeStylesByMemberKey() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeStylesByMemberKey()
                : java.util.Map.of();
    }

    /** 現在の UI 状態を直ちに session-state.json に保存する（タブ内の微調整の自動保存用）。 */
    public void persistDesktopSessionNow() {
        DesktopSessionStateStore.save(collectDesktopSession());
    }

    /** プッシュボタンのユーザー CSS をメインシーンに適用し直す（テーマ変更後も最後尾で上書き）。 */
    public void refreshPushButtonStylesheet() {
        if (primaryScene == null || pushButtonDesignTabController == null) {
            return;
        }
        PushButtonCssEmitter.applyToScene(primaryScene, pushButtonDesignTabController.snapshotPrefs());
    }

    private MainShellTabId mainShellTabId(Tab t) {
        if (t == null) {
            return null;
        }
        if (t == mainShellTabRun) {
            return MainShellTabId.RUN;
        }
        if (t == mainShellTabUiBadgeDesign) {
            return MainShellTabId.UI_BADGE_DESIGN;
        }
        if (t == mainShellTabPushButtonDesign) {
            return MainShellTabId.PUSH_BUTTON_DESIGN;
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
        if (t == mainShellTabSpecialRules) {
            return MainShellTabId.SPECIAL_RULES;
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
        if (t == mainShellTabEquipmentGanttGraphic) {
            return MainShellTabId.EQUIPMENT_GANTT_GRAPHIC;
        }
        if (t == mainShellTabGanttPersonBadgeDesign) {
            return MainShellTabId.GANTT_PERSON_BADGE_DESIGN;
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
            case UI_BADGE_DESIGN -> mainShellTabUiBadgeDesign;
            case PUSH_BUTTON_DESIGN -> mainShellTabPushButtonDesign;
            case ENV -> mainShellTabEnv;
            case MASTER_SUMMARY -> mainShellTabMasterSummary;
            case PLAN_INPUT -> mainShellTabPlanInput;
            case STAGE1_PREVIEW -> mainShellTabStage1Preview;
            case EXCLUDE_RULES -> mainShellTabExcludeRules;
            case SPECIAL_RULES -> mainShellTabSpecialRules;
            case ACTUALS_STATUS -> mainShellTabActualsStatus;
            case RESULT_DISPATCH -> mainShellTabResultDispatch;
            case MACHINE_CALENDAR_JSON -> mainShellTabMachineCalendarJson;
            case DISPATCH_INTERACTIVE -> mainShellTabDispatchInteractive;
            case PLAN_RESULT_VIEWER -> mainShellTabPlanResultViewer;
            case EQUIPMENT_GANTT_GRAPHIC -> mainShellTabEquipmentGanttGraphic;
            case GANTT_PERSON_BADGE_DESIGN -> mainShellTabGanttPersonBadgeDesign;
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

    /**
     * ui_ref_env_defaults.json と {@link #BOOTSTRAP_ORDER} にあるが、表に行が無い変数を同じ並びで追加する（既存行の値は保持）。
     */
    void addMissingReferenceEnvRows() {
        mergeMissingUiRefEnvRows();
    }

    /**
     * Same key order as {@link #populateEnvRows(ObservableList)}; inserts only keys not yet present (non-empty name).
     */
    private void mergeMissingUiRefEnvRows() {
        if (envRows == null) {
            return;
        }
        LinkedHashMap<String, EnvVarRow> sheetTemplates = new LinkedHashMap<>();
        for (WorkbookEnvSheetReader.RowEntry e : UiRefEnvDefaults.loadOrEmpty()) {
            EnvVarRow row = new EnvVarRow();
            row.setName(e.key());
            row.setValue(e.value() != null ? e.value() : "");
            row.setDescription(EnvVarDocs.mergeDescriptions(e.description(), e.key()));
            sheetTemplates.put(e.key(), row);
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
        List<String> refOrder = new ArrayList<>(BOOTSTRAP_ORDER.size() + sheetTemplates.size());
        for (String k : BOOTSTRAP_ORDER) {
            refOrder.add(k);
        }
        for (String k : sheetTemplates.keySet()) {
            if (!BOOTSTRAP_KEY_SET.contains(k)) {
                refOrder.add(k);
            }
        }
        ArrayList<EnvVarRow> out = new ArrayList<>(envRows.size() + refOrder.size());
        HashSet<String> placed = new HashSet<>();
        for (String k : refOrder) {
            EnvVarRow existing = byKey.get(k);
            if (existing != null) {
                if (BOOTSTRAP_KEY_SET.contains(k)) {
                    maybeFillEmptyBootstrap(existing, k, ui);
                }
                out.add(existing);
                placed.add(k);
            } else {
                EnvVarRow fromSheet = sheetTemplates.get(k);
                if (fromSheet != null) {
                    EnvVarRow copy = new EnvVarRow();
                    copy.setName(fromSheet.getName());
                    copy.setValue(fromSheet.getValue() != null ? fromSheet.getValue() : "");
                    copy.setDescription(fromSheet.getDescription());
                    if (BOOTSTRAP_KEY_SET.contains(k)) {
                        maybeFillEmptyBootstrap(copy, k, ui);
                    }
                    out.add(copy);
                    placed.add(k);
                } else if (BOOTSTRAP_KEY_SET.contains(k)) {
                    out.add(newBootstrapRow(k, ui));
                    placed.add(k);
                }
            }
        }
        HashSet<String> seen = new HashSet<>(placed);
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
        applyAlertStylesheetsFromOwner(alert);
        alert.setTitle("環境変数を初期値に戻す");
        alert.setHeaderText(null);
        alert.setContentText(
                "ui_ref_env_defaults.json の既定行に戻します。"
                        + "未保存の編集と、セッションに保存していた各タブの値（Python パス等）も失われます。"
                        + "続行しますか？");
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
                "[boot] " + PrismGpuBootstrapStatus.runTabSummary(), false);
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
                                + " が無いため、環境変数タブの値は未更新のままです。"
                                + " 段階1が配台除外ルール JSON を生成しているか、cwd/json の場所が一致しているか確認してください。");
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
                    "[env] PM_AI_EXCLUDE_RULES_JSON 行が見つからないため未更新のままです。");
        } catch (Exception ex) {
            appendLog("[env] PM_AI_EXCLUDE_RULES_JSON 更新に失敗: " + ex.getMessage());
        }
    }

    private void runStage(String script) {
        if (!runLock.compareAndSet(false, true)) {
            appendLog("[busy] already running (single flight).");
            return;
        }
        activeRunStageScript = script;
        applyRunTabGating();
        try {
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
            Map<String, String> childEnv = childEnvForPython(uiRun);
            if (lastNetworkSourceResolution != null) {
                for (String line : lastNetworkSourceResolution.logLines()) {
                    appendLog(line);
                }
            }
            if (STAGE1.equals(script)) {
                NetworkSourceDirResolver.Result res = lastNetworkSourceResolution;
                boolean show =
                        res != null && (res.taskInputFromCache() || res.actualDetailFromCache());
                mainRunTabController.setStage1NetworkCacheBadge(
                        show,
                        uiBadgeDesignTabController != null
                                ? uiBadgeDesignTabController.snapshotStage1NetworkCacheBadgeStyle()
                                : PersonBadgeStyle.networkSourceCacheBadgeDefault(),
                        uiBadgeDesignTabController != null
                                ? uiBadgeDesignTabController.snapshotStage1NetworkCacheBadgeLabel()
                                : "キャッシュ");
            }
            RunRequest req = new RunRequest(py, dir, script, wb, childEnv);
            mainRunTabController.getStatusLabel().setText("実行中…");

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
                            ex -> appendLog("[error] " + ex.getMessage()),
                            activeStageChildProcess::set)
                    .whenComplete(
                            (code, err) -> {
                                runLock.set(false);
                                activeRunStageScript = null;
                                activeStageChildProcess.set(null);
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
                                                mainRunTabController
                                                        .getStatusLabel()
                                                        .setText(exitCodeLegend(c));
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
                                                            "段階1 完了",
                                                            "段階1 の処理が正常終了しました。");
                                                }
                                                if (STAGE2.equals(script) && c == 0) {
                                                    refreshStage2OutputArtifacts();
                                                    if (dispatchInteractiveTabController != null) {
                                                        dispatchInteractiveTabController
                                                                .reloadTableFromDiskAfterExternalUpdate();
                                                    }
                                                    showStageCompletionDialog(
                                                            "段階2 完了",
                                                            "段階2 の処理が正常終了しました。");
                                                }
                                            }
                                        });
                            });
        } catch (Throwable t) {
            runLock.set(false);
            activeRunStageScript = null;
            activeStageChildProcess.set(null);
            appendLog("[error] runStage: " + t.getMessage());
            javafx.application.Platform.runLater(this::applyRunTabGating);
        }
    }

    /**
     * 段階1/2 実行中の Python 子プロセスを終了する（ツールバー・実行・ログの「中断」）。
     */
    void cancelActiveStageRun() {
        Process child = activeStageChildProcess.get();
        if (child != null && child.isAlive()) {
            appendLog("[interrupt] 段階1/2 の子プロセスを終了します…");
            try {
                child.destroyForcibly();
            } catch (Exception ex) {
                appendLog("[interrupt] 子プロセス終了に失敗: " + ex.getMessage());
            }
        } else {
            appendLog("[interrupt] 終了対象の子プロセスがありません。");
        }
    }

    @FXML
    private void onCancelStageRunAction() {
        cancelActiveStageRun();
    }

    /**
     * 段階1／段階2 実行中は「実行・ログ」以外のタブを無効化し、タブ切り替えを禁止する（ツールバーに進捗・中断）。
     */
    private void applyRunTabGating() {
        String script = activeRunStageScript;
        boolean stage1Running = STAGE1.equals(script);
        boolean stage2Running = STAGE2.equals(script);
        if (mainRunTabController != null) {
            mainRunTabController.setStageRunProgressVisible(stage1Running, stage2Running);
        }
        if (dispatchInteractiveTabController != null) {
            dispatchInteractiveTabController.setStageRunProgressVisible(stage1Running, stage2Running);
        }
        updateShellStageProgressOverlay(stage1Running, stage2Running);
        if (tabPane == null) {
            return;
        }
        ObservableList<Tab> tabs = tabPane.getTabs();
        if (tabs.isEmpty()) {
            return;
        }
        boolean stageBusy = stage1Running || stage2Running;
        for (Tab t : tabs) {
            t.setDisable(stageBusy && t != mainShellTabRun);
        }
        if (stageBusy) {
            Tab sel = tabPane.getSelectionModel().getSelectedItem();
            if (sel != mainShellTabRun) {
                tabPane.getSelectionModel().select(mainShellTabRun);
            }
        }
    }

    /**
     * メインウィンドウ上部ツールバーに段階1/2 実行中を表示する。
     * プログレスは {@link DispatchInteractiveTabController} の「機械 JSON 再読み」と同じ
     * {@link ProgressIndicator}（22×22）+ {@link ProgressBar}（prefWidth 220・不定）の組み合わせ。
     */
    private void updateShellStageProgressOverlay(boolean stage1Running, boolean stage2Running) {
        if (shellStageProgressBox == null) {
            return;
        }
        boolean show = stage1Running || stage2Running;
        if (show) {
            shellStageProgressBox.setManaged(true);
            shellStageProgressBox.setVisible(true);
            if (shellStageProgressBar != null) {
                shellStageProgressBar.setManaged(true);
                shellStageProgressBar.setVisible(true);
                shellStageProgressBar.setProgress(ProgressBar.INDETERMINATE_PROGRESS);
            }
            if (shellStageBusyIndicator != null) {
                shellStageBusyIndicator.setManaged(true);
                shellStageBusyIndicator.setVisible(true);
            }
            if (shellStageProgressLabel != null) {
                shellStageProgressLabel.setText(
                        stage1Running ? "段階1 実行中…" : "段階2 実行中…");
            }
            if (shellStageCancelButton != null) {
                shellStageCancelButton.setManaged(true);
                shellStageCancelButton.setVisible(true);
            }
        } else {
            if (shellStageProgressBar != null) {
                shellStageProgressBar.setProgress(0);
                shellStageProgressBar.setVisible(false);
                shellStageProgressBar.setManaged(false);
            }
            if (shellStageBusyIndicator != null) {
                shellStageBusyIndicator.setVisible(false);
                shellStageBusyIndicator.setManaged(false);
            }
            if (shellStageProgressLabel != null) {
                shellStageProgressLabel.setText("");
            }
            if (shellStageCancelButton != null) {
                shellStageCancelButton.setVisible(false);
                shellStageCancelButton.setManaged(false);
            }
            shellStageProgressBox.setVisible(false);
            shellStageProgressBox.setManaged(false);
        }
    }

    private void showStageCompletionDialog(String title, String contentText) {
        Alert alert = new Alert(AlertType.INFORMATION);
        alert.initOwner(primaryStage);
        applyAlertStylesheetsFromOwner(alert);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(contentText);
        alert.showAndWait();
    }

    /** メインウィンドウと同じテーマ CSS をダイアログに載せる（Alert は別 Scene のため未設定だと配色がずれる） */
    private void applyAlertStylesheetsFromOwner(Alert alert) {
        if (primaryStage == null) {
            return;
        }
        Scene ownerScene = primaryStage.getScene();
        if (ownerScene == null) {
            return;
        }
        var paneSheets = alert.getDialogPane().getStylesheets();
        for (String url : ownerScene.getStylesheets()) {
            if (!paneSheets.contains(url)) {
                paneSheets.add(url);
            }
        }
    }

    private static String exitCodeLegend(int code) {
        return "exit="
                + code
                + " （0=OK / 1=error / 2=fatal / 3=PlanningValidationError / 9=cancel）";
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
     * the 配台計画_タスク入力 tab are applied so that stage-2 uses the
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
        NetworkSourceDirResolver.Result netRes =
                NetworkSourceDirResolver.resolve(
                        m,
                        startupSkipTaskInputSourceDirListing,
                        startupSkipActualDetailSourceDirListing);
        lastNetworkSourceResolution = netRes;
        NetworkSourceDirResolver.applyToEnv(m, netRes);
        String pauseOnErr = m.get(AppPaths.KEY_PM_AI_CMD_PAUSE_ON_ERROR);
        if (pauseOnErr == null || pauseOnErr.isBlank()) {
            m.put(AppPaths.KEY_PM_AI_CMD_PAUSE_ON_ERROR, "0");
        }
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
     * Child-process env from the 環境変数 tab (same skip rules as workbook sheet: empty name, #).
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
     * 起動時: {@code PM_AI_TASK_INPUT_SOURCE_DIR} / {@code PM_AI_ACTUAL_DETAIL_SOURCE_DIR} のフォルダが一覧可能か調べ、
     * 不可なら以降の Python 向け env マージでネットワーク側の一覧を省略しキャッシュのみ試行する。
     */
    private void probeNetworkSourceDirsAtStartup() {
        Map<String, String> ui = collectUiEnv();
        boolean taskReach = NetworkSourceDirResolver.isTaskInputSourceDirReachable(ui);
        boolean actReach = NetworkSourceDirResolver.isActualDetailSourceDirReachable(ui);
        startupSkipTaskInputSourceDirListing = !taskReach;
        startupSkipActualDetailSourceDirListing = !actReach;
        Path taskDir = AppPaths.resolveTaskInputSourceDir(ui);
        Path actDir = AppPaths.resolveActualDetailSourceDir(ui);
        if (!taskReach) {
            appendLog(
                    "[startup] PM_AI_TASK_INPUT_SOURCE_DIR にアクセスできません（一覧不可）: "
                            + taskDir
                            + " — フォルダ参照を省略しキャッシュを優先します");
        }
        if (!actReach) {
            appendLog(
                    "[startup] PM_AI_ACTUAL_DETAIL_SOURCE_DIR にアクセスできません（一覧不可）: "
                            + actDir
                            + " — フォルダ参照を省略しキャッシュを優先します");
        }
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

    /**
     * メインシェルのタブを ID で選択する（配台試行ウィザードなどから）。
     */
    public void selectMainShellTab(MainShellTabId id) {
        if (tabPane == null || id == null) {
            return;
        }
        Tab t = mainShellTabFor(id);
        if (t != null) {
            tabPane.getSelectionModel().select(t);
        }
    }

    /** 計画結果ビューアを選択し、段階2成果のパスで JSON フィールドを埋める。 */
    public void navigatePlanResultViewerWithArtifacts(String productionPlanPath, String memberSchedulePath) {
        selectMainShellTab(MainShellTabId.PLAN_RESULT_VIEWER);
        String p = productionPlanPath != null ? productionPlanPath : "";
        String m = memberSchedulePath != null ? memberSchedulePath : "";
        planResultViewerTabController.tryAutofillJsonFromStage2Xlsx(p, m);
    }

    /** 設備ガントを選択し、同じ成果パスで読み込む。 */
    public void navigateEquipmentGanttWithArtifacts(String productionPlanPath, String memberSchedulePath) {
        selectMainShellTab(MainShellTabId.EQUIPMENT_GANTT_GRAPHIC);
        String p = productionPlanPath != null ? productionPlanPath : "";
        String m = memberSchedulePath != null ? memberSchedulePath : "";
        equipmentGanttGraphicTabController.tryAutofillJsonFromStage2Xlsx(p, m);
    }

    /** 配台計画手動修正タブへ切り替える。 */
    public void navigateDispatchInteractiveTab() {
        selectMainShellTab(MainShellTabId.DISPATCH_INTERACTIVE);
    }

    /** {@link AppPaths#defaultPlanningOutputDir} を OS のファイルマネージャで開く。 */
    public void openDefaultPlanningOutputFolderInOs() {
        try {
            Path dir = AppPaths.defaultPlanningOutputDir(collectUiEnv());
            if (!Files.isDirectory(dir)) {
                appendLog("[dispatch-wizard] 出力フォルダがありません: " + dir);
                return;
            }
            java.awt.Desktop.getDesktop().open(dir.toFile());
            appendLog("[dispatch-wizard] 出力フォルダを開きました: " + dir);
        } catch (Exception e) {
            appendLog(
                    "[dispatch-wizard] フォルダを開けませんでした: "
                            + (e.getMessage() != null ? e.getMessage() : e));
        }
    }

    /** Same-package tab controllers append run-tab log lines here. */
    void appendLog(String line) {
        mainRunTabController.appendLog(line);
    }

    Map<String, String> snapshotUiEnv() {
        return collectUiEnv();
    }

    /**
     * Environment for Python child processes (same as stage1/2): env tab + plan-input tab overlays,
     * {@code PM_AI_*} inheritance rules, UTF-8 stdio.
     */
    public Map<String, String> snapshotPythonChildEnv() {
        return childEnvForPython(collectUiEnv());
    }

    /**
     * Environment for {@code dispatch_interactive_trial.py}: same {@link AppPaths#KEY_PM_AI_STAGE2_WRITE_EXCEL} and
     * {@link AppPaths#KEY_PM_AI_RESULT_BOOK_FONT} overrides as running stage 2 from the run tab (unchecking Excel
     * there suppresses xlsx deliverables in planning_core for the trial as well).
     */
    public Map<String, String> snapshotDispatchTrialPythonEnv() {
        Map<String, String> ui = new HashMap<>(collectUiEnv());
        ui.put(
                AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL,
                mainRunTabController.snapshotStage2WriteExcel() ? "1" : "0");
        String resultFont = mainRunTabController.snapshotStage2ResultBookFont();
        if (resultFont != null && !resultFont.isBlank()) {
            ui.put(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT, resultFont.trim());
        } else {
            ui.remove(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT);
        }
        return childEnvForPython(ui);
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
    /**
     * 配台試行完了後など、出力フォルダに新しい段階2成果物があれば実行・ログタブのパス（production_plan /
     * member_schedule）と関連タブの自動反映を更新する。{@link #refreshStage2OutputArtifacts} と同じ処理。
     */
    void refreshRunTabStage2ArtifactLinks() {
        refreshStage2OutputArtifacts();
    }

    private void refreshStage2OutputArtifacts() {
        try {
            Map<String, String> ui = collectUiEnv();
            Path dir = AppPaths.defaultPlanningOutputDir(ui);
            if (!Files.isDirectory(dir)) {
                mainRunTabController.setStage2ArtifactPaths("", "");
                appendLog(
                        "[stage2-ui] "
                                + "出力フォルダがありません: "
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
            equipmentGanttGraphicTabController.tryAutofillJsonFromStage2Xlsx(planStr, memStr);
            operatorCardTabController.tryAutofillMemberJsonFromStage2(memStr);
            if (!planStr.isEmpty() || !memStr.isEmpty()) {
                appendLog(
                        "[stage2-ui] "
                                + "最新成果物: production_plan="
                                + planStr
                                + " | member_schedule="
                                + memStr);
            }
        } catch (Exception ex) {
            appendLog(
                    "[stage2-ui] "
                            + "成果パス更新エラー: "
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
