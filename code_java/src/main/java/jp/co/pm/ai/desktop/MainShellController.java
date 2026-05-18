package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.geometry.Rectangle2D;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ChoiceDialog;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Labeled;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.TreeItem;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Text;
import javafx.stage.Modality;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.util.Duration;
import javafx.util.StringConverter;

import com.fasterxml.jackson.databind.JsonNode;

import jp.co.pm.ai.desktop.runtime.FxJvmMemoryStatusBar;

import jp.co.pm.ai.desktop.audio.MacroCompleteChime;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.bridge.Stage2PythonChildEnv;
import jp.co.pm.ai.desktop.bridge.StagePythonExecutable;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.Stage1AiCacheClearer;
import jp.co.pm.ai.desktop.config.WorkspaceCacheArchiveStore;
import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.DispatchTrialLogUiStore;
import jp.co.pm.ai.desktop.config.JvmMemoryLogStore;
import jp.co.pm.ai.desktop.config.MainShellTabLayoutDefaults;
import jp.co.pm.ai.desktop.config.MainShellTabLayoutNode;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.config.DesktopTheme;
import jp.co.pm.ai.desktop.config.PushButtonCssEmitter;
import jp.co.pm.ai.desktop.config.PushButtonDesignPrefs;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.config.PortableBundleSelfUpdater;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.config.PlanWorkspaceSessionFragment;
import jp.co.pm.ai.desktop.config.PlanWorkspaceSnapshotStore;
import jp.co.pm.ai.desktop.config.EnvVarDocs;
import jp.co.pm.ai.desktop.config.InitSettingPersistence;
import jp.co.pm.ai.desktop.config.UiEnvRowSnapshot;
import jp.co.pm.ai.desktop.config.UiRefEnvDefaults;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.runtime.MemoryJvmRingLog;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchDocument;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPythonExport;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.SummaryAiDispatchWorkbookExporter;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;
import jp.co.pm.ai.desktop.ipc.IpcStdoutTap;
/**
 * Main window controller（従来は {@link PmAiFxApp} 内蔵だった業務ロジックを分離）。
 * Layout: {@code MainShell.fxml} and tab FXML files.
 */
public final class MainShellController {

    /**
     * {@link Tab#getProperties()} に登録済みかどうか。選択変更時に見出し chrome を再適用するリスナーを二重登録しない。
     */
    private static final String PROP_SHELL_TAB_SELECTION_CHROME_LISTENER =
            "pmShellTabSelectionChromeListener";

    private static final String STAGE1 = "task_extract_stage1.py";
    private static final String STAGE2 = "plan_simulation_stage2.py";

    /** 段階1実行前ログに出す「入力解決に関わる」環境変数キー（子プロセスへ渡る値）。 */
    private static final List<String> STAGE1_CHILD_INPUT_ENV_KEYS =
            List.of(
                    AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                    AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH,
                    "PM_AI_PROCESSING_PLAN_SHEET",
                    "PM_AI_PROCESSING_PLAN_HEADER_ROW",
                    AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
                    AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK,
                    AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SHEET,
                    AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                    AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                    AppPaths.KEY_MASTER_WORKBOOK_FILE,
                    AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
                    AppPaths.KEY_PM_AI_OUTPUT_DIR,
                    AppPaths.KEY_PM_AI_REPO_ROOT,
                    "PM_AI_DEBUG_LOG",
                    AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH);

    /** 段階2実行前ログに出す「入力解決に関わる」環境変数キー。 */
    private static final List<String> STAGE2_CHILD_INPUT_ENV_KEYS =
            List.of(
                    AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                    PlanInputTabController.ENV_TASK_PLAN_SHEET,
                    AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                    AppPaths.KEY_MASTER_WORKBOOK_FILE,
                    AppPaths.KEY_PM_AI_OUTPUT_DIR,
                    AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH,
                    AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK,
                    AppPaths.KEY_PM_AI_PLAN_WORKBOOK_JSON,
                    AppPaths.KEY_PM_AI_MEMBER_SCHEDULE_JSON,
                    AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
                    AppPaths.KEY_PM_AI_STAGE2_ENGINE,
                    AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL,
                    AppPaths.KEY_PM_AI_STAGE2_SKIP_TODAY_DISPATCH,
                    AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH,
                    "PM_AI_DEBUG_LOG");

    private static final String PREFIX_CHILD = "[child] ";
    private static final String NDJSON_START = PREFIX_CHILD + "{";

    /** 段階1／2 失敗ダイアログに載せる子プロセス出力の末尾行数上限（リングバッファ）。 */
    private static final int STAGE_CHILD_LOG_TAIL_MAX = 48;

    /**
     * Dropped from the env tab (defaults and session); not used in normal operation. Python still accepts
     * these if set in the real OS environment.
     */
    private static final Set<String> DROPPED_ENV_TAB_ROW_KEYS =
            Set.of(
                    "DEBUG_TASK_ID",
                    "TRACE_TEAM_ASSIGN_TASK_ID",
                    "EXCLUDE_RULES_TEST_E1234",
                    "EXCLUDE_RULES_TEST_E1234_ROW",
                    "STAGE2_SKIP_SHEET_VISIBILITY_APPLY",
                    "STAGE2_SKIP_SNAPSHOT_EXPORT",
                    "STAGE2_SKIP_MEMBER_SCHEDULE_IMPORT",
                    "STAGE12_CMD_HIDE_WINDOW",
                    "EXCLUDE_RULES_TRY_OPENPYXL_SAVE");

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
                    AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH,
                    AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
                    AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
                    AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON,
                    AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH,
                    AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR);

    /** Keys in {@link #BOOTSTRAP_ORDER} for quick membership checks. */
    private static final Set<String> BOOTSTRAP_KEY_SET = Set.copyOf(BOOTSTRAP_ORDER);

    private final Stage primaryStage;

    @FXML
    private TabPane tabPane;

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
    private Label jvmMemoryStatusLabel;

    @FXML
    private MainRunTabController mainRunTabController;

    @FXML
    private EnvTabController envTabController;

    @FXML
    private MemorySettingsTabController memorySettingsTabController;

    @FXML
    private GlobalSettingsTabController globalSettingsTabController;

    @FXML
    private SummaryAiDispatchExportCustomizeTabController summaryAiDispatchExportCustomizeTabController;

    @FXML
    private UserProfilesTabController userProfilesTabController;

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
    private DeliveryCalendarViewTabController deliveryCalendarViewTabController;

    @FXML
    private MasterReadSummaryTabController masterReadSummaryTabController;

    @FXML
    private ResultDispatchTableTabController resultDispatchTableTabController;

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
    private PlanWorkspaceHistoryTabController planWorkspaceHistoryTabController;

    private WorkspaceCacheHistoryTabController workspaceCacheHistoryTabController;

    @FXML
    private ApiModelBenchmarkTabController apiModelBenchmarkTabController;

    @FXML
    private CodeDispatchLookupTablesTabController codeDispatchLookupTablesTabController;

    @FXML
    private Tab mainShellTabRun;

    @FXML
    private Tab mainShellTabUiBadgeDesign;

    @FXML
    private Tab mainShellTabPushButtonDesign;

    @FXML
    private Tab mainShellTabEnv;

    @FXML
    private Tab mainShellTabMemorySettings;

    @FXML
    private Tab mainShellTabGlobalSettings;

    @FXML
    private Tab mainShellTabSummaryAiDispatchExportCustomize;

    @FXML
    private Tab mainShellTabUserProfiles;

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
    private Tab mainShellTabDeliveryCalendar;

    @FXML
    private Tab mainShellTabResultDispatch;

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

    @FXML
    private Tab mainShellTabPlanWorkspaceHistory;

    @FXML
    private Tab mainShellTabCacheHistory;

    @FXML
    private Tab mainShellTabApiModelBenchmark;

    @FXML
    private Tab mainShellTabCodeLookupTables;

    @FXML
    private Tab mainShellTabOrganizer;

    @FXML
    private MainShellTabOrganizerTabController mainShellTabOrganizerPaneController;

    /** 入れ子 {@link TabPane} の選択変化を監視する（メイン以外）。 */
    private final List<TabPane> wiredInnerMainShellTabPanes = new ArrayList<>();

    /** {@link #emitShellTabNavigation()} 用の直前リーフ（列フィルタ解除・実行タブ離脱保存）。 */
    private Tab lastEffectiveShellLeaf;

    private ObservableList<EnvVarRow> envRows;

    private final AtomicBoolean runLock = new AtomicBoolean(false);

    /** Non-null while a stage script is running; equals {@link #STAGE1} or {@link #STAGE2}. */
    private volatile String activeRunStageScript;

    /** Python child process while stage 1/2 is running; cleared on completion or interrupt. */
    private final AtomicReference<Process> activeStageChildProcess = new AtomicReference<>();

    /** {@link #childEnvForPython(Map)} の直近結果（実行タブのキャッシュ表示・ログ用）。 */
    private NetworkSourceDirResolver.Result lastNetworkSourceResolution;

    /**
     * ソースフォルダが一覧不可なら {@code true}。{@link NetworkSourceDirResolver#resolve(Map, boolean, boolean)} でネットワーク側の一覧を省略する。
     *
     * <p>起動時プローブで初期化し、段階1／段階2実行直前に {@link #refreshNetworkSourceDirListingSkipsBeforeStageRun(Map)} で再評価する（ネットワーク復旧後は一覧を再試行する）。
     */
    private volatile boolean startupSkipTaskInputSourceDirListing;

    private volatile boolean startupSkipActualDetailSourceDirListing;

    private final AtomicBoolean suppressEnvSessionPersistence = new AtomicBoolean(false);

    /** 納期管理ビュー再読み込み中のタブ差し戻しで {@link TabPane} の選択リスナーを再入しない。 */
    private final AtomicBoolean suppressDeliveryCalendarReloadTabGuard = new AtomicBoolean(false);

    /**
     * メインタブの組み替え中に {@link #refreshMainShellTabHeaderChromeFromStoredColors()} を抑止する。
     * タブ追加・削除と同期 {@link TabPane#layout()} が重なると {@code IndexOutOfBoundsException} になりやすい。
     */
    private final AtomicBoolean suppressMainShellTabChromeRefresh = new AtomicBoolean(false);

    /**
     * {@link #applyDesktopSession} でタブ構成を復元するセッション。{@link Stage#setOnShown} 前は
     * {@link TabPane} の再構築を遅延し、初回 {@link Scene#doLayoutPass} と競合しないようにする。
     */
    private DesktopSessionState pendingMainShellTabLayoutSession;

    /** 非選択タブの重い {@link Tab#setContent(Node)} を退避するときの {@link Tab#getProperties()} キー。 */
    private static final String PM_DEFERRED_TAB_CONTENT = "pmDeferredTabContent";

    private static final String PM_LAZY_TAB_PLACEHOLDER = "pmLazyTabPlaceholder";

    private final AtomicBoolean suppressLazyMainShellTabContentSwap = new AtomicBoolean(false);

    /** メインシェル見出しのユーザー色にドロップシャドウ風グローを付ける（タブ整理のチェック）。 */
    private final AtomicBoolean mainShellTabOrganizerHeaderGlowEnabled = new AtomicBoolean(true);

    /** 見出しグローの強さ（0〜1、1 が従来の既定ビジュアル）。 */
    private final AtomicReference<Double> mainShellTabOrganizerHeaderGlowStrength =
            new AtomicReference<>(1.0);

    private final PauseTransition uiEnvSaveDebounce = new PauseTransition(Duration.millis(400));
    /** Assigned in {@link #installUiEnvAutoSave()} for debounced {@link #scheduleDesktopSessionSave()}. */
    private Runnable uiEnvPersistSchedule;
    private final AtomicBoolean envResetInProgress = new AtomicBoolean(false);

    private DesktopTheme pendingTheme = DesktopTheme.LIGHT;

    /** FXML 読込直後に固定した既定見出し（内部 ID は {@link MainShellTabId#key()} のまま）。 */
    private final Map<MainShellTabId, String> mainShellTabBaselineTitles = new EnumMap<>(MainShellTabId.class);

    /** セッション保存する見出しエイリアス（キーは {@link MainShellTabId#key()}）。 */
    private final Map<String, String> mainShellTabTitleAliases = new LinkedHashMap<>();

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
            captureMainShellTabBaselineTitles();
            installMainShellTabPaneChromeHooks();
            installLazyMainShellTabContentForStartup();
            envRows = FXCollections.observableArrayList();
            populateEnvRows(envRows);
            applyBundledPortableDefaultsIfPresent();
            Map<String, String> ui0 = collectUiEnv();

            mainRunTabController.bindShell(this);
            envTabController.bindShell(this);
            memorySettingsTabController.bindShell(this);
            if (globalSettingsTabController != null) {
                globalSettingsTabController.bindShell(this);
            }
            if (summaryAiDispatchExportCustomizeTabController != null) {
                summaryAiDispatchExportCustomizeTabController.bindShell(this);
            }
            if (userProfilesTabController != null) {
                userProfilesTabController.bindShell(this);
            }
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
        if (codeDispatchLookupTablesTabController != null) {
            codeDispatchLookupTablesTabController.bindShell(this);
        }
        excludeRulesTabController.bindShell(this);
        specialRulesTabController.bindShell(this);
        actualsStatusTabController.bindShell(this);
        deliveryCalendarViewTabController.bindShell(this);
        resultDispatchTableTabController.bindShell(this);
        dispatchInteractiveTabController.bindShell(this);
        if (planWorkspaceHistoryTabController != null) {
            planWorkspaceHistoryTabController.bindShell(this);
        }
        if (workspaceCacheHistoryTabController != null) {
            workspaceCacheHistoryTabController.bindShell(this);
        }
        if (apiModelBenchmarkTabController != null) {
            apiModelBenchmarkTabController.bindShell(this);
        }

        primaryStage.setMinWidth(640);
        primaryStage.setMinHeight(480);

            applyDesktopSession(DesktopSessionStateStore.load());
            if (mainShellTabOrganizerPaneController != null) {
                mainShellTabOrganizerPaneController.bindShell(this);
                mainShellTabOrganizerPaneController.installTreeCellFactory();
            }
        } finally {
            suppressEnvSessionPersistence.set(false);
        }

        /* 起動時は常にメインウィンドウを最大化（セッションの幅・高さ・位置は復元後に上書き） */
        primaryStage.setMaximized(true);

        if (toolbarGrowSpacer != null) {
            HBox.setHgrow(toolbarGrowSpacer, Priority.ALWAYS);
        }

        FxJvmMemoryStatusBar.start(jvmMemoryStatusLabel, primaryStage);

        installUiEnvAutoSave();

        applyRepoFolderPathNormalization();
        maybePortableFirstLaunchEnvInit();

        probeNetworkSourceDirsAtStartup();

        primaryStage.setOnCloseRequest(
                e -> {
                    memorySettingsTabController.shutdown();
                    JvmMemoryLogStore.persistSnapshot(
                            MemoryJvmRingLog.getMaxLines(), MemoryJvmRingLog.snapshotLines());
                    DesktopSessionStateStore.save(collectDesktopSession());
                });

        primaryStage.setOnShown(
                e -> {
                    primaryStage.toFront();
                    primaryStage.requestFocus();
                    applyPendingMainShellTabLayoutFromSessionIfNeeded();
                    installLazyMainShellTabContentForStartup();
                    if (tabPane.getSelectionModel().getSelectedItem() == null
                            && !tabPane.getTabs().isEmpty()) {
                        tabPane.getSelectionModel().selectFirst();
                    }
                    activateMainShellTabHeavyContentRecursive(
                            tabPane.getSelectionModel().getSelectedItem());
                    Platform.runLater(
                            () ->
                                    Platform.runLater(
                                            () -> {
                                                refreshMainShellTabHeaderChromeFromStoredColors();
                                                if (dispatchInteractiveTabController != null) {
                                                    dispatchInteractiveTabController
                                                            .scheduleInitialReloadAfterMainWindowShown();
                                                }
                                            }));
                });

        lastEffectiveShellLeaf =
                resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        tabPane
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, prevTab, newTab) -> {
                            if (!suppressDeliveryCalendarReloadTabGuard.get()
                                    && deliveryCalendarViewTabController != null
                                    && mainShellTabDeliveryCalendar != null
                                    && deliveryCalendarViewTabController
                                            .isReloadBlockingMainShellTabNavigation()
                                    && newTab != mainShellTabDeliveryCalendar) {
                                suppressDeliveryCalendarReloadTabGuard.set(true);
                                try {
                                    tabPane.getSelectionModel().select(mainShellTabDeliveryCalendar);
                                    appendLog(
                                            "[delivery-calendar] 再読み込み完了まで他のメインタブへ切り替えできません");
                                } finally {
                                    suppressDeliveryCalendarReloadTabGuard.set(false);
                                }
                                return;
                            }
                            if (!suppressLazyMainShellTabContentSwap.get()) {
                                deferMainShellTabBranchHeavyContent(prevTab);
                                activateMainShellTabHeavyContentRecursive(newTab);
                            }
                            emitShellTabNavigation();
                            /* :selected 由来の -fx-text-fill がインラインより後勝ちになることがあるため再適用 */
                            if (!suppressMainShellTabChromeRefresh.get()) {
                                refreshMainShellTabHeaderChromeFromStoredColors();
                            }
                            if (newTab == mainShellTabEquipmentGanttGraphic
                                    && equipmentGanttGraphicTabController != null) {
                                equipmentGanttGraphicTabController
                                        .flushPendingGraphicRebuildAfterSessionApply();
                            }
                            if (newTab == mainShellTabDeliveryCalendar
                                    && deliveryCalendarViewTabController != null) {
                                deliveryCalendarViewTabController.collapseInnerSectionPanesOnShellSelect();
                            }
                            if (newTab == mainShellTabApiModelBenchmark
                                    && apiModelBenchmarkTabController != null) {
                                apiModelBenchmarkTabController.refreshShellDerivedLabels();
                            }
                            if (newTab == mainShellTabDispatchInteractive
                                    && dispatchInteractiveTabController != null) {
                                dispatchInteractiveTabController.onMainShellDispatchTabSelected();
                            }
                            if (newTab == mainShellTabOrganizer
                                    && mainShellTabOrganizerPaneController != null) {
                                mainShellTabOrganizerPaneController.refreshFromShell();
                            }
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
                            refreshMainShellTabHeaderChromeFromStoredColors();
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

    private void applyDesktopSession(DesktopSessionState s) {
        applyDesktopSession(s, true);
    }

    /**
     * @param restoreUiEnvRowsFromSession {@code false} のとき環境変数タブはセッションから復元せず、呼び出し元で構築済みの行を保持する（ポータル
     *     バージョンアップ直後のバンドル既定への初期化後など）。
     */
    private void applyDesktopSession(DesktopSessionState s, boolean restoreUiEnvRowsFromSession) {
        if (s == null) {
            return;
        }
        JvmMemoryLogStore.bootstrapRingFromDisk();
        setMainShellTabOrganizerHeaderGlowEnabled(s.mainShellTabOrganizerHeaderGlow());
        setMainShellTabOrganizerHeaderGlowStrength(
                clamp(s.mainShellTabOrganizerHeaderGlowStrength(), 0.0, 1.0));
        if (restoreUiEnvRowsFromSession) {
            applyUiEnvRowsFromSession(s);
        }
        memorySettingsTabController.applyMemorySettingsSession(s);
        planInputTabController.restoreDesktopSessionPaths(s.planInputPath(), s.planInputSheet());
        stage1PreviewTabController.restoreDesktopSessionPaths(s.stage1PreviewPath(), s.stage1PreviewSheet());
        excludeRulesTabController.restoreDesktopSessionPath(s.excludeRulesPath());
        if (nonBlank(s.mainRunWorkbook())) {
            mainRunTabController.getWorkbookField().setText(s.mainRunWorkbook());
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
        planInputTabController.applyStage2SkipTodayDispatchFromSession(s.mainRunStage2SkipTodayDispatch());
        mainRunTabController.applyStage2SkipInProgressDispatchFromSession(
                s.mainRunStage2SkipInProgressDispatch());
        mainRunTabController.applyStage2ResultBookFontFromSession(s.mainRunStage2ResultBookFont());
        /*
         * 設備ガントの apply は末尾で Canvas を再構築し personBadgeStyleResolverForGantt を参照する。
         * 担当バッジのセッション（グロー等）を先に適用しないと、起動直後の帯は既定スタイルで描かれる。
         */
        if (ganttPersonBadgeDesignTabController != null) {
            ganttPersonBadgeDesignTabController.applyPersonBadgeDesignSession(s);
        }
        equipmentGanttGraphicTabController.applyEquipmentGanttSession(s);
        if (uiBadgeDesignTabController != null) {
            uiBadgeDesignTabController.applyUiBadgeSession(s);
        }
        if (pushButtonDesignTabController != null) {
            pushButtonDesignTabController.applyPushButtonSession(s);
        }
        applyWindowGeometry(s);
        applyOrDeferMainShellTabLayoutFromSession(s);
        pendingTheme = DesktopTheme.fromStored(s.uiTheme());
        if (mainShellTabOrganizerPaneController != null) {
            mainShellTabOrganizerPaneController.syncHeaderGlowControlsFromShell();
        }
        mainRunTabController.refreshOpenWorkbookHintLabels();
        mainRunTabController.refreshFactorySiteLogo();
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
                planInputTabController.snapshotStage2SkipTodayDispatch(),
                mainRunTabController.snapshotStage2SkipInProgressDispatch(),
                mainRunTabController.snapshotStage2ResultBookFont(),
                snapshotUiEnvRows(),
                snapshotMainShellTabOrder(),
                snapshotMainShellTabLayout(),
                snapshotMainShellTabTitleAliases(),
                snapshotInnerTabSelectedIndexByShellTabKey(),
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
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeGapPx(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeBandVerticalOffsetPx(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttGraphicDataFingerprint(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttBadgeDragDeltas(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeDragAdjustEnabled(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeEnabled(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeWireEnabled(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeWireStrokeHex(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeWireWidthPx(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeWireDashStyleKey(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPersonBadgeWireMaxLengthPx(),
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
                snapshotPersonBadgeOpacity(),
                snapshotPersonBadgeStylesByLabel(),
                snapshotPersonBadgeStylesByMemberKey(),
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPlanJsonPath(),
                uiBadgeDesignTabController != null
                        ? uiBadgeDesignTabController.snapshotStage1NetworkCacheBadgeLabel()
                        : "",
                uiBadgeDesignTabController != null
                        ? uiBadgeDesignTabController.snapshotStage1NetworkCacheBadgeStyle()
                        : PersonBadgeStyle.networkSourceCacheBadgeDefault(),
                mainShellTabOrganizerHeaderGlowEnabled.get(),
                getMainShellTabOrganizerHeaderGlowStrength(),
                pushButtonDesignTabController != null
                        ? pushButtonDesignTabController.snapshotPrefs()
                        : PushButtonDesignPrefs.inactiveDefaults(),
                memorySettingsTabController.snapshotMemoryMonitorEnabled(),
                memorySettingsTabController.snapshotMemoryMonitorIntervalSec(),
                memorySettingsTabController.snapshotNextLaunchHeapMaxMiB());
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
                AppPaths.resolveMasterWorkbookPathForDesktopOpen(
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

    private double snapshotPersonBadgeOpacity() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.snapshotPersonBadgeOpacity()
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

    /** 配台ワークスペース用スナップショットに書き出す現在の配台表ドキュメント（未初期化時は {@code null}）。 */
    public ResultDispatchDocument snapshotDispatchDocumentForPlanWorkspace() {
        return dispatchInteractiveTabController != null
                ? dispatchInteractiveTabController.copyDispatchDocumentForSnapshot()
                : null;
    }

    /**
     * スナップショットの内容で正規の結果_配台表 JSON と関連 UI 状態（配台入力・ガント・列順断片）を復元する。
     *
     * @throws IOException 入出力エラー
     */
    /**
     * キャッシュ退避履歴を現在のワークスペースパスへ復元する。
     */
    public void restoreWorkspaceCacheArchive(WorkspaceCacheArchiveStore.WorkspaceCacheArchiveEntry entry)
            throws IOException {
        if (entry == null) {
            return;
        }
        for (String line : WorkspaceCacheArchiveStore.restoreToWorkspace(entry, collectUiEnv())) {
            appendLog(line);
        }
        appendLog("[cache-archive] キャッシュを復元しました（履歴 ID: " + entry.id() + "）。");
        if (dispatchInteractiveTabController != null) {
            dispatchInteractiveTabController.reloadTableFromDiskAfterExternalUpdate();
        }
        if (resultDispatchTableTabController != null) {
            resultDispatchTableTabController.reloadResultDispatchTableFromDisk();
        }
        invalidateDeliveryCalendarAfterPipelineRun();
    }

    public void restorePlanWorkspaceSnapshot(PlanWorkspaceSnapshotStore.PlanWorkspaceSnapshotEntry entry)
            throws IOException {
        if (entry == null) {
            return;
        }
        Path snapJson = PlanWorkspaceSnapshotStore.resultDispatchJsonPath(entry);
        if (!Files.isRegularFile(snapJson)) {
            throw new IOException(
                    "スナップショットに "
                            + AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME
                            + "（旧 result_dispatch.json）がありません");
        }
        JsonNode colPart = PlanWorkspaceSnapshotStore.readColumnOrderPartial(entry);
        TableColumnOrderPersistence.mergePlanWorkspaceColumnOrderPartial(colPart);

        Path canonical = AppPaths.resolveResultDispatchTableJsonPath(collectUiEnv());
        Path parent = canonical.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Files.copy(snapJson, canonical, StandardCopyOption.REPLACE_EXISTING);

        tryExportResultDispatchTableXlsxNearJson(canonical);

        PlanWorkspaceSessionFragment frag = PlanWorkspaceSnapshotStore.readSessionFragment(entry);
        DesktopSessionState merged = frag.mergeOnto(collectDesktopSession());
        applyDesktopSession(merged, false);
        if (dispatchInteractiveTabController != null) {
            dispatchInteractiveTabController.reloadTableFromDiskAfterExternalUpdate();
        }
        if (resultDispatchTableTabController != null) {
            resultDispatchTableTabController.reloadResultDispatchTableFromDisk();
        }
        persistDesktopSessionNow();
    }

    /**
     * {@code export_result_dispatch_from_json.py} 経由で、指定 JSON と同階層に {@code 結果_配台表.xlsx} を書き出す（段階2の
     * {@code planning_core._write_dispatch_table_standalone_xlsx} と同一経路）。失敗時はログのみ。
     */
    public void tryExportResultDispatchTableXlsxNearJson(Path jsonPath) {
        if (jsonPath == null) {
            return;
        }
        try {
            Path pyExe = resolveStagePythonExecutablePath();
            Path pyDir = AppPaths.resolvePythonScriptDir(collectUiEnv());
            String line = ResultDispatchPythonExport.exportXlsxNearJson(jsonPath, pyExe, pyDir);
            if (line != null && !line.isBlank()) {
                appendLog("[結果_配台表] xlsx 同期待ち（段階2と同一 export）: " + line.trim());
            } else {
                appendLog(
                        "[結果_配台表] xlsx 同期待ち: export_result_dispatch_from_json が失敗または未配置（JSON のみ更新）");
            }
        } catch (Exception ex) {
            appendLog(
                    "[結果_配台表] xlsx 同期スキップ: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.getClass().getSimpleName()));
        }
    }

    /** グローバル設定の「現在の状態をデフォルトとする」実行直前にローカル {@code session-state.json} を同期する。 */
    public void preparePackageDefaultsExport() {
        persistDesktopSessionNow();
    }

    /** {@link InitSettingPersistence} 用のセッションスナップショット。 */
    public DesktopSessionState snapshotDesktopSessionForExport() {
        return collectDesktopSession();
    }

    /**
     * ユーザープロファイル読み出し: 列順 JSON と {@link DesktopSessionState} を適用し、テーマ・ログ・ガント等を追従させる。
     *
     * @throws IOException 列順ファイルの書き込みに失敗したとき
     */
    public void applyUserProfileSnapshot(DesktopSessionState state, JsonNode tableColumnOrderRoot)
            throws IOException {
        if (state == null) {
            return;
        }
        TableColumnOrderPersistence.overwriteStoreRoot(tableColumnOrderRoot);
        applyDesktopSession(state, true);
        applyDesktopThemeFromSession(state);
        refreshDesktopSessionDependentUi();
        persistDesktopSessionNow();
    }

    /** セッションの {@code uiTheme} をツールバー・シーンへ反映する。 */
    private void applyDesktopThemeFromSession(DesktopSessionState state) {
        if (state == null) {
            return;
        }
        DesktopTheme t = DesktopTheme.fromStored(state.uiTheme());
        pendingTheme = t;
        if (themeCombo != null) {
            themeCombo.setValue(t);
        }
        if (primaryScene != null) {
            t.applyTo(primaryScene);
        }
        refreshThemeTrackedSecondaryScenes();
    }

    /** テーマ／プッシュボタン CSS／タブ見出し色など、セッション保存前に画面へ揃える。 */
    private void refreshDesktopSessionDependentUi() {
        refreshPushButtonStylesheet();
        refreshMainShellTabHeaderChromeFromStoredColors();
        if (equipmentGanttGraphicTabController != null) {
            equipmentGanttGraphicTabController.refreshGraphicForTheme();
        }
        mainRunTabController.refreshLogThemeCells();
    }

    /**
     * グローバル設定「デフォルトに戻す」適用後、現在の画面状態を {@code ~/.pm-ai-desktop/session-state.json} へ保存する。
     * タブ再構築・見出し色の再適用が終わってから書き込むため、FX スレッドで 2 パルス遅延する。
     */
    private void schedulePersistUserSessionAfterGlobalFactoryReset() {
        Platform.runLater(() -> Platform.runLater(this::persistUserSessionAfterGlobalFactoryReset));
    }

    private void persistUserSessionAfterGlobalFactoryReset() {
        if (ganttPersonBadgeDesignTabController != null) {
            ganttPersonBadgeDesignTabController.flushBadgeEditsBeforeSnapshot();
        }
        if (uiBadgeDesignTabController != null) {
            uiBadgeDesignTabController.flushEditsBeforeSnapshot();
        }
        if (pushButtonDesignTabController != null) {
            pushButtonDesignTabController.flushEditsBeforeSnapshot();
        }
        DesktopSessionStateStore.save(collectDesktopSession());
    }

    Stage primaryStageForDialogs() {
        return primaryStage;
    }

    /**
     * メインウィンドウと同じテーマ CSS をダイアログに載せる（{@link Alert} / {@link ChoiceDialog} 等）。
     */
    public void prepareDialogForMainTheme(Dialog<?> dialog) {
        if (dialog == null) {
            return;
        }
        if (primaryStage != null) {
            dialog.initOwner(primaryStage);
        }
        applyAlertStylesheetsFromOwner(dialog);
    }

    /** 保存・読込完了などの情報ダイアログ。 */
    public void showInformationDialog(String title, String message) {
        showThemedAlert(AlertType.INFORMATION, title, null, message);
    }

    /** ファイルなし・部分成功などの注意ダイアログ。 */
    public void showWarningDialog(String title, String message) {
        showThemedAlert(AlertType.WARNING, title, null, message);
    }

    /** 失敗時のエラーダイアログ。 */
    public void showErrorDialog(String title, String message) {
        showThemedAlert(AlertType.ERROR, title, null, message);
    }

    private void showThemedAlert(AlertType type, String title, String headerText, String message) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(headerText);
        alert.setContentText(message);
        prepareDialogForMainTheme(alert);
        alert.showAndWait();
    }

    /**
     * タブ・表・テーマ等をマージ済みバンドル既定へ戻し、環境タブはテンプレ既定へ戻す。実行パス・環境値のうちブートストラップ系は
     * リセット後の環境タブから再収集する。
     *
     * <p>適用完了後、{@link #schedulePersistUserSessionAfterGlobalFactoryReset()} でユーザーセッション
     * （{@code session-state.json}）へ保存する。
     */
    public void performGlobalUiFactoryReset() {
        TextInputDialog dialog = new TextInputDialog();
        if (primaryStage != null) {
            dialog.initOwner(primaryStage);
        }
        dialog.setTitle("確認");
        dialog.setHeaderText(null);
        dialog.setContentText(
                "タブ・表・テーマ等をバンドル既定に戻し、環境変数タブをテンプレートに戻します。"
                        + "誤操作防止のため、次のパスワードを入力してください。\nパスワード: 111");
        Optional<String> ans = dialog.showAndWait();
        if (ans.isEmpty() || !"111".equals(ans.get().trim())) {
            return;
        }

        performGlobalUiFactoryResetWithoutConfirmation();

        Alert done = new Alert(AlertType.INFORMATION);
        done.initOwner(primaryStage);
        applyAlertStylesheetsFromOwner(done);
        done.setTitle("完了");
        done.setHeaderText(null);
        done.setContentText("UI を既定に戻しました。");
        done.showAndWait();
    }

    /**
     * グローバル設定タブ「デフォルトに戻す」と同一の処理（確認ダイアログ・完了アラートなし）。
     *
     * <p>ポータブル自動バージョンアップ完了後に呼び出し、バンドル／{@code init_setting} 既定へ UI を揃える。
     */
    private void performGlobalUiFactoryResetWithoutConfirmation() {
        suppressEnvSessionPersistence.set(true);
        try {
            applyEnvRowsFullBundledResetAndPersist(false, FactorySite.KONAN);
            try {
                Files.deleteIfExists(TableColumnOrderPersistence.userHomeStorePath());
            } catch (IOException ignored) {
            }
            DispatchTrialLogUiStore.deleteStoreSilently();
            PlanWorkspaceSnapshotStore.deleteAllSilently();
            WorkspaceCacheArchiveStore.deleteAllSilently();
            PushButtonCssEmitter.deleteUserOverridesFileSilently();

            DesktopSessionState merged =
                    DesktopSessionStateStore.buildFactoryResetSession(collectDesktopSession(), collectUiEnv());
            /*
             * 環境変数タブは直前の applyEnvRowsFullBundledResetAndPersist で既にテンプレ＋ブートストラップ済み。
             * applyUiEnvRowsFromSession（true）でセッションを再適用すると、マージ JSON の uiEnvRows 欠落や
             * 早期 return と相性が悪く、アップグレード直後に「初期化されていない」ように見えることがあるため false。
             */
            applyDesktopSession(merged, false);
            TableColumnOrderPersistence.materializeTableColumnStoreAfterFactoryReset(collectUiEnv());
            applyDesktopThemeFromSession(merged);
            refreshDesktopSessionDependentUi();
            schedulePersistUserSessionAfterGlobalFactoryReset();
        } finally {
            suppressEnvSessionPersistence.set(false);
        }
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
        if (t == mainShellTabMemorySettings) {
            return MainShellTabId.MEMORY_SETTINGS;
        }
        if (t == mainShellTabGlobalSettings) {
            return MainShellTabId.GLOBAL_SETTINGS;
        }
        if (t == mainShellTabSummaryAiDispatchExportCustomize) {
            return MainShellTabId.SUMMARY_AI_DISPATCH_EXPORT_CUSTOMIZE;
        }
        if (t == mainShellTabUserProfiles) {
            return MainShellTabId.USER_PROFILES;
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
        if (t == mainShellTabCodeLookupTables) {
            return MainShellTabId.CODE_LOOKUP_TABLES;
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
        if (t == mainShellTabDeliveryCalendar) {
            return MainShellTabId.DELIVERY_CALENDAR_VIEW;
        }
        if (t == mainShellTabResultDispatch) {
            return MainShellTabId.RESULT_DISPATCH;
        }
        if (t == mainShellTabDispatchInteractive) {
            return MainShellTabId.DISPATCH_INTERACTIVE;
        }
        if (t == mainShellTabPlanWorkspaceHistory) {
            return MainShellTabId.PLAN_WORKSPACE_HISTORY;
        }
        if (t == mainShellTabCacheHistory) {
            return MainShellTabId.CACHE_HISTORY;
        }
        if (t == mainShellTabApiModelBenchmark) {
            return MainShellTabId.API_MODEL_BENCHMARK;
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
        if (t == mainShellTabOrganizer) {
            return MainShellTabId.TAB_ORGANIZER;
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
            case MEMORY_SETTINGS -> mainShellTabMemorySettings;
            case GLOBAL_SETTINGS -> mainShellTabGlobalSettings;
            case SUMMARY_AI_DISPATCH_EXPORT_CUSTOMIZE -> mainShellTabSummaryAiDispatchExportCustomize;
            case USER_PROFILES -> mainShellTabUserProfiles;
            case MASTER_SUMMARY -> mainShellTabMasterSummary;
            case PLAN_INPUT -> mainShellTabPlanInput;
            case STAGE1_PREVIEW -> mainShellTabStage1Preview;
            case CODE_LOOKUP_TABLES -> mainShellTabCodeLookupTables;
            case EXCLUDE_RULES -> mainShellTabExcludeRules;
            case SPECIAL_RULES -> mainShellTabSpecialRules;
            case ACTUALS_STATUS -> mainShellTabActualsStatus;
            case DELIVERY_CALENDAR_VIEW -> mainShellTabDeliveryCalendar;
            case RESULT_DISPATCH -> mainShellTabResultDispatch;
            case DISPATCH_INTERACTIVE -> mainShellTabDispatchInteractive;
            case PLAN_WORKSPACE_HISTORY -> mainShellTabPlanWorkspaceHistory;
            case CACHE_HISTORY -> mainShellTabCacheHistory;
            case API_MODEL_BENCHMARK -> mainShellTabApiModelBenchmark;
            case PLAN_RESULT_VIEWER -> mainShellTabPlanResultViewer;
            case EQUIPMENT_GANTT_GRAPHIC -> mainShellTabEquipmentGanttGraphic;
            case GANTT_PERSON_BADGE_DESIGN -> mainShellTabGanttPersonBadgeDesign;
            case OPERATOR_CARD -> mainShellTabOperatorCard;
            case TAB_ORGANIZER -> mainShellTabOrganizer;
        };
    }

    private List<String> snapshotMainShellTabOrder() {
        if (tabPane == null) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (Tab t : tabPane.getTabs()) {
            if (t == mainShellTabOrganizer) {
                continue;
            }
            flattenMainShellTabOrderKeys(t, out);
        }
        return List.copyOf(out);
    }

    private void flattenMainShellTabOrderKeys(Tab t, List<String> out) {
        if (t == null) {
            return;
        }
        if (t.getContent() instanceof TabPane inner) {
            for (Tab c : inner.getTabs()) {
                flattenMainShellTabOrderKeys(c, out);
            }
            return;
        }
        MainShellTabId id = mainShellTabId(t);
        if (id != null && id != MainShellTabId.TAB_ORGANIZER) {
            out.add(id.key());
        }
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

    private void emitShellTabNavigation() {
        Tab now = resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        Tab prev = lastEffectiveShellLeaf;
        lastEffectiveShellLeaf = now;
        if (prev == mainShellTabRun && now != mainShellTabRun) {
            DesktopSessionStateStore.save(collectDesktopSession());
        }
    }

    /**
     * ルートで選ばれているタブがグループのときは、その内側の選択タブまで辿ったリーフ（実タブ）を返す。
     */
    private Tab resolveEffectiveLeafTab(Tab rootSelected) {
        if (rootSelected == null) {
            return null;
        }
        if (rootSelected.getContent() instanceof TabPane inner) {
            Tab innerSel = inner.getSelectionModel().getSelectedItem();
            if (innerSel != null) {
                return resolveEffectiveLeafTab(innerSel);
            }
            if (!inner.getTabs().isEmpty()) {
                return resolveEffectiveLeafTab(inner.getTabs().getFirst());
            }
            return null;
        }
        return rootSelected;
    }

    private void captureMainShellTabBaselineTitles() {
        mainShellTabBaselineTitles.clear();
        for (MainShellTabId id : MainShellTabId.values()) {
            if (id == MainShellTabId.TAB_ORGANIZER) {
                continue;
            }
            Tab t = mainShellTabFor(id);
            if (t != null) {
                String tx = t.getText();
                mainShellTabBaselineTitles.put(
                        id, tx != null && !tx.isBlank() ? tx.strip() : id.name());
            }
        }
    }

    private Map<String, String> snapshotMainShellTabTitleAliases() {
        return Map.copyOf(mainShellTabTitleAliases);
    }

    private void refreshMainShellTabDisplayedTitles() {
        for (MainShellTabId id : MainShellTabId.values()) {
            if (id == MainShellTabId.TAB_ORGANIZER) {
                continue;
            }
            Tab t = mainShellTabFor(id);
            if (t != null) {
                t.setText(mainShellTabTitle(id));
            }
        }
    }

    private void applyMainShellTabTitleAliasesFromSession(Map<String, String> fromSession) {
        mainShellTabTitleAliases.clear();
        if (fromSession != null) {
            for (Map.Entry<String, String> e : fromSession.entrySet()) {
                if (e.getKey() != null
                        && !e.getKey().isBlank()
                        && e.getValue() != null
                        && !e.getValue().isBlank()) {
                    mainShellTabTitleAliases.put(e.getKey().trim(), e.getValue().strip());
                }
            }
        }
        refreshMainShellTabDisplayedTitles();
    }

    private Map<String, Integer> snapshotInnerTabSelectedIndexByShellTabKey() {
        LinkedHashMap<String, Integer> m = new LinkedHashMap<>();
        if (deliveryCalendarViewTabController != null) {
            int i = deliveryCalendarViewTabController.snapshotInnerTabSelectedIndex();
            if (i >= 0) {
                m.put(MainShellTabId.DELIVERY_CALENDAR_VIEW.key(), i);
            }
        }
        if (dispatchInteractiveTabController != null) {
            int i = dispatchInteractiveTabController.snapshotInnerTabSelectedIndex();
            if (i >= 0) {
                m.put(MainShellTabId.DISPATCH_INTERACTIVE.key(), i);
            }
        }
        if (codeDispatchLookupTablesTabController != null) {
            int i = codeDispatchLookupTablesTabController.snapshotInnerTabSelectedIndex();
            if (i >= 0) {
                m.put(MainShellTabId.CODE_LOOKUP_TABLES.key(), i);
            }
        }
        return Map.copyOf(m);
    }

    private void applyInnerTabSelectionsFromSession(Map<String, Integer> map) {
        if (map == null || map.isEmpty()) {
            return;
        }
        Platform.runLater(
                () -> {
                    Integer dc = map.get(MainShellTabId.DELIVERY_CALENDAR_VIEW.key());
                    if (dc != null && deliveryCalendarViewTabController != null) {
                        deliveryCalendarViewTabController.applyInnerTabSelectedIndex(dc.intValue());
                    }
                    Integer di = map.get(MainShellTabId.DISPATCH_INTERACTIVE.key());
                    if (di != null && dispatchInteractiveTabController != null) {
                        dispatchInteractiveTabController.applyInnerTabSelectedIndex(di.intValue());
                    }
                    Integer lk = map.get(MainShellTabId.CODE_LOOKUP_TABLES.key());
                    if (lk != null && codeDispatchLookupTablesTabController != null) {
                        codeDispatchLookupTablesTabController.applyInnerTabSelectedIndex(lk.intValue());
                    }
                });
    }

    private List<MainShellTabLayoutNode> snapshotMainShellTabLayout() {
        if (tabPane == null) {
            return List.of();
        }
        List<MainShellTabLayoutNode> top = new ArrayList<>();
        for (Tab t : tabPane.getTabs()) {
            if (t == mainShellTabOrganizer) {
                continue;
            }
            MainShellTabLayoutNode n = snapshotMainShellLayoutNode(t);
            if (n != null) {
                top.add(n);
            }
        }
        return List.copyOf(top);
    }

    private MainShellTabLayoutNode snapshotMainShellLayoutNode(Tab t) {
        if (t == null) {
            return null;
        }
        Node content = resolveMainShellTabContentForSnapshot(t);
        if (content instanceof TabPane inner) {
            List<MainShellTabLayoutNode> ch = new ArrayList<>();
            for (Tab c : inner.getTabs()) {
                MainShellTabLayoutNode cn = snapshotMainShellLayoutNode(c);
                if (cn != null) {
                    ch.add(cn);
                }
            }
            String title = t.getText() != null && !t.getText().isBlank() ? t.getText() : "グループ";
            return MainShellTabLayoutNode.groupNode(title, readShellTabColorHex(t), ch);
        }
        MainShellTabId id = mainShellTabId(t);
        if (id == null || id == MainShellTabId.TAB_ORGANIZER) {
            return null;
        }
        return MainShellTabLayoutNode.tabNode(id.key(), readShellTabColorHex(t));
    }

    /**
     * 遅延ロードで {@link Tab#setContent} がプレースホルダのとき、退避中の実コンテンツをスナップショットに使う。
     * これが無いとグループ見出しタブが単独リーフとして保存され、タブ整理ツリーにグループが現れない。
     */
    private Node resolveMainShellTabContentForSnapshot(Tab t) {
        if (t == null) {
            return null;
        }
        Node content = t.getContent();
        if (isLazyMainShellTabPlaceholder(content)) {
            Object detached = t.getProperties().get(PM_DEFERRED_TAB_CONTENT);
            if (detached instanceof Node node) {
                return node;
            }
        }
        return content;
    }

    private static String readShellTabColorHex(Tab t) {
        Object v = t.getProperties().get("pmShellTabColor");
        return v instanceof String s && !s.isBlank() ? s.strip() : "";
    }

    /** メインシェル {@link TabPane} の見出し CSS 再適用（スキン生成遅延・選択切替後の黒文字化対策）。 */
    private void installMainShellTabPaneChromeHooks() {
        if (tabPane == null) {
            return;
        }
        if (!tabPane.getStyleClass().contains("pm-main-shell-tab-pane")) {
            tabPane.getStyleClass().add("pm-main-shell-tab-pane");
        }
        if (Boolean.TRUE.equals(tabPane.getProperties().get("pmMainShellTabChromeHooksInstalled"))) {
            return;
        }
        tabPane.getProperties().put("pmMainShellTabChromeHooksInstalled", Boolean.TRUE);
        tabPane
                .skinProperty()
                .addListener(
                        (obs, oldSkin, newSkin) -> {
                            if (newSkin != null && !suppressMainShellTabChromeRefresh.get()) {
                                Platform.runLater(this::refreshMainShellTabHeaderChromeFromStoredColors);
                            }
                        });
    }

    /**
     * 選択／非選択の切り替えでテーマ CSS が見出しを塗り直し、インライン前景が潰れることがあるため、保存色があれば再適用する。
     */
    private void refreshShellTabChromeOnSelectionChange(Tab tab) {
        if (tab == null) {
            return;
        }
        String hex = readShellTabColorHex(tab);
        if (hex.isEmpty()) {
            return;
        }
        applyShellTabColor(tab, hex);
    }

    private void ensureShellTabSelectionChromeListener(Tab tab) {
        if (tab == null
                || Boolean.TRUE.equals(
                        tab.getProperties().get(PROP_SHELL_TAB_SELECTION_CHROME_LISTENER))) {
            return;
        }
        tab.getProperties().put(PROP_SHELL_TAB_SELECTION_CHROME_LISTENER, Boolean.TRUE);
        tab.selectedProperty()
                .addListener(
                        (obs, was, now) ->
                                Platform.runLater(() -> refreshShellTabChromeOnSelectionChange(tab)));
    }

    private void applyShellTabColor(Tab tab, String colorHex) {
        if (tab == null) {
            return;
        }
        TabPane pane = tab.getTabPane();
        if (colorHex != null && !colorHex.isBlank()) {
            String h = colorHex.strip();
            tab.getProperties().put("pmShellTabColor", h);
            String textFill = contrastingTabLabelTextFillHex(h);
            String glowEffect =
                    mainShellTabOrganizerHeaderGlowEnabled.get()
                            ? shellTabHeaderGlowEffectCss(h)
                            : "";
            tab.setStyle(shellTabHeaderChromeInlineStyle(h, textFill, glowEffect));
            pokeShellTabHeaderBackground(collectUiEnv(), pane, tab, h, textFill, glowEffect);
        } else {
            tab.getProperties().remove("pmShellTabColor");
            tab.setStyle("");
            pokeShellTabHeaderBackground(collectUiEnv(), pane, tab, null, null, null);
        }
        ensureShellTabSelectionChromeListener(tab);
    }

    public boolean isMainShellTabOrganizerHeaderGlowEnabled() {
        return mainShellTabOrganizerHeaderGlowEnabled.get();
    }

    public void setMainShellTabOrganizerHeaderGlowEnabled(boolean enabled) {
        mainShellTabOrganizerHeaderGlowEnabled.set(enabled);
    }

    public double getMainShellTabOrganizerHeaderGlowStrength() {
        Double v = mainShellTabOrganizerHeaderGlowStrength.get();
        double x = v != null ? v : 1.0;
        return clamp(x, 0.0, 1.0);
    }

    public void setMainShellTabOrganizerHeaderGlowStrength(double strength01) {
        mainShellTabOrganizerHeaderGlowStrength.set(clamp(strength01, 0.0, 1.0));
    }

    /** 保存済みの {@code pmShellTabColor} を踏まえて全タブ見出しのインラインスタイルを再適用（グロー切替時）。 */
    public void refreshMainShellTabHeaderChromeFromStoredColors() {
        if (tabPane == null || suppressMainShellTabChromeRefresh.get()) {
            return;
        }
        if (tabPane.getScene() == null) {
            Platform.runLater(this::refreshMainShellTabHeaderChromeFromStoredColors);
            return;
        }
        applyStoredShellTabColorsRecursive(tabPane.getTabs());
        layoutShellTabPanesRecursive(tabPane);
    }

    /**
     * 入れ子 {@link TabPane} まで {@code applyCss}/{@code layout} し、見出しセル（{@code .headers-region}）の取りこぼしを減らす。
     */
    private static void layoutShellTabPanesRecursive(TabPane pane) {
        if (pane == null) {
            return;
        }
        pane.applyCss();
        pane.requestLayout();
        for (Tab t : pane.getTabs()) {
            if (t.getContent() instanceof TabPane inner) {
                layoutShellTabPanesRecursive(inner);
            }
        }
    }

    private void applyOrDeferMainShellTabLayoutFromSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        if (primaryScene != null && tabPane != null && tabPane.getScene() != null) {
            applyMainShellTabLayoutFromSession(s);
            pendingMainShellTabLayoutSession = null;
        } else {
            pendingMainShellTabLayoutSession = s;
        }
    }

    private void applyPendingMainShellTabLayoutFromSessionIfNeeded() {
        if (pendingMainShellTabLayoutSession == null) {
            return;
        }
        DesktopSessionState s = pendingMainShellTabLayoutSession;
        pendingMainShellTabLayoutSession = null;
        applyMainShellTabLayoutFromSession(s);
    }

    private void applyMainShellTabLayoutFromSession(DesktopSessionState s) {
        if (s == null || tabPane == null) {
            return;
        }
        suppressMainShellTabChromeRefresh.set(true);
        suppressLazyMainShellTabContentSwap.set(true);
        try {
            if (s.mainShellTabLayout() != null && !s.mainShellTabLayout().isEmpty()) {
                if (!rebuildMainShellTabsFromLayout(s.mainShellTabLayout())
                        && !rebuildMainShellTabsFromLayout(
                                flatMainShellTabLayoutFromOrderKeys(s.mainShellTabOrder()))) {
                    rebuildMainShellTabsFromLayout(null);
                }
            } else if (!rebuildMainShellTabsFromLayout(
                            flatMainShellTabLayoutFromOrderKeys(s.mainShellTabOrder()))
                    && !rebuildMainShellTabsFromLayout(null)) {
                applyMainShellTabOrder(s.mainShellTabOrder());
            }
            applyMainShellTabTitleAliasesFromSession(s.mainShellTabTitleAliases());
            applyInnerTabSelectionsFromSession(s.innerTabSelectedIndexByShellTabKey());
            lastEffectiveShellLeaf =
                    resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        } finally {
            suppressLazyMainShellTabContentSwap.set(false);
            suppressMainShellTabChromeRefresh.set(false);
            installLazyMainShellTabContentForStartup();
        }
    }

    /**
     * 初回 {@link Scene#doLayoutPass} 前に、全作業タブの Spreadsheet 等をシーンから外す。
     * 非表示タブまで FXML 読込で載ると {@code IndexOutOfBoundsException}（index 19, length 19）になりやすい。
     */
    private void installLazyMainShellTabContentForStartup() {
        if (tabPane == null) {
            return;
        }
        suppressLazyMainShellTabContentSwap.set(true);
        try {
            for (Tab t : tabPane.getTabs()) {
                if (t == mainShellTabOrganizer) {
                    continue;
                }
                deferMainShellTabHeavyContentRecursive(t);
                if (t.getContent() instanceof TabPane inner) {
                    for (Tab innerTab : inner.getTabs()) {
                        deferMainShellTabHeavyContentRecursive(innerTab);
                    }
                }
            }
        } finally {
            suppressLazyMainShellTabContentSwap.set(false);
        }
    }

    private void deferMainShellTabBranchHeavyContent(Tab tab) {
        if (tab == null) {
            return;
        }
        deferMainShellTabHeavyContentRecursive(tab);
        if (tab.getContent() instanceof TabPane inner) {
            for (Tab innerTab : inner.getTabs()) {
                deferMainShellTabHeavyContentRecursive(innerTab);
            }
        }
    }

    private void deferMainShellTabHeavyContentRecursive(Tab tab) {
        if (tab == null || isLazyMainShellTabPlaceholder(tab.getContent())) {
            return;
        }
        Node content = tab.getContent();
        if (content == null) {
            return;
        }
        tab.getProperties().put(PM_DEFERRED_TAB_CONTENT, content);
        Region placeholder = new Region();
        placeholder.setMinSize(0, 0);
        placeholder.setPrefSize(0, 0);
        placeholder.getProperties().put(PM_LAZY_TAB_PLACEHOLDER, Boolean.TRUE);
        tab.setContent(placeholder);
    }

    /**
     * 配台計画手動修正タブで {@link SpreadsheetView#setGrid} する直前に呼ぶ。メインシェル遅延ロードで
     * プレースホルダに差し替えられていると、再構築しても画面に反映されない。
     */
    void ensureDispatchInteractiveReadyForGridRebuild() {
        if (mainShellTabDispatchInteractive == null) {
            return;
        }
        boolean prev = suppressLazyMainShellTabContentSwap.get();
        suppressLazyMainShellTabContentSwap.set(true);
        try {
            restoreDeferredTabContent(mainShellTabDispatchInteractive);
            if (dispatchInteractiveTabController != null) {
                dispatchInteractiveTabController.ensureInnerTabsMaterializedForRebuild();
            }
        } finally {
            suppressLazyMainShellTabContentSwap.set(prev);
        }
    }

    /**
     * 配台 Spreadsheet をシーングラフ上に載せてから {@code setGrid} する。未選択タブのコンテンツは
     * {@link javafx.scene.Node#getScene()} が null のままになり、オフシーンでの再構築は空表示・IOOBE の原因になる。
     *
     * @param forceSelectTab {@code true} のとき配台タブが未選択なら選択する（手動「再読み」向け）
     */
    void ensureDispatchInteractiveOnSceneForGridRebuild(boolean forceSelectTab) {
        ensureDispatchInteractiveReadyForGridRebuild();
        if (mainShellTabDispatchInteractive == null || tabPane == null) {
            return;
        }
        Tab effective = resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        if (effective != mainShellTabDispatchInteractive && forceSelectTab) {
            selectMainShellTab(MainShellTabId.DISPATCH_INTERACTIVE);
        }
        boolean prev = suppressLazyMainShellTabContentSwap.get();
        suppressLazyMainShellTabContentSwap.set(true);
        try {
            activateMainShellTabHeavyContentRecursive(mainShellTabDispatchInteractive);
            ensureDispatchInteractiveReadyForGridRebuild();
        } finally {
            suppressLazyMainShellTabContentSwap.set(prev);
        }
    }

    void restoreDeferredTabContent(Tab tab) {
        if (tab == null) {
            return;
        }
        Object detached = tab.getProperties().remove(PM_DEFERRED_TAB_CONTENT);
        if (detached instanceof Node node) {
            tab.setContent(node);
        }
    }

    private void activateMainShellTabHeavyContentRecursive(Tab tab) {
        if (tab == null) {
            return;
        }
        restoreDeferredTabContent(tab);
        Node content = tab.getContent();
        if (!(content instanceof TabPane inner)) {
            return;
        }
        Tab innerSelected = inner.getSelectionModel().getSelectedItem();
        if (innerSelected == null && !inner.getTabs().isEmpty()) {
            inner.getSelectionModel().select(0);
            innerSelected = inner.getTabs().getFirst();
        }
        for (Tab innerTab : inner.getTabs()) {
            if (innerTab != innerSelected) {
                deferMainShellTabHeavyContentRecursive(innerTab);
            }
        }
        activateMainShellTabHeavyContentRecursive(innerSelected);
    }

    private static boolean isLazyMainShellTabPlaceholder(Node content) {
        return content != null
                && Boolean.TRUE.equals(content.getProperties().get(PM_LAZY_TAB_PLACEHOLDER));
    }

    private void applyStoredShellTabColorsRecursive(ObservableList<Tab> tabs) {
        if (tabs == null) {
            return;
        }
        for (Tab t : tabs) {
            if (t == mainShellTabOrganizer) {
                continue;
            }
            applyShellTabColor(t, readShellTabColorHex(t));
            Node content = resolveMainShellTabContentForSnapshot(t);
            if (content instanceof TabPane inner) {
                applyStoredShellTabColorsRecursive(inner.getTabs());
            }
        }
    }

    /** タブ整理ツリーのミニプレビュー（チップ）のインラインスタイル。 */
    public String tabOrganizerPreviewChipSurfaceStyle(String colorHexOrEmpty) {
        if (colorHexOrEmpty == null || colorHexOrEmpty.isBlank()) {
            return "";
        }
        String h = colorHexOrEmpty.strip();
        StringBuilder sb =
                new StringBuilder()
                        .append("-fx-background-color: ")
                        .append(h)
                        .append("; -fx-background-radius: 5; -fx-border-radius: 5; -fx-border-width: 1; ")
                        .append("-fx-border-color: ")
                        .append(previewChipBorderRgba(h))
                        .append("; ");
        if (mainShellTabOrganizerHeaderGlowEnabled.get()) {
            String g = shellTabHeaderGlowEffectCss(h);
            if (!g.isBlank()) {
                sb.append("-fx-effect: ").append(g).append("; ");
            }
        }
        return sb.toString().trim();
    }

    /**
     * タブ整理ツリー上の色ピル用（メイン見出しのグロー設定に依存しないフラットな面スタイル）。
     */
    public String tabOrganizerTreePillSurfaceStyle(String colorHexOrEmpty) {
        if (colorHexOrEmpty == null || colorHexOrEmpty.isBlank()) {
            return "";
        }
        String h = colorHexOrEmpty.strip();
        return ("-fx-background-color: "
                        + h
                        + "; -fx-background-radius: 6; -fx-border-radius: 6; -fx-border-width: 1; "
                        + "-fx-border-color: "
                        + previewChipBorderRgba(h)
                        + ";")
                .trim();
    }

    public String tabOrganizerPreviewChipLabelTextFill(String colorHexOrEmpty) {
        if (colorHexOrEmpty == null || colorHexOrEmpty.isBlank()) {
            return "#94a3b8";
        }
        return contrastingTabLabelTextFillHex(colorHexOrEmpty.strip());
    }

    private static String previewChipBorderRgba(String bgHex) {
        try {
            Color c = Color.web(bgHex.strip());
            return String.format(
                    Locale.US,
                    "rgba(%d,%d,%d,0.40)",
                    clamp255((int) Math.round(c.getRed() * 255.0)),
                    clamp255((int) Math.round(c.getGreen() * 255.0)),
                    clamp255((int) Math.round(c.getBlue() * 255.0)));
        } catch (IllegalArgumentException ex) {
            return "rgba(148,163,184,0.65)";
        }
    }

    /**
     * 見出し背景に連動した半透明のガウシアン {@code dropshadow} でグロー風の縁取り。
     * 強さは {@link #getMainShellTabOrganizerHeaderGlowStrength()} でスケールする（0 で効果なし）。
     *
     * @return CSS の {@code -fx-effect} に渡す値（{@code dropshadow(...)}）。失敗時は空。
     */
    private String shellTabHeaderGlowEffectCss(String hexBg) {
        double strength = clamp(getMainShellTabOrganizerHeaderGlowStrength(), 0.0, 1.0);
        if (strength <= 1e-6) {
            return "";
        }
        try {
            Color c = Color.web(hexBg.strip());
            double alpha = 0.62 * strength;
            double radius = 14.0 * strength;
            double spread = 0.38 * strength;
            String rgba =
                    String.format(
                            Locale.US,
                            "rgba(%d,%d,%d,%.4f)",
                            clamp255((int) Math.round(c.getRed() * 255.0)),
                            clamp255((int) Math.round(c.getGreen() * 255.0)),
                            clamp255((int) Math.round(c.getBlue() * 255.0)),
                            alpha);
            return "dropshadow(gaussian, "
                    + rgba
                    + ", "
                    + String.format(Locale.US, "%.2f", radius)
                    + ", "
                    + String.format(Locale.US, "%.2f", spread)
                    + ", 0, 0)";
        } catch (IllegalArgumentException ex) {
            return "";
        }
    }

    private static String shellTabHeaderChromeInlineStyle(
            String bgHex, String labelFillHex, String glowEffectCssValue) {
        StringBuilder sb =
                new StringBuilder()
                        .append("-fx-background-color: ")
                        .append(bgHex.strip())
                        .append("; -fx-text-fill: ")
                        .append(labelFillHex.strip())
                        .append(";");
        if (glowEffectCssValue != null && !glowEffectCssValue.isBlank()) {
            sb.append(" -fx-effect: ").append(glowEffectCssValue.strip()).append(";");
        }
        return sb.toString();
    }

    /**
     * ユーザー指定のタブ背景に対し、WCAG 系の相対輝度で明暗を判定してラベル色を選ぶ（固定の白／グレー文字との衝突を避ける）。
     */
    private static String contrastingTabLabelTextFillHex(String bgHex) {
        try {
            Color c = Color.web(bgHex.strip());
            double lum =
                    relativeSrgbLuminance(
                            (int) Math.round(c.getRed() * 255.0),
                            (int) Math.round(c.getGreen() * 255.0),
                            (int) Math.round(c.getBlue() * 255.0));
            return lum > 0.45 ? "#0f172a" : "#f8fafc";
        } catch (IllegalArgumentException ex) {
            return "#f8fafc";
        }
    }

    /** sRGB の相対輝度（0〜1）。{@link Color} と同じ係数。 */
    private static double relativeSrgbLuminance(int r, int g, int b) {
        double rs = linearizeSrgbChannel(clamp255(r) / 255.0);
        double gs = linearizeSrgbChannel(clamp255(g) / 255.0);
        double bs = linearizeSrgbChannel(clamp255(b) / 255.0);
        return 0.2126 * rs + 0.7152 * gs + 0.0722 * bs;
    }

    private static int clamp255(int x) {
        return Math.max(0, Math.min(255, x));
    }

    private static double linearizeSrgbChannel(double channel01) {
        if (channel01 <= 0.03928) {
            return channel01 / 12.92;
        }
        return Math.pow((channel01 + 0.055) / 1.055, 2.4);
    }

    /**
     * タブ見出しラベル（{@code .tab-label}）以下の {@link Text} にも前景色を適用する。Modena の {@code .tab-label}
     * は {@code Labeled} に対する {@code -fx-text-fill} と子 {@link Text} の {@code -fx-fill} が一致しないことがあり、タブ整理のプレビュー（単純
     * {@link Label}）と実タブで文字色だけずれる原因になる。
     * <p>JavaFX 26 以降、タブ見出しの {@code LabeledText} などでは {@code fill} が CSS 側でバインドされることがあり、
     * {@link Text#setFill} が例外になる。そのため {@link Text} 系はインライン {@code -fx-fill} のみで指定する。
     */
    private static void applyShellTabHeaderForegroundRecursive(
            Node root, Color fillColor, String tfHex) {
        if (root == null || tfHex == null || tfHex.isBlank()) {
            return;
        }
        String tf = tfHex.strip();
        if (root instanceof Text textNode) {
            textNode.setStyle("-fx-fill: " + tf + ";");
        } else if (root instanceof Labeled labeled) {
            /* TabSkin のバインドと干渉しないよう、可能なときだけ直接指定（主に -fx-text-fill） */
            if (!labeled.textFillProperty().isBound()) {
                labeled.setTextFill(fillColor);
            }
            labeled.setStyle("-fx-text-fill: " + tf + ";");
        }
        if (root instanceof Parent p) {
            for (Node ch : p.getChildrenUnmodifiable()) {
                applyShellTabHeaderForegroundRecursive(ch, fillColor, tf);
            }
        }
    }

    /** 着色解除時に {@link #applyShellTabHeaderForegroundRecursive} で付けたインラインを除去する。 */
    private static void clearShellTabHeaderForegroundRecursive(Node root) {
        if (root == null) {
            return;
        }
        if (root instanceof Text textNode) {
            textNode.setStyle("");
        } else if (root instanceof Labeled labeled) {
            labeled.setStyle("");
            if (!labeled.textFillProperty().isBound()) {
                labeled.setTextFill(null);
            }
        }
        if (root instanceof Parent p) {
            for (Node ch : p.getChildrenUnmodifiable()) {
                clearShellTabHeaderForegroundRecursive(ch);
            }
        }
    }

    /** デバッグ計測：{@code .tab-label} サブツリー内の最初の {@link Text} の {@code fill}。 */
    private static String firstTabLabelDescendantTextFillString(Node root) {
        if (root instanceof Text t) {
            javafx.scene.paint.Paint f = t.getFill();
            return f != null ? f.toString() : "";
        }
        if (root instanceof Parent p) {
            for (Node ch : p.getChildrenUnmodifiable()) {
                String s = firstTabLabelDescendantTextFillString(ch);
                if (!s.isEmpty()) {
                    return s;
                }
            }
        }
        return "";
    }

    /**
     * テーマ CSS の {@code .tab-pane > ... > .tab:selected} 等が Tab のインラインより強く当たり色が変わらないことがあるため、
     * 見出し行のセル（{@code .headers-region} 直下の {@code .tab}）へ直接背景・文字色を指定する。
     */
    private static void applyShellTabHeaderCellChrome(
            Node tabHeaderCell,
            String bgHex,
            String labelFillHex,
            String glowEffectCssOrNull) {
        String tf = labelFillHex.strip();
        tabHeaderCell.setStyle(shellTabHeaderChromeInlineStyle(bgHex, tf, glowEffectCssOrNull));
        if (!tabHeaderCell.getStyleClass().contains("pm-shell-tab-colored")) {
            tabHeaderCell.getStyleClass().add("pm-shell-tab-colored");
        }
        String labelInline = "-fx-text-fill: " + tf + ";";
        Node lab = tabHeaderCell.lookup(".tab-label");
        if (lab != null) {
            lab.setStyle(labelInline);
        }
        try {
            applyShellTabHeaderForegroundRecursive(tabHeaderCell, Color.web(tf), tf);
        } catch (IllegalArgumentException ex) {
            if (lab instanceof Labeled labeled) {
                labeled.setStyle(labelInline);
            }
        }
    }

    private static void clearShellTabHeaderCellChrome(Node tabHeaderCell) {
        tabHeaderCell.setStyle("");
        tabHeaderCell.getStyleClass().remove("pm-shell-tab-colored");
        Node lab = tabHeaderCell.lookup(".tab-label");
        if (lab != null) {
            lab.setStyle("");
        }
        clearShellTabHeaderForegroundRecursive(tabHeaderCell);
    }

    private static void pokeShellTabHeaderBackground(
            Map<String, String> uiEnv,
            TabPane pane,
            Tab tab,
            String rgbHexOrNull,
            String labelFillHexOrNull,
            String glowEffectCssOrNull) {
        if (pane == null) {
            return;
        }
        Runnable op =
                () -> {
                    int idx = pane.getTabs().indexOf(tab);
                    if (idx < 0) {
                        return;
                    }
                    Node headersRegion = pane.lookup(".headers-region");
                    if (!(headersRegion instanceof Parent hp)) {
                        return;
                    }
                    int tabOrdinal = 0;
                    for (Node child : hp.getChildrenUnmodifiable()) {
                        if (!child.getStyleClass().contains("tab")) {
                            continue;
                        }
                        if (tabOrdinal == idx) {
                            if (rgbHexOrNull != null
                                    && !rgbHexOrNull.isBlank()
                                    && labelFillHexOrNull != null
                                    && !labelFillHexOrNull.isBlank()) {
                                applyShellTabHeaderCellChrome(
                                        child,
                                        rgbHexOrNull.strip(),
                                        labelFillHexOrNull.strip(),
                                        glowEffectCssOrNull);
                            } else {
                                clearShellTabHeaderCellChrome(child);
                            }
                            return;
                        }
                        tabOrdinal++;
                    }
                };
        op.run();
        Platform.runLater(op);
        Platform.runLater(() -> Platform.runLater(op));
    }

    /**
     * タブ整理ツリーで編集した見出し色を、メインシェル上部のタブへ即時反映する（並び替えはしない）。
     *
     * <p>作業タブ（リーフ）は {@link MainShellTabId} で一意に付け替え、グループ見出しの色は「そのグループに含まれる作業タブキーの集合」が一致する
     * メインシェル上のグループタブへ適用する（並びがツリーと異なっていてもインデックスでは突き合わせない）。
     */
    void syncMainShellTabHeaderColorsFromOrganizerTree(
            TreeItem<MainShellTabOrganizerTabController.OrgRow> invisibleRoot) {
        if (tabPane == null || invisibleRoot == null) {
            return;
        }
        syncLeafTabColorsFromOrganizerTree(invisibleRoot);
        syncGroupTabHeadersFromOrganizerTree(invisibleRoot);
        /* 同一フレームで見出しへ反映（runLater のみだと未レイアウトで poke が無効になることがある） */
        refreshMainShellTabHeaderChromeFromStoredColors();
        Platform.runLater(this::refreshMainShellTabHeaderChromeFromStoredColors);
    }

    private void syncLeafTabColorsFromOrganizerTree(TreeItem<MainShellTabOrganizerTabController.OrgRow> node) {
        if (node == null) {
            return;
        }
        MainShellTabOrganizerTabController.OrgRow r = node.getValue();
        if (r != null && r.kind == MainShellTabOrganizerTabController.OrgRow.Kind.TAB) {
            Tab t = mainShellTabFor(r.tabId);
            if (t != null) {
                applyShellTabColor(t, r.colorHex);
            }
        }
        for (TreeItem<MainShellTabOrganizerTabController.OrgRow> c : node.getChildren()) {
            syncLeafTabColorsFromOrganizerTree(c);
        }
    }

    private void syncGroupTabHeadersFromOrganizerTree(
            TreeItem<MainShellTabOrganizerTabController.OrgRow> invisibleRoot) {
        List<Tab> shellTop = new ArrayList<>();
        for (Tab t : tabPane.getTabs()) {
            if (t != mainShellTabOrganizer) {
                shellTop.add(t);
            }
        }
        syncGroupHeaderColorsForTreeLevel(invisibleRoot.getChildren(), shellTop);
    }

    /**
     * ツリー上の各グループ行に対し、メインシェル側の「同じ descendant タブキー集合」を持つグループ Tab を探して見出し色を適用する。
     */
    private void syncGroupHeaderColorsForTreeLevel(
            List<TreeItem<MainShellTabOrganizerTabController.OrgRow>> treeLevel, List<Tab> shellTabsAtLevel) {
        if (treeLevel == null || shellTabsAtLevel == null) {
            return;
        }
        List<Tab> unmatched = new ArrayList<>(shellTabsAtLevel);
        for (TreeItem<MainShellTabOrganizerTabController.OrgRow> ti : treeLevel) {
            MainShellTabOrganizerTabController.OrgRow r = ti.getValue();
            if (r == null || r.kind != MainShellTabOrganizerTabController.OrgRow.Kind.GROUP) {
                continue;
            }
            Tab match = findShellGroupTabWithSameLeafKeys(ti, unmatched);
            if (match != null
                    && match.getContent() instanceof TabPane inner) {
                applyShellTabColor(match, r.colorHex);
                unmatched.remove(match);
                syncGroupHeaderColorsForTreeLevel(ti.getChildren(), new ArrayList<>(inner.getTabs()));
            }
        }
    }

    /**
     * {@code candidates} のうち、配下の作業タブキー集合がツリー上のグループ {@code groupItem} と一致する TabPane 付きタブを返す。
     */
    private Tab findShellGroupTabWithSameLeafKeys(
            TreeItem<MainShellTabOrganizerTabController.OrgRow> groupItem, List<Tab> candidates) {
        Set<String> wanted = new HashSet<>();
        collectOrganizerDescendantTabKeys(groupItem, wanted);
        for (Tab t : candidates) {
            if (!(t.getContent() instanceof TabPane)) {
                continue;
            }
            Set<String> have = new HashSet<>();
            collectShellTabSubtreeLeafKeys(t, have);
            if (wanted.equals(have)) {
                return t;
            }
        }
        return null;
    }

    private static void collectOrganizerDescendantTabKeys(
            TreeItem<MainShellTabOrganizerTabController.OrgRow> node, Set<String> out) {
        if (node == null) {
            return;
        }
        MainShellTabOrganizerTabController.OrgRow r = node.getValue();
        if (r != null && r.kind == MainShellTabOrganizerTabController.OrgRow.Kind.TAB) {
            out.add(r.tabId.key());
        }
        for (TreeItem<MainShellTabOrganizerTabController.OrgRow> c : node.getChildren()) {
            collectOrganizerDescendantTabKeys(c, out);
        }
    }

    /** シェル上の Tab（リーフまたは入れ子グループ）の配下にあるすべての作業タブ ID キーを収集する。 */
    private void collectShellTabSubtreeLeafKeys(Tab t, Set<String> out) {
        if (t == null) {
            return;
        }
        MainShellTabId id = mainShellTabId(t);
        if (id != null && id != MainShellTabId.TAB_ORGANIZER) {
            out.add(id.key());
            return;
        }
        if (t.getContent() instanceof TabPane inner) {
            for (Tab c : inner.getTabs()) {
                collectShellTabSubtreeLeafKeys(c, out);
            }
        }
    }

    /**
     * @return レイアウトが検証されメイン {@link TabPane} が組み替えられたとき {@code true}。検証不一致などでスキップしたとき {@code false}
     */
    private boolean rebuildMainShellTabsFromLayout(List<MainShellTabLayoutNode> layout) {
        if (tabPane == null || mainShellTabOrganizer == null) {
            return false;
        }
        List<MainShellTabLayoutNode> prepared = prepareMainShellLayoutForRebuild(layout);
        HashSet<String> required = requiredShellTabKeys();
        HashSet<String> found = new HashSet<>();
        for (MainShellTabLayoutNode n : prepared) {
            collectLayoutLeafKeys(n, found);
        }
        if (!found.equals(required)) {
            return false;
        }
        suppressEnvSessionPersistence.set(true);
        suppressMainShellTabChromeRefresh.set(true);
        try {
            wiredInnerMainShellTabPanes.clear();
            tabPane.getTabs().clear();
            for (MainShellTabLayoutNode n : prepared) {
                Tab built = materializeLayoutNode(n);
                if (built != null) {
                    tabPane.getTabs().add(built);
                }
            }
            tabPane.getTabs().add(mainShellTabOrganizer);
            boolean nested = prepared.stream().anyMatch(MainShellTabLayoutNode::isGroup);
            tabPane.setTabDragPolicy(
                    nested
                            ? TabPane.TabDragPolicy.FIXED
                            : TabPane.TabDragPolicy.REORDER);
            for (TabPane inner : wiredInnerMainShellTabPanes) {
                inner.getSelectionModel()
                        .selectedItemProperty()
                        .addListener(
                                (o, p, n) -> {
                                    if (!suppressLazyMainShellTabContentSwap.get()) {
                                        deferMainShellTabBranchHeavyContent(p);
                                        activateMainShellTabHeavyContentRecursive(n);
                                    }
                                    emitShellTabNavigation();
                                    if (!suppressMainShellTabChromeRefresh.get()) {
                                        refreshMainShellTabHeaderChromeFromStoredColors();
                                    }
                                });
            }
        } finally {
            suppressMainShellTabChromeRefresh.set(false);
            suppressEnvSessionPersistence.set(false);
        }
        refreshMainShellTabDisplayedTitles();
        lastEffectiveShellLeaf =
                resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        Platform.runLater(this::refreshMainShellTabHeaderChromeFromStoredColors);
        return true;
    }

    private static void collectLayoutLeafKeys(MainShellTabLayoutNode n, Set<String> out) {
        if (n.isTab()) {
            out.add(n.id());
            return;
        }
        for (MainShellTabLayoutNode c : n.children()) {
            collectLayoutLeafKeys(c, out);
        }
    }

    private Tab materializeLayoutNode(MainShellTabLayoutNode n) {
        if (n.isTab()) {
            MainShellTabId id = MainShellTabId.fromKey(n.id());
            Tab t = id != null ? mainShellTabFor(id) : null;
            if (t != null) {
                applyShellTabColor(t, n.colorHex());
            }
            return t;
        }
        if (n.isGroup()) {
            Tab groupTab = new Tab(n.title().isBlank() ? "グループ" : n.title());
            TabPane inner = new TabPane();
            inner.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);
            inner.setTabDragPolicy(TabPane.TabDragPolicy.REORDER);
            inner.getStyleClass().add("pm-main-shell-inner-tab-pane");
            for (MainShellTabLayoutNode c : n.children()) {
                Tab ct = materializeLayoutNode(c);
                if (ct != null) {
                    inner.getTabs().add(ct);
                }
            }
            groupTab.setContent(inner);
            applyShellTabColor(groupTab, n.colorHex());
            wiredInnerMainShellTabPanes.add(inner);
            return groupTab;
        }
        return null;
    }

    /** 「タブ整理」タブから呼ばれ、既定のフラット構成に戻す（作業タブを1段に並べ替え）。 */
    void restoreDefaultFlatMainShellTabLayout() {
        if (tabPane == null || mainShellTabOrganizer == null) {
            return;
        }
        suppressEnvSessionPersistence.set(true);
        suppressMainShellTabChromeRefresh.set(true);
        try {
            wiredInnerMainShellTabPanes.clear();
            tabPane.getTabs().clear();
            for (String key : MainShellTabLayoutDefaults.completeFlatTabKeyOrder()) {
                MainShellTabId id = MainShellTabId.fromKey(key);
                Tab t = id != null ? mainShellTabFor(id) : null;
                if (t != null) {
                    applyShellTabColor(t, "");
                    tabPane.getTabs().add(t);
                }
            }
            tabPane.getTabs().add(mainShellTabOrganizer);
            tabPane.setTabDragPolicy(TabPane.TabDragPolicy.REORDER);
        } finally {
            suppressMainShellTabChromeRefresh.set(false);
            suppressEnvSessionPersistence.set(false);
        }
        refreshMainShellTabDisplayedTitles();
        lastEffectiveShellLeaf =
                resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
    }

    /**
     * ツリー編集結果を適用し、成功時のみセッション保存まで行う。
     *
     * @return メインタブの組み替えに成功したとき {@code true}
     */
    boolean applyMainShellTabLayoutFromOrganizer(List<MainShellTabLayoutNode> layout) {
        if (!rebuildMainShellTabsFromLayout(layout)) {
            return false;
        }
        DesktopSessionStateStore.save(collectDesktopSession());
        return true;
    }

    /** 現在のメインシェル構成をツリー編集用にエクスポート。 */
    List<MainShellTabLayoutNode> snapshotMainShellTabLayoutNodes() {
        return snapshotMainShellTabLayout();
    }

    /** {@link MainShellTabLayoutDefaults#completeFlatTabKeyOrder()} と同順の {@link MainShellTabId}。 */
    List<MainShellTabId> defaultMainShellTabIds() {
        List<MainShellTabId> out = new ArrayList<>();
        for (String k : MainShellTabLayoutDefaults.completeFlatTabKeyOrder()) {
            MainShellTabId id = MainShellTabId.fromKey(k);
            if (id != null) {
                out.add(id);
            }
        }
        return List.copyOf(out);
    }

    /** タブ整理オーガナイザ用の既定グループ構成（メインシェルが未構築のときのツリー表示）。 */
    List<MainShellTabLayoutNode> defaultMainShellTabLayoutGrouped() {
        return MainShellTabLayoutDefaults.groupedLayout();
    }

    private static HashSet<String> requiredShellTabKeys() {
        HashSet<String> r = new HashSet<>();
        for (MainShellTabId id : MainShellTabId.values()) {
            if (id != MainShellTabId.TAB_ORGANIZER) {
                r.add(id.key());
            }
        }
        return r;
    }

    /**
     * 従来の {@code mainShellTabOrder}（リーフキー列）からフラットな {@link MainShellTabLayoutNode} 列を組み立てる。
     * 欠落キーは {@link MainShellTabLayoutDefaults#DEFAULT_FLAT_TAB_KEY_ORDER} の順で末尾に足す。
     */
    private static List<MainShellTabLayoutNode> flatMainShellTabLayoutFromOrderKeys(List<String> orderKeys) {
        if (orderKeys == null || orderKeys.isEmpty()) {
            return List.of();
        }
        LinkedHashSet<String> keys = new LinkedHashSet<>();
        for (String key : orderKeys) {
            if (key == null || key.isBlank()) {
                continue;
            }
            MainShellTabId id = MainShellTabId.fromKey(key.trim());
            if (id != null && id != MainShellTabId.TAB_ORGANIZER) {
                keys.add(id.key());
            }
        }
        for (String key : MainShellTabLayoutDefaults.DEFAULT_FLAT_TAB_KEY_ORDER) {
            if (requiredShellTabKeys().contains(key)) {
                keys.add(key);
            }
        }
        List<MainShellTabLayoutNode> out = new ArrayList<>();
        for (String key : keys) {
            out.add(MainShellTabLayoutNode.tabNode(key, ""));
        }
        return List.copyOf(out);
    }

    /**
     * セッション由来やユーザー編集のレイアウトを、未知 ID の除去・欠落タブの末尾追記・重複時のフォールバックを行う。
     */
    private List<MainShellTabLayoutNode> prepareMainShellLayoutForRebuild(
            List<MainShellTabLayoutNode> raw) {
        if (raw == null || raw.isEmpty()) {
            return mergeMissingMainShellTabLeaves(MainShellTabLayoutDefaults.groupedLayout());
        }
        List<MainShellTabLayoutNode> sanitized = sanitizeMainShellTabLayoutNodes(raw);
        List<String> leaves = new ArrayList<>();
        for (MainShellTabLayoutNode n : sanitized) {
            collectLayoutLeafKeysToList(n, leaves);
        }
        Set<String> req = requiredShellTabKeys();
        Set<String> uniq = new HashSet<>(leaves);
        if (uniq.size() != leaves.size()) {
            return mergeMissingMainShellTabLeaves(MainShellTabLayoutDefaults.groupedLayout());
        }
        for (String leaf : uniq) {
            if (!req.contains(leaf)) {
                return mergeMissingMainShellTabLeaves(MainShellTabLayoutDefaults.groupedLayout());
            }
        }
        if (uniq.equals(req)) {
            return sanitized;
        }
        return mergeMissingMainShellTabLeaves(sanitized);
    }

    private static MainShellTabLayoutNode sanitizeLayoutNode(MainShellTabLayoutNode n) {
        if (n == null) {
            return null;
        }
        if (n.isTab()) {
            MainShellTabId id = MainShellTabId.fromKey(n.id());
            if (id == null || id == MainShellTabId.TAB_ORGANIZER) {
                return null;
            }
            return MainShellTabLayoutNode.tabNode(id.key(), n.colorHex());
        }
        List<MainShellTabLayoutNode> ch = new ArrayList<>();
        for (MainShellTabLayoutNode c : n.children()) {
            MainShellTabLayoutNode s = sanitizeLayoutNode(c);
            if (s != null) {
                ch.add(s);
            }
        }
        if (ch.isEmpty()) {
            return null;
        }
        String title = n.title().isBlank() ? "グループ" : n.title();
        return MainShellTabLayoutNode.groupNode(title, n.colorHex(), ch);
    }

    private static List<MainShellTabLayoutNode> sanitizeMainShellTabLayoutNodes(
            List<MainShellTabLayoutNode> top) {
        List<MainShellTabLayoutNode> out = new ArrayList<>();
        for (MainShellTabLayoutNode n : top) {
            MainShellTabLayoutNode s = sanitizeLayoutNode(n);
            if (s != null) {
                out.add(s);
            }
        }
        return out;
    }

    private static void collectLayoutLeafKeysToList(MainShellTabLayoutNode n, List<String> out) {
        if (n.isTab()) {
            out.add(n.id());
            return;
        }
        for (MainShellTabLayoutNode c : n.children()) {
            collectLayoutLeafKeysToList(c, out);
        }
    }

    private static List<MainShellTabLayoutNode> mergeMissingMainShellTabLeaves(
            List<MainShellTabLayoutNode> top) {
        Set<String> required = requiredShellTabKeys();
        Set<String> found = new HashSet<>();
        for (MainShellTabLayoutNode n : top) {
            collectLayoutLeafKeys(n, found);
        }
        LinkedHashSet<String> missing = new LinkedHashSet<>(required);
        missing.removeAll(found);
        if (missing.isEmpty()) {
            return List.copyOf(top);
        }
        List<MainShellTabLayoutNode> out = new ArrayList<>(top);
        for (String key : MainShellTabLayoutDefaults.DEFAULT_FLAT_TAB_KEY_ORDER) {
            if (missing.remove(key)) {
                out.add(MainShellTabLayoutNode.tabNode(key, ""));
            }
        }
        for (String key : missing) {
            out.add(MainShellTabLayoutNode.tabNode(key, ""));
        }
        return out;
    }

    String mainShellTabTitle(MainShellTabId id) {
        if (id == null) {
            return "";
        }
        String a = mainShellTabTitleAliases.get(id.key());
        if (a != null && !a.isBlank()) {
            return a;
        }
        String baseline = mainShellTabBaselineTitles.get(id);
        if (baseline != null && !baseline.isBlank()) {
            return baseline;
        }
        return id.name();
    }

    /** FXML 既定の見出し（エイリアス未設定時の説明・プレースホルダ用）。 */
    String mainShellTabBaselineTitle(MainShellTabId id) {
        if (id == null) {
            return "";
        }
        String baseline = mainShellTabBaselineTitles.get(id);
        return baseline != null && !baseline.isBlank() ? baseline : id.name();
    }

    /** セッションに保存されているエイリアス（未設定は空文字）。 */
    String mainShellTabTitleAliasStored(MainShellTabId id) {
        if (id == null) {
            return "";
        }
        return mainShellTabTitleAliases.getOrDefault(id.key(), "");
    }

    /**
     * メインタブ見出しの表示エイリアスを設定する。空ならエイリアスを解除し既定見出しに戻す。
     * 内部 ID（{@link MainShellTabId#key()}）やレイアウト JSON は変更しない。
     */
    void setMainShellTabDisplayAlias(MainShellTabId id, String alias) {
        if (id == null || id == MainShellTabId.TAB_ORGANIZER) {
            return;
        }
        String k = id.key();
        if (alias == null || alias.isBlank()) {
            mainShellTabTitleAliases.remove(k);
        } else {
            mainShellTabTitleAliases.put(k, alias.strip());
        }
        Tab tab = mainShellTabFor(id);
        if (tab != null) {
            tab.setText(mainShellTabTitle(id));
        }
    }

    /** セッション保存用スナップショット（同一プロセス内の子コントローラから）。 */
    DesktopSessionState collectDesktopSessionSnapshot() {
        return collectDesktopSession();
    }

    private static boolean omitEnvRowKey(String name) {
        String k = name != null ? name.trim() : "";
        return Stage2PythonChildEnv.LEGACY_WORKBOOK_KEYS_STRIPPED_FOR_PYTHON_CHILD.contains(k)
                || DROPPED_ENV_TAB_ROW_KEYS.contains(k);
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
        ensureBootstrapDefaultValuesVisible(collectUiEnv());
        ensureUiRefOptionalDisplayDefaultsVisible(collectUiEnv());
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
        ensureUiRefOptionalDisplayDefaultsVisible(collectUiEnv());
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
        ensureUiRefOptionalDisplayDefaultsVisible(collectUiEnv());
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
                    if (mainRunTabController != null) {
                        mainRunTabController.refreshOpenWorkbookHintLabels();
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
        Optional<FactorySite> siteOpt = promptFactorySiteChoiceForEnvDefaults();
        FactorySite site = siteOpt.orElse(FactorySite.KONAN);
        if (siteOpt.isEmpty()) {
            appendLog("[env] 工場既定の選択をキャンセルしたため湖南工場の既定を適用します。");
        }
        applyEnvRowsFullBundledResetAndPersist(true, site);
    }

    /**
     * 湖南／国分工場の環境タブ既定を選択する共通 {@link javafx.scene.control.ChoiceDialog}。
     *
     * @return OK 時は選択した工場。キャンセル時は empty。
     */
    private Optional<FactorySite> promptFactorySiteChoice(String title, String contentText) {
        if (primaryStage == null) {
            return Optional.of(FactorySite.KONAN);
        }
        FactorySite pref = GlobalInitSettingTarget.load();
        ChoiceDialog<FactorySite> d = new ChoiceDialog<>(pref, List.of(FactorySite.values()));
        d.initOwner(primaryStage);
        applyAlertStylesheetsFromOwner(d);
        d.setTitle(title);
        d.setHeaderText(null);
        d.setContentText(contentText);
        d.setSelectedItem(pref);
        return d.showAndWait();
    }

    /**
     * 環境変数初期化（バンドル既定へ戻す）直前: 湖南／国分の工場既定を選ばせる。
     *
     * @return OK 時は選択した工場。キャンセル時は empty（呼び出し側で湖南とみなす）。
     */
    private Optional<FactorySite> promptFactorySiteChoiceForEnvDefaults() {
        return promptFactorySiteChoice(
                "環境変数を初期値に戻す",
                "ネットワークの計画／実績フォルダ・自動バージョンアップ用 ZIP・マスタの既定を、利用する工場に合わせて選んでください。\n"
                        + "（キャンセルした場合は湖南工場の既定を適用します。）");
    }

    /**
     * 環境タブをバンドル既定で再構築し永続化する（確認ダイアログなし）。初回起動マーカー・工場出荷 UI リセット・ポータル
     * アップグレード直後などから利用。
     *
     * @param persistSession false のとき {@code session-state.json} には書かない（工場出荷 UI リセットの途中で利用）。
     * @param factorySite テンプレ再構築後に適用する工場別ネットワーク／マスタ既定（湖南＝従来のコード既定）
     */
    private void applyEnvRowsFullBundledResetAndPersist(boolean persistSession, FactorySite factorySite) {
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
                    .getScriptDirField()
                    .setText(
                            firstNonBlank(
                                    ui.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                    AppPaths.resolvePythonScriptDir(ui).toString()));
        } finally {
            envResetInProgress.set(false);
            suppressEnvSessionPersistence.set(false);
        }
        // テンプレ再構築だけでは ui_ref 空行等で欠ける場合があるため、工場別の共有 UNC 等を確実に入れる
        applyFactorySitePortableAndNetworkDefaults(factorySite);
        ensureBootstrapDefaultValuesVisible(collectUiEnv());
        ensureUiRefOptionalDisplayDefaultsVisible(collectUiEnv());
        applyRepoFolderPathNormalization();
        if (persistSession) {
            DesktopSessionStateStore.save(collectDesktopSession());
        }
        mainRunTabController.refreshOpenWorkbookHintLabels();
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
     * 起動スプラッシュ（APPLICATION_MODAL・常に前面）を閉じたあとにポータブル自動バージョンアップを走らせる。
     * スプラッシュ表示中に {@link Alert#showAndWait()} すると確認ダイアログが背面に隠れて見えないことがある。
     */
    void schedulePortableBundleSelfUpdateAfterSplash() {
        Platform.runLater(this::maybePortableBundleSelfUpdate);
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
        if (STAGE2.equals(script) && dispatchInteractiveTabController != null) {
            Runnable clearDispatch =
                    () -> dispatchInteractiveTabController.resetTableDisplayForStage2Run();
            if (Platform.isFxApplicationThread()) {
                clearDispatch.run();
            } else {
                Platform.runLater(clearDispatch);
            }
        }
        try {
            Map<String, String> uiRun = collectUiEnv();
            if (STAGE1.equals(script)) {
                uiRun.put(
                        AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH,
                        mainRunTabController.snapshotStage2SkipInProgressDispatch() ? "1" : "0");
            }
            if (STAGE2.equals(script)) {
                uiRun.put(
                        AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL,
                        mainRunTabController.snapshotStage2WriteExcel() ? "1" : "0");
                uiRun.put(
                        AppPaths.KEY_PM_AI_STAGE2_SKIP_TODAY_DISPATCH,
                        planInputTabController.snapshotStage2SkipTodayDispatch() ? "1" : "0");
                uiRun.put(
                        AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH,
                        mainRunTabController.snapshotStage2SkipInProgressDispatch() ? "1" : "0");
                String resultFont = mainRunTabController.snapshotStage2ResultBookFont();
                if (resultFont != null && !resultFont.isBlank()) {
                    uiRun.put(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT, resultFont.trim());
                } else {
                    uiRun.remove(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT);
                }
            }
            String wb = effectiveTaskInputWorkbookPath();
            appendLog("--- start: " + script + " ---");
            if (STAGE1.equals(script)) {
                if (mainRunTabController.snapshotStage1ClearCacheAndRun()) {
                    appendLog("[stage1] キャッシュをクリアして実行します。");
                    try {
                        Stage1AiCacheClearer.ClearResult cacheClear =
                                Stage1AiCacheClearer.archiveAndClearBeforeStage1Run(
                                        uiRun, "段階1実行前");
                        for (String line : cacheClear.detailLines()) {
                            appendLog(line);
                        }
                        if (cacheClear.anyFailed()) {
                            appendLog("[stage1] キャッシュの一部を削除できませんでした。");
                        } else {
                            appendLog("[stage1] キャッシュをクリアしました。");
                        }
                        if (workspaceCacheHistoryTabController != null) {
                            workspaceCacheHistoryTabController.refreshListQuietly();
                        }
                    } catch (IOException archiveEx) {
                        appendLog(
                                "[stage1] キャッシュ退避に失敗したためクリアを中止しました: "
                                        + archiveEx.getMessage());
                        runLock.set(false);
                        activeRunStageScript = null;
                        activeStageChildProcess.set(null);
                        mainRunTabController.getStatusLabel().setText("キャッシュ退避失敗");
                        applyRunTabGating();
                        return;
                    }
                } else {
                    List<String> cacheKinds = Stage1AiCacheClearer.existingCacheKindLabelsJa(uiRun);
                    if (!cacheKinds.isEmpty()) {
                        appendLog(
                                "[stage1] キャッシュを使用します（"
                                        + String.join("・", cacheKinds)
                                        + "）。");
                    }
                }
            }
            if (STAGE1.equals(script) || STAGE2.equals(script)) {
                refreshNetworkSourceDirListingSkipsBeforeStageRun(uiRun);
            }
            Map<String, String> childEnv = childEnvForPython(uiRun);
            if (lastNetworkSourceResolution != null) {
                for (String line : lastNetworkSourceResolution.logLines()) {
                    appendLog(line);
                }
            }
            if (STAGE1.equals(script)) {
                NetworkSourceDirResolver.Result res = lastNetworkSourceResolution;
                boolean networkFromCache =
                        res != null && (res.taskInputFromCache() || res.actualDetailFromCache());
                if (networkFromCache) {
                    appendLog("[stage1] キャッシュを使用します（加工計画DATA／実績明細のネットワーク代替）。");
                }
                mainRunTabController.setStage1NetworkCacheBadge(
                        networkFromCache,
                        uiBadgeDesignTabController != null
                                ? uiBadgeDesignTabController.snapshotStage1NetworkCacheBadgeStyle()
                                : PersonBadgeStyle.networkSourceCacheBadgeDefault(),
                        uiBadgeDesignTabController != null
                                ? uiBadgeDesignTabController.snapshotStage1NetworkCacheBadgeLabel()
                                : "キャッシュ");
            }
            Path py = resolveStagePythonExecutablePath(uiRun);
            Path dir =
                    Path.of(
                            firstNonBlank(
                                    uiRun.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                    mainRunTabController.getScriptDirField().getText().trim()));
            appendStageChildResolvedEnvForRun(script, childEnv);
            RunRequest req = new RunRequest(py, dir, script, wb, childEnv);
            mainRunTabController.getStatusLabel().setText("実行中…");

            ArrayDeque<String> recentChildLines = new ArrayDeque<>(STAGE_CHILD_LOG_TAIL_MAX + 4);

            PythonProcessRunner.runAsync(
                            req,
                            line -> {
                                synchronized (recentChildLines) {
                                    while (recentChildLines.size() >= STAGE_CHILD_LOG_TAIL_MAX) {
                                        recentChildLines.removeFirst();
                                    }
                                    recentChildLines.addLast(line);
                                }
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
                                final List<String> tailSnap;
                                synchronized (recentChildLines) {
                                    tailSnap = new ArrayList<>(recentChildLines);
                                }
                                runLock.set(false);
                                activeRunStageScript = null;
                                activeStageChildProcess.set(null);
                                javafx.application.Platform.runLater(
                                        () -> completeStageRunOnFx(script, code, err, tailSnap));
                            });
        } catch (Throwable t) {
            runLock.set(false);
            activeRunStageScript = null;
            activeStageChildProcess.set(null);
            appendLog("[error] runStage: " + t.getMessage());
            boolean stage2 = STAGE2.equals(script);
            boolean stage1 = STAGE1.equals(script);
            Platform.runLater(
                    () -> {
                        applyRunTabGating();
                        if (stage2 && dispatchInteractiveTabController != null) {
                            dispatchInteractiveTabController.reloadTableFromDiskAfterExternalUpdate();
                        }
                        if (stage1) {
                            mainRunTabController.resetStage1ClearCacheAndRunCheckbox();
                        }
                        if (stage1 || stage2) {
                            selectMainShellTab(MainShellTabId.RUN);
                            showStageFailureDialog(script, null, t, List.of());
                        }
                    });
        }
    }

    private void completeStageRunOnFx(String script, Integer code, Throwable err, List<String> tailSnap) {
        applyRunTabGating();
        if (err != null) {
            mainRunTabController
                    .getStatusLabel()
                    .setText(
                            "failed: "
                                    + (err.getMessage() != null ? err.getMessage() : err.toString()));
            appendLog(
                    "[end] exceptional exit: "
                            + (err.getMessage() != null ? err.getMessage() : err.toString()));
            if (STAGE2.equals(script) && dispatchInteractiveTabController != null) {
                dispatchInteractiveTabController.reloadTableFromDiskAfterExternalUpdate();
            }
        } else {
            int c = code != null ? code : -1;
            mainRunTabController.getStatusLabel().setText(exitCodeLegend(c));
            appendLog("[end] exitCode=" + c + " " + exitHint(c));
            if (STAGE1.equals(script) && c == 0) {
                applyStage1ExcludeRulesJsonToEnvTab();
                try {
                    CodeDispatchLookupTablesMerge.MergeSummary ms =
                            CodeDispatchLookupTablesMerge.mergeAfterStage1(collectUiEnv());
                    if (ms.totalAdded() > 0) {
                        appendLog("[stage1] 材料・製品種類情報(code/) 自動追記: " + ms.summaryJa());
                    }
                } catch (Exception ex) {
                    appendLog("[stage1] 材料・製品種類情報(code/) 自動追記失敗: " + ex.getMessage());
                }
                if (codeDispatchLookupTablesTabController != null) {
                    Platform.runLater(() -> codeDispatchLookupTablesTabController.reloadAllFromDisk());
                }
                if (reloadAfterStage1Preview != null) {
                    reloadAfterStage1Preview.run();
                }
                if (reloadAfterStage1PlanInput != null) {
                    reloadAfterStage1PlanInput.run();
                }
                invalidateDeliveryCalendarAfterPipelineRun();
                refreshEquipmentGanttGraphicAfterPipelineRun();
                MacroCompleteChime.playIfAvailable(collectUiEnv());
                selectMainShellTab(MainShellTabId.PLAN_INPUT);
                showStageCompletionDialog("段階1 完了", "段階1 の処理が正常終了しました。");
            }
            if (STAGE2.equals(script)) {
                if (c == 0) {
                    refreshStage2OutputArtifacts();
                    exportSummaryAiDispatchWorkbookAfterPipelineStageRun(true);
                    refreshEquipmentGanttGraphicAfterPipelineRun();
                    Runnable afterDispatchReload =
                            () -> {
                                MacroCompleteChime.playIfAvailable(collectUiEnv());
                                selectMainShellTab(MainShellTabId.DISPATCH_INTERACTIVE);
                                showStageCompletionDialog(
                                        "段階2 完了", "段階2 の処理が正常終了しました。");
                            };
                    if (dispatchInteractiveTabController != null) {
                        dispatchInteractiveTabController.reloadTableFromDiskAfterStage2Success(
                                afterDispatchReload);
                    } else {
                        afterDispatchReload.run();
                    }
                } else if (dispatchInteractiveTabController != null) {
                    dispatchInteractiveTabController.reloadTableFromDiskAfterExternalUpdate();
                }
            }
        }
        boolean stage12 = STAGE1.equals(script) || STAGE2.equals(script);
        boolean failed = err != null || (code != null && code.intValue() != 0);
        if (stage12 && failed) {
            appendLog("[ui] 段階処理が異常終了しました。エラーダイアログを表示します。");
            selectMainShellTab(MainShellTabId.RUN);
            showStageFailureDialog(script, err != null ? null : code, err, tailSnap);
        }
        if (STAGE1.equals(script)) {
            mainRunTabController.resetStage1ClearCacheAndRunCheckbox();
        }
    }

    /**
     * 段階1/2 実行中の Python 子プロセスを終了する（ツールバー・実行・ログの「中断」）。
     */
    void cancelActiveStageRun() {
        boolean didSomething = false;
        Process child = activeStageChildProcess.get();
        if (child != null && child.isAlive()) {
            appendLog("[interrupt] 段階1/2 の子プロセスを終了します…");
            try {
                child.destroyForcibly();
            } catch (Exception ex) {
                appendLog("[interrupt] 子プロセス終了に失敗: " + ex.getMessage());
            }
            didSomething = true;
        }
        if (!didSomething) {
            appendLog("[interrupt] 終了対象の子プロセスがありません。");
        }
    }

    @FXML
    private void onCancelStageRunAction() {
        cancelActiveStageRun();
    }

    /** ツールバー「配台の使い方」… リポジトリ直下の Word 手順書を既定アプリで開く。 */
    @FXML
    private void onOpenDispatchUsageGuideDocxAction() {
        Path p = AppPaths.resolveDispatchUsageGuideDocx(collectUiEnv());
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[dispatch-usage-docx] file not found: "
                            + p
                            + " (expected "
                            + AppPaths.DISPATCH_USAGE_GUIDE_DOCX
                            + " under "
                            + AppPaths.KEY_PM_AI_REPO_ROOT
                            + ")");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[dispatch-usage-docx] opened: " + p);
        } catch (IOException e) {
            appendLog("[dispatch-usage-docx] open failed: " + e.getMessage());
        }
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
        if (planInputTabController != null) {
            planInputTabController.setStageRunProgressVisible(stage1Running, stage2Running);
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
                shellStageProgressLabel.setText(stage1Running ? "段階1 実行中…" : "段階2 実行中…");
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

    /**
     * 段階1／段階2の子プロセスに渡す直前に、入力解決に効く環境変数をログへ列挙する（ネットワーク解決ログの直後）。
     */
    private void appendStageChildResolvedEnvForRun(String script, Map<String, String> childEnv) {
        List<String> keys =
                STAGE1.equals(script)
                        ? STAGE1_CHILD_INPUT_ENV_KEYS
                        : (STAGE2.equals(script) ? STAGE2_CHILD_INPUT_ENV_KEYS : List.of());
        if (keys.isEmpty()) {
            return;
        }
        String ja = STAGE1.equals(script) ? "段階1" : "段階2";
        appendLog("--- " + ja + " 子プロセス入力（環境変数キー → 渡す値）---");
        for (String k : keys) {
            String v = childEnv != null ? childEnv.get(k) : null;
            if (v == null || v.isBlank()) {
                appendLog("[" + ja + "-input] " + k + " = （未設定または空）");
            } else {
                appendLog("[" + ja + "-input] " + k + " = " + v);
            }
        }
        if (STAGE1.equals(script)) {
            appendLog(
                    "[段階1-input] 加工計画DATAの実ファイルは "
                            + AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH
                            + "（未設定時は "
                            + AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR
                            + " 直下の最新表から解決）。実績明細は "
                            + AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK
                            + " または "
                            + AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR
                            + "。");
        } else if (STAGE2.equals(script)) {
            appendLog(
                    "[段階2-input] 配台計画の入力は "
                            + AppPaths.KEY_PM_AI_PLAN_INPUT_PATH
                            + " とシート名 "
                            + PlanInputTabController.ENV_TASK_PLAN_SHEET
                            + "。マスタは "
                            + AppPaths.KEY_PM_AI_MASTER_WORKBOOK
                            + " / "
                            + AppPaths.KEY_MASTER_WORKBOOK_FILE
                            + "。");
        }
    }

    private static String exitHintJa(int code) {
        return switch (code) {
            case 0 -> "正常終了しました。";
            case 1 -> "一般エラーです（データや設定の不整合など）。";
            case 2 -> "致命的エラー、またはマスタ・入力ファイルの欠如などです。";
            case 3 -> "計画データの検証エラーです（必須列の不足など）。";
            case 9 -> "ユーザーによる中断です。";
            default -> "終了コード " + code + " です。";
        };
    }

    /**
     * 段階1／段階2が異常終了したときにエラーダイアログを出す。{@code tailLines} は子の標準出力に付いた行（先頭に {@code
     * [child] } を含む）の末尾スナップショット。
     */
    private void showStageFailureDialog(
            String script, Integer code, Throwable err, List<String> tailLines) {
        String stageJa = STAGE1.equals(script) ? "段階1" : "段階2";
        Alert alert = new Alert(AlertType.ERROR);
        alert.initOwner(primaryStage);
        applyAlertStylesheetsFromOwner(alert);
        alert.setTitle(stageJa + " 失敗");
        alert.setHeaderText(null);
        StringBuilder body = new StringBuilder();
        if (err != null) {
            body.append("子プロセスの起動または実行中に例外が発生しました。\n");
            body.append(err.getMessage() != null ? err.getMessage() : err.toString());
        } else {
            int c = code != null ? code : -1;
            body.append(exitCodeLegend(c)).append("\n");
            body.append(exitHintJa(c));
        }
        body.append("\n\n詳細は「実行・ログ」タブのログを確認してください。");
        if (tailLines != null && !tailLines.isEmpty()) {
            body.append("\n\n【直近の子プロセス出力】\n");
            int start = Math.max(0, tailLines.size() - 14);
            for (int i = start; i < tailLines.size(); i++) {
                String ln = tailLines.get(i);
                if (ln.length() > 220) {
                    ln = ln.substring(0, 217) + "...";
                }
                body.append(ln).append('\n');
            }
        }
        alert.setContentText(body.toString());
        alert.showAndWait();
    }

    /** メインウィンドウと同じテーマ CSS をダイアログに載せる（Alert / ChoiceDialog は別 Scene のため未設定だと配色がずれる） */
    private void applyAlertStylesheetsFromOwner(Dialog<?> dialog) {
        if (primaryStage == null || dialog == null) {
            return;
        }
        Scene ownerScene = primaryStage.getScene();
        if (ownerScene == null) {
            return;
        }
        var paneSheets = dialog.getDialogPane().getStylesheets();
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
     * Stage 1/2 child processes do not receive legacy {@link
     * jp.co.pm.ai.desktop.bridge.Stage2PythonChildEnv#LEGACY_WORKBOOK_KEYS_STRIPPED_FOR_PYTHON_CHILD}; use
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
        Path py = resolveStagePythonExecutablePath(uiRun);
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
        Path py = resolveStagePythonExecutablePath(uiRun);
        Path dir =
                Path.of(
                        firstNonBlank(
                                uiRun.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                mainRunTabController.getScriptDirField().getText().trim()));
        String wb = effectiveTaskInputWorkbookPath();
        return new RunRequest(py, dir, "pm_ai_actuals_status.py", wb, childEnvForPython(uiRun));
    }

    /** pm_ai_delivery_calendar_view.py: same env merge as stage1/2 / actuals status. */
    RunRequest buildDeliveryCalendarRequest() {
        Map<String, String> uiRun = collectUiEnv();
        Path py = resolveStagePythonExecutablePath(uiRun);
        Path dir =
                Path.of(
                        firstNonBlank(
                                uiRun.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR),
                                mainRunTabController.getScriptDirField().getText().trim()));
        String wb = effectiveTaskInputWorkbookPath();
        return new RunRequest(py, dir, "pm_ai_delivery_calendar_view.py", wb, childEnvForPython(uiRun));
    }

    /**
     * Env tab keys passed to Python; strips legacy workbook keys（{@link
     * jp.co.pm.ai.desktop.bridge.Stage2PythonChildEnv#LEGACY_WORKBOOK_KEYS_STRIPPED_FOR_PYTHON_CHILD}）。
     * If {@code PM_AI_PLAN_INPUT_PATH} / {@code TASK_PLAN_SHEET} are unset in the env tab, values from
     * the 配台計画_タスク入力 tab are applied so that stage-2 uses the
     * file the user is editing there.
     */
    private Map<String, String> childEnvForPython(Map<String, String> ui) {
        Map<String, String> m = new HashMap<>(ui);
        Stage2PythonChildEnv.stripLegacyWorkbookKeys(m);
        Stage2PythonChildEnv.ensureSkipWorkbookEnvSheetDefault(m);
        overlayPlanInputTabPathsIfEnvBlank(m);
        lastNetworkSourceResolution =
                Stage2PythonChildEnv.applyNetworkSourceAndChildPause(
                        m,
                        startupSkipTaskInputSourceDirListing,
                        startupSkipActualDetailSourceDirListing);
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
     * 段階1／段階2の実行直前に、ネットワークソースフォルダの一覧可否を再評価する。起動時は未到達だったが実行時に回復していれば
     * {@link #startupSkipTaskInputSourceDirListing} / {@link #startupSkipActualDetailSourceDirListing} を下げ、
     * {@link NetworkSourceDirResolver#resolve(Map, boolean, boolean)} がネットワーク側の最新ファイル検出を再度行う。
     */
    private void refreshNetworkSourceDirListingSkipsBeforeStageRun(Map<String, String> ui) {
        boolean wasTaskSkip = startupSkipTaskInputSourceDirListing;
        boolean wasActSkip = startupSkipActualDetailSourceDirListing;
        boolean taskReach = NetworkSourceDirResolver.isTaskInputSourceDirReachable(ui);
        boolean actReach = NetworkSourceDirResolver.isActualDetailSourceDirReachable(ui);
        startupSkipTaskInputSourceDirListing = !taskReach;
        startupSkipActualDetailSourceDirListing = !actReach;
        if (wasTaskSkip && taskReach) {
            appendLog(
                    "[network-source] PM_AI_TASK_INPUT_SOURCE_DIR が再び一覧可能になりました。ネットワークから最新を検出します: "
                            + AppPaths.resolveTaskInputSourceDir(ui));
        }
        if (wasActSkip && actReach) {
            appendLog(
                    "[network-source] PM_AI_ACTUAL_DETAIL_SOURCE_DIR が再び一覧可能になりました。ネットワークから最新を検出します: "
                            + AppPaths.resolveActualDetailSourceDir(ui));
        }
    }

    /**
     * フォルダ系 {@code PM_AI_*} のうち、リポジトリ基準へ補正できるものを更新する（{@link AppPaths#normalizedFolderEnvOverrides(Map)}）。
     *
     * <p>{@code PM_AI_TASK_INPUT_SOURCE_DIR} / {@code PM_AI_ACTUAL_DETAIL_SOURCE_DIR} は {@link AppPaths#normalizedFolderEnvOverrides(Map)}
     * の対象外のためここでは変更しない（バージョンアップ完了時の {@link #applyFactorySitePortableAndNetworkDefaults(FactorySite)} とフォルダ選択のみで更新）。
     */
    private void applyRepoFolderPathNormalization() {
        if (envRows == null) {
            return;
        }
        suppressEnvSessionPersistence.set(true);
        try {
            Map<String, String> ui = collectUiEnv();
            Map<String, String> overrides = AppPaths.normalizedFolderEnvOverrides(ui);
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

    /**
     * {@code pm-ai-data/code/python/task_extract_stage1.py} がある同梱レイアウトか。
     *
     * @see #applyBundledPortableDefaultsIfPresent()
     */
    private boolean bundledPortableStage1MarkerPresent() {
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        Path marker =
                cwd.resolve("pm-ai-data")
                        .resolve("code")
                        .resolve("python")
                        .resolve("task_extract_stage1.py");
        return Files.isRegularFile(marker);
    }

    /**
     * 工場別のネットワークソース・バージョンアップ正本 ZIP・マスタ basename・サマリ用ブック絶対パスを環境タブへ書き込む（UNC は {@link Path} 経由にしない）。
     *
     * <p>環境タブでこれらをコードから書き換えるのは、ポータル自動バージョンアップ完了時・
     * {@link #applyEnvRowsFullBundledResetAndPersist(boolean, FactorySite)}（環境変数を初期化）とする。
     *
     * @param site 選択された工場（湖南＝従来既定）
     */
    /**
     * サマリ用ブックの出力先を環境タブ {@link AppPaths#KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK} に反映する。
     * 既に非空のときも、出力した絶対パスへ揃える（固定ファイル名の上書き運用）。
     */
    void ensureSummaryAiDispatchWorkbookEnvPath(Path absoluteOutputPath) {
        if (envRows == null || absoluteOutputPath == null) {
            return;
        }
        String target = absoluteOutputPath.toAbsolutePath().normalize().toString();
        suppressEnvSessionPersistence.set(true);
        try {
            for (EnvVarRow row : envRows) {
                String name = row.getName() != null ? row.getName().trim() : "";
                if (AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK.equals(name)) {
                    row.setValue(target);
                    break;
                }
            }
        } finally {
            suppressEnvSessionPersistence.set(false);
        }
        scheduleDesktopSessionSave();
        if (mainRunTabController != null) {
            mainRunTabController.refreshOpenWorkbookHintLabels();
            mainRunTabController.refreshFactorySiteLogo();
        }
    }

    private void applyFactorySitePortableAndNetworkDefaults(FactorySite site) {
        if (envRows == null || site == null) {
            return;
        }
        String task = site.taskInputSourceDir();
        String actual = site.actualDetailSourceDir();
        String portable = site.portableBundleSourceDir();
        String masterBasename = site.masterWorkbookFileBasename();
        String pmAiMaster = site.pmAiMasterWorkbookEnvValue(collectUiEnv());
        String pmAiSummary = site.pmAiSummaryAiDispatchWorkbookEnvValue(collectUiEnv());
        for (EnvVarRow r : envRows) {
            String name = r.getName() != null ? r.getName().trim() : "";
            if (AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR.equals(name)) {
                r.setValue(task);
            } else if (AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR.equals(name)) {
                r.setValue(actual);
            } else if (AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR.equals(name)) {
                r.setValue(portable);
            } else if (AppPaths.KEY_MASTER_WORKBOOK_FILE.equals(name)) {
                r.setValue(masterBasename != null ? masterBasename : "");
            } else if (AppPaths.KEY_PM_AI_MASTER_WORKBOOK.equals(name)) {
                r.setValue(pmAiMaster != null ? pmAiMaster : "");
            } else if (AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK.equals(name)) {
                r.setValue(pmAiSummary != null ? pmAiSummary : "");
            }
        }
        GlobalInitSettingTarget.save(site);
        if (globalSettingsTabController != null) {
            globalSettingsTabController.refreshInitSettingTargetComboFromStore();
        }
        if (mainRunTabController != null) {
            mainRunTabController.refreshFactorySiteLogo();
        }
    }

    /**
     * ポータブル自動バージョンアップ完了直後: 湖南／国分の環境タブ既定をユーザーに選ばせる。
     *
     * @return OK 時は選択した工場。キャンセル時は empty（呼び出し側で湖南とみなす）。
     */
    private Optional<FactorySite> promptFactorySiteAfterPortableUpgrade() {
        return promptFactorySiteChoice(
                "自動バージョンアップ",
                "バージョンアップが完了しました。\n"
                        + "ネットワークの計画／実績フォルダ・自動バージョンアップ用 ZIP・マスタファイル名の既定を、利用する工場に合わせて選んでください。\n"
                        + "（キャンセルした場合は湖南工場の既定のままです。）");
    }

    /**
     * 納期管理ビュー再読み込み中に、メインシェル最上段の「納期管理ビュー」以外のタブを無効化してグレーアウトする。
     *
     * @param greyOut {@code true} で他タブを無効化、{@code false} で通常に戻す
     */
    void setDeliveryCalendarReloadGreyOutOtherMainTabs(boolean greyOut) {
        if (tabPane == null || mainShellTabDeliveryCalendar == null) {
            return;
        }
        for (Tab t : tabPane.getTabs()) {
            if (t != mainShellTabDeliveryCalendar) {
                t.setDisable(greyOut);
            }
        }
    }

    /**
     * メインシェルのタブを ID で選択する（配台試行ウィザードなどから）。
     */
    public void selectMainShellTab(MainShellTabId id) {
        if (tabPane == null || id == null) {
            return;
        }
        selectMainShellTabRecursive(tabPane, id);
    }

    private boolean selectMainShellTabRecursive(TabPane pane, MainShellTabId id) {
        for (Tab t : pane.getTabs()) {
            if (mainShellTabId(t) == id) {
                pane.getSelectionModel().select(t);
                return true;
            }
        }
        for (Tab t : pane.getTabs()) {
            if (t.getContent() instanceof TabPane inner) {
                if (selectMainShellTabRecursive(inner, id)) {
                    pane.getSelectionModel().select(t);
                    return true;
                }
            }
        }
        return false;
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

    /** グローバル設定の工場切替などで実行・ログタブ上部ロゴを更新する。 */
    void refreshMainRunTabFactoryLogo() {
        if (mainRunTabController != null) {
            mainRunTabController.refreshFactorySiteLogo();
        }
    }

    Map<String, String> snapshotUiEnv() {
        return collectUiEnv();
    }

    /** 実行タブに表示中の段階2計画ブックパス（設備ガントの兄弟 JSON オートフィル用）。 */
    public String mainRunStage2ProductionPlanPathOrEmpty() {
        if (mainRunTabController == null) {
            return "";
        }
        String p = mainRunTabController.snapshotStage2ProductionPlanPath();
        return p != null ? p.strip() : "";
    }

    /**
     * Environment for Python child processes (same as stage1/2): env tab + plan-input tab overlays,
     * {@code PM_AI_*} inheritance rules, UTF-8 stdio.
     */
    public Map<String, String> snapshotPythonChildEnv() {
        return childEnvForPython(collectUiEnv());
    }

    /**
     * Environment for {@code dispatch_interactive_trial.py}: same stage-2 overrides as {@link #runStage}:
     * {@link AppPaths#KEY_PM_AI_STAGE2_WRITE_EXCEL} and {@link AppPaths#KEY_PM_AI_RESULT_BOOK_FONT} from the run tab,
     * {@link AppPaths#KEY_PM_AI_STAGE2_SKIP_TODAY_DISPATCH} from the plan-input tab, and
     * {@link AppPaths#KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH} from the run tab.
     */
    public Map<String, String> snapshotDispatchTrialPythonEnv() {
        Map<String, String> ui = new HashMap<>(collectUiEnv());
        ui.put(
                AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL,
                mainRunTabController.snapshotStage2WriteExcel() ? "1" : "0");
        ui.put(
                AppPaths.KEY_PM_AI_STAGE2_SKIP_TODAY_DISPATCH,
                planInputTabController.snapshotStage2SkipTodayDispatch() ? "1" : "0");
        ui.put(
                AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH,
                mainRunTabController.snapshotStage2SkipInProgressDispatch() ? "1" : "0");
        String resultFont = mainRunTabController.snapshotStage2ResultBookFont();
        if (resultFont != null && !resultFont.isBlank()) {
            ui.put(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT, resultFont.trim());
        } else {
            ui.remove(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT);
        }
        String trialSid = AgentDebugLog.resolveDispatchTrialSessionId(ui);
        ui.put("PM_AI_AGENT_DEBUG_SESSION", trialSid);
        String dbg = ui.get("PM_AI_DEBUG_LOG");
        if (dbg == null || dbg.isBlank()) {
            ui.put(
                    "PM_AI_DEBUG_LOG",
                    AgentDebugLog.resolveNdjsonPath(ui, trialSid).toAbsolutePath().toString());
        }
        return childEnvForPython(ui);
    }

    void acceptReloadAfterStage1PlanInput(Runnable r) {
        this.reloadAfterStage1PlanInput = r;
    }

    void acceptReloadAfterStage1Preview(Runnable r) {
        this.reloadAfterStage1Preview = r;
    }

    /** 手動修正タブの未保存状態をタスク入力の段階2ボタンへ反映する（起動時・bind 直後用）。 */
    void syncPlanInputStage2ButtonFromDispatchDirty() {
        boolean dirty =
                dispatchInteractiveTabController != null
                        && dispatchInteractiveTabController.isDispatchDocDirtySinceSave();
        onDispatchInteractiveTableDirtyChanged(dirty);
    }

    void onDispatchInteractiveTableDirtyChanged(boolean dispatchTableDirty) {
        if (planInputTabController != null) {
            planInputTabController.setStage2BlockedByUnsavedDispatchEdit(dispatchTableDirty);
        }
    }

    void triggerStage1() {
        runStage(STAGE1);
    }

    void triggerStage2() {
        if (dispatchInteractiveTabController != null
                && dispatchInteractiveTabController.isDispatchDocDirtySinceSave()) {
            appendLog(
                    "[stage2] 配台計画手動修正に未保存の変更があります。JSON を「保存」するか「再読み」後に実行してください。");
            return;
        }
        if (planInputTabController != null
                && planInputTabController.isPlanInputTableDirtySinceSave()) {
            appendLog(
                    "[stage2] 配台計画_タスク入力タブの表に未保存の変更があります。「保存」または「再読み」で確定してから実行してください。");
            return;
        }
        runStage(STAGE2);
    }

    /**
     * 配台試行完了後など、出力フォルダに新しい段階2成果物があれば実行・ログタブのパス（production_plan /
     * member_schedule）と関連タブの自動反映を更新する。{@link #refreshStage2OutputArtifacts} と同じ処理。
     */
    void refreshRunTabStage2ArtifactLinks() {
        refreshStage2OutputArtifacts();
    }

    /**
     * 段階1/2/3 実行後に納期管理ビューを「再読み込みボタンまで全面オーバーレイ」にする（古い表の誤閲覧防止）。
     */
    void invalidateDeliveryCalendarAfterPipelineRun() {
        if (deliveryCalendarViewTabController != null) {
            deliveryCalendarViewTabController.markStaleUntilManualReload();
        }
    }

    /** 材料・製品種類情報タブで {@code code/} ルックアップ表がディスクと同期したあと、配台計画_タスク入力のロール長ハイライトを更新する。 */
    void invalidatePlanInputRollUnitHighlightCache() {
        if (planInputTabController != null) {
            planInputTabController.invalidateRollUnitHighlightCacheAndRefresh();
        }
    }

    /**
     * 配台試行（段階3）正常終了後: 納期管理ビューは段階3前・段階3後（配台結果）のみ反映し、サマリ xlsx を更新する。
     */
    void reloadDeliveryCalendarInBackgroundAfterDispatchTrialSuccess() {
        if (deliveryCalendarViewTabController != null) {
            deliveryCalendarViewTabController.reloadInBackgroundAfterStage3DispatchTrialSuccess(
                    this::exportSummaryAiDispatchWorkbookAfterStage3DispatchReload);
        } else {
            exportSummaryAiDispatchWorkbookAfterStage3DispatchReload();
        }
    }

    /** 段階3 配台のみ反映後のサマリ xlsx 出力（メイン表スナップショットは反映済み UI を使用）。 */
    private void exportSummaryAiDispatchWorkbookAfterStage3DispatchReload() {
        Map<String, String> ui = collectUiEnv();
        try {
            PlanInputTabularIo.TabularSheet main =
                    deliveryCalendarViewTabController != null
                            ? deliveryCalendarViewTabController.snapshotMainCompareForExport()
                            : new PlanInputTabularIo.TabularSheet(List.of(), List.of());
            Path out = SummaryAiDispatchWorkbookExporter.writeFromPipelineArtifacts(ui, main);
            ensureSummaryAiDispatchWorkbookEnvPath(out);
            appendLog("[summary-ai-dispatch] 段階3後エクセル出力: " + out);
        } catch (Exception ex) {
            appendLog(
                    "[summary-ai-dispatch] 段階3後エクセル出力失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
    }

    /**
     * 段階2 正常完了後: 成果物 JSON でサマリ xlsx を更新し、納期管理ビューをフル再読込する。
     */
    private void exportSummaryAiDispatchWorkbookAfterPipelineStageRun(boolean scheduleDeliveryReload) {
        Map<String, String> ui = collectUiEnv();
        try {
            PlanInputTabularIo.TabularSheet main =
                    deliveryCalendarViewTabController != null
                            ? deliveryCalendarViewTabController.snapshotMainCompareForExport()
                            : new PlanInputTabularIo.TabularSheet(List.of(), List.of());
            Path out = SummaryAiDispatchWorkbookExporter.writeFromPipelineArtifacts(ui, main);
            ensureSummaryAiDispatchWorkbookEnvPath(out);
            appendLog("[summary-ai-dispatch] パイプライン後エクセル出力: " + out);
        } catch (Exception ex) {
            appendLog(
                    "[summary-ai-dispatch] パイプライン後エクセル出力失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
        if (scheduleDeliveryReload && deliveryCalendarViewTabController != null) {
            deliveryCalendarViewTabController.reloadInBackgroundAfterStage2Success();
        }
    }

    /**
     * 既定出力の最新計画 JSON で設備ガント（グラフィック）を再読み込みする。段階2・配台試行完了後に呼ぶ。
     */
    void refreshEquipmentGanttGraphicAfterPipelineRun() {
        if (equipmentGanttGraphicTabController != null) {
            equipmentGanttGraphicTabController.syncLatestPlanJsonFromOutputDirAndReload(false);
        }
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
            Path newestPlan = Stage2OutputNaming.newestPrimaryPlanXlsx(dir);
            if (newestPlan == null) {
                newestPlan = Stage2OutputNaming.newestPrimaryPlanJson(dir);
            }
            Path newestMember = Stage2OutputNaming.newestPrimaryMemberXlsx(dir);
            if (newestMember == null) {
                newestMember = Stage2OutputNaming.newestPrimaryMemberJson(dir);
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

    /**
     * 段階1/2・プローブスクリプト起動時の Python 実行ファイル。
     *
     * @see StagePythonExecutable#resolve(Map)
     */
    public Path resolveStagePythonExecutablePath(Map<String, String> ui) {
        return StagePythonExecutable.resolve(ui);
    }

    /** {@link #resolveStagePythonExecutablePath(Map)} を現在の環境変数タブの値で解決する。 */
    public Path resolveStagePythonExecutablePath() {
        return resolveStagePythonExecutablePath(collectUiEnv());
    }

    /**
     * シェル未結線など {@link MainShellController} が無いときのフォールバック（テスト・退避経路）。
     *
     * @see StagePythonExecutable#defaultPythonPathWhenShellMissing()
     */
    public static Path defaultPythonPathWhenShellMissing() {
        return StagePythonExecutable.defaultPythonPathWhenShellMissing();
    }

    private void maybePortableFirstLaunchEnvInit() {
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        if (!PortableBundleSelfUpdater.isPortableBundleLayout(cwd)) {
            return;
        }
        Path marker = cwd.resolve(AppPaths.PORTABLE_FIRST_LAUNCH_MARKER_FILE);
        if (!Files.isRegularFile(marker)) {
            return;
        }
        try {
            FactorySite firstLaunchSite =
                    FactorySite.inferFromPortableBundleInitSetting(cwd)
                            .orElseGet(GlobalInitSettingTarget::load);
            GlobalInitSettingTarget.save(firstLaunchSite);
            applyEnvRowsFullBundledResetAndPersist(true, firstLaunchSite);
            applyBundledPortableDefaultsIfPresent();
            applyRepoFolderPathNormalization();
            DesktopSessionStateStore.save(collectDesktopSession());
            Files.deleteIfExists(marker);
            appendLog(
                    "[startup] 初回起動: 環境変数を初期化し（工場既定="
                            + firstLaunchSite.displayLabelJa()
                            + "）、"
                            + AppPaths.PORTABLE_FIRST_LAUNCH_MARKER_FILE
                            + " を削除しました。");
        } catch (Exception ex) {
            appendLog(
                    "[startup] 初回起動の環境変数初期化に失敗（"
                            + AppPaths.PORTABLE_FIRST_LAUNCH_MARKER_FILE
                            + " は残します）: "
                            + ex.getMessage());
        }
    }

    /**
     * ポータブル配布: 正本が {@link AppPaths#KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR} に設定され、{@code version.txt} がローカルより新しいときに
     * {@code pm-ai-data} を同期する。正本はディレクトリ（リポジトリルート）または {@code .zip}（ZIP 隣に外付け {@code version.txt}）。
     */
    private void maybePortableBundleSelfUpdate() {
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        if (!PortableBundleSelfUpdater.isPortableBundleLayout(cwd)) {
            appendLog(
                    "[startup] 自動バージョンアップは対象外（PMD.exe 直下に pm-ai-data のポータブル配布レイアウトがありません）。"
                            + " user.dir="
                            + PortableBundleSelfUpdater.safePathForLog(cwd));
            return;
        }
        appendLog("[startup] 自動バージョンアップ: ポータブル配布を検出しました。user.dir=" + PortableBundleSelfUpdater.safePathForLog(cwd));
        Map<String, String> ui = collectUiEnv();
        String raw = ui.get(AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR);
        if (raw == null || raw.isBlank()) {
            Alert a = new Alert(AlertType.INFORMATION);
            a.initOwner(primaryStage);
            applyAlertStylesheetsFromOwner(a);
            a.setTitle("自動バージョンアップ");
            a.setHeaderText(null);
            a.setContentText(
                    "ポータブル配布では "
                            + AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR
                            + "（正本フォルダまたはバージョンアップ用 ZIP のパス）が空です。\n"
                            + "自動バージョンアップは行いません。そのまま続行します。");
            a.show();
            appendLog(
                    "[startup] "
                            + AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR
                            + " が未設定のためポータル同期をスキップしました。");
            return;
        }
        Path canonical = Path.of(raw.trim()).toAbsolutePath().normalize();
        Path localData = cwd.resolve("pm-ai-data").normalize();
        if (!PortableBundleSelfUpdater.isValidPortableBundleCanonical(canonical)) {
            appendLog(
                    "[startup] 正本パスにアクセスできません: "
                            + PortableBundleSelfUpdater.safePathForLog(canonical));
            Alert w = new Alert(AlertType.WARNING);
            w.initOwner(primaryStage);
            applyAlertStylesheetsFromOwner(w);
            w.setTitle("自動バージョンアップ");
            w.setHeaderText(null);
            w.setContentText(
                    "正本フォルダまたは ZIP を開けませんでした。自動バージョンアップはスキップします。\n"
                            + PortableBundleSelfUpdater.safePathForLog(canonical));
            w.show();
            return;
        }
        Optional<BigDecimal> cv = PortableBundleSelfUpdater.readCanonicalPortableBundleVersion(canonical);
        Optional<BigDecimal> lv = PortableBundleSelfUpdater.readLocalBundleVersion(cwd, localData);
        if (!PortableBundleSelfUpdater.shouldUpdate(cv, lv)) {
            String reason =
                    cv.isEmpty()
                            ? "正本の version.txt が読めません（ZIP の隣、または pm-ai-package-release 直下）。"
                            : "ローカル版が正本以上です（更新不要）。";
            appendLog(
                    "[startup] 自動バージョンアップはスキップ: "
                            + reason
                            + " 正本="
                            + cv.map(BigDecimal::toPlainString).orElse("（なし）")
                            + " ローカル="
                            + lv.map(BigDecimal::toPlainString).orElse("（なし・0扱い）")
                            + " 正本パス="
                            + PortableBundleSelfUpdater.safePathForLog(canonical));
            return;
        }
        appendLog(
                "[startup] 自動バージョンアップ: 正本の方が新しいため確認ダイアログを表示します。正本="
                        + cv.map(BigDecimal::toPlainString).orElse("?")
                        + " ローカル="
                        + lv.map(BigDecimal::toPlainString).orElse("（なし・初回）"));
        String canonVerStr = cv.map(BigDecimal::toPlainString).orElse("?");
        String localVerStr = lv.map(BigDecimal::toPlainString).orElse("（なし・初回）");
        Optional<Path> upgradeZip = PortableBundleSelfUpdater.resolveEffectiveUpgradeZip(canonical);
        String syncHint =
                upgradeZip.isPresent()
                        ? "ZIP を展開して pm-ai-data に同期します。"
                        : "正本フォルダから pm-ai-data へファイルを同期します。";
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.initOwner(primaryStage);
        applyAlertStylesheetsFromOwner(confirm);
        confirm.setTitle("自動バージョンアップ");
        confirm.setHeaderText(null);
        confirm.setContentText(
                "正本のバージョン（"
                        + canonVerStr
                        + "）がローカル pm-ai-data（"
                        + localVerStr
                        + "）より新しいです。\n"
                        + syncHint
                        + " 実行してよいですか？");
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            appendLog("[startup] ポータル同期はユーザー操作によりスキップしました（版 " + canonVerStr + " → 保留）。");
            return;
        }

        Stage wait = new Stage();
        wait.initModality(Modality.APPLICATION_MODAL);
        wait.initOwner(primaryStage);
        wait.setTitle("自動バージョンアップ");
        wait.setMinWidth(520);
        wait.setMinHeight(220);
        VBox root = new VBox(20);
        root.setAlignment(Pos.CENTER);
        root.setStyle("-fx-padding: 28;");
        Label msg = new Label("正本から pm-ai-data を更新しています…");
        msg.setWrapText(true);
        msg.setAlignment(Pos.CENTER);
        msg.setMaxWidth(460);
        ProgressIndicator pi = new ProgressIndicator();
        pi.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
        root.getChildren().addAll(msg, pi);
        wait.setScene(new Scene(root, 520, 220));
        wait.show();

        final Path[] extractedHolder = new Path[1];
        Task<Void> task =
                new Task<>() {
                    @Override
                    protected Void call() throws Exception {
                        Path syncSource;
                        Optional<Path> zipForSync =
                                PortableBundleSelfUpdater.resolveEffectiveUpgradeZip(canonical);
                        if (zipForSync.isPresent()) {
                            Path tmp =
                                    PortableBundleSelfUpdater.extractUpgradeZipToTempDirectory(
                                            zipForSync.get(),
                                            line -> Platform.runLater(() -> appendLog(line)));
                            extractedHolder[0] = tmp;
                            syncSource = tmp.resolve("pm-ai-data");
                            if (!Files.isDirectory(syncSource)) {
                                throw new IOException(
                                        "ZIP 内に pm-ai-data フォルダがありません: " + zipForSync.get());
                            }
                        } else {
                            syncSource = PortableBundleSelfUpdater.resolveSyncSourceRoot(canonical);
                        }
                        PortableBundleSelfUpdater.syncFromCanonical(
                                syncSource,
                                localData,
                                line -> Platform.runLater(() -> appendLog(line)));
                        PortableBundleSelfUpdater.copyOuterVersionTxtToLocal(canonical, cwd, localData);
                        return null;
                    }
                };
        task.setOnSucceeded(
                e -> {
                    if (extractedHolder[0] != null) {
                        PortableBundleSelfUpdater.deleteDirectoryRecursive(
                                extractedHolder[0],
                                line -> Platform.runLater(() -> appendLog(line)));
                    }
                    wait.close();
                    try {
                        InitSettingPersistence.applyPortableUpgradeOverwriteFromPmAiData(
                                localData, collectUiEnv());
                        DesktopSessionStateStore.applyPortableUpgradeBundledPolicyToSessionStore(collectUiEnv());
                        TableColumnOrderPersistence.overwriteTableColumnOrderStoreAfterPortableUpgrade(
                                collectUiEnv());
                    } catch (IOException ex) {
                        appendLog(
                                "[startup] バージョンアップ後のバンドル既定（タブ／列順／配台不要 JSON パス）の上書きに失敗: "
                                        + ex.getMessage());
                    }
                    performGlobalUiFactoryResetWithoutConfirmation();
                    applyBundledPortableDefaultsIfPresent();
                    mainRunTabController.clearMainRunTabLog();
                    applyRepoFolderPathNormalization();
                    Optional<FactorySite> chosenOpt = promptFactorySiteAfterPortableUpgrade();
                    FactorySite siteAfterUpgrade = chosenOpt.orElse(FactorySite.KONAN);
                    if (chosenOpt.isEmpty()) {
                        appendLog(
                                "[startup] 工場既定の選択をキャンセルしたため湖南工場の既定を適用します。");
                    }
                    applyFactorySitePortableAndNetworkDefaults(siteAfterUpgrade);
                    ensureBootstrapDefaultValuesVisible(collectUiEnv());
                    ensureUiRefOptionalDisplayDefaultsVisible(collectUiEnv());
                    applyRepoFolderPathNormalization();
                    DesktopSessionStateStore.save(collectDesktopSession());
                    mainRunTabController.refreshAppVersionLabel();
                    mainRunTabController.refreshOpenWorkbookHintLabels();
                    mainRunTabController.refreshFactorySiteLogo();
                    appendLog(
                            "[startup] ポータル同期が完了しました（version.txt・pm-ai-data／init_setting をリポジトリへ反映）。"
                                    + "グローバル設定「デフォルトに戻す」相当で UI をバンドル既定へ揃えました。"
                                    + " 工場既定: "
                                    + siteAfterUpgrade.displayLabelJa()
                                    + "。");
                });
        task.setOnFailed(
                e -> {
                    if (extractedHolder[0] != null) {
                        PortableBundleSelfUpdater.deleteDirectoryRecursive(
                                extractedHolder[0],
                                line -> Platform.runLater(() -> appendLog(line)));
                    }
                    wait.close();
                    Throwable ex = task.getException();
                    String detail = ex != null ? ex.getMessage() : "不明なエラー";
                    appendLog("[startup] ポータル同期に失敗: " + detail);
                    Alert er = new Alert(AlertType.WARNING);
                    er.initOwner(primaryStage);
                    applyAlertStylesheetsFromOwner(er);
                    er.setTitle("自動バージョンアップ");
                    er.setHeaderText(null);
                    er.setContentText("正本からの同期に失敗しました。\n" + detail);
                    er.show();
                });
        Thread t = new Thread(task, "pm-ai-portable-sync");
        t.setDaemon(true);
        t.start();
    }

    /**
     * jpackage 配布の {@code pm-ai-data/}（{@code package_app.ps1} が同梱）があるとき、{@link AppPaths#KEY_PM_AI_OUTPUT_DIR} をインストール直下の
     * {@code pm-ai-data/output} に寄せる。ネットワークソース正本（{@code PM_AI_TASK_INPUT_SOURCE_DIR} / {@code PM_AI_ACTUAL_DETAIL_SOURCE_DIR}）は上書きしない。
     */
    private void applyBundledPortableDefaultsIfPresent() {
        if (envRows == null) {
            return;
        }
        if (!bundledPortableStage1MarkerPresent()) {
            return;
        }
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        Path repo = cwd.resolve("pm-ai-data").toAbsolutePath().normalize();
        Path outDir = repo.resolve("output");
        try {
            Files.createDirectories(outDir);
        } catch (IOException ignored) {
            /* UI にはパスだけ反映；作成失敗はユーザー環境で対応 */
        }
        for (EnvVarRow r : envRows) {
            String name = r.getName() != null ? r.getName().trim() : "";
            if (AppPaths.KEY_PM_AI_OUTPUT_DIR.equals(name)) {
                r.setValue(outDir.toString());
            }
        }
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
        if (rows == envRows) {
            ensureUiRefOptionalDisplayDefaultsVisible(collectUiEnv());
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
        String v = bootstrapDefaultValueForKey(k, ui);
        if (!v.isBlank()) {
            r.setValue(v);
        }
    }

    /**
     * 環境変数タブ「値」列に出すブートストラップ既定（新規行・空欄補完・初期化と同一ソース）。
     *
     * @param ui リポジトリ根などの解決に使う（空マップ可）
     */
    private static String bootstrapDefaultValueForKey(String k, Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        if (k == null || k.isBlank()) {
            return "";
        }
        switch (k) {
            case AppPaths.KEY_PM_AI_PYTHON -> {
                return StagePythonExecutable.defaultPmAiPythonForBootstrap();
            }
            case AppPaths.KEY_PM_AI_REPO_ROOT -> {
                return AppPaths.resolveRepoRoot(u).toString();
            }
            case AppPaths.KEY_PM_AI_CODE_PYTHON_DIR -> {
                return AppPaths.resolvePythonScriptDir(u).toString();
            }
            case AppPaths.KEY_PM_AI_WORKSPACE -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR -> {
                return AppPaths.DEFAULT_PM_AI_TASK_INPUT_SOURCE_DIR;
            }
            case AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR -> {
                return AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_SOURCE_DIR;
            }
            case AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR -> {
                return AppPaths.resolveResultDispatchTableDir(u).toString();
            }
            case AppPaths.KEY_PM_AI_OUTPUT_DIR -> {
                return AppPaths.resolveDefaultOutputDir(u).toString();
            }
            case AppPaths.KEY_GEMINI_CREDENTIALS_JSON -> {
                Path root = AppPaths.resolveRepoRoot(u);
                Path underCode =
                        root.resolve("code").resolve("gemini_credentials.encrypted.json");
                if (Files.isRegularFile(underCode)) {
                    return underCode.toAbsolutePath().normalize().toString();
                }
                Path atRoot = root.resolve("gemini_credentials.encrypted.json");
                return Files.isRegularFile(atRoot)
                        ? atRoot.toAbsolutePath().normalize().toString()
                        : "";
            }
            case AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON -> {
                return AppPaths.resolveDefaultExcludeRulesJsonPath(u).map(Path::toString).orElse("");
            }
            case AppPaths.KEY_PM_AI_MASTER_WORKBOOK -> {
                return AppPaths.resolveMasterWorkbookCandidate(u).map(Path::toString).orElse("");
            }
            case AppPaths.KEY_PM_AI_COLUMN_CONFIG_WORKBOOK,
                    AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK,
                    AppPaths.KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET -> {
                return "1";
            }
            case AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON,
                    AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR -> {
                return AppPaths.DEFAULT_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR;
            }
            default -> {
                return "";
            }
        }
    }

    /**
     * {@link #BOOTSTRAP_ORDER} に載る行のうち値が空のものへ {@link #bootstrapDefaultValueForKey} を適用する。
     * セッション復元後など、テーブルに空セルが残る場合の再補完に使う。
     */
    private void ensureBootstrapDefaultValuesVisible(Map<String, String> ui) {
        if (envRows == null) {
            return;
        }
        Map<String, String> ctx = ui != null ? ui : Map.of();
        for (EnvVarRow row : envRows) {
            String k = row.getName() != null ? row.getName().trim() : "";
            if (k.isEmpty() || !BOOTSTRAP_KEY_SET.contains(k)) {
                continue;
            }
            String cur = row.getValue();
            if (cur != null && !cur.isBlank()) {
                continue;
            }
            if (AppPaths.KEY_PM_AI_MASTER_WORKBOOK.equals(k)) {
                String mf = envTabValueTrimmed(AppPaths.KEY_MASTER_WORKBOOK_FILE);
                if (!mf.isEmpty()) {
                    continue;
                }
            }
            String v = bootstrapDefaultValueForKey(k, ctx);
            if (!v.isBlank()) {
                row.setValue(v);
            }
        }
    }

    /** 環境タブの現在行からキーに対応する値を trim して返す（無ければ空）。 */
    private String envTabValueTrimmed(String key) {
        if (envRows == null || key == null || key.isBlank()) {
            return "";
        }
        for (EnvVarRow row : envRows) {
            String n = row.getName() != null ? row.getName().trim() : "";
            if (key.equals(n)) {
                String v = row.getValue();
                return v != null ? v.trim() : "";
            }
        }
        return "";
    }

    /**
     * {@code ui_ref_env_defaults.json} 由来の行のうち、ブートストラップ以外で「空欄＝planning_core / AppPaths の既定と同じ意味」のキーへ、
     * 値列に解決済みの既定を表示する（子プロセスへ渡す意味は従来どおりで、未設定と同じキーは空のままにするものは触らない）。
     */
    private void ensureUiRefOptionalDisplayDefaultsVisible(Map<String, String> ui) {
        if (envRows == null) {
            return;
        }
        Map<String, String> ctx = ui != null ? ui : Map.of();
        for (EnvVarRow row : envRows) {
            String k = row.getName() != null ? row.getName().trim() : "";
            if (k.isEmpty()) {
                continue;
            }
            String cur = row.getValue();
            if (cur != null && !cur.isBlank()) {
                continue;
            }
            String v = optionalUiRefDisplayDefaultForKey(k, ctx);
            if (!v.isBlank()) {
                row.setValue(v);
            }
        }
    }

    /**
     * {@link #ensureUiRefOptionalDisplayDefaultsVisible} 用。キーごとの表示既定（実ファイルがあるときのみパスを返すものあり）。
     */
    private static String optionalUiRefDisplayDefaultForKey(String k, Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        if (k == null || k.isBlank()) {
            return "";
        }
        Path codeDir = AppPaths.resolveRepoRoot(u).resolve("code");
        return switch (k) {
            case AppPaths.KEY_MASTER_WORKBOOK_FILE -> "master.xlsm";
            case PlanInputTabController.ENV_TASK_PLAN_SHEET ->
                    PlanInputTabController.DEFAULT_PLAN_INPUT_SHEET_NAME;
            case "MASTER_SPEED_SHEET_NAME" -> "speed";
            case "MASTER_SPEED_FIRST_EXCEL_COL" -> "4";
            case AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK ->
                    AppPaths.summaryAiDispatchXlsxPath(u).toString();
            case "RAW_FABRIC_WIDTH_TABLE_PATH" -> {
                Path p = codeDir.resolve("使用原反, 加工幅.txt");
                yield Files.isRegularFile(p)
                        ? p.toAbsolutePath().normalize().toString()
                        : "";
            }
            case "ROLL_UNIT_BY_USED_RAW_TABLE_PATH" -> {
                String out = "";
                for (String fn :
                        List.of("使用原反,ロール単位の長さ.txt", "使用原反, ロール単位の長さ.txt")) {
                    Path p = codeDir.resolve(fn);
                    if (Files.isRegularFile(p)) {
                        out = p.toAbsolutePath().normalize().toString();
                        break;
                    }
                }
                yield out;
            }
            case "PRODUCT_WIDTH_TABLE_PATH" -> {
                Path p = codeDir.resolve("製品名, 製品幅.txt");
                yield Files.isRegularFile(p)
                        ? p.toAbsolutePath().normalize().toString()
                        : "";
            }
            case "PRODUCT_LENGTH_TABLE_PATH" -> {
                Path p = codeDir.resolve("製品名,製品長.txt");
                yield Files.isRegularFile(p)
                        ? p.toAbsolutePath().normalize().toString()
                        : "";
            }
            case "PRODUCT_THICKNESS_TABLE_PATH" -> {
                Path p = codeDir.resolve("製品名,製品厚み.txt");
                yield Files.isRegularFile(p)
                        ? p.toAbsolutePath().normalize().toString()
                        : "";
            }
            case "DISPATCH_TRIAL_PATTERN_LIST_SHEET" -> "配台試行順_パターン一覧";
            case "DISPATCH_PATTERN_STAGE2_SUMMARY_SHEET" -> "配台試行順_パターン別段階2";
            default -> "";
        };
    }

    private static EnvVarRow newBootstrapRow(String k, Map<String, String> ui) {
        EnvVarRow r = new EnvVarRow();
        r.setName(k);
        r.setDescription(EnvVarDocs.mergeDescriptions("", k));
        r.setValue(bootstrapDefaultValueForKey(k, ui));
        return r;
    }
}
