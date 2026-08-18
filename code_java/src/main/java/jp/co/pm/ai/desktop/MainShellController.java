package jp.co.pm.ai.desktop;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDate;
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
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
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
import javafx.scene.control.PasswordField;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.TreeItem;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.image.ImageView;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.Modality;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.stage.WindowEvent;
import javafx.util.Duration;
import javafx.util.StringConverter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;


import jp.co.pm.ai.desktop.ui.SevenDigitChallenge;
import jp.co.pm.ai.desktop.ui.SevenDigitChallengeDialog;
import jp.co.pm.ai.desktop.ui.ThemedAlertContentSupport;
import jp.co.pm.ai.desktop.ui.TodayDispatchSourceSelectionDialog;
import jp.co.pm.ai.planning.stage2.source.Stage1SourceBundle;
import jp.co.pm.ai.planning.stage2.source.Stage1SourceBundleCompletionGate;
import jp.co.pm.ai.planning.stage2.source.Stage1SourceBundleIo;
import jp.co.pm.ai.planning.stage2.source.Stage1SourcePairMatcher;
import jp.co.pm.ai.planning.stage2.source.Stage2SkipTodayDispatchPolicy;
import jp.co.pm.ai.planning.stage2.source.Stage2SourceConsistencyGuard;
import jp.co.pm.ai.planning.stage2.source.Stage2SourceGuardCoordinator;
import jp.co.pm.ai.planning.stage2.source.Stage2SourceGuardSnapshot;

import jp.co.pm.ai.desktop.audio.MacroCompleteChime;
import jp.co.pm.ai.desktop.audio.UiClickSound;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.bridge.Stage2PythonChildEnv;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreviewPython;
import jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths;
import jp.co.pm.ai.desktop.dispatch.rules.stage.DispatchRuleBuilderRunContext;
import jp.co.pm.ai.desktop.dispatch.rules.stage.DispatchRuleStageRunOverlay;
import jp.co.pm.ai.desktop.dispatch.rules.trace.DispatchRuleTraceLoader;
import jp.co.pm.ai.desktop.bridge.StagePythonExecutable;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.PipelineDownstreamResultsClearer;
import jp.co.pm.ai.desktop.config.PipelineLocalResultsPolicy;
import jp.co.pm.ai.desktop.config.PlanningCoreMaterialTableAppendProbe;
import jp.co.pm.ai.desktop.config.RemoteDesktopEnvRows;
import jp.co.pm.ai.desktop.config.RemoteSupportLogArchive;
import jp.co.pm.ai.desktop.config.SharedPipelineResultsCleaner;
import jp.co.pm.ai.desktop.config.Stage1AiCacheClearer;
import jp.co.pm.ai.desktop.config.WorkspaceCacheArchiveStore;
import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;
import jp.co.pm.ai.desktop.config.DispatchTrialLogUiStore;
import jp.co.pm.ai.desktop.config.JvmMemoryLogStore;
import jp.co.pm.ai.desktop.config.MainShellTabLayoutDefaults;
import jp.co.pm.ai.desktop.config.MainShellTabLayoutNode;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.OperatorActionLogStore;
import jp.co.pm.ai.desktop.config.OperatorUserPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.FactorySiteOperatorAccess;
import jp.co.pm.ai.desktop.config.FactorySiteWorkspaceMigrator;
import jp.co.pm.ai.desktop.config.FactorySiteWorkspaceSnapshot;
import jp.co.pm.ai.desktop.config.FactorySiteWorkspaceStore;
import jp.co.pm.ai.desktop.config.PortableBundleUpgradeUiSnapshot;
import jp.co.pm.ai.desktop.config.GeminiDispatchModelTryOrderDefaults;
import jp.co.pm.ai.desktop.config.EnvVarsInitializedAtStore;
import jp.co.pm.ai.desktop.config.EnvVarsInitialTemplate;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.config.LastLaunchedFactorySiteStore;
import jp.co.pm.ai.desktop.config.StartupFactorySiteResolver;
import jp.co.pm.ai.desktop.config.DesktopTheme;
import jp.co.pm.ai.desktop.config.PushButtonCssEmitter;
import jp.co.pm.ai.desktop.config.PushButtonDesignPrefs;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.gemini.GeminiFreeTierModelsRefreshService;
import jp.co.pm.ai.desktop.config.PortableBundleBuildManifest;
import jp.co.pm.ai.desktop.config.PortableBundlePendingUpdate;
import jp.co.pm.ai.desktop.config.PortableBundleSelfUpdater;
import jp.co.pm.ai.desktop.config.PortableBundleUpdateLauncher;
import jp.co.pm.ai.desktop.config.PortableBundleUpgradeFollowUp;
import jp.co.pm.ai.desktop.config.PortableBundleUpgradeLog;
import jp.co.pm.ai.desktop.config.PortableBundleUpgradeProgress;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.config.PlanWorkspaceSessionFragment;
import jp.co.pm.ai.desktop.config.PlanWorkspaceSnapshotStore;
import jp.co.pm.ai.desktop.config.EnvVarDocs;
import jp.co.pm.ai.desktop.config.InitSettingPersistence;
import jp.co.pm.ai.desktop.config.UiEnvRowSnapshot;
import jp.co.pm.ai.desktop.config.UiRefEnvDefaults;
import jp.co.pm.ai.desktop.ui.AttendanceGridCellSizing;
import jp.co.pm.ai.desktop.ui.UiRowHoverDimmingSettings;
import jp.co.pm.ai.desktop.ui.EnvVarsStartupCheckBusyDialog;
import jp.co.pm.ai.desktop.ui.FactorySiteSwitchBusyDialog;
import jp.co.pm.ai.desktop.ui.FactorySiteSwitchBusySupport;
import jp.co.pm.ai.desktop.ui.GlobalAppStatusBar;
import jp.co.pm.ai.desktop.ui.ShellFactoryOperatorToolbar;
import jp.co.pm.ai.desktop.ui.StageRunBusyDialog;
import jp.co.pm.ai.desktop.ui.StageRunLogProgressParser;
import jp.co.pm.ai.desktop.ui.Stage1NewMaterialLookupDialog;
import jp.co.pm.ai.desktop.ui.Stage1EcSideUnknownDialog;
import jp.co.pm.ai.desktop.ui.Stage1EcSideUnknownDialogResult;
import jp.co.pm.ai.desktop.ui.MissingSkillsSheetColumnDialog;
import jp.co.pm.ai.desktop.ui.Stage2UnknownMasterCombinationDialog;
import jp.co.pm.ai.desktop.ui.Stage2UnknownMasterCombinationDialogResult;
import jp.co.pm.ai.desktop.ui.ButtonPressFeedback;
import jp.co.pm.ai.desktop.ui.MainStageScreenGeometry;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.runtime.MemoryJvmRingLog;
import jp.co.pm.ai.desktop.dispatch.RawInputMorningDispatchRateAnalyzer;
import jp.co.pm.ai.desktop.dispatch.RawInputMorningDispatchRateWarning;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchDocument;
import jp.co.pm.ai.desktop.io.Stage2EquipmentGanttContractPaths;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPythonExport;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.ProcessOwnedLockFiles;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;
import jp.co.pm.ai.desktop.ipc.IpcStdoutTap;
/**
 * Main window controller（従来は {@link PmAiFxApp} 内蔵だった業務ロジックを分離）。
 * Layout: {@code MainShell.fxml} and tab FXML files.
 */
public final class MainShellController
        implements DesktopShellHost, EnvTabShellHost, StartupTabBackgroundLoadCoordinator.Host {

    /**
     * {@link Tab#getProperties()} に登録済みかどうか。選択変更時に見出し chrome を再適用するリスナーを二重登録しない。
     */
    private static final String PROP_SHELL_TAB_SELECTION_CHROME_LISTENER =
            "pmShellTabSelectionChromeListener";

    private static final String STAGE1 = "task_extract_stage1.py";
    private static final String STAGE2 = "plan_simulation_stage2.py";
    private static final String STAGE2_1 = "plan_simulation_stage2_1.py";

    private static PipelineExecutionTimingKind pipelineTimingKindForStageScript(String script) {
        if (STAGE1.equals(script)) {
            return PipelineExecutionTimingKind.STAGE1;
        }
        if (STAGE2.equals(script)) {
            return PipelineExecutionTimingKind.STAGE2_0;
        }
        if (STAGE2_1.equals(script)) {
            return PipelineExecutionTimingKind.STAGE2_1;
        }
        return null;
    }

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
                    AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
                    AppPaths.KEY_PM_AI_OUTPUT_DIR,
                    AppPaths.KEY_PM_AI_REPO_ROOT,
                    AppPaths.KEY_PM_AI_CODE_PYTHON_DIR,
                    AppPaths.KEY_PM_AI_CODE_DIR,
                    "PM_AI_AGENT_DEBUG_SESSION",
                    "PM_AI_DEBUG_LOG",
                    AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH);

    /** 段階2実行前ログに出す「入力解決に関わる」環境変数キー。 */
    private static final List<String> STAGE2_CHILD_INPUT_ENV_KEYS =
            List.of(
                    AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                    PlanInputTabController.ENV_TASK_PLAN_SHEET,
                    AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
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
                    AppPaths.KEY_PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON,
                    "PM_AI_AGENT_DEBUG_SESSION",
                    "PM_AI_DEBUG_LOG",
                    AppPaths.KEY_PM_AI_CURSOR_DEBUG_LOG,
                    AppPaths.KEY_PM_AI_DEBUG_LOG_MIRROR);

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
                    AppPaths.KEY_MASTER_WORKBOOK_FILE,
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
                    AppPaths.KEY_PM_AI_DAILY_REPORT_SOURCE_DIR,
                    AppPaths.KEY_PM_AI_ORDER_DETAIL_SOURCE_DIR,
                    AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR,
                    AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
                    AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE,
                    AppPaths.KEY_PM_AI_MACHINE_DELIVERY_MANAGEMENT_XLSM,
                    AppPaths.KEY_PM_AI_REQUEST_FORM_TPI_PDF_DIR,
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
    private StackPane shellFactoryLogoHost;

    @FXML
    private ImageView shellFactoryLogoImageView;

    @FXML
    private Label shellFactoryLogoCaptionLabel;

    @FXML
    private ComboBox<FactorySite> shellFactorySiteCombo;

    @FXML
    private Label shellOperatorUserLabel;

    @FXML
    private Button shellChangeSessionOperatorButton;

    @FXML
    private Button shellChangeOperatorPinButton;

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
    private Label envVarsInitializedAtLabel;

    @FXML
    private Label globalStatusMessageLabel;

    @FXML
    private ProgressIndicator globalStatusProgressIndicator;

    @FXML
    private ProgressBar globalStatusProgressBar;

    @FXML
    private Label globalStatusTabLabel;

    @FXML
    private Label globalStatusOperatorLabel;

    @FXML
    private Label globalStatusFactoryLabel;

    @FXML
    private Label globalStatusAttendanceLabel;

    @FXML
    private Label globalStatusMemoryLabel;

    @FXML
    private Button dispatchUsageGuideButton;

    @FXML
    private MainRunTabController mainRunTabController;

    @FXML
    private EnvTabController envTabController;

    @FXML
    private MemorySettingsTabController memorySettingsTabController;

    @FXML
    private GlobalSettingsTabController globalSettingsTabController;

    @FXML
    private UserProfilesTabController userProfilesTabController;

    @FXML
    private OperatorUserManagementTabController operatorUserManagementTabController;

    @FXML
    private CompanyCalendarTabController companyCalendarTabController;

    @FXML
    private MemberAttendanceTabController memberAttendanceTabController;

    @FXML
    private MachineCalendarTabController machineCalendarTabController;

    @FXML
    private PlanInputTabController planInputTabController;

    @FXML
    private RequestFormInputTabController requestFormInputTabController;

    @FXML
    private RequestFormPipelineCheckTabController requestFormPipelineCheckTabController;

    @FXML
    private RemoteDesktopTabController remoteDesktopTabController;

    @FXML
    private Stage1PreviewTabController stage1PreviewTabController;

    @FXML
    private ExcludeRulesTabController excludeRulesTabController;

    @FXML
    private SpecialRulesTabController specialRulesTabController;

    @FXML
    private ActualsStatusTabController actualsStatusTabController;

    @FXML
    private DailyReportCsvTabController dailyReportCsvTabController;

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
    private RequestFormPreviewBadgeDesignTabController requestFormPreviewBadgeDesignTabController;

    @FXML
    private PushButtonDesignTabController pushButtonDesignTabController;

    @FXML
    private OperatorCardTabController operatorCardTabController;

    @FXML
    private PlanWorkspaceHistoryTabController planWorkspaceHistoryTabController;

    private WorkspaceCacheHistoryTabController workspaceCacheHistoryTabController;

    @FXML
    private OperatorActionLogTabController operatorActionLogTabController;

    @FXML
    private ApiModelBenchmarkTabController apiModelBenchmarkTabController;

    @FXML
    private PipelineExecutionTimingTabController pipelineExecutionTimingTabController;

    @FXML
    private EquipmentStatusDashboardTabController equipmentStatusDashboardTabController;

    @FXML
    private CodeDispatchLookupTablesTabController codeDispatchLookupTablesTabController;

    @FXML
    private Tab mainShellTabEquipmentStatusDashboard;

    @FXML
    private Tab mainShellTabRun;

    @FXML
    private Tab mainShellTabPipelineExecutionTiming;

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
    private Tab mainShellTabUserProfiles;

    @FXML
    private Tab mainShellTabOperatorUserManagement;

    @FXML
    private Tab mainShellTabCompanyCalendar;

    @FXML
    private Tab mainShellTabMemberAttendance;

    @FXML
    private Tab mainShellTabMachineCalendar;

    @FXML
    private Tab mainShellTabMasterSummary;

    @FXML
    private Tab mainShellTabPlanInput;

    @FXML
    private Tab mainShellTabRequestFormInput;

    @FXML
    private Tab mainShellTabRequestFormPipelineCheck;

    @FXML
    private Tab mainShellTabRemoteDesktop;

    @FXML
    private Tab mainShellTabStage1Preview;

    @FXML
    private Tab mainShellTabExcludeRules;

    @FXML
    private Tab mainShellTabSpecialRules;

    @FXML
    private Tab mainShellTabActualsStatus;

    @FXML
    private Tab mainShellTabDailyReportCsvView;

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
    private Tab mainShellTabRequestFormPreviewBadgeDesign;

    @FXML
    private Tab mainShellTabOperatorCard;

    @FXML
    private Tab mainShellTabPlanWorkspaceHistory;

    @FXML
    private Tab mainShellTabCacheHistory;

    @FXML
    private Tab mainShellTabOperatorActionLog;

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

    /** ユーザー管理者タブをこのセッションで解錠済みか。 */
    private boolean operatorUserAdminTabUnlocked;

    private final AtomicBoolean suppressOperatorUserAdminTabGuard = new AtomicBoolean(false);

    private ObservableList<EnvVarRow> envRows;

    private final AtomicBoolean runLock = new AtomicBoolean(false);

    /** {@link Platform#exit()} 等、確認なしで閉じる内部終了用。 */
    private volatile boolean suppressCloseConfirmation;

    private final Stage2IdentityCloseGate stage2IdentityCloseGate = new Stage2IdentityCloseGate();

    /** 終了確認済み。2回目の WINDOW_CLOSE_REQUEST はゲートせず閉じる。 */
    private volatile boolean applicationCloseProceeding;

    private final PipelineExecutionTimingHistoryStore pipelineExecutionTimingHistory =
            new PipelineExecutionTimingHistoryStore();

    /** Non-null while a stage script is running; equals {@link #STAGE1} or {@link #STAGE2}. */
    private volatile String activeRunStageScript;

    /** Python child process while stage 1/2 is running; cleared on completion or interrupt. */
    private final AtomicReference<Process> activeStageChildProcess = new AtomicReference<>();

    /** 「中断」ボタンで {@link Process#destroyForcibly()} した直後は true（子の exit=1 を cancel 扱いにする）。 */
    private final AtomicBoolean activeStageRunUserCancelled = new AtomicBoolean(false);

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

    /** Gemini 無料枠 Flash-Lite モデル一覧の日次 {@code models.list} 更新。 */
    private GeminiFreeTierModelsRefreshService geminiFreeTierModelsRefreshService;

    /** 納期管理ビュー再読み込み中のタブ差し戻しで {@link TabPane} の選択リスナーを再入しない。 */
    private final AtomicBoolean suppressDeliveryCalendarReloadTabGuard = new AtomicBoolean(false);

    /** メンバー勤怠の未保存確認でタブ差し戻し時に選択リスナーを再入しない。 */
    private final AtomicBoolean suppressMemberAttendanceUnsavedTabGuard = new AtomicBoolean(false);
    private final AtomicBoolean suppressCompanyCalendarUnsavedTabGuard = new AtomicBoolean(false);

    private volatile boolean memberAttendanceDirtySinceSave = false;
    private volatile boolean machineCalendarDirtySinceSave = false;
    private volatile boolean companyCalendarDirtySinceSave = false;

    /** 納期管理ビュー再読み込み中は段階1～段階3.5 の実行ボタンを無効化する。 */
    private final AtomicBoolean deliveryCalendarReloadBlockingStageRuns = new AtomicBoolean(false);

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

    /** OFF中も段階3タブの保存済み位置・グループを失わないための完全レイアウト（レガシー復元用）。 */
    private List<MainShellTabLayoutNode> completeMainShellTabLayout = List.of();

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

    /** 工場切替後の session-state.json 保存 debounce（§7）。 */
    private final PauseTransition sessionPersistDebounce = new PauseTransition(Duration.millis(300));

    private volatile boolean factorySiteSwitchInProgress;
    /** Assigned in {@link #installUiEnvAutoSave()} for debounced {@link #scheduleDesktopSessionSave()}. */
    private Runnable uiEnvPersistSchedule;
    private final AtomicBoolean envResetInProgress = new AtomicBoolean(false);

    /**
     * 環境変数初期化未記録時に他タブへ遷移しようとしたとき、{@link #ensureMainShellEnvTabSelected()} の再入を防ぐ。
     */
    private final AtomicBoolean suppressEnvVarsInitTabGuard = new AtomicBoolean(false);

    /** ゲスト操作者時に他タブへ遷移しようとしたとき、{@link #ensureMainShellRunTabSelected()} の再入を防ぐ。 */
    private final AtomicBoolean suppressGuestSessionTabGuard = new AtomicBoolean(false);

    /**
     * ポータル自動バージョンアップ実行中は、起動時の操作者ダイアログと依頼書原本フォルダ案内を出さない。
     */
    private final AtomicBoolean deferOperatorPromptForPortableUpgrade = new AtomicBoolean(false);

    /** 起動時点で環境変数タブの値が初期化テンプレートと一致しなかったとき {@code true}（セッション中は再評価しない）。 */
    private final AtomicBoolean envVarsDifferFromInitialAtStartup = new AtomicBoolean(false);

    /** 起動時の環境変数テンプレート照合が完了したら {@code true}（完了前は初期化済みでも保守的にブロック）。 */
    private final AtomicBoolean envVarsStartupCheckCompleted = new AtomicBoolean(false);

    /** 環境変数未初期化によるタブブロックログをセッション中 1 回に抑える。 */
    private final AtomicBoolean envInitTabBlockLogEmitted = new AtomicBoolean(false);
    private GlobalAppStatusBar globalAppStatusBar;
    private ShellFactoryOperatorToolbar factoryOperatorToolbar;
    private StartupTabBackgroundLoadCoordinator startupTabBackgroundLoad;
    private volatile String startupBackgroundLoadMessage = "";
    private volatile String globalLongTaskDetail = "";
    private volatile Double globalLongTaskProgress;
    private volatile boolean startupTabBackgroundLoadActive;
    private String lastGlobalLogLine = "";

    /** 起動時の環境変数確認モーダル（表示中のみ非 null）。 */
    private EnvVarsStartupCheckBusyDialog envVarsStartupCheckBusy;

    /** 起動シーケンス（ワークスペース復元〜環境照合〜BG 読込）の進行中。 */
    private volatile boolean startupSequenceActive;
    private volatile boolean startupRestoredFactorySite;

    /** 起動 BG 読込完了後に進捗モーダルを閉じる。 */
    private volatile boolean startupAwaitingBackgroundLoadBeforeModalClose;

    /** 工場切替中の進捗モーダル（表示中のみ非 null）。 */
    private FactorySiteSwitchBusyDialog factorySiteSwitchBusy;

    /** 工場切替後のタブ再読込完了まで進捗モーダルを維持する。 */
    private volatile boolean factorySwitchAwaitingBackgroundLoadBeforeModalClose;

    private FactorySite factorySwitchBusyFrom;
    private FactorySite factorySwitchBusyTo;

    /** 段階1／2 実行中の進捗モーダル（表示中のみ非 null）。 */
    private StageRunBusyDialog stageRunBusyDialog;

    /**
     * バージョンアップ後処理で操作者を復元済みなら、同起動での操作者ダイアログと依頼書原本フォルダ案内を省略する。
     */
    private final AtomicBoolean skipOperatorPromptAfterPortableUpgrade = new AtomicBoolean(false);

    private DesktopTheme pendingTheme = DesktopTheme.LIGHT;

    /** FXML 読込直後に固定した既定見出し（内部 ID は {@link MainShellTabId#key()} のまま）。 */
    private final Map<MainShellTabId, String> mainShellTabBaselineTitles = new EnumMap<>(MainShellTabId.class);

    /** セッション保存する見出しエイリアス（キーは {@link MainShellTabId#key()}）。 */
    private final Map<String, String> mainShellTabTitleAliases = new LinkedHashMap<>();

    /** 固定子タブ見出し色（{@link jp.co.pm.ai.desktop.config.MainShellInnerTabColorKeys}）。 */
    private final Map<String, String> innerTabHeaderColorByKey = new LinkedHashMap<>();

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

            pipelineExecutionTimingHistory.configureFromUi(ui0);
            pipelineExecutionTimingHistory.setPersistLog(this::appendLog);
            FactoryOperatorUserStore.configureFromUi(ui0);

            factoryOperatorToolbar =
                    new ShellFactoryOperatorToolbar(
                            shellFactoryLogoHost,
                            shellFactoryLogoImageView,
                            shellFactoryLogoCaptionLabel,
                            shellFactorySiteCombo,
                            shellOperatorUserLabel,
                            shellChangeSessionOperatorButton,
                            shellChangeOperatorPinButton);
            factoryOperatorToolbar.wire(this);

            mainRunTabController.bindShell(this);
            mainRunTabController.setCalendarReadinessBlocked(
                    true, "勤怠・カレンダーの準備状態を確認中…");
            if (equipmentStatusDashboardTabController != null) {
                equipmentStatusDashboardTabController.bindShell(this);
            }
            envTabController.bindShell(this);
            memorySettingsTabController.bindShell(this);
            if (globalSettingsTabController != null) {
                globalSettingsTabController.bindShell(this);
            }
            if (userProfilesTabController != null) {
                userProfilesTabController.bindShell(this);
            }
            if (operatorUserManagementTabController != null) {
                operatorUserManagementTabController.bindShell(this);
            }
            if (companyCalendarTabController != null) {
                companyCalendarTabController.bindShell(this);
            }
            if (memberAttendanceTabController != null) {
                memberAttendanceTabController.bindShell(this);
            }
            if (machineCalendarTabController != null) {
                machineCalendarTabController.bindShell(this);
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
            if (requestFormPreviewBadgeDesignTabController != null) {
                requestFormPreviewBadgeDesignTabController.bindShell(this);
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
        if (requestFormInputTabController != null) {
            requestFormInputTabController.bindShell(this);
        }
        if (requestFormPipelineCheckTabController != null) {
            requestFormPipelineCheckTabController.bindShell(this);
        }
        refreshStage1PipelineCheckGate();
        if (remoteDesktopTabController != null) {
            remoteDesktopTabController.bindShell(this);
        }
        startupTabBackgroundLoad = new StartupTabBackgroundLoadCoordinator(this);
        stage1PreviewTabController.bindShell(this);
        if (codeDispatchLookupTablesTabController != null) {
            codeDispatchLookupTablesTabController.bindShell(this);
        }
        excludeRulesTabController.bindShell(this);
        specialRulesTabController.bindShell(this);
        actualsStatusTabController.bindShell(this);
        if (dailyReportCsvTabController != null) {
            dailyReportCsvTabController.bindShell(this);
        }
        deliveryCalendarViewTabController.bindShell(this);
        resultDispatchTableTabController.bindShell(this);
        dispatchInteractiveTabController.bindShell(this);
        if (planWorkspaceHistoryTabController != null) {
            planWorkspaceHistoryTabController.bindShell(this);
        }
        if (workspaceCacheHistoryTabController != null) {
            workspaceCacheHistoryTabController.bindShell(this);
        }
        if (operatorActionLogTabController != null) {
            operatorActionLogTabController.bindShell(this);
        }
        if (apiModelBenchmarkTabController != null) {
            apiModelBenchmarkTabController.bindShell(this);
        }
        if (pipelineExecutionTimingTabController != null) {
            pipelineExecutionTimingTabController.bindShell(this);
        }

        primaryStage.setMinWidth(640);
        primaryStage.setMinHeight(480);

            applyDesktopSession(DesktopSessionStateStore.load());
            FactorySite startupFactory = StartupFactorySiteResolver.resolve();
            GlobalInitSettingTarget.save(startupFactory);
            FactoryOperatorUserStore.configureFromUi(collectUiEnv(), startupFactory);
            if (globalSettingsTabController != null) {
                globalSettingsTabController.refreshInitSettingTargetComboFromStore();
            }
            if (mainRunTabController != null) {
                factoryOperatorToolbar.refreshFactorySiteLogo();
            }
            refreshEnvVarsInitializedAtToolbarLabel();
            if (equipmentStatusDashboardTabController != null) {
                equipmentStatusDashboardTabController.resetDashboardDatesToToday();
            }
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

        globalAppStatusBar =
                new GlobalAppStatusBar(
                        globalStatusMessageLabel,
                        globalStatusProgressIndicator,
                        globalStatusProgressBar,
                        globalStatusTabLabel,
                        globalStatusOperatorLabel,
                        globalStatusFactoryLabel,
                        globalStatusAttendanceLabel,
                        globalStatusMemoryLabel);
        globalAppStatusBar.startMemoryMonitor(primaryStage);
        refreshGlobalStatusBar();

        installUiEnvAutoSave();
        installMainStageGeometryAutoSave();

        geminiFreeTierModelsRefreshService =
                new GeminiFreeTierModelsRefreshService(
                        this::snapshotUiEnv, this::onGeminiFreeTierModelsRefreshFinished);
        geminiFreeTierModelsRefreshService.start();

        applyRepoFolderPathNormalization();
        maybePortableFirstLaunchEnvInit();
        maybeForceEnvInitAfterPortableUpgradeRestart();
        applyRunTabGating();

        probeNetworkSourceDirsAtStartup();

        primaryStage.setOnCloseRequest(
                e -> {
                    if (applicationCloseProceeding) {
                        return;
                    }
                    e.consume();
                    if (!confirmAttendanceTabsUnsavedBeforeLeave("終了")) {
                        return;
                    }
                    beginApplicationCloseSequence();
                });

        primaryStage.setOnShown(
                e -> {
                    primaryStage.toFront();
                    primaryStage.requestFocus();
                    applyPendingMainShellTabLayoutFromSessionIfNeeded();
                    installLazyMainShellTabContentForStartup();
                    if (isEnvVarsInitializationPending()) {
                        ensureMainShellEnvTabSelected();
                    }
                    activateMainShellTabHeavyContentRecursive(
                            tabPane.getSelectionModel().getSelectedItem());
                    applyRunTabGating();
                    if (!isEnvVarsInitializationPending()
                            && tabPane.getSelectionModel().getSelectedItem() == null
                            && !tabPane.getTabs().isEmpty()) {
                        tabPane.getSelectionModel().selectFirst();
                    }
                    Platform.runLater(
                            () ->
                                    Platform.runLater(
                                            () -> {
                                                refreshMainShellTabHeaderChromeFromStoredColors();
                                                scheduleEquipmentStatusDashboardInitialReloadIfSelected();
                                                scheduleRequestFormPipelineCheckInitialRefreshIfSelected();
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
                            if (blockMainShellTabSelectionIfEnvInitPending()) {
                                return;
                            }
                            if (blockMainShellTabSelectionIfGuestSessionOnly()) {
                                return;
                            }
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
                            if (blockMemberAttendanceUnsavedTabNavigation(prevTab, newTab)) {
                                return;
                            }
                            if (blockCompanyCalendarUnsavedTabNavigation(prevTab, newTab)) {
                                return;
                            }
                            if (!suppressLazyMainShellTabContentSwap.get()) {
                                deferMainShellTabBranchHeavyContent(prevTab);
                                activateMainShellTabHeavyContentRecursive(newTab);
                            }
                            emitShellTabNavigation();
                            refreshGlobalStatusBar();
                            /* :selected 由来の -fx-text-fill がインラインより後勝ちになることがあるため再適用 */
                            if (!suppressMainShellTabChromeRefresh.get()) {
                                refreshMainShellTabHeaderChromeFromStoredColors();
                            }
                            if (newTab == mainShellTabEquipmentStatusDashboard
                                    && equipmentStatusDashboardTabController != null) {
                                equipmentStatusDashboardTabController.onMainShellTabSelected();
                            }
                            if (newTab == mainShellTabRequestFormInput
                                    && requestFormInputTabController != null
                                    && !startupTabBackgroundLoadActive) {
                                Platform.runLater(
                                        requestFormInputTabController::onMainShellTabSelected);
                            }
                            if (newTab == mainShellTabRequestFormPipelineCheck
                                    && requestFormPipelineCheckTabController != null
                                    && !startupTabBackgroundLoadActive) {
                                Platform.runLater(
                                        requestFormPipelineCheckTabController
                                                ::onMainShellTabSelected);
                            }
                            if (newTab == mainShellTabRemoteDesktop
                                    && remoteDesktopTabController != null) {
                                Platform.runLater(
                                        remoteDesktopTabController::onMainShellTabSelected);
                            }
                            if (newTab == mainShellTabCompanyCalendar
                                    && companyCalendarTabController != null
                                    && !startupTabBackgroundLoadActive) {
                                Platform.runLater(
                                        companyCalendarTabController::onMainShellTabSelected);
                            }
                            if (newTab == mainShellTabMemberAttendance
                                    && memberAttendanceTabController != null
                                    && !startupTabBackgroundLoadActive) {
                                Platform.runLater(
                                        memberAttendanceTabController::onMainShellTabSelected);
                            }
                            if (newTab == mainShellTabMachineCalendar
                                    && machineCalendarTabController != null
                                    && !startupTabBackgroundLoadActive) {
                                Platform.runLater(
                                        machineCalendarTabController::onMainShellTabSelected);
                            }
                            if (prevTab == mainShellTabEquipmentStatusDashboard
                                    && equipmentStatusDashboardTabController != null) {
                                equipmentStatusDashboardTabController.onMainShellTabDeselected();
                            }
                            if (prevTab == mainShellTabRequestFormInput
                                    && requestFormInputTabController != null) {
                                requestFormInputTabController.onMainShellTabDeselected();
                            }
                            if (prevTab == mainShellTabRemoteDesktop
                                    && remoteDesktopTabController != null) {
                                remoteDesktopTabController.onMainShellTabDeselected();
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
                            if (newTab == mainShellTabPipelineExecutionTiming
                                    && pipelineExecutionTimingTabController != null) {
                                pipelineExecutionTimingTabController.refreshFromStore();
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
        ButtonPressFeedback.installOnScene(scene);
        UiClickSound.warmUp(collectUiEnv());
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
        ButtonPressFeedback.installOnScene(scene);
    }

    public void unregisterThemeTrackedScene(Scene scene) {
        themeTrackedSecondaryScenes.remove(scene);
    }

    /**
     * 残業シミュレーションウィザード用。グローバル {@link DesktopTheme} は適用せず、紙面風の固定 CSS のみ載せる
     * （ダークテーマ時に TableView が紺色化するのを防ぐ）。
     */
    public void registerOvertimeWizardScene(Scene scene) {
        if (scene == null) {
            return;
        }
        if (!scene.getStylesheets().contains(PM_AI_DESKTOP_CSS)) {
            scene.getStylesheets().add(PM_AI_DESKTOP_CSS);
        }
        ButtonPressFeedback.installOnScene(scene);
    }

    private void refreshThemeTrackedSecondaryScenes() {
        DesktopTheme t = currentDesktopTheme();
        for (Scene s : themeTrackedSecondaryScenes) {
            t.applyTo(s);
        }
    }

    private void applyDesktopSession(DesktopSessionState s) {
        applyDesktopSession(s, true, false, false);
    }

    private void applyDesktopSession(
            DesktopSessionState s, boolean restoreUiEnvRowsFromSession, boolean restoreMainRunLogLines) {
        applyDesktopSession(s, restoreUiEnvRowsFromSession, restoreMainRunLogLines, false);
    }

    /**
     * @param restoreUiEnvRowsFromSession {@code false} のとき環境変数タブはセッションから復元せず、呼び出し元で構築済みの行を保持する（ポータル
     *     バージョンアップ直後のバンドル既定への初期化後など）。
     * @param restoreMainRunLogLines {@code true} のとき実行・ログタブのログ行・スクロールをセッションから復元する。アプリ起動時は {@code false}（ログは空で開始）。
     * @param globalInitSettingOnly {@code true} のとき環境変数初期化の正本（環境タブ・{@code PM_AI_*} 派生パス）を適用しない。
     */
    private void applyDesktopSession(
            DesktopSessionState s,
            boolean restoreUiEnvRowsFromSession,
            boolean restoreMainRunLogLines,
            boolean globalInitSettingOnly) {
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
        if (!globalInitSettingOnly) {
            planInputTabController.restoreDesktopSessionPaths(s.planInputPath(), s.planInputSheet());
            stage1PreviewTabController.restoreDesktopSessionPaths(
                    s.stage1PreviewPath(), s.stage1PreviewSheet());
            excludeRulesTabController.restoreDesktopSessionPath(s.excludeRulesPath());
            if (nonBlank(s.mainRunWorkbook())) {
                mainRunTabController.getWorkbookField().setText(s.mainRunWorkbook());
            }
            if (nonBlank(s.mainRunScriptDir())) {
                mainRunTabController.getScriptDirField().setText(s.mainRunScriptDir());
            }
            if (nonBlank(s.mainRunStage2ProductionPlan())
                    || nonBlank(s.mainRunStage2MemberSchedule())) {
                mainRunTabController.setStage2ArtifactPaths(
                        nz(s.mainRunStage2ProductionPlan()),
                        nz(s.mainRunStage2MemberSchedule()));
            }
        }
        mainRunTabController.applyLogFontFromSession(s.logFontFamily(), s.logFontSize());
        List<String> runLogLines = restoreMainRunLogLines ? s.mainRunLogLines() : List.of();
        double runLogScroll =
                restoreMainRunLogLines ? s.mainRunLogScroll() : Double.NaN;
        mainRunTabController.restoreRunLogUiFromSession(
                s.mainRunLogFilter(), runLogLines, runLogScroll);
        if (!globalInitSettingOnly) {
            mainRunTabController.applyTodayDispatchModeFromSession(
                    s.mainRunStage2SkipTodayDispatch(), s.planInputTodayDispatch());
            planInputTabController.applyStage2SkipGeminiApiFromSession(s.planInputStage2SkipGeminiApi());
            planInputTabController.refreshNextDayDialogRadioCoupling();
            planInputTabController.applyStage2NextDayDialogModeFromSession(
                    s.planInputStage2NextDayDialogMode());
            planInputTabController.applyComboSheetMayExceedNeedFromSession(
                    s.planInputComboSheetMayExceedNeed());
            mainRunTabController.applyStage2ResultBookFontFromSession(s.mainRunStage2ResultBookFont());
            mainRunTabController.applySkipGeminiApiFromSession(s.mainRunSkipGeminiApi());
        }
        /*
         * 設備ガントの apply は末尾で Canvas を再構築し personBadgeStyleResolverForGantt を参照する。
         * 担当バッジのセッション（グロー等）を先に適用しないと、起動直後の帯は既定スタイルで描かれる。
         */
        if (ganttPersonBadgeDesignTabController != null) {
            ganttPersonBadgeDesignTabController.applyPersonBadgeDesignSession(s);
        }
        equipmentGanttGraphicTabController.applyEquipmentGanttSession(s);
        if (equipmentStatusDashboardTabController != null) {
            equipmentStatusDashboardTabController.applyDashboardSession(s);
        }
        setTableRowHoverDimmingEnabled(s.tableRowHoverDimmingEnabled());
        if (globalSettingsTabController != null) {
            globalSettingsTabController.syncTableRowHoverDimmingCheckbox();
        }
        if (uiBadgeDesignTabController != null) {
            uiBadgeDesignTabController.applyUiBadgeSession(s);
        }
        if (requestFormPreviewBadgeDesignTabController != null) {
            requestFormPreviewBadgeDesignTabController.applyRequestFormPreviewBadgeSession(s);
        }
        if (requestFormInputTabController != null) {
            requestFormInputTabController.applyComboChoicesFromSession(
                    resolveRequestFormComboChoices(s));
            requestFormInputTabController.reloadJuchuHeaderAliasRegistry(
                    GlobalInitSettingTarget.loadEffective(collectUiEnv()), collectUiEnv(), false);
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
        factoryOperatorToolbar.refreshFactorySiteLogo();
        Platform.runLater(() -> excludeRulesTabController.tryStartupLoadFromPathField());
    }

    private void applyWindowGeometry(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        MainStageScreenGeometry.applyToStage(
                primaryStage, MainStageScreenGeometry.fromSessionState(s));
    }

    private void installMainStageGeometryAutoSave() {
        Runnable schedule = () -> schedulePersistSessionDebounced();
        primaryStage.xProperty().addListener((obs, oldV, newV) -> schedule.run());
        primaryStage.yProperty().addListener((obs, oldV, newV) -> schedule.run());
        primaryStage.widthProperty().addListener((obs, oldV, newV) -> schedule.run());
        primaryStage.heightProperty().addListener((obs, oldV, newV) -> schedule.run());
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
        if (requestFormPreviewBadgeDesignTabController != null) {
            requestFormPreviewBadgeDesignTabController.flushEditsBeforeSnapshot();
        }
        if (pushButtonDesignTabController != null) {
            pushButtonDesignTabController.flushEditsBeforeSnapshot();
        }
        MainStageScreenGeometry.Snapshot windowGeometry =
                MainStageScreenGeometry.snapshotFromStage(primaryStage);
        return new DesktopSessionState(
                planInputTabController.snapshotPlanInputPath(),
                planInputTabController.snapshotPlanInputSheet(),
                stage1PreviewTabController.snapshotStage1PreviewPath(),
                stage1PreviewTabController.snapshotStage1PreviewSheet(),
                excludeRulesTabController.snapshotExcludeRulesPath(),
                nz(mainRunTabController.getWorkbookField().getText()),
                nz(mainRunTabController.getScriptDirField().getText()),
                windowGeometry.width(),
                windowGeometry.height(),
                windowGeometry.x(),
                windowGeometry.y(),
                windowGeometry.screenVisualMinX(),
                windowGeometry.screenVisualMinY(),
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
                mainRunTabController.snapshotStage2SkipTodayDispatch(),
                planInputTabController.snapshotStage2NextDayDialogMode().name(),
                planInputTabController.snapshotComboSheetMayExceedNeed(),
                planInputTabController.snapshotStage2SkipGeminiApi(),
                mainRunTabController.snapshotTodayDispatch(),
                mainRunTabController.snapshotStage2ResultBookFont(),
                mainRunTabController.snapshotSkipGeminiApi(),
                false,
                snapshotUiEnvRows(),
                snapshotMainShellTabOrder(),
                snapshotMainShellTabLayout(),
                snapshotMainShellTabTitleAliases(),
                snapshotInnerTabSelectedIndexByShellTabKey(),
                snapshotInnerTabHeaderColorByKey(),
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
                equipmentGanttGraphicTabController.snapshotEquipmentGanttPrepTimeLabelsEnabled(),
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
                requestFormPreviewBadgeDesignTabController != null
                        ? requestFormPreviewBadgeDesignTabController.snapshotPreviewBadgeLabel()
                        : "更新",
                requestFormPreviewBadgeDesignTabController != null
                        ? requestFormPreviewBadgeDesignTabController.snapshotPreviewBadgeStyle()
                        : PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault(),
                mainShellTabOrganizerHeaderGlowEnabled.get(),
                getMainShellTabOrganizerHeaderGlowStrength(),
                pushButtonDesignTabController != null
                        ? pushButtonDesignTabController.snapshotPrefs()
                        : PushButtonDesignPrefs.inactiveDefaults(),
                memorySettingsTabController.snapshotMemoryMonitorEnabled(),
                memorySettingsTabController.snapshotMemoryMonitorIntervalSec(),
                memorySettingsTabController.snapshotNextLaunchHeapFixed(),
                memorySettingsTabController.snapshotNextLaunchHeapMinMiB(),
                memorySettingsTabController.snapshotNextLaunchHeapMaxMiB(),
                equipmentStatusDashboardTabController != null
                        ? equipmentStatusDashboardTabController.snapshotActualDateIso()
                        : "",
                equipmentStatusDashboardTabController != null
                        ? equipmentStatusDashboardTabController.snapshotPlanDateIso()
                        : "",
                0,
                0,
                equipmentStatusDashboardTabController == null
                        || equipmentStatusDashboardTabController.snapshotAutoRefreshEnabled(),
                equipmentStatusDashboardTabController != null
                        ? equipmentStatusDashboardTabController.snapshotAutoRefreshIntervalSec()
                        : DesktopSessionState.DEFAULT_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                equipmentStatusDashboardTabController == null
                        || equipmentStatusDashboardTabController.snapshotShowAladdinPlans(),
                equipmentStatusDashboardTabController == null
                        || equipmentStatusDashboardTabController.snapshotShowDispatchPlans(),
                equipmentStatusDashboardTabController != null
                        ? equipmentStatusDashboardTabController.snapshotAppearancePrefs()
                        : EquipmentStatusDashboardAppearancePrefs.defaults(),
                tableRowHoverDimmingEnabled,
                requestFormInputTabController != null
                        ? requestFormInputTabController.snapshotComboChoices()
                        : jp.co.pm.ai.desktop.reconciliation.RequestFormComboChoices.empty());
    }

    /** 設備ガントのプレビュー用に、バッジ「既定」スタイルを返す。 */
    public PersonBadgeStyle currentPersonBadgeStyleForGantt() {
        return ganttPersonBadgeDesignTabController != null
                ? ganttPersonBadgeDesignTabController.previewStyleForGantt()
                : PersonBadgeStyle.defaultStyle();
    }

    /** 依頼書プレビュー・原本更新バッジの表示設定。 */
    public jp.co.pm.ai.desktop.reconciliation.RequestFormPreviewBadgeConfig requestFormPreviewBadgeConfig() {
        if (requestFormPreviewBadgeDesignTabController != null) {
            return requestFormPreviewBadgeDesignTabController.snapshotPreviewBadgeConfig();
        }
        return jp.co.pm.ai.desktop.reconciliation.RequestFormPreviewBadgeConfig.defaults();
    }

    /** 依頼書入力タブのプレビュー上部バッジ見た目を再描画する。 */
    public void refreshRequestFormPreviewBadgeAppearance() {
        if (requestFormInputTabController != null) {
            requestFormInputTabController.refreshPreviewBadgeAppearance();
        }
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

    /**
     * master.xlsm を OS 既定アプリ（Excel）で開く。
     *
     * @param logPrefix ログ行の接頭辞（例: {@code "[company-calendar]"}）
     * @return 開いたとき {@code true}
     */
    public boolean openMasterWorkbookInDesktop(String logPrefix) {
        return openMasterWorkbookInDesktop(logPrefix, false);
    }

    /**
     * master.xlsm を OS 既定アプリ（Excel）で開く。
     *
     * @param logPrefix ログ行の接頭辞（例: {@code "[company-calendar]"}）
     * @param readOnly 読み取り専用で開く（会社カレンダー・メンバー勤怠の閲覧向け）
     * @return 開いたとき {@code true}
     */
    public boolean openMasterWorkbookInDesktop(String logPrefix, boolean readOnly) {
        Path target = resolveMasterWorkbookIfPresent();
        if (target == null) {
            Path attempted =
                    AppPaths.resolveMasterWorkbookPathForDesktopOpen(
                            collectUiEnv(), nz(mainRunTabController.getWorkbookField().getText()));
            appendLog(logPrefix + " master not found: " + attempted);
            return false;
        }
        try {
            if (readOnly) {
                DesktopFileOpener.openFileReadOnly(target);
                appendLog(logPrefix + " opened master (read-only): " + target);
            } else {
                DesktopFileOpener.openFile(target);
                appendLog(logPrefix + " opened master: " + target);
            }
            return true;
        } catch (Exception e) {
            appendLog(logPrefix + " open master failed: " + e.getMessage());
            return false;
        }
    }

    /**
     * 勤怠・機械カレンダー.xlsx を OS 既定アプリ（Excel）で読み取り専用で開く。
     *
     * @param logPrefix ログ行の接頭辞（例: {@code "[company-calendar]"}）
     * @return 開いたとき {@code true}
     */
    public boolean openAttendanceCalendarXlsxInDesktop(String logPrefix) {
        Path target = AppPaths.attendanceCalendarXlsxPath(collectUiEnv());
        if (!Files.isRegularFile(target)) {
            appendLog(logPrefix + " attendance calendar xlsx not found: " + target);
            return false;
        }
        try {
            DesktopFileOpener.openFileReadOnly(target);
            appendLog(logPrefix + " opened attendance calendar (read-only): " + target);
            return true;
        } catch (Exception e) {
            appendLog(logPrefix + " open attendance calendar failed: " + e.getMessage());
            return false;
        }
    }

    /** 環境変数・工場ワークスペース確定後に勤怠 readiness を更新し、表示済みタブ（会社・メンバー・機械）のみ再読込する。 */
    private void reloadAttendanceTabsFromJson() {
        reloadAttendanceTabsFromJson(false);
    }

    private void reloadAttendanceTabsFromJson(boolean force) {
        Path jsonPath = AppPaths.attendanceDataJsonPath(collectUiEnv());
        appendLog("[attendance] JSON 再読込: " + jsonPath);
        if (companyCalendarTabController != null) {
            companyCalendarTabController.reloadAttendanceDataFromJsonIfEnabled();
        }
        if (memberAttendanceTabController != null) {
            memberAttendanceTabController.reloadAttendanceDataFromJsonIfEnabled();
        }
        if (machineCalendarTabController != null) {
            machineCalendarTabController.reloadMachineCalendarDataIfEnabled();
        }
        refreshAttendanceReadiness(force);
    }

    /**
     * 起動後バックグラウンド読込（勤怠 JSON 再読込・タブ順次プリロード）を許可するか。
     * {@link #isEnvVarsInitializationPending()} より厳格（ポータブル版アップグレード待ち中も抑制）。
     */
    private boolean isStartupBackgroundLoadAllowed() {
        if (!envVarsStartupCheckCompleted.get()) {
            return false;
        }
        if (!EnvVarsInitializedAtStore.isRecorded()) {
            return false;
        }
        return !envVarsDifferFromInitialAtStartup.get();
    }

    /**
     * 工場切替完了後のバックグラウンド読込を許可するか。
     * 起動時とは異なり env 差分ブロック中でもリモート・原本等を再読込する。
     */
    private boolean isFactorySwitchBackgroundLoadAllowed() {
        if (!envVarsStartupCheckCompleted.get()) {
            return false;
        }
        if (FactoryOperatorUserStore.isGuestSession()) {
            return true;
        }
        return EnvVarsInitializedAtStore.isRecorded();
    }

    /**
     * 環境変数・工場ワークスペース確定後にタブを順次バックグラウンド読込する。
     */
    private void maybeReloadAttendanceTabsAfterEnvReady() {
        if (!isStartupBackgroundLoadAllowed()) {
            return;
        }
        reloadAttendanceTabsFromJson();
        runAfterUiPulse(
                () -> {
                    if (startupTabBackgroundLoad != null) {
                        startupTabBackgroundLoad.resetAndSchedule();
                    }
                });
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
        if (requestFormInputTabController != null) {
            requestFormInputTabController.persistInputSettings();
        }
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

        Path canonical = AppPaths.resolveResultDispatchTableStage2JsonPath(collectUiEnv());
        Path parent = canonical.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Files.copy(snapJson, canonical, StandardCopyOption.REPLACE_EXISTING);

        tryExportResultDispatchTableXlsxNearJson(canonical);

        PlanWorkspaceSessionFragment frag = PlanWorkspaceSnapshotStore.readSessionFragment(entry);
        DesktopSessionState merged = frag.mergeOnto(collectDesktopSession());
        applyDesktopSession(merged, false, true, false);
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

    public jp.co.pm.ai.desktop.reconciliation.JuchuHeaderAliasRegistry
            snapshotJuchuHeaderAliasRegistryForExport() {
        if (requestFormInputTabController != null) {
            return requestFormInputTabController.snapshotJuchuHeaderAliasRegistry();
        }
        return jp.co.pm.ai.desktop.reconciliation.JuchuHeaderAliasRegistry.loadForFactory(
                GlobalInitSettingTarget.load(), collectUiEnv());
    }

    /**
     * 工場別グローバル既定（{@code init_setting}）から依頼書フォーム候補・列定義を反映する。
     *
     * @param restoreJuchuFromInitSetting {@code true} のとき列定義を init_setting の JSON で上書き（デフォルトに戻す用）
     */
    public void applyFactoryRequestFormGlobalSettings(
            FactorySite site, boolean restoreJuchuFromInitSetting) {
        FactorySite effective = site != null ? site : GlobalInitSettingTarget.load();
        if (requestFormInputTabController != null) {
            requestFormInputTabController.applyComboChoicesFromSession(
                    jp.co.pm.ai.desktop.reconciliation.RequestFormInputSettingsStore.loadComboChoices(
                            collectUiEnv(), effective));
            requestFormInputTabController.reloadJuchuHeaderAliasRegistry(
                    effective, collectUiEnv(), restoreJuchuFromInitSetting);
        }
    }

    private jp.co.pm.ai.desktop.reconciliation.RequestFormComboChoices resolveRequestFormComboChoices(
            DesktopSessionState session) {
        jp.co.pm.ai.desktop.reconciliation.RequestFormComboChoices fromSummary =
                jp.co.pm.ai.desktop.reconciliation.RequestFormInputSettingsStore.loadComboChoices(
                        collectUiEnv(), GlobalInitSettingTarget.load());
        if (fromSummary != null && !fromSummary.isEmpty()) {
            return fromSummary;
        }
        jp.co.pm.ai.desktop.reconciliation.RequestFormComboChoices saved =
                session != null
                        ? session.requestFormComboChoices()
                        : jp.co.pm.ai.desktop.reconciliation.RequestFormComboChoices.empty();
        if (saved != null && !saved.isEmpty()) {
            return saved.mergedWithDefaults();
        }
        return DesktopSessionStateStore.loadFactoryRequestFormComboChoices(collectUiEnv())
                .mergedWithDefaults();
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
        applyDesktopSession(state, true, true, false);
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
        if (requestFormPreviewBadgeDesignTabController != null) {
            requestFormPreviewBadgeDesignTabController.flushEditsBeforeSnapshot();
        }
        if (pushButtonDesignTabController != null) {
            pushButtonDesignTabController.flushEditsBeforeSnapshot();
        }
        DesktopSessionStateStore.save(collectDesktopSessionForGlobalPersistence());
    }

    @Override
    public Stage primaryStageForDialogs() {
        return primaryStage;
    }

    /**
     * メインウィンドウと同じテーマ CSS をダイアログに載せる（{@link Alert} / {@link ChoiceDialog} 等）。
     */
    public void prepareDialogForMainTheme(Dialog<?> dialog) {
        if (dialog == null) {
            return;
        }
        initDialogOwnerIfSceneReady(dialog);
        applyAlertStylesheetsFromOwner(dialog);
    }

    /** {@link Dialog#initOwner} はオーナー Stage に Scene が無いと JavaFX 26 で NPE になる。 */
    private void initDialogOwnerIfSceneReady(Dialog<?> dialog) {
        if (dialog == null || primaryStage == null) {
            return;
        }
        if (primaryStage.getScene() != null) {
            dialog.initOwner(primaryStage);
        }
    }

    /** ダイアログ表示直後に入力欄へフォーカスし、すぐタイピングできるようにする。 */
    private static void focusInputWhenDialogShown(Dialog<?> dialog, javafx.scene.Node input) {
        if (dialog == null || input == null) {
            return;
        }
        dialog.setOnShown(e -> Platform.runLater(input::requestFocus));
    }

    /** 保存・読込完了などの情報ダイアログ。 */
    public void showInformationDialog(String title, String message) {
        showThemedAlert(AlertType.INFORMATION, title, null, message);
    }

    /** ファイルなし・部分成功などの注意ダイアログ。 */
    public void showWarningDialog(String title, String message) {
        showThemedAlert(AlertType.WARNING, title, null, message);
    }

    /** 段階2正常終了後: 原反投入日制約で午前配台率が 50% 未満の暦日があれば警告。 */
    void showRawInputMorningDispatchRateWarningAfterStage2() {
        if (planInputTabController == null) {
            return;
        }
        Map<String, LocalDate> rawDates = planInputTabController.collectEffectiveRawInputDateByTaskId();
        if (rawDates.isEmpty()) {
            return;
        }
        Path json = AppPaths.resolveResultDispatchTableJsonPath(collectUiEnv());
        logRawInputMorningDispatchRateWarningIfAny(json, rawDates);
        RawInputMorningDispatchRateWarning.showIfNeeded(this, primaryStage, json, rawDates);
    }

    private void logRawInputMorningDispatchRateWarningIfAny(
            Path resultDispatchJson, Map<String, LocalDate> rawInputByTaskId) {
        Path contract = Stage2EquipmentGanttContractPaths.resolveNearResultDispatchJson(resultDispatchJson);
        if (contract == null) {
            return;
        }
        try {
            var result = RawInputMorningDispatchRateAnalyzer.analyze(contract, rawInputByTaskId);
            if (!result.hasWarnings()) {
                return;
            }
            appendLog(
                    "[原反投入日・午前配台率] 警告: "
                            + result.lowRateDays().size()
                            + " 暦日で午前配台率が "
                            + (int) (RawInputMorningDispatchRateAnalyzer.RATE_THRESHOLD * 100)
                            + "% 未満");
        } catch (Exception ignored) {
            // ダイアログ側で再分析する
        }
    }

    /** 失敗時のエラーダイアログ。 */
    public void showErrorDialog(String title, String message) {
        showThemedAlert(AlertType.ERROR, title, null, message);
    }

    private boolean showNormalApplicationCloseConfirm() {
        Alert alert = new Alert(AlertType.CONFIRMATION);
        initDialogOwnerIfSceneReady(alert);
        applyAlertStylesheetsFromOwner(alert);
        alert.setTitle("終了確認");
        alert.setHeaderText(null);
        StringBuilder msg =
                new StringBuilder("工程管理 AI 配台を終了します。よろしいですか？");
        if (isPipelineRunLocked()) {
            msg.append("\n\n段階処理が実行中です。終了すると処理は中断されます。");
        }
        alert.setContentText(msg.toString());
        Optional<ButtonType> ans = alert.showAndWait();
        return ans.isPresent() && ans.get() == ButtonType.OK;
    }

    private void beginApplicationCloseSequence() {
        if (suppressCloseConfirmation) {
            finishApplicationCloseAfterConfirm();
            return;
        }
        if (!stage2IdentityCloseGate.stage2CompletedThisLaunch()) {
            if (showNormalApplicationCloseConfirm()) {
                finishApplicationCloseAfterConfirm();
            }
            return;
        }
        Stage wait = showIdentityCompareWaitDialog();
        Map<String, String> ui = snapshotUiEnv();
        Thread worker =
                new Thread(
                        () -> {
                            Stage2IdentityCloseGate.Decision decision;
                            try {
                                decision = stage2IdentityCloseGate.decide(ui);
                            } catch (RuntimeException ex) {
                                decision =
                                        new Stage2IdentityCloseGate.Decision(
                                                true,
                                                "比較失敗",
                                                "比較中にエラーが発生しました。同一化チェックで内容を確認してください。");
                            }
                            Stage2IdentityCloseGate.Decision decided = decision;
                            Platform.runLater(
                                    () -> {
                                        wait.close();
                                        if (!decided.required()) {
                                            if (showNormalApplicationCloseConfirm()) {
                                                finishApplicationCloseAfterConfirm();
                                            }
                                            return;
                                        }
                                        recordOperatorAction(
                                                "close_warning", "shown", decided.detail());
                                        SevenDigitChallengeDialog.Outcome outcome =
                                                SevenDigitChallengeDialog.showAndConfirm(
                                                        primaryStage,
                                                        SevenDigitChallenge.generate(),
                                                        decided.detail(),
                                                        decided.dialogBody());
                                        if (outcome
                                                == SevenDigitChallengeDialog.Outcome
                                                        .RETURN_TO_CHECK) {
                                            navigateToIdentityCheckAfterCloseGate();
                                            return;
                                        }
                                        if (outcome
                                                        == SevenDigitChallengeDialog.Outcome
                                                                .CONFIRMED
                                                && showNormalApplicationCloseConfirm()) {
                                            finishApplicationCloseAfterConfirm();
                                        }
                                    });
                        },
                        "stage2-identity-close-gate");
        worker.setDaemon(true);
        worker.start();
    }

    private Stage showIdentityCompareWaitDialog() {
        Stage wait = new Stage();
        wait.initModality(Modality.APPLICATION_MODAL);
        if (primaryStage != null) {
            wait.initOwner(primaryStage);
        }
        wait.setTitle("同一化チェック");
        wait.setOnCloseRequest(WindowEvent::consume);
        Label label = new Label("配台計画と加工計画を比較しています…");
        label.setPadding(new Insets(20));
        Scene scene = new Scene(label, 360, 80);
        if (primaryStage != null && primaryStage.getScene() != null) {
            scene.getStylesheets().setAll(primaryStage.getScene().getStylesheets());
        }
        wait.setScene(scene);
        wait.setResizable(false);
        wait.show();
        return wait;
    }

    private void finishApplicationCloseAfterConfirm() {
        applicationCloseProceeding = true;
        performApplicationShutdownOnClose();
        if (primaryStage != null) {
            primaryStage.close();
        }
    }

    private void navigateToIdentityCheckAfterCloseGate() {
        selectMainShellTab(MainShellTabId.DELIVERY_CALENDAR_VIEW);
        if (deliveryCalendarViewTabController != null) {
            deliveryCalendarViewTabController.selectDispatchResultInnerTab(false);
        }
        if (resultDispatchTableTabController != null) {
            resultDispatchTableTabController.promptIdentityCheckAttention();
        }
    }

    /** 配台重要操作を共有フォルダの操作ログへ追記する。失敗時は実行・ログに1行。 */
    public void recordOperatorAction(String action, String result, String detail) {
        Map<String, String> ui = snapshotUiEnv();
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        if (operator.isBlank()) {
            operator = OperatorUserPaths.resolveOperatorUser(ui);
        }
        boolean ok = OperatorActionLogStore.append(ui, operator, action, result, detail);
        if (!ok) {
            appendLog("[operator-action-log] 書き込みに失敗しました");
        }
    }

    public void markStage2CompletedThisLaunch(boolean excelExportSucceeded) {
        stage2IdentityCloseGate.markStage2Completed(excelExportSucceeded);
        recordOperatorAction(
                "stage2_complete",
                "ok",
                excelExportSucceeded ? "段階2完了" : "段階2完了（Excel出力失敗）");
    }

    private void markStage2PipelineAwaitingExcelThisLaunch() {
        if (stage2IdentityCloseGate.stage2CompletedThisLaunch()) {
            return;
        }
        stage2IdentityCloseGate.markStage2PipelineAwaitingExcel();
        recordOperatorAction("stage2_complete", "ok", "段階2完了");
    }

    /** 終了確認後、または内部終了時のクリーンアップ（セッション保存・ロック解放）。 */
    private void performApplicationShutdownOnClose() {
        try {
            ProcessOwnedLockFiles.releaseAllOwnedQuietly();
        } catch (LinkageError ignored) {
            // 増分 compile で target/classes が欠けると NoClassDefFoundError になるため終了自体は続行する
        }
        if (geminiFreeTierModelsRefreshService != null) {
            geminiFreeTierModelsRefreshService.shutdown();
        }
        if (memorySettingsTabController != null) {
            memorySettingsTabController.shutdown();
        }
        JvmMemoryLogStore.persistSnapshot(
                MemoryJvmRingLog.getMaxLines(), MemoryJvmRingLog.snapshotLines());
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        FactorySite current = GlobalInitSettingTarget.load();
        if (!operator.isBlank() && current != null && current != FactorySite.RDP_LAUNCHER) {
            FactorySiteWorkspaceStore.save(operator, current, buildFactorySiteWorkspaceSnapshot());
            FactorySiteWorkspaceStore.saveLastFactorySite(operator, current);
            FactorySiteWorkspaceStore.flushMemoryCacheToDisk(operator);
        }
        LastLaunchedFactorySiteStore.save(current);
        DesktopSessionStateStore.save(collectDesktopSessionForGlobalPersistence());
    }

    private void showThemedAlert(AlertType type, String title, String headerText, String message) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(headerText);
        ThemedAlertContentSupport.applyContent(alert, message);
        prepareDialogForMainTheme(alert);
        alert.showAndWait();
    }

    /**
     * タブ・表・テーマ等を {@code init_setting} 既定へ戻す（環境変数タブは対象外）。
     *
     * <p>適用完了後、{@link #schedulePersistUserSessionAfterGlobalFactoryReset()} でユーザーセッション
     * （{@code session-state.json}）へ保存する。
     */
    public void performGlobalUiFactoryReset() {
        TextInputDialog dialog = new TextInputDialog();
        initDialogOwnerIfSceneReady(dialog);
        dialog.setTitle("確認");
        dialog.setHeaderText(null);
        dialog.setContentText(
                "タブ・表・テーマ等を init_setting の既定に戻します。"
                        + "環境変数タブは変更しません。"
                        + "誤操作防止のため、次のパスワードを入力してください。\nパスワード: 111");
        Optional<String> ans = dialog.showAndWait();
        if (ans.isEmpty() || !"111".equals(ans.get().trim())) {
            return;
        }

        performGlobalUiFactoryResetWithoutConfirmation(
                GlobalInitSettingTarget.loadEffective(collectUiEnv()));

        Alert done = new Alert(AlertType.INFORMATION);
        initDialogOwnerIfSceneReady(done);
        applyAlertStylesheetsFromOwner(done);
        done.setTitle("完了");
        done.setHeaderText(null);
        done.setContentText("UI を既定に戻しました。");
        done.showAndWait();
    }

    /**
     * グローバル設定タブ「デフォルトに戻す」と同一の処理（確認ダイアログ・完了アラートなし）。
     *
     * <p>環境変数タブは {@link #applyFactoryScopedGlobalAndEnvReset} では触らない。
     */
    private void performGlobalUiFactoryResetWithoutConfirmation(FactorySite factorySite) {
        FactorySite site = factorySite != null ? factorySite : FactorySite.KONAN;
        suppressEnvSessionPersistence.set(true);
        try {
            try {
                Files.deleteIfExists(TableColumnOrderPersistence.userHomeStorePath());
            } catch (IOException ignored) {
            }
            DispatchTrialLogUiStore.deleteStoreSilently();
            PlanWorkspaceSnapshotStore.deleteAllSilently();
            WorkspaceCacheArchiveStore.deleteAllSilently();
            PushButtonCssEmitter.deleteUserOverridesFileSilently();
            applyGlobalInitSettingBeforeEnvReset(site);
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
        if (t == mainShellTabEquipmentStatusDashboard) {
            return MainShellTabId.EQUIPMENT_STATUS_DASHBOARD;
        }
        if (t == mainShellTabPipelineExecutionTiming) {
            return MainShellTabId.PIPELINE_EXECUTION_TIMING;
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
        if (t == mainShellTabUserProfiles) {
            return MainShellTabId.USER_PROFILES;
        }
        if (t == mainShellTabOperatorUserManagement) {
            return MainShellTabId.OPERATOR_USER_MANAGEMENT;
        }
        if (t == mainShellTabCompanyCalendar) {
            return MainShellTabId.COMPANY_CALENDAR;
        }
        if (t == mainShellTabMemberAttendance) {
            return MainShellTabId.MEMBER_ATTENDANCE;
        }
        if (t == mainShellTabMachineCalendar) {
            return MainShellTabId.MACHINE_CALENDAR;
        }
        if (t == mainShellTabMasterSummary) {
            return MainShellTabId.MASTER_SUMMARY;
        }
        if (t == mainShellTabPlanInput) {
            return MainShellTabId.PLAN_INPUT;
        }
        if (t == mainShellTabRequestFormInput) {
            return MainShellTabId.REQUEST_FORM_INPUT;
        }
        if (t == mainShellTabRequestFormPipelineCheck) {
            return MainShellTabId.REQUEST_FORM_PIPELINE_CHECK;
        }
        if (t == mainShellTabRemoteDesktop) {
            return MainShellTabId.REMOTE_DESKTOP;
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
        if (t == mainShellTabDailyReportCsvView) {
            return MainShellTabId.DAILY_REPORT_CSV_VIEW;
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
        if (t == mainShellTabOperatorActionLog) {
            return MainShellTabId.OPERATOR_ACTION_LOG;
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
        if (t == mainShellTabRequestFormPreviewBadgeDesign) {
            return MainShellTabId.REQUEST_FORM_PREVIEW_BADGE_DESIGN;
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
            case EQUIPMENT_STATUS_DASHBOARD -> mainShellTabEquipmentStatusDashboard;
            case RUN -> mainShellTabRun;
            case PIPELINE_EXECUTION_TIMING -> mainShellTabPipelineExecutionTiming;
            case UI_BADGE_DESIGN -> mainShellTabUiBadgeDesign;
            case PUSH_BUTTON_DESIGN -> mainShellTabPushButtonDesign;
            case ENV -> mainShellTabEnv;
            case MEMORY_SETTINGS -> mainShellTabMemorySettings;
            case GLOBAL_SETTINGS -> mainShellTabGlobalSettings;
            case USER_PROFILES -> mainShellTabUserProfiles;
            case OPERATOR_USER_MANAGEMENT -> mainShellTabOperatorUserManagement;
            case COMPANY_CALENDAR -> mainShellTabCompanyCalendar;
            case MEMBER_ATTENDANCE -> mainShellTabMemberAttendance;
            case MACHINE_CALENDAR -> mainShellTabMachineCalendar;
            case MASTER_SUMMARY -> mainShellTabMasterSummary;
            case PLAN_INPUT -> mainShellTabPlanInput;
            case REQUEST_FORM_INPUT -> mainShellTabRequestFormInput;
            case REQUEST_FORM_PIPELINE_CHECK -> mainShellTabRequestFormPipelineCheck;
            case REMOTE_DESKTOP -> mainShellTabRemoteDesktop;
            case STAGE1_PREVIEW -> mainShellTabStage1Preview;
            case CODE_LOOKUP_TABLES -> mainShellTabCodeLookupTables;
            case EXCLUDE_RULES -> mainShellTabExcludeRules;
            case SPECIAL_RULES -> mainShellTabSpecialRules;
            case ACTUALS_STATUS -> mainShellTabActualsStatus;
            case DAILY_REPORT_CSV_VIEW -> mainShellTabDailyReportCsvView;
            case DELIVERY_CALENDAR_VIEW -> mainShellTabDeliveryCalendar;
            case RESULT_DISPATCH -> mainShellTabResultDispatch;
            case DISPATCH_INTERACTIVE -> mainShellTabDispatchInteractive;
            case PLAN_WORKSPACE_HISTORY -> mainShellTabPlanWorkspaceHistory;
            case CACHE_HISTORY -> mainShellTabCacheHistory;
            case OPERATOR_ACTION_LOG -> mainShellTabOperatorActionLog;
            case API_MODEL_BENCHMARK -> mainShellTabApiModelBenchmark;
            case PLAN_RESULT_VIEWER -> mainShellTabPlanResultViewer;
            case EQUIPMENT_GANTT_GRAPHIC -> mainShellTabEquipmentGanttGraphic;
            case GANTT_PERSON_BADGE_DESIGN -> mainShellTabGanttPersonBadgeDesign;
            case REQUEST_FORM_PREVIEW_BADGE_DESIGN -> mainShellTabRequestFormPreviewBadgeDesign;
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
        if (now == mainShellTabOperatorUserManagement
                && !operatorUserAdminTabUnlocked
                && !suppressOperatorUserAdminTabGuard.get()) {
            if (!promptOperatorUserAdminTabUnlock()) {
                suppressOperatorUserAdminTabGuard.set(true);
                try {
                    if (prev != null && prev != mainShellTabOperatorUserManagement) {
                        selectShellTabLeaf(prev);
                    } else {
                        selectMainShellTab(MainShellTabId.RUN);
                    }
                } finally {
                    suppressOperatorUserAdminTabGuard.set(false);
                }
                return;
            }
            operatorUserAdminTabUnlocked = true;
        }
        lastEffectiveShellLeaf = now;
        if (prev == mainShellTabRun && now != mainShellTabRun) {
            DesktopSessionStateStore.save(collectDesktopSession());
        }
    }

    private boolean selectShellTabLeaf(Tab leaf) {
        if (leaf == null || tabPane == null) {
            return false;
        }
        return selectShellTabLeafRecursive(tabPane, leaf);
    }

    private boolean selectShellTabLeafRecursive(TabPane pane, Tab target) {
        for (Tab t : pane.getTabs()) {
            if (t == target) {
                pane.getSelectionModel().select(t);
                return true;
            }
        }
        for (Tab t : pane.getTabs()) {
            if (t.getContent() instanceof TabPane inner) {
                if (selectShellTabLeafRecursive(inner, target)) {
                    pane.getSelectionModel().select(t);
                    return true;
                }
            }
        }
        return false;
    }

    private boolean promptOperatorUserAdminTabUnlock() {
        return AdminTabUnlockSupport.ensureUnlocked(primaryStage, this::prepareDialogForMainTheme);
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

    Map<String, String> innerTabHeaderColorByKeySnapshot() {
        return Map.copyOf(innerTabHeaderColorByKey);
    }

    /** 依頼書入力など遅延構築タブの子 TabPane へ、保存済み見出し色を再適用する。 */
    public void refreshInnerTabHeaderColorsForShellTab(MainShellTabId id) {
        if (id == null) {
            return;
        }
        applyInnerTabHeaderColorsForShellTab(id, innerTabHeaderColorByKey);
        refreshMainShellTabHeaderChromeFromStoredColors();
        Platform.runLater(this::refreshMainShellTabHeaderChromeFromStoredColors);
    }

    private Map<String, String> snapshotInnerTabHeaderColorByKey() {
        return Map.copyOf(innerTabHeaderColorByKey);
    }

    private void applyInnerTabHeaderColorsFromSession(Map<String, String> fromSession) {
        innerTabHeaderColorByKey.clear();
        if (fromSession != null) {
            for (Map.Entry<String, String> e : fromSession.entrySet()) {
                if (e.getKey() != null
                        && !e.getKey().isBlank()
                        && e.getValue() != null
                        && !e.getValue().isBlank()) {
                    innerTabHeaderColorByKey.put(e.getKey().trim(), e.getValue().strip());
                }
            }
        }
        applyInnerTabHeaderColorsToLiveUi(innerTabHeaderColorByKey);
    }

    private void syncInnerTabHeaderColorsFromOrganizerTree(
            TreeItem<MainShellTabOrganizerTabController.OrgRow> invisibleRoot) {
        LinkedHashMap<String, String> next = new LinkedHashMap<>();
        collectInnerTabHeaderColorsFromOrganizerTree(invisibleRoot, next);
        innerTabHeaderColorByKey.clear();
        innerTabHeaderColorByKey.putAll(next);
        applyInnerTabHeaderColorsToLiveUi(innerTabHeaderColorByKey);
    }

    private static void collectInnerTabHeaderColorsFromOrganizerTree(
            TreeItem<MainShellTabOrganizerTabController.OrgRow> node,
            Map<String, String> out) {
        if (node == null) {
            return;
        }
        MainShellTabOrganizerTabController.OrgRow r = node.getValue();
        if (r != null) {
            if (r.kind == MainShellTabOrganizerTabController.OrgRow.Kind.INNER_TAB) {
                String key =
                        jp.co.pm.ai.desktop.config.MainShellInnerTabColorKeys.innerKey(
                                r.tabId, r.groupTitle);
                putOrRemoveInnerTabColor(out, key, r.colorHex);
            } else if (r.kind == MainShellTabOrganizerTabController.OrgRow.Kind.INNER_NESTED_TAB) {
                String key =
                        jp.co.pm.ai.desktop.config.MainShellInnerTabColorKeys.nestedKey(
                                r.tabId, r.anchorInnerTabLabel, r.groupTitle);
                putOrRemoveInnerTabColor(out, key, r.colorHex);
            }
        }
        for (TreeItem<MainShellTabOrganizerTabController.OrgRow> c : node.getChildren()) {
            collectInnerTabHeaderColorsFromOrganizerTree(c, out);
        }
    }

    private static void putOrRemoveInnerTabColor(
            Map<String, String> out, String key, String colorHex) {
        if (key == null || key.isBlank()) {
            return;
        }
        String h = colorHex != null ? colorHex.strip() : "";
        if (h.isBlank()) {
            out.remove(key);
        } else {
            out.put(key, h);
        }
    }

    private void applyInnerTabHeaderColorsToLiveUi(Map<String, String> colors) {
        Map<String, String> safe = colors != null ? colors : Map.of();
        for (MainShellTabId id : MainShellTabId.values()) {
            if (id == MainShellTabId.TAB_ORGANIZER) {
                continue;
            }
            if (jp.co.pm.ai.desktop.config.MainShellInnerTabCatalog.labelsFor(id).isEmpty()) {
                continue;
            }
            applyInnerTabHeaderColorsForShellTab(id, safe);
        }
    }

    private void applyInnerTabHeaderColorsForShellTab(
            MainShellTabId shellTabId, Map<String, String> colors) {
        Tab mainTab = mainShellTabFor(shellTabId);
        if (mainTab == null || isLazyMainShellTabPlaceholder(mainTab.getContent())) {
            return;
        }
        applyInnerTabHeaderColorsUnderNode(mainTab.getContent(), shellTabId, colors);
    }

    private void applyInnerTabHeaderColorsUnderNode(
            javafx.scene.Node node, MainShellTabId shellTabId, Map<String, String> colors) {
        javafx.scene.control.TabPane outer = findFirstTabPane(node);
        if (outer == null) {
            return;
        }
        List<String> labels =
                jp.co.pm.ai.desktop.config.MainShellInnerTabCatalog.labelsFor(shellTabId);
        for (javafx.scene.control.Tab tab : outer.getTabs()) {
            String text = nz(tab.getText());
            String innerKey =
                    jp.co.pm.ai.desktop.config.MainShellInnerTabColorKeys.innerKey(shellTabId, text);
            applyShellTabColor(tab, colors.getOrDefault(innerKey, ""));
            int idx = labels.indexOf(text);
            if (idx >= 0) {
                List<String> nestedLabels =
                        jp.co.pm.ai.desktop.config.MainShellInnerTabCatalog
                                .nestedInnerTabLabelsUnderInnerTab(shellTabId, idx);
                if (!nestedLabels.isEmpty() && tab.getContent() != null) {
                    javafx.scene.control.TabPane inner = findFirstTabPane(tab.getContent());
                    if (inner != null) {
                        for (javafx.scene.control.Tab nestedTab : inner.getTabs()) {
                            String nestedText = nz(nestedTab.getText());
                            String nestedKey =
                                    jp.co.pm.ai.desktop.config.MainShellInnerTabColorKeys.nestedKey(
                                            shellTabId, text, nestedText);
                            applyShellTabColor(nestedTab, colors.getOrDefault(nestedKey, ""));
                        }
                    }
                }
            }
        }
    }

    private static javafx.scene.control.TabPane findFirstTabPane(javafx.scene.Node node) {
        if (node instanceof javafx.scene.control.TabPane tp) {
            return tp;
        }
        if (node instanceof javafx.scene.Parent parent) {
            for (javafx.scene.Node child : parent.getChildrenUnmodifiable()) {
                javafx.scene.control.TabPane found = findFirstTabPane(child);
                if (found != null) {
                    return found;
                }
            }
        }
        return null;
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
        return snapshotLiveMainShellTabLayout();
    }

    private List<MainShellTabLayoutNode> snapshotLiveMainShellTabLayout() {
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
            applyInnerTabHeaderColorsFromSession(s.innerTabHeaderColorByKey());
            lastEffectiveShellLeaf =
                    resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        } finally {
            suppressLazyMainShellTabContentSwap.set(false);
            suppressMainShellTabChromeRefresh.set(false);
            installLazyMainShellTabContentForStartup();
            restoreActiveMainShellTabHeavyContentAfterLazyInstall();
            applyRunTabGating();
        }
    }

    /**
     * {@link #installLazyMainShellTabContentForStartup()} 直後に、現在選択中タブの実コンテンツを載せ直す。
     * 起動時 {@link Stage#setOnShown} 以外（工場切替・セッション再適用等）でも空白表示にならないようにする。
     */
    private void restoreActiveMainShellTabHeavyContentAfterLazyInstall() {
        if (tabPane == null) {
            return;
        }
        if (isEnvVarsInitializationPending()) {
            ensureMainShellEnvTabSelected();
        }
        Tab selected = tabPane.getSelectionModel().getSelectedItem();
        if (selected == null && !tabPane.getTabs().isEmpty()) {
            tabPane.getSelectionModel().selectFirst();
            selected = tabPane.getSelectionModel().getSelectedItem();
        }
        if (selected == null) {
            return;
        }
        boolean prev = suppressLazyMainShellTabContentSwap.get();
        suppressLazyMainShellTabContentSwap.set(true);
        try {
            activateMainShellTabHeavyContentRecursive(selected);
        } finally {
            suppressLazyMainShellTabContentSwap.set(false);
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

    private void scheduleEquipmentStatusDashboardInitialReloadIfSelected() {
        if (tabPane == null
                || mainShellTabEquipmentStatusDashboard == null
                || equipmentStatusDashboardTabController == null) {
            return;
        }
        if (tabPane.getSelectionModel().getSelectedItem() == mainShellTabEquipmentStatusDashboard) {
            equipmentStatusDashboardTabController.onMainShellTabSelected();
        }
    }

    private void scheduleRequestFormPipelineCheckInitialRefreshIfSelected() {
        if (tabPane == null
                || mainShellTabRequestFormPipelineCheck == null
                || requestFormPipelineCheckTabController == null) {
            return;
        }
        if (tabPane.getSelectionModel().getSelectedItem() == mainShellTabRequestFormPipelineCheck) {
            requestFormPipelineCheckTabController.onMainShellTabSelected();
        }
    }

    private void activateMainShellTabHeavyContentRecursive(Tab tab) {
        if (tab == null) {
            return;
        }
        restoreDeferredTabContent(tab);
        MainShellTabId shellId = mainShellTabId(tab);
        if (shellId != null) {
            applyInnerTabHeaderColorsForShellTab(shellId, innerTabHeaderColorByKey);
        }
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
        syncInnerTabHeaderColorsFromOrganizerTree(invisibleRoot);
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
        MainShellTabId selectedLeafBefore =
                mainShellTabId(
                        resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem()));
        completeMainShellTabLayout = prepared;
        suppressEnvSessionPersistence.set(true);
        suppressMainShellTabChromeRefresh.set(true);
        try {
            wiredInnerMainShellTabPanes.clear();
            tabPane.getTabs().clear();
            List<MainShellTabLayoutNode> displayed = prepared;
            for (MainShellTabLayoutNode n : displayed) {
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
                                    if (blockMainShellTabSelectionIfEnvInitPending()) {
                                        return;
                                    }
                                    if (blockMemberAttendanceUnsavedInnerTabNavigation(p, n, inner)) {
                                        return;
                                    }
                                    if (blockCompanyCalendarUnsavedInnerTabNavigation(p, n, inner)) {
                                        return;
                                    }
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
        if (isEnvVarsInitializationPending()) {
            ensureMainShellEnvTabSelected();
        } else if (isPipelineRunLocked()
                || activeRunStageScript != null) {
            ensureMainShellRunTabSelected();
        } else if (selectedLeafBefore != null) {
            selectMainShellTabRecursive(tabPane, selectedLeafBefore);
            if (!suppressLazyMainShellTabContentSwap.get()) {
                activateMainShellTabHeavyContentRecursive(
                        tabPane.getSelectionModel().getSelectedItem());
            }
        }
        lastEffectiveShellLeaf =
                resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        Platform.runLater(this::refreshMainShellTabHeaderChromeFromStoredColors);
        applyRunTabGating();
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
        List<MainShellTabLayoutNode> flat = new ArrayList<>();
        for (String key : MainShellTabLayoutDefaults.completeFlatTabKeyOrder()) {
            flat.add(MainShellTabLayoutNode.tabNode(key, ""));
        }
        rebuildMainShellTabsFromLayout(flat);
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
        applyInnerTabHeaderColorsToLiveUi(innerTabHeaderColorByKey);
        Platform.runLater(this::refreshMainShellTabHeaderChromeFromStoredColors);
        DesktopSessionStateStore.save(collectDesktopSession());
        return true;
    }

    /** 現在のメインシェル構成をツリー編集用にエクスポート。 */
    List<MainShellTabLayoutNode> snapshotMainShellTabLayoutNodes() {
        return snapshotLiveMainShellTabLayout();
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

    /** 環境変数初期化フィンガープリント照合に含めるキー（RDP・パイプライン実行時同期は除外）。 */
    private static boolean includeInEnvInitFingerprint(String name) {
        return !omitEnvRowKey(name)
                && !RemoteDesktopEnvRows.excludedFromMainShellEnvInitFingerprint(name)
                && !AppPaths.isPipelineRuntimeSyncedEnvKey(name);
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
        if (s == null || s.uiEnvRows() == null || s.uiEnvRows().isEmpty()) {
            return;
        }
        applyUiEnvRowSnapshots(s.uiEnvRows());
    }

    private void applyUiEnvRowSnapshots(List<UiEnvRowSnapshot> snapshots) {
        if (snapshots == null || snapshots.isEmpty() || envRows == null) {
            return;
        }
        List<EnvVarRow> restored = new ArrayList<>(snapshots.size());
        for (UiEnvRowSnapshot snap : snapshots) {
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
        migrateLegacyMasterWorkbookFileEnvRows();
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
    public void addMissingReferenceEnvRows() {
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
        schedulePersistSessionDebounced();
    }

    /** 実行・ログの当日配台排他2択（段階1ボタン横）。 */
    boolean snapshotTodayDispatch() {
        return mainRunTabController != null && mainRunTabController.snapshotTodayDispatch();
    }

    /** 「当日は配台しない」または当日配台モード中の skip 上書き。未初期化時は true（しない側）。 */
    boolean snapshotStage2SkipTodayDispatch() {
        return mainRunTabController == null || mainRunTabController.snapshotStage2SkipTodayDispatch();
    }

    /** 当日配台モード変更時に配台計画_タスク入力の翌日配台ダイアログ連動を更新する。 */
    void refreshPlanInputNextDayDialogCoupling() {
        if (planInputTabController != null) {
            planInputTabController.refreshNextDayDialogRadioCoupling();
        }
    }

    private void schedulePersistSessionDebounced() {
        if (!suppressEnvSessionPersistence.get()) {
            sessionPersistDebounce.playFromStart();
        }
    }

    /** session-state.json 用: 工場スコープ項目を omit し端末共通 UI のみ保存（§10e）。 */
    private DesktopSessionState collectDesktopSessionForGlobalPersistence() {
        return collectDesktopSession()
                .mergeFactoryScopedFrom(DesktopSessionState.empty().extractFactoryScopedFields());
    }

    private FactorySiteWorkspaceSnapshot buildFactorySiteWorkspaceSnapshot() {
        return new FactorySiteWorkspaceSnapshot(
                snapshotUiEnvRows(), collectDesktopSession().extractFactoryScopedFields());
    }

    /**
     * 工場切替・起動復元: 保存済み env 行（または ui_ref 既定）を先に載せ替え、続けて {@code init_setting} と session 断片を適用する。
     */
    private void applyFactorySiteWorkspaceRestore(
            FactorySite site, Optional<FactorySiteWorkspaceSnapshot> workspace) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        if (workspace.isPresent() && workspace.get().hasUiEnvRows()) {
            suppressEnvSessionPersistence.set(true);
            envResetInProgress.set(true);
            try {
                applyFactoryWorkspaceEnvSnapshot(workspace.get());
                applyFactorySitePortableAndNetworkDefaults(effective);
                applyRepoFolderPathNormalization();
            } finally {
                envResetInProgress.set(false);
                suppressEnvSessionPersistence.set(false);
                uiEnvSaveDebounce.stop();
            }
        } else {
            applyEnvRowsFullBundledResetAndPersist(false, effective);
        }
        applyGlobalInitSettingBeforeEnvReset(effective);
        workspace.ifPresent(this::applyFactoryWorkspaceSessionFragment);
    }

    /** 工場ワークスペースの環境変数行のみ復元（session は別途）。 */
    private void applyFactoryWorkspaceEnvSnapshot(FactorySiteWorkspaceSnapshot snapshot) {
        if (snapshot == null || !snapshot.hasUiEnvRows()) {
            return;
        }
        applyUiEnvRowSnapshots(snapshot.uiEnvRows());
        mergeMissingBootstrapEnvRows();
        stripRemovedEnvVarRows(envRows);
        ensureBootstrapDefaultValuesVisible(collectUiEnv());
        ensureUiRefOptionalDisplayDefaultsVisible(collectUiEnv());
    }

    /**
     * 環境変数初期化完了時点のタブ値を工場ワークスペースと session-state に保存する。
     * {@link #recordEnvInitializationBaseline()} の直前に呼び、フィンガープリントと復元元を揃える。
     */
    private void persistOperatorWorkspaceForEnvInitBaseline(FactorySite site) {
        stabilizeEnvRowsForInitializationBaseline();
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        if (!operator.isBlank() && site != null && site != FactorySite.RDP_LAUNCHER) {
            FactorySiteWorkspaceStore.save(operator, site, buildFactorySiteWorkspaceSnapshot());
        }
        DesktopSessionStateStore.save(collectDesktopSessionForGlobalPersistence());
    }

    /** 工場ワークスペースの session 断片のみ適用（環境変数行は {@link #applyEnvRowsFullBundledResetAndPersist} 後に使う）。 */
    private void applyFactoryWorkspaceSessionFragment(FactorySiteWorkspaceSnapshot snapshot) {
        applyFactoryWorkspaceSessionFragment(snapshot, false);
    }

    /**
     * @param afterEnvInitialization {@code true} のとき環境変数初期化直後。ワークスペース内の env 系パス・{@code uiEnvRows} は載せない。
     */
    private void applyFactoryWorkspaceSessionFragment(
            FactorySiteWorkspaceSnapshot snapshot, boolean afterEnvInitialization) {
        if (snapshot == null) {
            return;
        }
        DesktopSessionState fragment = snapshot.sessionFragment();
        if (afterEnvInitialization) {
            fragment = fragment.withoutEnvInitializationFields();
        }
        DesktopSessionState merged =
                afterEnvInitialization
                        ? collectDesktopSession()
                                .mergeFactoryScopedFromPreservingEnvInitialization(fragment)
                        : collectDesktopSession().mergeFactoryScopedFrom(fragment);
        applyDesktopSession(merged, false, false, afterEnvInitialization);
    }

    private void notifyActiveMainShellTabAfterWorkspaceChange() {
        if (tabPane == null) {
            return;
        }
        restoreActiveMainShellTabHeavyContentAfterLazyInstall();
        refreshMainShellTabHeaderChromeFromStoredColors();
        refreshDesktopSessionDependentUi();
        Tab effective = resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        if (effective == mainShellTabEquipmentStatusDashboard
                && equipmentStatusDashboardTabController != null) {
            equipmentStatusDashboardTabController.onMainShellTabSelected();
        }
        if (effective == mainShellTabRequestFormInput && requestFormInputTabController != null) {
            requestFormInputTabController.onMainShellTabSelected();
        }
        if (effective == mainShellTabRequestFormPipelineCheck
                && requestFormPipelineCheckTabController != null) {
            requestFormPipelineCheckTabController.onMainShellTabSelected();
        }
        if (effective == mainShellTabRemoteDesktop && remoteDesktopTabController != null) {
            remoteDesktopTabController.onMainShellTabSelected();
        }
        if (effective == mainShellTabEquipmentGanttGraphic
                && equipmentGanttGraphicTabController != null) {
            equipmentGanttGraphicTabController.flushPendingGraphicRebuildAfterSessionApply();
        }
        if (effective == mainShellTabDispatchInteractive && dispatchInteractiveTabController != null) {
            dispatchInteractiveTabController.onMainShellDispatchTabSelected();
        }
    }

    private void refreshFactoryDependentTabs(FactorySite site, boolean lightweight) {
        if (requestFormInputTabController != null) {
            requestFormInputTabController.onFactorySiteChanged(lightweight);
        }
        if (requestFormPipelineCheckTabController != null) {
            requestFormPipelineCheckTabController.onFactorySiteChanged(lightweight);
        }
        if (remoteDesktopTabController != null) {
            remoteDesktopTabController.onFactorySiteChanged(lightweight);
        }
        if (mainRunTabController != null) {
            mainRunTabController.refreshOpenWorkbookHintLabels();
            factoryOperatorToolbar.refreshFactorySiteLogo();
        }
        if (globalSettingsTabController != null) {
            globalSettingsTabController.refreshInitSettingTargetComboFromStore();
        }
        refreshFactorySiteComboPresentation();
    }

    void refreshFactorySiteComboPresentation() {
        if (factoryOperatorToolbar != null) {
            factoryOperatorToolbar.refreshFactorySiteComboPresentation();
        }
        if (globalSettingsTabController != null) {
            globalSettingsTabController.refreshInitSettingTargetComboPresentation();
        }
    }

    void refreshShellFactoryOperatorToolbar() {
        if (factoryOperatorToolbar != null) {
            factoryOperatorToolbar.refreshFactorySiteLogo();
        }
    }

    void refreshShellFactorySiteComboPresentation() {
        refreshFactorySiteComboPresentation();
    }

    void setShellFactorySiteComboDisabled(boolean disabled) {
        setFactorySiteCombosDisabled(disabled);
    }

    void refreshShellFactorySiteComboFromStore() {
        if (factoryOperatorToolbar != null) {
            factoryOperatorToolbar.refreshFactorySiteComboFromStore();
        }
    }

    void refreshShellOperatorUserLabel() {
        if (factoryOperatorToolbar != null) {
            factoryOperatorToolbar.refreshOperatorUserLabel();
        }
    }

    void setGuestSessionFactoryToolbar(boolean guestOnly) {
        if (factoryOperatorToolbar != null) {
            factoryOperatorToolbar.setGuestSessionFactorySwitchOnly(guestOnly);
        }
    }

    private void setFactorySiteCombosDisabled(boolean disabled) {
        if (factoryOperatorToolbar != null) {
            factoryOperatorToolbar.setFactorySiteComboDisabled(disabled);
        }
        if (globalSettingsTabController != null) {
            globalSettingsTabController.setInitSettingTargetComboDisabled(disabled);
        }
    }

    /**
     * スプラッシュで採用した工場を本番に載せる。操作者の前回工場や UNC 推定では上書きしない。
     *
     * @return 起動時の工場切替を開始したとき {@code true}（呼び出し元は後続の環境照合を工場切替完了後に続行する）
     */
    private boolean finalizeOperatorLocalWorkspaceAfterSessionEstablished() {
        FactorySite adopted = StartupFactorySiteResolver.resolve();
        LastLaunchedFactorySiteStore.save(adopted);
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        if (operator.isBlank() || FactoryOperatorUserStore.isGuestOperator(operator)) {
            if (StartupFactorySiteResolver.requiresStartupSwitch(
                    GlobalInitSettingTarget.load(), adopted)) {
                startupRestoredFactorySite = true;
                switchActiveFactorySite(adopted, true);
                return true;
            }
            GlobalInitSettingTarget.save(adopted);
            return false;
        }
        FactorySiteWorkspaceMigrator.migrateIfNeeded(
                operator,
                adopted,
                snapshotUiEnvRows(),
                DesktopSessionStateStore.load(),
                collectUiEnv());
        FactorySiteWorkspaceStore.warmMemoryCacheFromDisk(operator);
        if (StartupFactorySiteResolver.requiresStartupSwitch(
                GlobalInitSettingTarget.load(), adopted)) {
            startupRestoredFactorySite = true;
            switchActiveFactorySite(adopted, true);
            return true;
        }
        Optional<FactorySiteWorkspaceSnapshot> ws =
                FactorySiteWorkspaceStore.load(operator, adopted);
        if (ws.isPresent()) {
            applyFactorySiteWorkspaceRestore(adopted, ws);
        }
        return false;
    }

    private void applyPortableUpgradeShellUiSnapshotIfPresent() {
        DesktopSessionState snap = PortableBundleUpgradeUiSnapshot.loadIfPresent();
        if (snap == null) {
            return;
        }
        DesktopSessionState merged = collectDesktopSession().mergeShellTabUiFrom(snap);
        applyDesktopSession(merged, false, false);
        refreshDesktopSessionDependentUi();
        PortableBundleUpgradeUiSnapshot.clear();
        appendLog("[startup] バージョンアップ: メインシェルタブ配置を前回状態から復元しました。");
    }

    private void installUiEnvAutoSave() {
        sessionPersistDebounce.setOnFinished(
                e -> {
                    if (!suppressEnvSessionPersistence.get()) {
                        DesktopSessionStateStore.save(collectDesktopSessionForGlobalPersistence());
                    }
                });
        uiEnvSaveDebounce.setOnFinished(
                e -> {
                    if (!suppressEnvSessionPersistence.get()) {
                        DesktopSessionStateStore.save(collectDesktopSessionForGlobalPersistence());
                    }
                    if (mainRunTabController != null) {
                        mainRunTabController.refreshOpenWorkbookHintLabels();
                    }
                    pipelineExecutionTimingHistory.configureFromUi(collectUiEnv());
                    FactoryOperatorUserStore.configureFromUi(
                            collectUiEnv(), GlobalInitSettingTarget.load());
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

    public Stage getPrimaryStage() {
        return primaryStage;
    }

    public ObservableList<EnvVarRow> getEnvRows() {
        return envRows;
    }

    /**
     * Resets the env-var table to bundled defaults ({@link UiRefEnvDefaults}) and reapplies bootstrap fills.
     * Shows a confirmation dialog first.
     *
     * <p>順序: 現在選択工場のグローバル設定（{@code init_setting} 全体）を適用 → 環境変数を初期化 →
     * 依頼書原本フォルダ案内 → 実行・ログタブへ遷移。
     */
    public void confirmAndResetEnvRowsToDefaults() {
        FactorySite site = GlobalInitSettingTarget.load();
        if (site == null || site == FactorySite.RDP_LAUNCHER) {
            site = FactorySite.KONAN;
        }
        Alert alert = new Alert(AlertType.CONFIRMATION);
        initDialogOwnerIfSceneReady(alert);
        applyAlertStylesheetsFromOwner(alert);
        alert.setTitle("環境変数を初期値に戻す");
        alert.setHeaderText(null);
        alert.setContentText(
                site.displayLabelJa()
                        + "（現在の利用工場）のグローバル設定（init_setting）を適用したうえで、"
                        + "環境変数を ui_ref 既定に戻します。"
                        + "未保存の編集と、セッションに保存していた各タブの値（Python パス等）も失われます。"
                        + "続行しますか？");
        Optional<ButtonType> ans = alert.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        GlobalInitSettingTarget.save(site);
        applyFactoryScopedGlobalAndEnvReset(site, true);
        maybePromptRequestFormOriginalDirIfUnset("[env]", site);
        persistOperatorWorkspaceForEnvInitBaseline(site);
        recordEnvInitializationBaseline();
        // 4) 実行・ログタブへ遷移（ダイアログ後・タブ再構築後の自動遷移を打ち消す）
        ensureMainShellRunTabSelected();
        requireOperatorSelectionForFactory(site, false);
        Platform.runLater(this::ensureMainShellRunTabSelected);
    }

    /**
     * 環境変数初期化の直前に、指定工場のグローバル設定（{@code init_setting} 全体: タブ構成・テーマ・列順・依頼書設定など）を適用する。
     * {@link #confirmAndResetEnvRowsToDefaults()} の手順 1 と同一。環境変数行は {@code applyUiEnvRowsFromSession=false} で載せ替えない。
     */
    /**
     * 工場スコープのグローバル設定（{@code init_setting}）適用のあと、環境変数タブを ui_ref 既定へ戻す。
     * バージョンアップ後・環境タブ「環境変数を初期化」・グローバル設定「デフォルトに戻す」で共通。
     */
    private void applyFactoryScopedGlobalAndEnvReset(FactorySite site, boolean persistSession) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        applyGlobalInitSettingBeforeEnvReset(effective);
        applyEnvRowsFullBundledResetAndPersist(persistSession, effective);
    }

    private void applyGlobalInitSettingBeforeEnvReset(FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        suppressEnvSessionPersistence.set(true);
        try {
            DesktopSessionState merged =
                    DesktopSessionStateStore.buildFactoryResetSessionFromInitSettingOnly(
                            collectUiEnv(), effective);
            applyDesktopSession(merged, false, false, true);
            applyFactoryRequestFormGlobalSettings(effective, true);
            TableColumnOrderPersistence.materializeTableColumnStoreAfterFactoryReset(collectUiEnv());
            applyDesktopThemeFromSession(merged);
            refreshDesktopSessionDependentUi();
            if (globalSettingsTabController != null) {
                globalSettingsTabController.refreshInitSettingTargetComboFromStore();
            }
            if (mainRunTabController != null) {
                factoryOperatorToolbar.refreshFactorySiteComboFromStore();
                factoryOperatorToolbar.refreshFactorySiteLogo();
            }
        } finally {
            suppressEnvSessionPersistence.set(false);
        }
        logGlobalInitSettingLoaded(effective);
    }

    private void logGlobalInitSettingLoaded(FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        appendLog(
                "[global-settings] グローバル設定を読み込みました（init_setting）。工場: "
                        + effective.displayLabelJa());
    }

    private void logEnvVarsBundledReset(FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        appendLog(
                "[env] 環境変数を ui_ref 既定に初期化しました。工場: "
                        + effective.displayLabelJa());
    }

    /**
     * 環境変数初期化後: 配台除外 JSON・タスク入力ブック等を環境タブの値へ揃える（{@code session_defaults} の
     * {@code excludeRulesPath} / {@code mainRunWorkbook} と環境タブの二重管理を解消）。
     */
    private void syncDesktopSessionPathFieldsFromEnvTab() {
        Map<String, String> ui = collectUiEnv();
        String exclude = envTabValueTrimmed(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON);
        if (!exclude.isEmpty() && excludeRulesTabController != null) {
            excludeRulesTabController.restoreDesktopSessionPath(exclude);
        }
        if (mainRunTabController != null) {
            AppPaths.resolveTaskInputWorkbook(ui)
                    .map(Path::toString)
                    .filter(s -> !s.isBlank())
                    .ifPresent(mainRunTabController.getWorkbookField()::setText);
        }
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
        FactorySite pref = GlobalInitSettingTarget.loadEffective(collectUiEnv());
        ChoiceDialog<FactorySite> d = new ChoiceDialog<>(pref, List.of(FactorySite.values()));
        initDialogOwnerIfSceneReady(d);
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
        syncDesktopSessionPathFieldsFromEnvTab();
        if (persistSession) {
            DesktopSessionStateStore.save(collectDesktopSessionForGlobalPersistence());
        }
        refreshPersonBadgeSkillsMembersFromMaster();
        mainRunTabController.refreshOpenWorkbookHintLabels();
        uiEnvSaveDebounce.stop();
        EnvVarsInitializedAtStore.recordNow();
        envVarsDifferFromInitialAtStartup.set(false);
        envInitTabBlockLogEmitted.set(false);
        envVarsStartupCheckCompleted.set(true);
        refreshEnvVarsInitializedAtToolbarLabel();
        applyRunTabGating();
        logEnvVarsBundledReset(factorySite);
        maybeReloadAttendanceTabsAfterEnvReady();
    }

    private void refreshEnvVarsInitializedAtToolbarLabel() {
        if (envVarsInitializedAtLabel == null) {
            return;
        }
        envVarsInitializedAtLabel.setText(
                "環境変数初期化: " + EnvVarsInitializedAtStore.formatForToolbar());
        refreshGlobalStatusBar();
    }

    private boolean isEnvVarsInitializationPending() {
        if (FactoryOperatorUserStore.isGuestSession()) {
            return false;
        }
        if (deferOperatorPromptForPortableUpgrade.get()) {
            return false;
        }
        if (!EnvVarsInitializedAtStore.isRecorded()) {
            return true;
        }
        if (!envVarsStartupCheckCompleted.get()) {
            return true;
        }
        return envVarsDifferFromInitialAtStartup.get();
    }

    /**
     * 操作者選択・工場ワークスペース復元の後に、環境変数タブの値を初期化テンプレートと一度だけ照合する。
     *
     * <p>進捗モーダル表示中は状況文言だけ更新する。未表示なら短いモーダルを出して閉じる。
     */
    private void completeEnvVarsStartupCheck() {
        completeEnvVarsStartupCheck(true);
    }

    private void completeEnvVarsStartupCheck(boolean schedulePostStartupWork) {
        if (envVarsStartupCheckCompleted.get()) {
            return;
        }
        boolean ownBusy = !isEnvVarsStartupCheckBusyShowing() && !isFactorySiteSwitchBusyShowing();
        if (ownBusy && !startupSequenceActive) {
            beginEnvVarsStartupCheckBusy(EnvVarsStartupCheckBusyDialog.STATUS_STABILIZE);
        } else if (isFactorySiteSwitchBusyShowing()) {
            updateFactorySiteSwitchBusy(EnvVarsStartupCheckBusyDialog.STATUS_STABILIZE);
        } else if (isEnvVarsStartupCheckBusyShowing()) {
            updateEnvVarsStartupCheckBusy(EnvVarsStartupCheckBusyDialog.STATUS_STABILIZE);
        }
        try {
            stabilizeEnvRowsForInitializationBaseline();
            if (isFactorySiteSwitchBusyShowing()) {
                updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_MATCH);
            } else {
                updateEnvVarsStartupCheckBusy(EnvVarsStartupCheckBusyDialog.STATUS_MATCH);
            }
            evaluateEnvVarsDifferFromInitialAtStartup();
            applyRunTabGating();
            if (schedulePostStartupWork && !startupSequenceActive) {
                if (!isEnvVarsInitializationPending()) {
                    ensureMainShellRunTabSelected();
                }
                maybeReloadAttendanceTabsAfterEnvReady();
            }
            if (ownBusy && !startupSequenceActive) {
                updateEnvVarsStartupCheckBusy(EnvVarsStartupCheckBusyDialog.STATUS_DONE);
            } else if (isFactorySiteSwitchBusyShowing()) {
                updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_DONE);
            }
        } finally {
            if (ownBusy && !startupSequenceActive) {
                endEnvVarsStartupCheckBusy();
            }
        }
    }

    private boolean isEnvVarsStartupCheckBusyShowing() {
        return envVarsStartupCheckBusy != null && envVarsStartupCheckBusy.isShowing();
    }

    private void beginEnvVarsStartupCheckBusy(String status) {
        if (isEnvVarsStartupCheckBusyShowing()) {
            updateEnvVarsStartupCheckBusy(status);
            return;
        }
        if (primaryStage == null) {
            return;
        }
        envVarsStartupCheckBusy = EnvVarsStartupCheckBusyDialog.show(primaryStage, status);
        envVarsStartupCheckBusy.setStep(resolveStartupCheckDialogStep(status));
    }

    private void updateEnvVarsStartupCheckBusy(String status) {
        if (envVarsStartupCheckBusy == null) {
            return;
        }
        Runnable update =
                () -> {
                    envVarsStartupCheckBusy.setHeader(resolveStartupCheckDialogHeader(status));
                    envVarsStartupCheckBusy.setStep(resolveStartupCheckDialogStep(status));
                    envVarsStartupCheckBusy.setStatus(status);
                };
        if (Platform.isFxApplicationThread()) {
            update.run();
        } else {
            Platform.runLater(update);
        }
    }

    private static String resolveStartupCheckDialogHeader(String status) {
        if (status == null || status.isBlank()) {
            return EnvVarsStartupCheckBusyDialog.HEADER;
        }
        if (EnvVarsStartupCheckBusyDialog.STATUS_BACKGROUND_LOAD.equals(status)
                || status.startsWith("起動後読込")) {
            return EnvVarsStartupCheckBusyDialog.HEADER_BACKGROUND_LOAD;
        }
        return EnvVarsStartupCheckBusyDialog.HEADER;
    }

    private static String resolveStartupCheckDialogStep(String status) {
        if (status == null || status.isBlank()) {
            return "";
        }
        if (EnvVarsStartupCheckBusyDialog.STATUS_RESTORE_WORKSPACE.equals(status)
                || EnvVarsStartupCheckBusyDialog.STATUS_FACTORY_SWITCH.equals(status)) {
            return EnvVarsStartupCheckBusyDialog.STEP_RESTORE_WORKSPACE;
        }
        if (EnvVarsStartupCheckBusyDialog.STATUS_STABILIZE.equals(status)
                || EnvVarsStartupCheckBusyDialog.STATUS_MATCH.equals(status)
                || EnvVarsStartupCheckBusyDialog.STATUS_DONE.equals(status)) {
            return EnvVarsStartupCheckBusyDialog.STEP_ENV_MATCH;
        }
        if (EnvVarsStartupCheckBusyDialog.STATUS_BACKGROUND_LOAD.equals(status)
                || status.startsWith("起動後読込")) {
            return EnvVarsStartupCheckBusyDialog.STEP_TAB_LOAD;
        }
        return "";
    }

    private void endEnvVarsStartupCheckBusy() {
        if (envVarsStartupCheckBusy != null) {
            envVarsStartupCheckBusy.close();
            envVarsStartupCheckBusy = null;
        }
    }

    private boolean isFactorySiteSwitchBusyShowing() {
        return factorySiteSwitchBusy != null && factorySiteSwitchBusy.isShowing();
    }

    private void beginFactorySiteSwitchBusy(FactorySite from, FactorySite to) {
        factorySwitchBusyFrom = from;
        factorySwitchBusyTo = to;
        if (startupSequenceActive && isEnvVarsStartupCheckBusyShowing()) {
            updateEnvVarsStartupCheckBusy(EnvVarsStartupCheckBusyDialog.STATUS_FACTORY_SWITCH);
            return;
        }
        if (isFactorySiteSwitchBusyShowing()) {
            updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_SAVING);
            return;
        }
        showFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_SAVING);
    }

    private void beginFactorySiteSwitchTabLoadBusy() {
        if (startupSequenceActive && isEnvVarsStartupCheckBusyShowing()) {
            updateEnvVarsStartupCheckBusy(FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD);
            return;
        }
        if (isFactorySiteSwitchBusyShowing()) {
            updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD);
            return;
        }
        showFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD);
    }

    private void showFactorySiteSwitchBusy(String initialStatus) {
        if (primaryStage == null) {
            return;
        }
        FactorySite from = factorySwitchBusyFrom;
        FactorySite to = factorySwitchBusyTo;
        String header =
                (from != null ? from.displayLabelJa() : "—")
                        + " → "
                        + (to != null ? to.displayLabelJa() : "—");
        factorySiteSwitchBusy = FactorySiteSwitchBusyDialog.show(primaryStage, header, initialStatus);
        Scene scene = factorySiteSwitchBusy.scene();
        if (scene != null) {
            registerThemeTrackedScene(scene);
        }
    }

    private void updateFactorySiteSwitchBusy(String status) {
        if (startupSequenceActive && isEnvVarsStartupCheckBusyShowing()) {
            updateEnvVarsStartupCheckBusy(status);
            return;
        }
        if (factorySiteSwitchBusy != null) {
            factorySiteSwitchBusy.setStatus(status);
        }
    }

    private void endFactorySiteSwitchBusy() {
        if (startupSequenceActive) {
            return;
        }
        factorySwitchAwaitingBackgroundLoadBeforeModalClose = false;
        if (factorySiteSwitchBusy != null) {
            Scene scene = factorySiteSwitchBusy.scene();
            if (scene != null) {
                unregisterThemeTrackedScene(scene);
            }
            factorySiteSwitchBusy.close();
            factorySiteSwitchBusy = null;
        }
    }

    private static boolean usesStageRunBusyModal(String script) {
        return STAGE1.equals(script) || STAGE2.equals(script);
    }

    private void beginStageRunBusyDialog(String script) {
        if (!usesStageRunBusyModal(script)) {
            return;
        }
        endStageRunBusyDialog();
        String title = STAGE1.equals(script) ? "段階1 実行中" : "段階2 実行中";
        String header =
                STAGE1.equals(script)
                        ? "段階1（成形）を実行しています"
                        : "段階2（配台計画）を実行しています";
        stageRunBusyDialog =
                StageRunBusyDialog.show(
                        primaryStage, title, header, "準備中…", this::cancelActiveStageRun);
    }

    private void updateStageRunBusyPhase(String phase) {
        if (stageRunBusyDialog != null && stageRunBusyDialog.isShowing() && phase != null) {
            stageRunBusyDialog.setPhase(phase);
        }
    }

    private void onStageRunChildLogLine(String line) {
        if (stageRunBusyDialog == null || !stageRunBusyDialog.isShowing()) {
            return;
        }
        StageRunLogProgressParser.extractDetail(line).ifPresent(stageRunBusyDialog::setDetail);
    }

    void syncStageRunBusyFromStage2Progress(MainRunStage2Progress.State state, String detail) {
        if (state == null) {
            return;
        }
        if (shouldCloseStageRunBusyForPostStage2AsyncWork(state)) {
            endStageRunBusyDialog();
            return;
        }
        if (stageRunBusyDialog == null || !stageRunBusyDialog.isShowing()) {
            return;
        }
        stageRunBusyDialog.setPhase(state.message());
        if (detail != null && !detail.isBlank()) {
            stageRunBusyDialog.setDetail(detail.strip());
        }
    }

    /**
     * 段階2 Python 完了後の配台表再読込・納期管理更新・Excel 生成は非同期が長い。モーダルを閉じ、各タブの進捗 UI に任せる。
     */
    static boolean shouldCloseStageRunBusyForPostStage2AsyncWork(MainRunStage2Progress.State state) {
        return state == MainRunStage2Progress.State.DISPATCH_RELOADING
                || state == MainRunStage2Progress.State.DELIVERY_RELOADING
                || state == MainRunStage2Progress.State.EXCEL_GENERATING;
    }

    private void endStageRunBusyDialog() {
        if (stageRunBusyDialog != null) {
            stageRunBusyDialog.close();
            stageRunBusyDialog = null;
        }
    }

    /** スプラッシュ／操作者ダイアログのあと、進捗が見えるよう短いパルスを空けてから処理する。 */
    private void runAfterUiPulse(Runnable action) {
        PauseTransition pause = new PauseTransition(Duration.millis(50));
        pause.setOnFinished(e -> Platform.runLater(action));
        pause.play();
    }

    /**
     * 工場ワークスペース復元 → 環境変数起動照合 → 起動後 BG 読込まで進捗モーダル付きで実行し、完了後に依頼書原本案内を出す。
     */
    private void runOperatorStartupWorkspaceAndEnvCheckWithProgress() {
        startupSequenceActive = true;
        beginEnvVarsStartupCheckBusy(EnvVarsStartupCheckBusyDialog.STATUS_RESTORE_WORKSPACE);
        runAfterUiPulse(
                () -> {
                    try {
                        if (finalizeOperatorLocalWorkspaceAfterSessionEstablished()) {
                            return;
                        }
                        continueStartupEnvCheckAfterWorkspaceReady();
                    } catch (RuntimeException ex) {
                        finishStartupSequenceProgressAndPrompt();
                        throw ex;
                    }
                });
    }

    private void continueStartupEnvCheckAfterWorkspaceReady() {
        if (envVarsStartupCheckCompleted.get()) {
            finishStartupSequenceAfterEnvCheck();
            return;
        }
        updateEnvVarsStartupCheckBusy(EnvVarsStartupCheckBusyDialog.STATUS_STABILIZE);
        runAfterUiPulse(
                () -> {
                    try {
                        completeEnvVarsStartupCheck(false);
                        finishStartupSequenceAfterEnvCheck();
                    } catch (RuntimeException ex) {
                        finishStartupSequenceProgressAndPrompt();
                        throw ex;
                    }
                });
    }

    private void finishStartupSequenceAfterEnvCheck() {
        if (!isEnvVarsInitializationPending()) {
            ensureMainShellRunTabSelected();
        }
        boolean allowBackgroundLoad =
                isStartupBackgroundLoadAllowed()
                        || (startupRestoredFactorySite && isFactorySwitchBackgroundLoadAllowed());
        if (!allowBackgroundLoad) {
            runAfterUiPulse(this::finishStartupSequenceProgressAndPrompt);
            return;
        }
        updateEnvVarsStartupCheckBusy(EnvVarsStartupCheckBusyDialog.STATUS_BACKGROUND_LOAD);
        startupAwaitingBackgroundLoadBeforeModalClose = true;
        reloadAttendanceTabsFromJson(true);
        runAfterUiPulse(
                () -> {
                    if (startupTabBackgroundLoad != null) {
                        if (startupRestoredFactorySite) {
                            startupTabBackgroundLoad.resetAndScheduleAfterFactorySwitch();
                        } else {
                            startupTabBackgroundLoad.resetAndSchedule();
                        }
                    } else {
                        finishStartupSequenceProgressAndPrompt();
                    }
                });
    }

    private void finishStartupSequenceProgressAndPrompt() {
        startupSequenceActive = false;
        startupRestoredFactorySite = false;
        startupAwaitingBackgroundLoadBeforeModalClose = false;
        endFactorySiteSwitchBusy();
        endEnvVarsStartupCheckBusy();
        if (!isEnvVarsInitializationPending()
                && !shouldSuppressStartupRequestFormOriginalDirPrompt()) {
            maybePromptRequestFormOriginalDirAtStartup();
        }
    }

    /** 初期化記録・起動時照合の直前に、環境変数タブの表示値を安定化する（ブートストラップ補完・表示既定・フォルダ正規化）。 */
    private void stabilizeEnvRowsForInitializationBaseline() {
        Map<String, String> ui = collectUiEnv();
        ensureBootstrapDefaultValuesVisible(ui);
        ensureUiRefOptionalDisplayDefaultsVisible(ui);
        applyRepoFolderPathNormalization();
    }

    private void recordEnvInitializationBaseline() {
        stabilizeEnvRowsForInitializationBaseline();
        EnvVarsInitializedAtStore.recordEnvFingerprint(
                collectUiEnv(), MainShellController::includeInEnvInitFingerprint);
    }

    /**
     * 起動時（工場ワークスペース復元後）に、環境変数タブの値が初期化テンプレートと一致するか一度だけ判定する。
     */
    private void evaluateEnvVarsDifferFromInitialAtStartup() {
        try {
            if (deferOperatorPromptForPortableUpgrade.get()) {
                envVarsDifferFromInitialAtStartup.set(false);
                return;
            }
            String operator = FactoryOperatorUserStore.sessionOperatorName();
            if (FactoryOperatorUserStore.isGuestOperator(operator)) {
                envVarsDifferFromInitialAtStartup.set(false);
                return;
            }
            FactorySite site = GlobalInitSettingTarget.load();
            stabilizeEnvRowsForInitializationBaseline();
            Map<String, String> current = collectUiEnv();
            java.util.function.Predicate<String> keyFilter = MainShellController::includeInEnvInitFingerprint;
            boolean matches;
            if (EnvVarsInitializedAtStore.loadEnvFingerprint().isPresent()) {
                matches = EnvVarsInitializedAtStore.envFingerprintMatches(current, keyFilter);
                if (!matches
                        && EnvVarsInitializedAtStore.matchesRecordedBaselineForKeys(current, keyFilter)) {
                    recordEnvInitializationBaseline();
                    matches = true;
                }
            } else if (EnvVarsInitializedAtStore.isRecorded()) {
                matches = true;
            } else {
                Map<String, String> expected = buildExpectedEnvMapAfterFullInit(site);
                matches = EnvVarsInitialTemplate.matches(current, expected, keyFilter);
            }
            envVarsDifferFromInitialAtStartup.set(!matches);
            if (!matches && EnvVarsInitializedAtStore.isRecorded()) {
                appendLog(
                        "[env] 起動時チェック: 環境変数の値が初期値と異なります。"
                                + "環境変数タブの「環境変数を初期化」を実行するまで、他タブは操作できません。");
            }
        } finally {
            envVarsStartupCheckCompleted.set(true);
        }
    }

    private Map<String, String> buildExpectedEnvMapAfterFullInit(FactorySite site) {
        Map<String, String> ui = collectUiEnv();
        AppPaths.ensureAllDispatchLookupTablesFromRepoIfMissing(ui);
        return EnvVarsInitialTemplate.buildExpectedMap(
                BOOTSTRAP_ORDER,
                site,
                MainShellController::bootstrapDefaultValueForKey,
                MainShellController::optionalUiRefDisplayDefaultForKey,
                MainShellController::includeInEnvInitFingerprint);
    }

    /**
     * 環境変数初期化未記録時、環境変数葉以外へ遷移しようとしたら戻す。
     *
     * @return ブロックして環境変数タブへ戻したとき {@code true}
     */
    private boolean blockMainShellTabSelectionIfEnvInitPending() {
        if (suppressEnvVarsInitTabGuard.get()
                || !isEnvVarsInitializationPending()
                || mainShellTabEnv == null
                || tabPane == null) {
            return false;
        }
        Tab effective =
                resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        Tab envLeaf = mainShellTabFor(MainShellTabId.ENV);
        if (effective == envLeaf) {
            return false;
        }
        suppressEnvVarsInitTabGuard.set(true);
        try {
            ensureMainShellEnvTabSelected();
            if (envInitTabBlockLogEmitted.compareAndSet(false, true)) {
                String reason =
                        EnvVarsInitializedAtStore.isRecorded()
                                ? "環境変数の値が初期値と異なるため"
                                : "環境変数の初期化が未完了のため";
                appendLog(
                        "[env] "
                                + reason
                                + "、環境変数タブ以外は操作できません。"
                                + "環境変数タブの「環境変数を初期化」を実行してください。");
            }
        } finally {
            suppressEnvVarsInitTabGuard.set(false);
        }
        return true;
    }

    /**
     * ゲスト操作者時、実行・ログ葉以外へ遷移しようとしたら戻す。
     *
     * @return ブロックして実行・ログタブへ戻したとき {@code true}
     */
    private boolean blockMainShellTabSelectionIfGuestSessionOnly() {
        if (suppressGuestSessionTabGuard.get()
                || !FactoryOperatorUserStore.isGuestSession()
                || mainShellTabRun == null
                || tabPane == null) {
            return false;
        }
        Tab effective =
                resolveEffectiveLeafTab(tabPane.getSelectionModel().getSelectedItem());
        if (effective == mainShellTabRun) {
            return false;
        }
        suppressGuestSessionTabGuard.set(true);
        try {
            ensureMainShellRunTabSelected();
            appendLog("[guest] ゲスト操作者は工場切替のみ利用できます。");
        } finally {
            suppressGuestSessionTabGuard.set(false);
        }
        return true;
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
     *
     * <p>バージョンアップ確認直後〜同期中に依頼書原本フォルダ案内が重ならないよう、VU 実行中／完了後処理済みのときは
     * {@link #maybePromptRequestFormOriginalDirAtStartup()} をスキップする（未設定時は工場既定へフォールバック）。
     */
    void schedulePortableBundleSelfUpdateAfterSplash() {
        Platform.runLater(
                () -> {
                    maybePortableBundleSelfUpdate();
                    maybePromptOperatorUserAtStartup();
                });
    }

    /**
     * ポータル VU 実行中、または VU 完了後処理で操作者復元済みのときは、起動時の依頼書原本フォルダ案内を出さない。
     */
    private boolean shouldSuppressStartupRequestFormOriginalDirPrompt() {
        return deferOperatorPromptForPortableUpgrade.get()
                || skipOperatorPromptAfterPortableUpgrade.get();
    }

    private void maybePromptRequestFormOriginalDirAtStartup() {
        if (shouldSuppressStartupRequestFormOriginalDirPrompt()) {
            return;
        }
        maybePromptRequestFormOriginalDirIfUnset("[startup]", null);
    }

    /**
     * {@link AppPaths#KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR} が空のとき、BOX ドライブ上の依頼書原本フォルダ選択を1回案内する。
     * キャンセル・未選択でも起動を続行する（実行時は {@link AppPaths#resolveRequestFormOriginalDir} の工場既定へフォールバック）。
     *
     * @param factorySiteHint 環境変数初期化直後など、工場既定の説明を出すときに渡す（起動時案内では {@code null}）
     */
    private void maybePromptRequestFormOriginalDirIfUnset(String logPrefix, FactorySite factorySiteHint) {
        if (primaryStage == null || envRows == null) {
            return;
        }
        if (!envTabValueTrimmed(AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR).isEmpty()) {
            return;
        }
        Alert intro = new Alert(AlertType.INFORMATION);
        initDialogOwnerIfSceneReady(intro);
        applyAlertStylesheetsFromOwner(intro);
        intro.setTitle("依頼書原本フォルダ（任意）");
        intro.setHeaderText(null);
        String unsetReason =
                factorySiteHint != null
                        ? "環境変数の初期化により "
                                + AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR
                                + " が未設定になりました。\n"
                        : AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR + " が未設定です。\n";
        String factoryFallback =
                factorySiteHint != null
                        ? "スキップした場合は "
                                + factorySiteHint.displayLabelJa()
                                + " の工場既定（"
                                + factorySiteHint.requestFormOriginalDir()
                                + "）で動作します。"
                        : "スキップした場合は工場既定（受注ファイル既定の親フォルダ）で動作します。";
        intro.setContentText(
                unsetReason
                        + "BOX ドライブ上の依頼書原本フォルダ（*加工依頼書*.xlsm 等を含むフォルダ）を指定できます。\n"
                        + factoryFallback);
        intro.showAndWait();

        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle("BOX ドライブの依頼書原本フォルダを選択（任意）");
        resolveBoxDriveInitialDirectory()
                .filter(Files::isDirectory)
                .ifPresent(p -> dc.setInitialDirectory(p.toFile()));
        File selected = dc.showDialog(primaryStage);
        if (selected == null) {
            appendLog(
                    logPrefix
                            + " "
                            + AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR
                            + " は未設定のまま起動します（実行時は工場既定を使用）。");
            return;
        }
        String abs = selected.toPath().toAbsolutePath().normalize().toString();
        updateEnvTabValue(AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, abs);
        DesktopSessionStateStore.save(collectDesktopSession());
        appendLog(logPrefix + " " + AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR + " を設定: " + abs);
    }

    /** BOX 同期フォルダ（{@code %USERPROFILE%\\Box} 等）があれば DirectoryChooser の初期ディレクトリ候補にする。 */
    private static Optional<Path> resolveBoxDriveInitialDirectory() {
        LinkedHashSet<String> candidates = new LinkedHashSet<>();
        String userProfile = System.getenv("USERPROFILE");
        if (userProfile != null && !userProfile.isBlank()) {
            candidates.add(Path.of(userProfile.trim(), "Box").toString());
        }
        String home = System.getProperty("user.home");
        if (home != null && !home.isBlank()) {
            candidates.add(Path.of(home.trim(), "Box").toString());
        }
        for (String raw : candidates) {
            try {
                Path p = Path.of(raw).toAbsolutePath().normalize();
                if (Files.isDirectory(p)) {
                    return Optional.of(p);
                }
            } catch (Exception ignored) {
                // try next candidate
            }
        }
        return Optional.empty();
    }

    /**
     * After stage 1 writes {@code stage1_exclude_rules.json} under the local output dir, mirror the path into the env tab so
     * {@code PM_AI_EXCLUDE_RULES_JSON} matches the next child-process run.
     */
    private void applyStage1ExcludeRulesJsonToEnvTab() {
        if (envRows == null) {
            return;
        }
        try {
            Map<String, String> ui = collectUiEnv();
            AppPaths.ensureStage1ExcludeRulesJsonFromRepoIfMissing(ui);
            Path p = AppPaths.stage1ExcludeRulesJsonPath(ui);
            if (!Files.isRegularFile(p)) {
                Path legacyCode = AppPaths.stage1ExcludeRulesJsonPathLegacyUnderCodeJson(ui);
                if (Files.isRegularFile(legacyCode)) {
                    p = legacyCode;
                } else {
                    Path legacy = AppPaths.stage1ExcludeRulesJsonPathLegacyUnderPython(ui);
                    if (Files.isRegularFile(legacy)) {
                        p = legacy;
                    }
                }
            }
            if (!Files.isRegularFile(p)) {
                appendLog(
                        "[env] PM_AI_EXCLUDE_RULES_JSON: "
                                + p
                                + " が無いため、環境変数タブの値は未更新のままです。"
                                + " 段階1が配台除外ルール JSON を生成しているか、サマリ Excel と同一フォルダの場所が一致しているか確認してください。");
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

    /** 開発用: 段階1正常終了後、配台計画_タスク入力の全行「配台不要」を yes にする。 */
    private void applyStage1DevMarkAllExcludeAfterRunIfEnabled() {
        if (mainRunTabController == null
                || !mainRunTabController.snapshotStage1MarkAllExcludeAfterRun()) {
            return;
        }
        try {
            Stage1DevMarkAllExcludeAfterRun.ApplySummary summary =
                    Stage1DevMarkAllExcludeAfterRun.applyToPlanInput(collectUiEnv());
            appendLog(
                    "[dev] 段階1後: 配台計画_タスク入力の全行を配台不要 yes に更新しました（"
                            + summary.updatedRows()
                            + "/"
                            + summary.totalRows()
                            + " 行変更、"
                            + summary.planPath()
                            + "）。");
        } catch (Exception ex) {
            appendLog(
                    "[dev] 段階1後: 全行配台不要への更新に失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
    }

    private void runStage(String script) {
        if (STAGE1.equals(script) && blockIfPlanningStagesCalendarNotReady("段階1")) {
            return;
        }
        if (stage2SourceGuardCoordinator.isRunning() && !stage2SourceGuardRunHandoff) {
            appendLog("[busy] 固定ソース確認中のため段階処理を開始できません。");
            return;
        }
        if (isDeliveryCalendarReloadBlockingStageRuns()) {
            String stageJa =
                    STAGE1.equals(script)
                            ? "段階1"
                            : (STAGE2_1.equals(script) ? "段階2.1" : "段階2");
            appendLog("[busy] 納期管理ビュー再読み込み中のため " + stageJa + " を開始できません。");
            return;
        }
        if (STAGE2.equals(script) && blockIfMaterialLookupTablesHaveBlankValues("段階2")) {
            return;
        }
        if (STAGE2_1.equals(script) && blockIfMaterialLookupTablesHaveBlankValues("段階2.1")) {
            return;
        }
        if (STAGE2.equals(script) && !confirmStage2UnknownMasterCombinationsBeforeRun()) {
            return;
        }
        if (STAGE2_1.equals(script) && !confirmStage2UnknownMasterCombinationsBeforeRun()) {
            return;
        }
        if (STAGE2.equals(script) && !confirmStage2MissingSkillsColumnsBeforeRun()) {
            return;
        }
        if (STAGE2_1.equals(script) && !confirmStage2MissingSkillsColumnsBeforeRun()) {
            return;
        }
        if (!runLock.compareAndSet(false, true)) {
            appendLog("[busy] already running (single flight).");
            return;
        }
        activeRunStageScript = script;
        activeStageRunUserCancelled.set(false);
        selectMainShellTab(MainShellTabId.RUN);
        if (STAGE2.equals(script)) {
            mainRunTabController.updateStage2Progress(
                    MainRunStage2Progress.State.RUNNING, "");
        }
        applyRunTabGating();
        if (STAGE1.equals(script) || STAGE2.equals(script)) {
            beginStageRunBusyDialog(script);
            if (STAGE2.equals(script)) {
                syncStageRunBusyFromStage2Progress(MainRunStage2Progress.State.RUNNING, "");
            }
        }
        mainRunTabController.beginLogTailFollowForRun();
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
            if (STAGE1.equals(script) || STAGE2.equals(script) || STAGE2_1.equals(script)) {
                // 段階1／2（および2.1）コア成果物はローカルのみ。共有上の残骸は削除（アラジンExcelは対象外）。
                // 削除は env 書き換え前に行い、共有を指していた OUTPUT_DIR も走査対象に含める。
                for (Path removed : SharedPipelineResultsCleaner.deletePipelineArtifactsFromShared(uiRun)) {
                    appendLog("[cleanup] 共有上の段階成果物を削除: " + removed);
                }
                if (PipelineLocalResultsPolicy.rewritePipelineOutputEnvToLocal(uiRun)) {
                    syncEnvTabValue(
                            AppPaths.KEY_PM_AI_OUTPUT_DIR,
                            uiRun.getOrDefault(AppPaths.KEY_PM_AI_OUTPUT_DIR, ""));
                    syncEnvTabValue(
                            AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
                            uiRun.getOrDefault(AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR, ""));
                    if (uiRun.containsKey(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH)) {
                        syncEnvTabValue(
                                AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                                uiRun.get(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH));
                    }
                    if (uiRun.containsKey(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON)) {
                        syncEnvTabValue(
                                AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
                                uiRun.get(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON));
                    }
                    appendLog("[policy] 段階成果物の出力先をローカルディスクに固定しました。");
                }
            }
            overlayWorkingExcludeRulesJsonPathForStageRun(uiRun);
            overlayDispatchSpecialRulesForStageRun(uiRun, script);
            overlayDispatchLookupTablePathsForStageRun(uiRun);
            overlayMainRunSkipGeminiApiEnv(uiRun);
            if (STAGE1.equals(script)) {
                uiRun.put(AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH, "0");
                /*
                 * 当日配台ソース束は invalidate（開始時）/ persist（成功時）で管理する。
                 * ここで消すと、完了時保存前に束が消えたまま段階2へ進む事故の温床になる。
                 */
                PipelineDownstreamResultsClearer.ClearResult downstreamClear =
                        PipelineDownstreamResultsClearer.clearStage2Downstream(
                                uiRun, true /* preserveTodayDispatchSourceBundle */);
                for (String line : downstreamClear.detailLines()) {
                    appendLog(line);
                }
                if (downstreamClear.anyFailed()) {
                    appendLog("[stage1] 段階2〜段階2.1 成果物の一部を削除できませんでした。");
                }
                pendingStage21OvertimeJsonPath = null;
                pendingStage2InProgressNextDayJsonPath = null;
                pendingStage2AladdinTodayExcludeJsonPath = null;
                syncUiAfterDownstreamPipelineResultsCleared();
            }
            if (STAGE2.equals(script)) {
                uiRun.put(AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL, "1");
                uiRun.put(
                        AppPaths.KEY_PM_AI_STAGE2_SKIP_TODAY_DISPATCH,
                        mainRunTabController.snapshotStage2SkipTodayDispatch() ? "1" : "0");
                uiRun.put(AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH, "0");
                applyStage2NextDayDialogEnvs(uiRun);
                overlayPlanInputComboSheetMayExceedNeedEnv(uiRun);
                overlayPlanInputStage2SkipGeminiApiEnv(uiRun);
                String resultFont = mainRunTabController.snapshotStage2ResultBookFont();
                if (resultFont != null && !resultFont.isBlank()) {
                    uiRun.put(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT, resultFont.trim());
                } else {
                    uiRun.remove(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT);
                }
            }
            if (STAGE2_1.equals(script)) {
                java.nio.file.Path ot = pendingStage21OvertimeJsonPath;
                Map<String, String> stage21Snap = snapshotStage21PythonEnv(ot);
                uiRun.putAll(stage21Snap);
                overlayPlanInputStage2SkipGeminiApiEnv(uiRun);
            }
            String wb = effectiveTaskInputWorkbookPath();
            appendLog("--- start: " + script + " ---");
            if (STAGE1.equals(script) && mainRunTabController.snapshotSkipGeminiApi()) {
                appendLog(
                        "[run] PM_AI_SKIP_GEMINI_API=1 — Gemini API 呼び出しをスキップします（段階1）。");
            }
            if ((STAGE2.equals(script) || STAGE2_1.equals(script))
                    && planInputTabController != null
                    && planInputTabController.snapshotStage2SkipGeminiApi()) {
                appendLog(
                        "[run] PM_AI_SKIP_GEMINI_API=1 — Gemini API 呼び出しをスキップします（段階2）。");
            }
            if (STAGE1.equals(script)
                    && mainRunTabController.snapshotStage1MarkAllExcludeAfterRun()) {
                appendLog(
                        "[dev] 段階1正常終了後、配台計画_タスク入力の全行を配台不要 yes に更新します（開発用）。");
            }
            if (STAGE1.equals(script)) {
                updateStageRunBusyPhase("キャッシュをクリアしています…");
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
                    clearPlanInputTableForStage1CacheClear();
                } catch (IOException archiveEx) {
                    appendLog(
                            "[stage1] キャッシュ退避に失敗したためクリアを中止しました: "
                                    + archiveEx.getMessage());
                    runLock.set(false);
                    activeRunStageScript = null;
                    activeStageChildProcess.set(null);
                    mainRunTabController.getStatusLabel().setText("キャッシュ退避失敗");
                    applyRunTabGating();
                    endStageRunBusyDialog();
                    return;
                }
            }
            if (STAGE1.equals(script) || STAGE2.equals(script) || STAGE2_1.equals(script)) {
                refreshNetworkSourceDirListingSkipsBeforeStageRun(uiRun);
            }
            overlayTodayDispatchSourcesForStageRun(uiRun, script);
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
            Path dir = AppPaths.resolvePythonScriptDir(uiRun);
            Path codeDir = AppPaths.resolveCodeDir(uiRun);
            childEnv.put(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR, dir.toAbsolutePath().normalize().toString());
            childEnv.put(AppPaths.KEY_PM_AI_CODE_DIR, codeDir.toAbsolutePath().normalize().toString());
            if (STAGE1.equals(script)) {
                appendLog("[stage1] Python script dir (resolved): " + dir.toAbsolutePath().normalize());
                appendLog("[stage1] code dir (material tables): " + codeDir.toAbsolutePath().normalize());
                Path corePy = dir.resolve("planning_core").resolve("_core.py");
                appendLog(
                        "[stage1] planning_core/_core.py exists="
                                + java.nio.file.Files.isRegularFile(corePy)
                                + " path="
                                + corePy.toAbsolutePath().normalize());
                try {
                    PlanningCoreMaterialTableAppendProbe.Result materialAppend =
                            PlanningCoreMaterialTableAppendProbe.detect(dir);
                    appendLog(
                            "[stage1] planning_core 材料テーブル追記仕様="
                                    + materialAppend.spec().logLabel()
                                    + materialAppend
                                            .buildId()
                                            .map(id -> " build=" + id)
                                            .orElse(""));
                    if (materialAppend.spec()
                            == PlanningCoreMaterialTableAppendProbe.Spec.LEGACY) {
                        appendLog(
                                "[stage1] 警告: 同梱 pm-ai-data 等の古い planning_core です。"
                                        + "英字開始の製品名は plan_input_tasks から除外され、材料テーブルにも追記されません。"
                                        + "環境変数 PM_AI_CODE_PYTHON_DIR をリポジトリの code\\python に設定し、JavaFX を再ビルドしてください。");
                    }
                } catch (Exception ex) {
                    appendLog("[stage1] planning_core 材料テーブル追記仕様の判定失敗: " + ex.getMessage());
                }
            }
            appendStageChildResolvedEnvForRun(script, childEnv);
            RunRequest req = new RunRequest(py, dir, script, wb, childEnv);
            mainRunTabController.getStatusLabel().setText("実行中…");
            if (usesStageRunBusyModal(script)) {
                updateStageRunBusyPhase("Python 実行中…");
            }
            PipelineExecutionTimingKind stageTimingKind = pipelineTimingKindForStageScript(script);
            if (stageTimingKind != null) {
                beginPipelineExecutionTiming(stageTimingKind);
            }

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
                                    onStageRunChildLogLine(line);
                                    if (STAGE1.equals(script)
                                            && line.contains("製品厚みを決定できずスキップ")) {
                                        appendLog(
                                                "[stage1] 古い planning_core を実行中: 英字開始製品はタスク行ごとスキップされます。"
                                                        + " PM_AI_CODE_PYTHON_DIR="
                                                        + childEnv.getOrDefault(
                                                                AppPaths.KEY_PM_AI_CODE_PYTHON_DIR,
                                                                "（未設定）"));
                                    }
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
                        endStageRunBusyDialog();
                        applyRunTabGating();
                        if (stage2 && dispatchInteractiveTabController != null) {
                            dispatchInteractiveTabController.reloadTableFromDiskAfterExternalUpdate();
                        }
                        if (stage1) {
                            mainRunTabController.resetDevCheckboxesAfterStage1Run();
                        }
                        if (stage1 || stage2) {
                            selectMainShellTab(MainShellTabId.RUN);
                            showStageFailureDialog(script, null, t, List.of());
                        }
                    });
        }
    }

    private void completeStageRunOnFx(String script, Integer code, Throwable err, List<String> tailSnap) {
        mainRunTabController.flushPendingLogAppends();
        if (activeStageRunUserCancelled.getAndSet(false)) {
            code = 9;
            err = null;
            appendLog("[interrupt] ユーザー操作により段階処理を中断しました。");
        }
        DispatchRuleBuilderRunContext.get().clearActiveRun();
        PipelineExecutionTimingKind stageTimingKind = pipelineTimingKindForStageScript(script);
        if (stageTimingKind != null) {
            endPipelineExecutionTiming(stageTimingKind);
        }
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
            if (STAGE2_1.equals(script) && dispatchInteractiveTabController != null) {
                dispatchInteractiveTabController.reloadTableFromDiskAfterExternalUpdate();
            }
        } else {
            int c = code != null ? code : -1;
            mainRunTabController.getStatusLabel().setText(exitCodeLegend(c));
            appendLog("[end] exitCode=" + c + " " + exitHint(c));
            if (STAGE1.equals(script) && c == 0) {
                Stage1SourceBundleCompletionGate.Result bundleResult =
                        persistStage1SourceBundleAfterSuccess();
                if (!bundleResult.completionAllowed()) {
                    mainRunTabController.getStatusLabel().setText("bundle保存失敗");
                    appendLog("[stage1] " + bundleResult.message());
                    showErrorDialog(
                            "段階1",
                            bundleResult.message() + "\n段階1は完了扱いにしません。再実行してください。");
                    mainRunTabController.flushPendingLogAppends();
                    endStageRunBusyDialog();
                    maybeArchiveRemoteSupportLogAfterStage(script, code, err);
                    return;
                }
                applyStage1ExcludeRulesJsonToEnvTab();
                try {
                    CodeDispatchLookupTablesMerge.MergeSummary ms =
                            CodeDispatchLookupTablesMerge.mergeAfterStage1(collectUiEnv());
                    if (ms.totalAdded() > 0) {
                        appendLog("[stage1] 材料・製品種類情報(code/) 自動追記: " + ms.summaryJa());
                    } else {
                        appendLog("[stage1] 材料・製品種類情報(code/) 自動追記: 追記なし");
                    }
                    Path codeDirAfter = AppPaths.resolveCodeDir(collectUiEnv());
                    Path thickPath =
                            codeDirAfter.resolve(CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK);
                    appendLog("[stage1] 製品厚みテーブル(正本): " + thickPath.toAbsolutePath().normalize());
                } catch (Exception ex) {
                    appendLog("[stage1] 材料・製品種類情報(code/) 自動追記失敗: " + ex.getMessage());
                }
                if (codeDispatchLookupTablesTabController != null) {
                    codeDispatchLookupTablesTabController.reloadAllFromDisk();
                }
                promptStage1NewMaterialLookupsAfterMerge();
                warnStage1MissingSkillsColumnsAfterSuccess();
                if (reloadAfterStage1Preview != null) {
                    reloadAfterStage1Preview.run();
                }
                applyStage1DevMarkAllExcludeAfterRunIfEnabled();
                Path stage1PlanPath =
                        AppPaths.defaultStage1PlanTasksPath(collectUiEnv())
                                .toAbsolutePath()
                                .normalize();
                updateEnvTabValue(
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH, stage1PlanPath.toString());
                if (planInputTabController != null) {
                    planInputTabController.restoreDesktopSessionPaths(
                            stage1PlanPath.toString(), AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
                }
                if (reloadAfterStage1PlanInput != null) {
                    reloadAfterStage1PlanInput.run();
                }
                promptStage1EcSideUnknownAfterSuccess();
                syncRequestFormFeedLocFromStage1Plan();
                invalidateDeliveryCalendarAfterPipelineRun();
                refreshEquipmentGanttGraphicAfterPipelineRun();
                MacroCompleteChime.playIfAvailable(collectUiEnv());
                selectMainShellTab(MainShellTabId.PLAN_INPUT);
                String completionMsg = buildStage1CompletionMessage();
                endStageRunBusyDialog();
                showStageCompletionDialog("段階1 完了", completionMsg);
            }
            if (STAGE2.equals(script)) {
                if (c == 0) {
                    markStage2PipelineAwaitingExcelThisLaunch();
                    mainRunTabController.updateStage2Progress(
                            MainRunStage2Progress.State.DISPATCH_RELOADING, "");
                    refreshStage2OutputArtifacts();
                    final Integer stage2ExitCode = code;
                    final Throwable stage2Err = err;
                    Platform.runLater(
                            () -> {
                                refreshEquipmentGanttGraphicAfterPipelineRun();
                                refreshOperatorCardAfterPipelineRun();
                                if (dispatchInteractiveTabController != null) {
                                    dispatchInteractiveTabController.reloadSpecialRuleBadges();
                                }
                                Runnable afterDispatchReload =
                                        () -> {
                                            if (planInputTabController != null) {
                                                planInputTabController
                                                        .reloadQuietlyFromDiskAfterStage2IfClean();
                                            }
                                            // 反映漏れダイアログ／[配台整合] ログの後に remote_log を残す
                                            mainRunTabController.flushPendingLogAppends();
                                            maybeArchiveRemoteSupportLogAfterStage(
                                                    STAGE2, stage2ExitCode, stage2Err);
                                            mainRunTabController.updateStage2Progress(
                                                    MainRunStage2Progress.State.DELIVERY_RELOADING,
                                                    "");
                                            selectMainShellTab(
                                                    MainShellTabId.DELIVERY_CALENDAR_VIEW);
                                            Runnable afterDeliveryCalendarReload =
                                                    () -> {
                                                        mainRunTabController
                                                                .updateStage2Progress(
                                                                        MainRunStage2Progress.State
                                                                                .EXCEL_GENERATING,
                                                                        "");
                                                        exportSharedAladdinEntryWorkbookAfterStage2(
                                                                outcome -> {
                                                                    updateStage2ExcelProgress(
                                                                            outcome);
                                                                    endStageRunBusyDialog();
                                                                    if (outcome != null
                                                                            && outcome
                                                                                    .succeeded()) {
                                                                        stage2IdentityCloseGate
                                                                                .markExcelExportSucceeded();
                                                                    }
                                                                    MacroCompleteChime
                                                                            .playIfAvailable(
                                                                                    collectUiEnv());
                                                                    showStageCompletionDialogAndWait(
                                                                            "段階2 完了",
                                                                            stage2CompletionHeader(
                                                                                    outcome),
                                                                            stage2CompletionContent(
                                                                                    outcome),
                                                                            () -> {
                                                                                selectMainShellTab(
                                                                                        MainShellTabId
                                                                                                .DELIVERY_CALENDAR_VIEW);
                                                                                if (deliveryCalendarViewTabController
                                                                                        != null) {
                                                                                    deliveryCalendarViewTabController
                                                                                            .selectDispatchResultInnerTab(
                                                                                                    false);
                                                                                }
                                                                            });
                                                                    showRawInputMorningDispatchRateWarningAfterStage2();
                                                                });
                                                    };
                                            if (deliveryCalendarViewTabController != null) {
                                                deliveryCalendarViewTabController
                                                        .reloadInBackgroundAfterStage2Success(
                                                                afterDeliveryCalendarReload);
                                            } else {
                                                afterDeliveryCalendarReload.run();
                                            }
                                        };
                                if (dispatchInteractiveTabController != null) {
                                    dispatchInteractiveTabController.reloadTableFromDiskAfterStage2Success(
                                            afterDispatchReload);
                                } else {
                                    afterDispatchReload.run();
                                }
                                if (specialRulesTabController != null) {
                                    specialRulesTabController.reloadTraceFromDisk();
                                }
                            });
                    // サマリ xlsx は段階2 exit 0 直後には作らない。納期管理ビュー再読込完了後に出力する。
                } else if (dispatchInteractiveTabController != null) {
                    dispatchInteractiveTabController.reloadTableFromDiskAfterExternalUpdate();
                }
            }
            if (STAGE2_1.equals(script)) {
                if (c == 0) {
                    Platform.runLater(
                            () -> {
                                Map<String, String> ui = collectUiEnv();
                                java.nio.file.Path mainJson =
                                        AppPaths.resolveResultDispatchTableStage2JsonPath(ui);
                                java.nio.file.Path overtimeJson =
                                        pendingStage21OvertimeJsonPath;
                                try {
                                    jp.co.pm.ai.desktop.dispatch.Stage21OutputPromoter.Result promoted =
                                            jp.co.pm.ai.desktop.dispatch.Stage21OutputPromoter
                                                    .promoteToMainOutput(ui);
                                    appendLog(
                                            "[stage2.1] メイン output へ "
                                                    + promoted.filesCopied()
                                                    + " 件を正本反映しました");
                                    java.nio.file.Path mainOvertime =
                                            AppPaths.resolveResultDispatchTableDir(ui)
                                                    .resolve("overtime_simulation_overrides.json");
                                    java.nio.file.Path overridesForMeta =
                                            java.nio.file.Files.isRegularFile(mainOvertime)
                                                    ? mainOvertime
                                                    : overtimeJson;
                                    java.nio.file.Path stage21Json =
                                            AppPaths.resolveStage21ResultDispatchJsonPath(ui);
                                    if (dispatchInteractiveTabController != null) {
                                        dispatchInteractiveTabController
                                                .finalizeStage21PromotedWithComparisonAfterRunSuccess(
                                                        mainJson,
                                                        stage21Json,
                                                        overridesForMeta);
                                    } else {
                                        jp.co.pm.ai.desktop.dispatch.Stage21TrialSnapshotStore
                                                .writePromotedWithComparison(
                                                        mainJson,
                                                        java.util.Map.of(),
                                                        stage21Json,
                                                        overridesForMeta,
                                                        jp.co.pm.ai.desktop.dispatch
                                                                .OvertimeSimulationOverridesReader
                                                                .summarize(overridesForMeta));
                                    }
                                    refreshStage2OutputArtifacts();
                                    if (promoted.mainPlanJson() != null
                                            && equipmentGanttGraphicTabController != null) {
                                        equipmentGanttGraphicTabController.syncPlanJsonPathAndReload(
                                                promoted.mainPlanJson(), false);
                                    } else {
                                        refreshEquipmentGanttGraphicAfterPipelineRun();
                                    }
                                    refreshOperatorCardAfterPipelineRun();
                                    if (dispatchInteractiveTabController != null) {
                                        dispatchInteractiveTabController
                                                .reloadTableFromDiskAfterStage21PromotedSuccess(
                                                        () ->
                                                                notifyStage21OvertimeSimulationSuccess(
                                                                        promoted));
                                    } else {
                                        notifyStage21OvertimeSimulationSuccess(promoted);
                                    }
                                } catch (Exception ex) {
                                    appendLog(
                                            "[stage2.1] 正本への反映に失敗: "
                                                    + (ex.getMessage() != null
                                                            ? ex.getMessage()
                                                            : ex));
                                    showErrorDialog(
                                            "段階2.1 反映エラー",
                                            "段階2.1 は正常終了しましたが、メイン output への正本反映に失敗しました。\n\n"
                                                    + (ex.getMessage() != null
                                                            ? ex.getMessage()
                                                            : ex.toString())
                                                    + "\n\noutput/stage21/ の成果物を手動で確認してください。");
                                    if (dispatchInteractiveTabController != null) {
                                        dispatchInteractiveTabController
                                                .reloadTableFromDiskAfterExternalUpdate();
                                    }
                                }
                            });
                } else if (dispatchInteractiveTabController != null) {
                    dispatchInteractiveTabController.reloadTableFromDiskAfterExternalUpdate();
                }
            }
        }
        boolean stage12 =
                STAGE1.equals(script) || STAGE2.equals(script) || STAGE2_1.equals(script);
        boolean failed =
                err != null
                        || (code != null && code.intValue() != 0 && code.intValue() != 9);
        if (stage12 && failed) {
            appendLog("[ui] 段階処理が異常終了しました。エラーダイアログを表示します。");
            selectMainShellTab(MainShellTabId.RUN);
            if (STAGE2.equals(script)) {
                String detail =
                        err != null
                                ? err.getMessage()
                                : code != null ? "exit=" + code : "";
                mainRunTabController.updateStage2Progress(
                        MainRunStage2Progress.State.STAGE2_FAILED, detail);
            }
            endStageRunBusyDialog();
            if (STAGE2.equals(script) && err == null && code != null && code.intValue() == 3) {
                showStage2FailureWithUnknownMasterComboRetry(code, tailSnap);
            } else {
                showStageFailureDialog(script, err != null ? null : code, err, tailSnap);
            }
        }
        if (STAGE1.equals(script)) {
            mainRunTabController.resetDevCheckboxesAfterStage1Run();
        }
        mainRunTabController.flushPendingLogAppends();
        // 段階2 正常終了は手動修正表の再読込・反映漏れチェック後にアーカイブする
        boolean deferRemoteLogForStage2Success =
                STAGE2.equals(script) && err == null && code != null && code.intValue() == 0;
        if (!deferRemoteLogForStage2Success) {
            maybeArchiveRemoteSupportLogAfterStage(script, code, err);
        }
        boolean keepBusyForStage2PostProcess =
                STAGE2.equals(script) && err == null && code != null && code.intValue() == 0;
        if (usesStageRunBusyModal(script) && !keepBusyForStage2PostProcess) {
            endStageRunBusyDialog();
        }
    }

    /**
     * 段階1／2／2.1 終了時: 共有 {@code remote_log/<操作者>/} へ実行ログを非同期保存する。
     */
    private void maybeArchiveRemoteSupportLogAfterStage(
            String script, Integer code, Throwable err) {
        String stageId =
                RemoteSupportLogArchive.stageIdForMainShellScript(
                        script, STAGE1, STAGE2, STAGE2_1);
        if (stageId == null || mainRunTabController == null) {
            return;
        }
        String uiLog = mainRunTabController.snapshotAllLogText();
        RemoteSupportLogArchive.archiveAfterStageAsync(
                collectUiEnv(), stageId, code, err, uiLog, this::appendLog);
    }

    /**
     * 段階1／2.0～3.2 実行中の Python 子プロセスを終了する（ツールバー・実行・ログの「中断」）。
     */
    void cancelActiveStageRun() {
        boolean didSomething = false;
        Process child = activeStageChildProcess.get();
        if (child != null && child.isAlive()) {
            activeStageRunUserCancelled.set(true);
            appendLog("[interrupt] 段階処理の子プロセスを終了します…");
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
     * 配台試行（段階3／段階3.5）開始時に {@link #runLock} を取得し、シェル全体の操作制限をかける。
     *
     * @return 他処理が実行中などで開始できないとき {@code false}
     */
    boolean tryBeginDispatchTrialGating(PipelineExecutionTimingKind kind) {
        return true;
    }

    void endDispatchTrialGating(PipelineExecutionTimingKind kind) {
        DispatchRuleBuilderRunContext.get().clearActiveRun();
        runLock.set(false);
        applyRunTabGating();
    }

    void endActiveDispatchTrialGatingIfAny() {
        runLock.set(false);
        applyRunTabGating();
    }

    public boolean isPlanningPipelineStageRunning() {
        String script = activeRunStageScript;
        return STAGE1.equals(script) || STAGE2.equals(script) || STAGE2_1.equals(script);
    }

    /**
     * 段階1／2.0～3.2／配台試行 実行中は「実行・ログ」以外のタブを無効化し、タブ切り替えを禁止する（ツールバーに進捗・中断）。
     * 依頼書入力・リモートデスクトップは段階処理と並行操作できるよう除外する。
     */
    private void applyRunTabGating() {
        String script = activeRunStageScript;
        boolean stage1Running = STAGE1.equals(script);
        boolean stageScriptRunning = pipelineTimingKindForStageScript(script) != null;
        boolean dispatchTrialBusy = false;
        boolean sourceGuardBusy = stage2SourceGuardCoordinator.isRunning();
        boolean pipelineBusy = stageScriptRunning || dispatchTrialBusy || sourceGuardBusy;
        if (mainRunTabController != null) {
            mainRunTabController.setStageRunProgressVisible(stage1Running, pipelineBusy);
        }
        if (dispatchInteractiveTabController != null) {
            dispatchInteractiveTabController.setStageRunProgressVisible(stage1Running, pipelineBusy);
        }
        if (planInputTabController != null) {
            planInputTabController.setStageRunProgressVisible(stage1Running, pipelineBusy);
        }
        updateShellStageProgressOverlay(script, null);
        boolean envInitPending = isEnvVarsInitializationPending();
        boolean guestSessionOnly = FactoryOperatorUserStore.isGuestSession();
        applyEnvVarsInitToolbarGating(envInitPending);
        applyGuestSessionToolbarGating(guestSessionOnly);
        if (tabPane == null) {
            if (mainRunTabController != null) {
                mainRunTabController.setGuestSessionFactorySwitchOnly(guestSessionOnly);
            }
            return;
        }
        ObservableList<Tab> tabs = tabPane.getTabs();
        if (tabs.isEmpty()) {
            if (mainRunTabController != null) {
                mainRunTabController.setGuestSessionFactorySwitchOnly(guestSessionOnly);
            }
            return;
        }
        if (envInitPending) {
            Tab envLeaf = mainShellTabFor(MainShellTabId.ENV);
            MainShellRunTabGating.applyEnvInitPending(tabPane, envLeaf);
            ensureMainShellEnvTabSelected();
            if (mainRunTabController != null) {
                mainRunTabController.setGuestSessionFactorySwitchOnly(false);
            }
            return;
        }
        if (guestSessionOnly) {
            Tab runLeaf = mainShellTabFor(MainShellTabId.RUN);
            MainShellRunTabGating.applyGuestSessionOnly(tabPane, runLeaf);
            if (mainRunTabController != null) {
                mainRunTabController.setGuestSessionFactorySwitchOnly(true);
            }
            ensureMainShellRunTabSelected();
            refreshGlobalStatusBar();
            return;
        }
        if (mainRunTabController != null) {
            mainRunTabController.setGuestSessionFactorySwitchOnly(false);
        }
        MainShellRunTabGating.clearDisableRecursive(tabPane);
        MainShellRunTabGating.apply(
                tabPane,
                pipelineBusy,
                this::isMainShellLeafOperableDuringPipelineRun,
                pipelineBusy ? mainShellTabRun : null);
        if (pipelineBusy) {
            // リモートデスクトップ等「操作可能」葉への自動遷移を許容せず実行・ログへ固定する
            ensureMainShellRunTabSelected();
        }
        refreshGlobalStatusBar();
    }

    private void refreshGlobalStatusBar() {
        if (!Platform.isFxApplicationThread()) {
            Platform.runLater(this::refreshGlobalStatusBar);
            return;
        }
        if (globalAppStatusBar == null) {
            return;
        }
        Tab tab =
                tabPane != null && tabPane.getSelectionModel() != null
                        ? tabPane.getSelectionModel().getSelectedItem()
                        : null;
        globalAppStatusBar.setTabName(tab != null ? tab.getText() : "—");
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        globalAppStatusBar.setOperator(operator.isBlank() ? "（未選択）" : operator);
        FactorySite site = GlobalInitSettingTarget.load();
        globalAppStatusBar.setFactory(site != null ? site.displayLabelJa() : "—");
        globalAppStatusBar.setAttendanceReady(attendanceStage2Ready, attendanceReadinessTooltip);
        globalAppStatusBar.setTaskProgress(globalLongTaskProgress);
        globalAppStatusBar.setMessage(resolveGlobalStatusMessage());
    }

    @Override
    public void setGlobalLongTaskProgress(double fraction, String detail) {
        globalLongTaskProgress = fraction;
        globalLongTaskDetail = detail != null ? detail : "";
        refreshGlobalStatusBar();
    }

    @Override
    public void clearGlobalLongTaskProgress() {
        globalLongTaskProgress = null;
        globalLongTaskDetail = "";
        refreshGlobalStatusBar();
    }

    private static final int STARTUP_BACKGROUND_LOAD_STEP_COUNT = 6;
    private static final int STARTUP_BACKGROUND_LOAD_STEP_REQUEST_FORM = 5;

    @Override
    public void setStartupBackgroundLoadStatus(String message) {
        startupBackgroundLoadMessage = message != null ? message : "";
        if (startupBackgroundLoadMessage.isBlank()) {
            clearGlobalLongTaskProgress();
        } else if (isEnvVarsStartupCheckBusyShowing()
                && startupAwaitingBackgroundLoadBeforeModalClose) {
            updateEnvVarsStartupCheckBusy(startupBackgroundLoadMessage);
        } else if (isFactorySiteSwitchBusyShowing()
                && factorySwitchAwaitingBackgroundLoadBeforeModalClose) {
            updateFactorySiteSwitchBusy(
                    FactorySiteSwitchBusySupport.resolveTabLoadStatus(startupBackgroundLoadMessage));
        }
        refreshGlobalStatusBar();
    }

    /** 起動後読込の「原本転記」段階で、依頼書照合の詳細進捗をダイアログへ反映する。 */
    void reportStartupRequestFormReloadProgress(String detail) {
        if (!startupTabBackgroundLoadActive) {
            return;
        }
        String body = detail != null ? detail.strip() : "";
        String message;
        if (body.isBlank()) {
            message =
                    "起動後読込 ("
                            + STARTUP_BACKGROUND_LOAD_STEP_REQUEST_FORM
                            + "/"
                            + STARTUP_BACKGROUND_LOAD_STEP_COUNT
                            + "): 原本転記…";
        } else {
            message =
                    "起動後読込 ("
                            + STARTUP_BACKGROUND_LOAD_STEP_REQUEST_FORM
                            + "/"
                            + STARTUP_BACKGROUND_LOAD_STEP_COUNT
                            + "): 原本転記\n"
                            + body;
        }
        setStartupBackgroundLoadStatus(message);
    }

    @Override
    public void appendStartupBackgroundLog(String line) {
        appendLog(line);
    }

    @Override
    public RemoteDesktopTabController remoteDesktopTab() {
        return remoteDesktopTabController;
    }

    @Override
    public CompanyCalendarTabController companyCalendarTab() {
        return companyCalendarTabController;
    }

    @Override
    public MemberAttendanceTabController memberAttendanceTab() {
        return memberAttendanceTabController;
    }

    @Override
    public MachineCalendarTabController machineCalendarTab() {
        return machineCalendarTabController;
    }

    @Override
    public RequestFormInputTabController requestFormInputTab() {
        return requestFormInputTabController;
    }

    @Override
    public RequestFormPipelineCheckTabController requestFormPipelineCheckTab() {
        return requestFormPipelineCheckTabController;
    }

    @Override
    public void onStartupBackgroundLoadFinished() {
        if (startupAwaitingBackgroundLoadBeforeModalClose) {
            finishStartupSequenceProgressAndPrompt();
        }
        if (factorySwitchAwaitingBackgroundLoadBeforeModalClose) {
            updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_DONE);
            endFactorySiteSwitchBusy();
        }
        clearGlobalLongTaskProgress();
        refreshAttendanceReadiness();
        refreshStage1PipelineCheckGate();
        refreshGlobalStatusBar();
    }

    @Override
    public void setStartupTabBackgroundLoadActive(boolean active) {
        startupTabBackgroundLoadActive = active;
    }

    @Override
    public boolean isStartupTabBackgroundLoadActive() {
        return startupTabBackgroundLoadActive;
    }

    @Override
    public boolean canScheduleStartupBackgroundLoad() {
        return isStartupBackgroundLoadAllowed();
    }

    @Override
    public boolean canScheduleFactorySwitchBackgroundLoad() {
        return isFactorySwitchBackgroundLoadAllowed();
    }

    private String resolveGlobalStatusMessage() {
        if (isEnvVarsInitializationPending()) {
            return "環境変数の初期化が必要です。環境変数タブで確認してください。";
        }
        String base;
        if (startupBackgroundLoadMessage != null && !startupBackgroundLoadMessage.isBlank()) {
            base = startupBackgroundLoadMessage;
        } else if (globalLongTaskDetail != null && !globalLongTaskDetail.isBlank()) {
            base = "計画確認";
        } else {
            base = resolveActiveRunStageStatusMessage();
        }
        if (globalLongTaskDetail != null && !globalLongTaskDetail.isBlank()) {
            return base + " — " + globalLongTaskDetail;
        }
        return base;
    }

    private String resolveActiveRunStageStatusMessage() {
        String script = activeRunStageScript;
        if (STAGE1.equals(script)) {
            return "段階1 実行中…";
        }
        if (STAGE2.equals(script)) {
            return "段階2.0 実行中…";
        }
        if (STAGE2_1.equals(script)) {
            return "段階2.1 実行中…";
        }
        if (stage2SourceGuardCoordinator.isRunning()) {
            return "配台パイプライン実行中…";
        }
        if (lastGlobalLogLine != null && !lastGlobalLogLine.isBlank()) {
            return lastGlobalLogLine;
        }
        return "準備完了";
    }

    private void applyEnvVarsInitToolbarGating(boolean pending) {
        if (themeCombo != null) {
            themeCombo.setDisable(pending);
        }
        if (dispatchUsageGuideButton != null) {
            dispatchUsageGuideButton.setDisable(pending);
        }
        if (envTabController != null) {
            envTabController.setEnvInitAttention(pending);
        }
    }

    private void applyGuestSessionToolbarGating(boolean guestOnly) {
        if (guestOnly) {
            if (themeCombo != null) {
                themeCombo.setDisable(true);
            }
            if (dispatchUsageGuideButton != null) {
                dispatchUsageGuideButton.setDisable(true);
            }
        }
    }

    /** 環境変数 ui_ref 既定への初期化が未記録のときのみ操作可能なメインシェル葉タブ。 */
    private boolean isMainShellLeafOperableDuringEnvVarsInit(Tab t) {
        return t == mainShellTabEnv;
    }

    /**
     * 「環境変数」葉タブを確実に選択する（初期化未記録時の操作制限用）。
     */
    private void ensureMainShellEnvTabSelected() {
        Tab envLeaf = mainShellTabFor(MainShellTabId.ENV);
        if (tabPane == null || envLeaf == null) {
            return;
        }
        Runnable select =
                () -> {
                    if (!selectShellTabLeaf(envLeaf)) {
                        selectMainShellTabRecursive(tabPane, MainShellTabId.ENV);
                    }
                    Tab effective =
                            resolveEffectiveLeafTab(
                                    tabPane.getSelectionModel().getSelectedItem());
                    if (effective != envLeaf) {
                        selectShellTabLeaf(envLeaf);
                    }
                    if (!suppressLazyMainShellTabContentSwap.get()) {
                        activateMainShellTabHeavyContentRecursive(
                                tabPane.getSelectionModel().getSelectedItem());
                    }
                    MainShellRunTabGating.enableOperableSubtree(envLeaf);
                    lastEffectiveShellLeaf = envLeaf;
                };
        if (Platform.isFxApplicationThread()) {
            select.run();
            Platform.runLater(select);
        } else {
            Platform.runLater(select);
        }
    }

    /** 段階1／2.0～3.2／配台試行 実行中も切り替え・操作を許可するメインシェル最上段タブ。 */
    private boolean isMainShellLeafOperableDuringPipelineRun(Tab t) {
        return t == mainShellTabRun
                || t == mainShellTabSpecialRules
                || t == mainShellTabRequestFormInput
                || t == mainShellTabRemoteDesktop;
    }

    /** 段階スクリプト／配台試行のいずれかが {@link #runLock} 保持中なら true。 */
    private boolean isPipelineRunLocked() {
        return runLock.get() || stage2SourceGuardCoordinator.isRunning();
    }

    private boolean blockIfPipelineRunLocked(String stageJa) {
        if (!isPipelineRunLocked()) {
            return false;
        }
        appendLog("[busy] 他の処理が実行中のため " + stageJa + " を開始できません。");
        return true;
    }

    private boolean blockIfStage2SourceGuardBusy(String stageJa) {
        if (stage2SourceGuardCoordinator.allowsRelatedStart()) {
            return false;
        }
        appendLog("[busy] 固定ソース確認中のため " + stageJa + " を開始できません。");
        return true;
    }

    void refreshStage1PipelineCheckGate() {
        if (mainRunTabController == null) {
            return;
        }
        RequestFormPipelineCheckTabController.Stage1GateStatus status =
                requestFormPipelineCheckTabController != null
                        ? requestFormPipelineCheckTabController.evaluateStage1Gate()
                        : RequestFormPipelineCheckTabController.Stage1GateStatus.blocked(
                                "原本転記・計画確認で「更新」を実行し、問題の有無を確認してください。",
                                "原本転記: 未走査");
        mainRunTabController.setStage1BlockedByPipelineCheck(
                !status.permitted(), status.message(), status.badgeMessage());
    }

    private boolean blockIfPipelineCheckBlocksStage1() {
        RequestFormPipelineCheckTabController.Stage1GateStatus status =
                requestFormPipelineCheckTabController != null
                        ? requestFormPipelineCheckTabController.evaluateStage1Gate()
                        : RequestFormPipelineCheckTabController.Stage1GateStatus.blocked(
                                "原本転記・計画確認で「更新」を実行し、問題の有無を確認してください。",
                                "原本転記: 未走査");
        if (status.permitted()) {
            return false;
        }
        appendLog("[stage1] " + status.message());
        showWarningDialog("段階1", status.message());
        return true;
    }

    /**
     * メインウィンドウ上部ツールバーに段階1/2／配台試行 実行中を表示する。
     * プログレスは {@link DispatchInteractiveTabController} の「機械 JSON 再読み」と同じ
     * {@link ProgressIndicator}（22×22）+ {@link ProgressBar}（prefWidth 220・不定）の組み合わせ。
     */
    private void updateShellStageProgressOverlay(
            String script, PipelineExecutionTimingKind dispatchTrialKind) {
        if (shellStageProgressBox == null) {
            return;
        }
        boolean stageScriptBusy = pipelineTimingKindForStageScript(script) != null;
        boolean dispatchTrialBusy = dispatchTrialKind != null;
        boolean show = stageScriptBusy || dispatchTrialBusy;
        if (show
                && stageRunBusyDialog != null
                && stageRunBusyDialog.isShowing()
                && (STAGE1.equals(script) || STAGE2.equals(script))) {
            show = false;
        }
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
                if (STAGE1.equals(script)) {
                    shellStageProgressLabel.setText("段階1 実行中…");
                } else if (STAGE2.equals(script)) {
                    shellStageProgressLabel.setText("段階2.0 実行中…");
                } else if (STAGE2_1.equals(script)) {
                    shellStageProgressLabel.setText("段階2.1 実行中…");
                }
            }
            if (shellStageCancelButton != null) {
                boolean showCancel = stageScriptBusy;
                shellStageCancelButton.setManaged(showCancel);
                shellStageCancelButton.setVisible(showCancel);
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

    /**
     * 段階1～3.5 正常終了の通知。OK 待ちで後続処理を止めないため {@link Alert#show()} を使う。
     * 納期管理ビュー再読込・サマリ Excel 等は呼び出し側が本ダイアログより先に開始すること。
     */
    private void showStageCompletionDialog(String title, String contentText) {
        Alert alert = new Alert(AlertType.INFORMATION);
        initDialogOwnerIfSceneReady(alert);
        applyAlertStylesheetsFromOwner(alert);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(contentText);
        alert.show();
    }

    /**
     * 段階2 完了など: 納期管理ビュー更新後に OK 待ちし、閉じたあと {@code afterOk} を実行する。
     */
    private void showStageCompletionDialogAndWait(
            String title, String headerText, String contentText, Runnable afterOk) {
        Alert alert = new Alert(AlertType.INFORMATION);
        initDialogOwnerIfSceneReady(alert);
        applyAlertStylesheetsFromOwner(alert);
        alert.setTitle(title);
        alert.setHeaderText(headerText);
        alert.setContentText(contentText);
        alert.showAndWait();
        if (afterOk != null) {
            afterOk.run();
        }
    }

    private void exportSharedAladdinEntryWorkbookAfterStage2(
            Consumer<ResultDispatchTableTabController.AladdinEntryExportOutcome> completion) {
        if (deliveryCalendarViewTabController != null
                && deliveryCalendarViewTabController.exportSharedAladdinEntryWorkbookAfterStage2(
                        completion)) {
            return;
        }
        if (resultDispatchTableTabController != null) {
            resultDispatchTableTabController.exportSharedAladdinEntryWorkbookAfterStage2(completion);
            return;
        }
        completion.accept(
                new ResultDispatchTableTabController.AladdinEntryExportOutcome(
                        null,
                        List.of(),
                        new IllegalStateException("配台結果画面を初期化できませんでした。")));
    }

    private void updateStage2ExcelProgress(
            ResultDispatchTableTabController.AladdinEntryExportOutcome outcome) {
        if (outcome != null && outcome.succeeded()) {
            Path latest = outcome.result().latestPath();
            String detail =
                    latest != null && latest.getFileName() != null
                            ? latest.getFileName().toString()
                            : "";
            mainRunTabController.updateStage2Progress(
                    MainRunStage2Progress.State.COMPLETED, detail);
            return;
        }
        Exception error = outcome != null ? outcome.error() : null;
        String detail =
                error != null && error.getMessage() != null
                        ? error.getMessage()
                        : "原因不明";
        mainRunTabController.updateStage2Progress(
                MainRunStage2Progress.State.FAILED, detail);
    }

    static String stage2CompletionHeader(
            ResultDispatchTableTabController.AladdinEntryExportOutcome outcome) {
        return outcome != null && outcome.succeeded()
                ? "アラジン入力用Excelを生成しました"
                : "Excel自動生成に失敗しました";
    }

    static String stage2CompletionContent(
            ResultDispatchTableTabController.AladdinEntryExportOutcome outcome) {
        StringBuilder text = new StringBuilder("段階2 の処理自体は正常終了しました。");
        if (outcome != null && outcome.succeeded()) {
            text.append("\n\n最新: ").append(outcome.result().latestPath());
            if (outcome.warnings() != null && !outcome.warnings().isEmpty()) {
                text.append("\n\n警告:\n").append(String.join("\n", outcome.warnings()));
            }
            return text.toString();
        }
        Exception error = outcome != null ? outcome.error() : null;
        String reason =
                error == null
                        ? "原因不明"
                        : error.getMessage() != null && !error.getMessage().isBlank()
                                ? error.getMessage()
                                : error.toString();
        return text.append("\n\nExcel自動生成の原因: ").append(reason).toString();
    }

    /** 配台試行 正常終了後: 完了音・配台タブへ切替・完了ダイアログ。 */
    void notifyDispatchTrialSuccess() {
        appendLog("[end] 配台試行 正常終了");
        refreshOperatorCardAfterPipelineRun();
        MacroCompleteChime.playIfAvailable(collectUiEnv());
        selectMainShellTab(MainShellTabId.DISPATCH_INTERACTIVE);
        showStageCompletionDialog("配台試行 完了", "配台試行の処理が正常終了しました。");
    }

    /** 配台試行 異常終了後: 実行・ログタブへ切替・失敗ダイアログ。 */
    void notifyDispatchTrialFailure(String detailMessage) {
        selectMainShellTab(MainShellTabId.RUN);
        Alert alert = new Alert(AlertType.ERROR);
        initDialogOwnerIfSceneReady(alert);
        applyAlertStylesheetsFromOwner(alert);
        alert.setTitle("配台試行 失敗");
        alert.setHeaderText(null);
        StringBuilder body = new StringBuilder();
        body.append("配台試行が異常終了しました。\n");
        if (detailMessage != null && !detailMessage.isBlank()) {
            body.append(detailMessage.trim()).append('\n');
        }
        body.append("\n詳細は「実行・ログ」タブのログを確認してください。");
        applyScrollableAlertBody(alert, body.toString());
        alert.showAndWait();
    }

    /**
     * 長文 Alert で OK ボタンが画面外に押し出されないよう、本文を ScrollPane 内に収める。
     */
    private void applyScrollableAlertBody(Alert alert, String bodyText) {
        TextArea area = new TextArea(bodyText != null ? bodyText : "");
        area.setEditable(false);
        area.setWrapText(true);
        area.setPrefRowCount(10);
        ScrollPane scroll = new ScrollPane(area);
        scroll.setFitToWidth(true);
        scroll.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        scroll.setVbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        double maxViewport = alertScrollMaxViewportHeight();
        scroll.setPrefViewportHeight(Math.min(320, maxViewport));
        scroll.setMaxHeight(maxViewport);
        alert.getDialogPane().setContent(scroll);
        alert.setResizable(true);
        alert.getDialogPane().setPrefWidth(680);
    }

    /** ダイアログのタイトル・ボタン行分を除いた ScrollPane 高さ上限。 */
    private static double alertScrollMaxViewportHeight() {
        Rectangle2D bounds =
                Screen.getPrimary() != null
                        ? Screen.getPrimary().getVisualBounds()
                        : new Rectangle2D(0, 0, 1280, 800);
        return Math.max(200, bounds.getHeight() * 0.42);
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
                            + "。");
        }
    }

    /** 子プロセス末尾ログから PlanningValidationError 相当の1行を拾う（段階2の勤怠終端エラー等）。 */
    private static String planningValidationDetailFromTail(List<String> tailLines) {
        if (tailLines == null || tailLines.isEmpty()) {
            return null;
        }
        for (int i = tailLines.size() - 1; i >= 0; i--) {
            String ln = tailLines.get(i);
            if (ln == null || ln.isBlank()) {
                continue;
            }
            String stripped = ln.startsWith("[child] ") ? ln.substring(8).trim() : ln.trim();
            if (stripped.contains("勤怠カレンダー")
                    || stripped.contains("配台しきれません")
                    || stripped.contains("必須列")
                    || stripped.contains("検証エラー")
                    || stripped.startsWith("段階1:")
                    || stripped.startsWith("段階2:")
                    || stripped.startsWith("段階3:")) {
                return stripped;
            }
        }
        return null;
    }

    private static String exitHintJa(int code) {
        return switch (code) {
            case 0 -> "正常終了しました。";
            case 1 -> "一般エラーです（データや設定の不整合など）。";
            case 2 -> "致命的エラー、またはマスタ・入力ファイルの欠如などです。";
            case 3 -> "計画データの検証エラーです（計画期間内に配台しきれない・必須列不足など）。";
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
        initDialogOwnerIfSceneReady(alert);
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
            String validationDetail = planningValidationDetailFromTail(tailLines);
            if (validationDetail != null) {
                body.append("\n\n【検証エラー詳細】\n").append(validationDetail);
            }
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
        applyScrollableAlertBody(alert, body.toString());
        alert.showAndWait();
    }

    /** メインウィンドウと同じテーマ CSS をダイアログに載せる。 */
    public void applyAlertStylesheets(Dialog<?> dialog) {
        applyAlertStylesheetsFromOwner(dialog);
    }

    /** カスタム Stage / Scene にメインテーマ CSS を適用する。 */
    public void applyStylesheetsToScene(javafx.scene.Scene scene) {
        if (primaryStage == null || scene == null) {
            return;
        }
        Scene ownerScene = primaryStage.getScene();
        if (ownerScene == null) {
            return;
        }
        for (String url : ownerScene.getStylesheets()) {
            if (!scene.getStylesheets().contains(url)) {
                scene.getStylesheets().add(url);
            }
        }
    }

    void onMemberAttendanceDirtyChanged(boolean dirty) {
        memberAttendanceDirtySinceSave = dirty;
        if (mainShellTabMemberAttendance != null) {
            mainShellTabMemberAttendance.setText(dirty ? "メンバー勤怠 *" : "メンバー勤怠");
        }
    }

    void onCompanyCalendarDirtyChanged(boolean dirty) {
        companyCalendarDirtySinceSave = dirty;
        if (mainShellTabCompanyCalendar != null) {
            mainShellTabCompanyCalendar.setText(dirty ? "会社カレンダー *" : "会社カレンダー");
        }
    }

    void onMachineCalendarDirtyChanged(boolean dirty) {
        machineCalendarDirtySinceSave = dirty;
        if (mainShellTabMachineCalendar != null) {
            mainShellTabMachineCalendar.setText(dirty ? "機械カレンダー *" : "機械カレンダー");
        }
    }

    /** 会社カレンダー正本更新後、機械カレンダータブのミニカレンダーを再読込する。 */
    void refreshMachineCalendarCompanyMiniCalendar() {
        if (machineCalendarTabController != null) {
            machineCalendarTabController.refreshCompanyCalendarMiniCalendar();
        }
    }

    /** 会社カレンダータブの会計年度ラベル（メンバー勤怠セットアップ等で共有）。 */
    public int attendanceFiscalYearLabel() {
        if (companyCalendarTabController != null) {
            return companyCalendarTabController.getFiscalYearLabel();
        }
        LocalDate today = LocalDate.now();
        return jp.co.pm.ai.desktop.ui.FiscalYearPeriod.fiscalYearLabelFor(
                today, jp.co.pm.ai.desktop.ui.FiscalYearPeriod.DEFAULT_APRIL_MARCH);
    }

    /** 会社カレンダータブの会計期間開始（未表示時は 4/1 既定）。 */
    public jp.co.pm.ai.desktop.ui.FiscalYearPeriod attendanceFiscalPeriod() {
        if (companyCalendarTabController != null) {
            return companyCalendarTabController.getFiscalPeriod();
        }
        return jp.co.pm.ai.desktop.ui.FiscalYearPeriod.DEFAULT_APRIL_MARCH;
    }

    private boolean confirmAttendanceTabsUnsavedBeforeLeave(String actionDescription) {
        if (!confirmCompanyCalendarUnsavedBeforeLeave(actionDescription)) {
            return false;
        }
        return confirmMemberAttendanceUnsavedBeforeLeave(actionDescription);
    }

    private boolean confirmMemberAttendanceUnsavedBeforeLeave(String actionDescription) {
        if (memberAttendanceTabController == null || !memberAttendanceTabController.hasUnsavedEdits()) {
            return true;
        }
        MemberAttendanceTabController.UnsavedPromptResult result =
                memberAttendanceTabController.promptUnsavedChanges(actionDescription);
        if (result == MemberAttendanceTabController.UnsavedPromptResult.CANCELLED) {
            return false;
        }
        if (result == MemberAttendanceTabController.UnsavedPromptResult.DISCARDED) {
            if ("終了".equals(actionDescription)) {
                memberAttendanceTabController.clearUnsavedWithoutReload();
            } else {
                memberAttendanceTabController.discardUnsavedEdits();
            }
            return true;
        }
        memberAttendanceTabController.saveEditsAsync(
                success -> {
                    if (!success) {
                        return;
                    }
                    Platform.runLater(
                            () -> {
                                if ("終了".equals(actionDescription)) {
                                    if (!confirmCompanyCalendarUnsavedBeforeLeave("終了")) {
                                        return;
                                    }
                                    if (primaryStage != null) {
                                        primaryStage.close();
                                    }
                                } else if (pendingMemberAttendanceTabAfterSave != null) {
                                    MainShellTabId target = pendingMemberAttendanceTabAfterSave;
                                    pendingMemberAttendanceTabAfterSave = null;
                                    suppressMemberAttendanceUnsavedTabGuard.set(true);
                                    try {
                                        selectMainShellTabRecursive(tabPane, target);
                                    } finally {
                                        suppressMemberAttendanceUnsavedTabGuard.set(false);
                                    }
                                }
                            });
                });
        return false;
    }

    private MainShellTabId pendingMemberAttendanceTabAfterSave = null;
    private MainShellTabId pendingCompanyCalendarTabAfterSave = null;

    private boolean confirmCompanyCalendarUnsavedBeforeLeave(String actionDescription) {
        if (companyCalendarTabController == null || !companyCalendarTabController.hasUnsavedEdits()) {
            return true;
        }
        CompanyCalendarTabController.UnsavedPromptResult result =
                companyCalendarTabController.promptUnsavedChanges(actionDescription);
        if (result == CompanyCalendarTabController.UnsavedPromptResult.CANCELLED) {
            return false;
        }
        if (result == CompanyCalendarTabController.UnsavedPromptResult.DISCARDED) {
            if ("終了".equals(actionDescription)) {
                companyCalendarTabController.clearUnsavedWithoutReload();
            } else {
                companyCalendarTabController.discardUnsavedEdits();
            }
            return true;
        }
        companyCalendarTabController.saveEditsAsync(
                success -> {
                    if (!success) {
                        return;
                    }
                    Platform.runLater(
                            () -> {
                                if ("終了".equals(actionDescription)) {
                                    if (!confirmMemberAttendanceUnsavedBeforeLeave("終了")) {
                                        return;
                                    }
                                    if (primaryStage != null) {
                                        primaryStage.close();
                                    }
                                } else if (pendingCompanyCalendarTabAfterSave != null) {
                                    MainShellTabId target = pendingCompanyCalendarTabAfterSave;
                                    pendingCompanyCalendarTabAfterSave = null;
                                    suppressCompanyCalendarUnsavedTabGuard.set(true);
                                    try {
                                        selectMainShellTabRecursive(tabPane, target);
                                    } finally {
                                        suppressCompanyCalendarUnsavedTabGuard.set(false);
                                    }
                                }
                            });
                });
        return false;
    }

    private boolean blockCompanyCalendarUnsavedTabNavigation(Tab prevTab, Tab newTab) {
        if (suppressCompanyCalendarUnsavedTabGuard.get()) {
            return false;
        }
        Tab prevLeaf = resolveEffectiveLeafTab(prevTab);
        Tab newLeaf = resolveEffectiveLeafTab(newTab);
        if (prevLeaf == null || newLeaf == null) {
            return false;
        }
        MainShellTabId prevId = mainShellTabId(prevLeaf);
        MainShellTabId newId = mainShellTabId(newLeaf);
        if (prevId != MainShellTabId.COMPANY_CALENDAR || newId == MainShellTabId.COMPANY_CALENDAR) {
            return false;
        }
        if (companyCalendarTabController == null || !companyCalendarTabController.hasUnsavedEdits()) {
            return false;
        }
        CompanyCalendarTabController.UnsavedPromptResult result =
                companyCalendarTabController.promptUnsavedChanges("タブを切り替える");
        if (result == CompanyCalendarTabController.UnsavedPromptResult.CANCELLED) {
            suppressCompanyCalendarUnsavedTabGuard.set(true);
            try {
                selectMainShellTabRecursive(tabPane, MainShellTabId.COMPANY_CALENDAR);
            } finally {
                suppressCompanyCalendarUnsavedTabGuard.set(false);
            }
            return true;
        }
        if (result == CompanyCalendarTabController.UnsavedPromptResult.DISCARDED) {
            companyCalendarTabController.discardUnsavedEdits();
            return false;
        }
        pendingCompanyCalendarTabAfterSave = newId;
        suppressCompanyCalendarUnsavedTabGuard.set(true);
        try {
            selectMainShellTabRecursive(tabPane, MainShellTabId.COMPANY_CALENDAR);
        } finally {
            suppressCompanyCalendarUnsavedTabGuard.set(false);
        }
        companyCalendarTabController.saveEditsAsync(
                success -> {
                    if (!success) {
                        pendingCompanyCalendarTabAfterSave = null;
                        return;
                    }
                    Platform.runLater(
                            () -> {
                                if (pendingCompanyCalendarTabAfterSave != null) {
                                    MainShellTabId target = pendingCompanyCalendarTabAfterSave;
                                    pendingCompanyCalendarTabAfterSave = null;
                                    suppressCompanyCalendarUnsavedTabGuard.set(true);
                                    try {
                                        selectMainShellTabRecursive(tabPane, target);
                                    } finally {
                                        suppressCompanyCalendarUnsavedTabGuard.set(false);
                                    }
                                }
                            });
                });
        return true;
    }

    private boolean blockCompanyCalendarUnsavedInnerTabNavigation(
            Tab prevInner, Tab newInner, TabPane innerPane) {
        if (suppressCompanyCalendarUnsavedTabGuard.get()) {
            return false;
        }
        MainShellTabId prevId = mainShellTabId(prevInner);
        MainShellTabId newId = mainShellTabId(newInner);
        if (prevId != MainShellTabId.COMPANY_CALENDAR || newId == MainShellTabId.COMPANY_CALENDAR) {
            return false;
        }
        if (companyCalendarTabController == null || !companyCalendarTabController.hasUnsavedEdits()) {
            return false;
        }
        CompanyCalendarTabController.UnsavedPromptResult result =
                companyCalendarTabController.promptUnsavedChanges("タブを切り替える");
        if (result == CompanyCalendarTabController.UnsavedPromptResult.CANCELLED) {
            suppressCompanyCalendarUnsavedTabGuard.set(true);
            try {
                innerPane.getSelectionModel().select(prevInner);
            } finally {
                suppressCompanyCalendarUnsavedTabGuard.set(false);
            }
            return true;
        }
        if (result == CompanyCalendarTabController.UnsavedPromptResult.DISCARDED) {
            companyCalendarTabController.discardUnsavedEdits();
            return false;
        }
        suppressCompanyCalendarUnsavedTabGuard.set(true);
        try {
            innerPane.getSelectionModel().select(prevInner);
        } finally {
            suppressCompanyCalendarUnsavedTabGuard.set(false);
        }
        final MainShellTabId targetId = newId;
        companyCalendarTabController.saveEditsAsync(
                success -> {
                    if (!success) {
                        return;
                    }
                    Platform.runLater(
                            () -> {
                                suppressCompanyCalendarUnsavedTabGuard.set(true);
                                try {
                                    innerPane.getSelectionModel().select(tabForMainShellTabId(targetId));
                                } finally {
                                    suppressCompanyCalendarUnsavedTabGuard.set(false);
                                }
                            });
                });
        return true;
    }

    private boolean blockMemberAttendanceUnsavedTabNavigation(Tab prevTab, Tab newTab) {
        if (suppressMemberAttendanceUnsavedTabGuard.get()) {
            return false;
        }
        Tab prevLeaf = resolveEffectiveLeafTab(prevTab);
        Tab newLeaf = resolveEffectiveLeafTab(newTab);
        if (prevLeaf == null || newLeaf == null) {
            return false;
        }
        MainShellTabId prevId = mainShellTabId(prevLeaf);
        MainShellTabId newId = mainShellTabId(newLeaf);
        if (prevId != MainShellTabId.MEMBER_ATTENDANCE || newId == MainShellTabId.MEMBER_ATTENDANCE) {
            return false;
        }
        if (memberAttendanceTabController == null || !memberAttendanceTabController.hasUnsavedEdits()) {
            return false;
        }
        MemberAttendanceTabController.UnsavedPromptResult result =
                memberAttendanceTabController.promptUnsavedChanges("タブを切り替える");
        if (result == MemberAttendanceTabController.UnsavedPromptResult.CANCELLED) {
            suppressMemberAttendanceUnsavedTabGuard.set(true);
            try {
                selectMainShellTabRecursive(tabPane, MainShellTabId.MEMBER_ATTENDANCE);
            } finally {
                suppressMemberAttendanceUnsavedTabGuard.set(false);
            }
            return true;
        }
        if (result == MemberAttendanceTabController.UnsavedPromptResult.DISCARDED) {
            memberAttendanceTabController.discardUnsavedEdits();
            return false;
        }
        pendingMemberAttendanceTabAfterSave = newId;
        suppressMemberAttendanceUnsavedTabGuard.set(true);
        try {
            selectMainShellTabRecursive(tabPane, MainShellTabId.MEMBER_ATTENDANCE);
        } finally {
            suppressMemberAttendanceUnsavedTabGuard.set(false);
        }
        memberAttendanceTabController.saveEditsAsync(
                success -> {
                    if (!success) {
                        pendingMemberAttendanceTabAfterSave = null;
                        return;
                    }
                    Platform.runLater(
                            () -> {
                                if (pendingMemberAttendanceTabAfterSave != null) {
                                    MainShellTabId target = pendingMemberAttendanceTabAfterSave;
                                    pendingMemberAttendanceTabAfterSave = null;
                                    suppressMemberAttendanceUnsavedTabGuard.set(true);
                                    try {
                                        selectMainShellTabRecursive(tabPane, target);
                                    } finally {
                                        suppressMemberAttendanceUnsavedTabGuard.set(false);
                                    }
                                }
                            });
                });
        return true;
    }

    private boolean blockMemberAttendanceUnsavedInnerTabNavigation(
            Tab prevInner, Tab newInner, TabPane innerPane) {
        if (suppressMemberAttendanceUnsavedTabGuard.get()) {
            return false;
        }
        MainShellTabId prevId = mainShellTabId(prevInner);
        MainShellTabId newId = mainShellTabId(newInner);
        if (prevId != MainShellTabId.MEMBER_ATTENDANCE || newId == MainShellTabId.MEMBER_ATTENDANCE) {
            return false;
        }
        if (memberAttendanceTabController == null || !memberAttendanceTabController.hasUnsavedEdits()) {
            return false;
        }
        MemberAttendanceTabController.UnsavedPromptResult result =
                memberAttendanceTabController.promptUnsavedChanges("タブを切り替える");
        if (result == MemberAttendanceTabController.UnsavedPromptResult.CANCELLED) {
            suppressMemberAttendanceUnsavedTabGuard.set(true);
            try {
                innerPane.getSelectionModel().select(prevInner);
            } finally {
                suppressMemberAttendanceUnsavedTabGuard.set(false);
            }
            return true;
        }
        if (result == MemberAttendanceTabController.UnsavedPromptResult.DISCARDED) {
            memberAttendanceTabController.discardUnsavedEdits();
            return false;
        }
        suppressMemberAttendanceUnsavedTabGuard.set(true);
        try {
            innerPane.getSelectionModel().select(prevInner);
        } finally {
            suppressMemberAttendanceUnsavedTabGuard.set(false);
        }
        final MainShellTabId targetId = newId;
        memberAttendanceTabController.saveEditsAsync(
                success -> {
                    if (!success) {
                        return;
                    }
                    Platform.runLater(
                            () -> {
                                suppressMemberAttendanceUnsavedTabGuard.set(true);
                                try {
                                    innerPane.getSelectionModel().select(tabForMainShellTabId(targetId));
                                } finally {
                                    suppressMemberAttendanceUnsavedTabGuard.set(false);
                                }
                            });
                });
        return true;
    }

    private Tab tabForMainShellTabId(MainShellTabId id) {
        if (id == null) {
            return null;
        }
        return mainShellTabFor(id);
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
                AppPaths.resolvePythonScriptDirForScript(
                        uiRun, MasterReadSummaryTabController.scriptName());
        String wb = effectiveTaskInputWorkbookPath();
        return new RunRequest(
                py,
                dir,
                MasterReadSummaryTabController.scriptName(),
                wb,
                childEnvForPython(uiRun));
    }

    /** {@code machine_calendar_io.py} — 機械カレンダー JSON。 */
    RunRequest buildMachineCalendarIoRequest(String... scriptArgs) {
        Map<String, String> uiRun = collectUiEnv();
        Path py = resolveStagePythonExecutablePath(uiRun);
        String scriptName = "machine_calendar_io.py";
        Path dir = AppPaths.resolvePythonScriptDirForScript(uiRun, scriptName);
        String wb = effectiveTaskInputWorkbookPath();
        return new RunRequest(
                py,
                dir,
                scriptName,
                wb,
                childEnvForPython(uiRun),
                List.of(scriptArgs));
    }

    /** {@code attendance_data_io.py} — 勤怠 JSON / 会社カレンダー / master 新シート出力。 */
    RunRequest buildAttendanceDataIoRequest(String... scriptArgs) {
        Map<String, String> uiRun = collectUiEnv();
        Path py = resolveStagePythonExecutablePath(uiRun);
        String scriptName = "attendance_data_io.py";
        Path dir = AppPaths.resolvePythonScriptDirForScript(uiRun, scriptName);
        String wb = effectiveTaskInputWorkbookPath();
        return new RunRequest(
                py,
                dir,
                scriptName,
                wb,
                childEnvForPython(uiRun),
                List.of(scriptArgs));
    }

    private static final ObjectMapper ATTENDANCE_JSON = new ObjectMapper();

    private int attendanceGridCellSizePx = AttendanceGridCellSizing.DEFAULT_CELL_PX;
    private boolean tableRowHoverDimmingEnabled =
            DesktopSessionState.DEFAULT_TABLE_ROW_HOVER_DIMMING_ENABLED;
    private volatile boolean attendanceStage2Ready = false;
    private volatile boolean attendanceReadinessResolved = false;
    private volatile String attendanceReadinessTooltip = "";

    /**
     * {@code attendance_data_io.py} を非同期実行し、stdout 末尾 JSON をコールバックへ渡す。
     */
    public void runAttendanceDataIoAsync(
            PythonProcessRunner.RunRequest req,
            Consumer<JsonNode> onOk,
            Consumer<String> onError) {
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            if (err != null) {
                                                if (onError != null) {
                                                    onError.accept(err.getMessage());
                                                }
                                                return;
                                            }
                                            if (cap == null) {
                                                if (onError != null) {
                                                    onError.accept("プロセス結果なし");
                                                }
                                                return;
                                            }
                                            try {
                                                JsonNode node =
                                                        ATTENDANCE_JSON.readTree(
                                                                AttendanceOvertimePreview
                                                                        .MasterReadSummaryJson
                                                                        .extractLastJsonLine(
                                                                                cap.stdout()));
                                                if (!node.path("ok").asBoolean(false)) {
                                                    if (onError != null) {
                                                        onError.accept(
                                                                node.path("error")
                                                                        .asText("失敗"));
                                                    }
                                                    markAttendanceReadinessUnknown();
                                                    return;
                                                }
                                                if (cap.exitCode() != 0) {
                                                    if (onError != null) {
                                                        onError.accept(
                                                                "exit=" + cap.exitCode());
                                                    }
                                                    markAttendanceReadinessUnknown();
                                                    return;
                                                }
                                                if (onOk != null) {
                                                    onOk.accept(node);
                                                }
                                            } catch (Exception e) {
                                                if (onError != null) {
                                                    onError.accept(e.getMessage());
                                                }
                                                markAttendanceReadinessUnknown();
                                            }
                                        }));
    }

    /**
     * {@code machine_calendar_io.py} を非同期実行し、stdout 末尾 JSON をコールバックへ渡す。
     */
    public void runMachineCalendarIoAsync(
            PythonProcessRunner.RunRequest req,
            Consumer<JsonNode> onOk,
            Consumer<String> onError) {
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            if (err != null) {
                                                if (onError != null) {
                                                    onError.accept(err.getMessage());
                                                }
                                                return;
                                            }
                                            if (cap == null) {
                                                if (onError != null) {
                                                    onError.accept("プロセス結果なし");
                                                }
                                                return;
                                            }
                                            try {
                                                JsonNode node =
                                                        ATTENDANCE_JSON.readTree(
                                                                AttendanceOvertimePreview
                                                                        .MasterReadSummaryJson
                                                                        .extractLastJsonLine(
                                                                                cap.stdout()));
                                                if (!node.path("ok").asBoolean(false)) {
                                                    if (onError != null) {
                                                        onError.accept(
                                                                node.path("error")
                                                                        .asText("失敗"));
                                                    }
                                                    return;
                                                }
                                                if (cap.exitCode() != 0) {
                                                    if (onError != null) {
                                                        onError.accept(
                                                                "exit=" + cap.exitCode());
                                                    }
                                                    return;
                                                }
                                                if (onOk != null) {
                                                    onOk.accept(node);
                                                }
                                            } catch (Exception e) {
                                                if (onError != null) {
                                                    onError.accept(e.getMessage());
                                                }
                                            }
                                        }));
    }

    public int attendanceGridCellSizePx() {
        return attendanceGridCellSizePx;
    }

    public boolean tableRowHoverDimmingEnabled() {
        return tableRowHoverDimmingEnabled;
    }

    public void setTableRowHoverDimmingEnabled(boolean enabled) {
        if (tableRowHoverDimmingEnabled == enabled) {
            return;
        }
        tableRowHoverDimmingEnabled = enabled;
        UiRowHoverDimmingSettings.setEnabled(enabled);
        refreshTableRowHoverDimmingPresentation();
    }

    public void persistGlobalDesktopSession() {
        DesktopSessionStateStore.save(collectDesktopSessionForGlobalPersistence());
    }

    private void refreshTableRowHoverDimmingPresentation() {
        if (memberAttendanceTabController != null) {
            memberAttendanceTabController.refreshRowHoverDimming();
        }
        if (machineCalendarTabController != null) {
            machineCalendarTabController.refreshRowHoverDimming();
        }
    }

  /** 会社カレンダー・メンバー勤怠のセルグリッド寸法（px）。両タブで共有。 */
    public void setAttendanceGridCellSizePx(int px) {
        int clamped = AttendanceGridCellSizing.clamp(px);
        if (attendanceGridCellSizePx == clamped) {
            return;
        }
        attendanceGridCellSizePx = clamped;
        if (companyCalendarTabController != null) {
            companyCalendarTabController.applyGridCellSizeToPane(clamped);
            companyCalendarTabController.syncGridCellSizeSpinner(clamped);
        }
        if (memberAttendanceTabController != null) {
            memberAttendanceTabController.applyGridCellSizeToPane(clamped);
            memberAttendanceTabController.syncGridCellSizeSpinner(clamped);
        }
        if (machineCalendarTabController != null) {
            machineCalendarTabController.applyGridCellSize(clamped);
            machineCalendarTabController.syncGridCellSizeSpinner(clamped);
        }
    }

    public void refreshAttendanceReadiness() {
        refreshAttendanceReadiness(false);
    }

    private void refreshAttendanceReadiness(boolean force) {
        if (!force && !isStartupBackgroundLoadAllowed()) {
            return;
        }
        LocalDate today = LocalDate.now();
        runAttendanceDataIoAsync(
                buildAttendanceDataIoRequest(
                        "readiness",
                        Integer.toString(today.getYear()),
                        Integer.toString(today.getMonthValue())),
                this::applyAttendanceReadinessFromJson,
                msg -> {
                    appendLog("[attendance-readiness] " + msg);
                    markAttendanceReadinessUnknown();
                });
    }

    private void markAttendanceReadinessUnknown() {
        attendanceStage2Ready = false;
        attendanceReadinessResolved = true;
        attendanceReadinessTooltip = "勤怠状態の取得に失敗しました。再読込してください。";
        applyAttendanceReadinessFromJson(null);
    }

    public void applyAttendanceReadinessFromJson(JsonNode node) {
        if (node == null) {
            String blockTooltip = attendanceReadinessTooltip;
            PersonBadgeStyle badgeStyle = PersonBadgeStyle.networkSourceCacheBadgeDefault();
            if (mainRunTabController != null) {
                mainRunTabController.setAttendanceReadinessBadge(
                        true, badgeStyle, "勤怠未確認", blockTooltip);
            }
            if (mainRunTabController != null) {
                mainRunTabController.setCalendarReadinessBlocked(true, blockTooltip);
            }
            if (planInputTabController != null) {
                planInputTabController.setAttendanceReadinessBadge(
                        true, "勤怠未確認", blockTooltip);
                planInputTabController.setAttendanceReadinessBlocked(true, blockTooltip);
            }
            refreshGlobalStatusBar();
            return;
        }
        attendanceReadinessResolved = true;
        attendanceStage2Ready = node.path("stage2_ready").asBoolean(false);
        StringBuilder issues = new StringBuilder();
        if (node.path("issues").isArray()) {
            for (JsonNode issue : node.path("issues")) {
                if (issues.length() > 0) {
                    issues.append('\n');
                }
                issues.append(issue.asText(""));
            }
        }
        attendanceReadinessTooltip = issues.toString();
        String badgeLabel = attendanceStage2Ready ? "" : "カレンダー未準備";
        String blockTooltip =
                attendanceReadinessTooltip.isBlank()
                        ? "カレンダー正本 JSON が未準備です。"
                            + " attendance-data.json（会社カレンダー・メンバー勤怠）と"
                            + " machine-calendar-data.json をセットアップしてください。"
                        : attendanceReadinessTooltip;
        PersonBadgeStyle badgeStyle = PersonBadgeStyle.networkSourceCacheBadgeDefault();
        if (mainRunTabController != null) {
            mainRunTabController.setAttendanceReadinessBadge(
                    !attendanceStage2Ready,
                    badgeStyle,
                    badgeLabel,
                    blockTooltip);
            mainRunTabController.setCalendarReadinessBlocked(
                    !attendanceStage2Ready, blockTooltip);
        }
        if (planInputTabController != null) {
            planInputTabController.setAttendanceReadinessBadge(
                    !attendanceStage2Ready, badgeLabel, blockTooltip);
            planInputTabController.setAttendanceReadinessBlocked(
                    !attendanceStage2Ready, blockTooltip);
        }
        refreshGlobalStatusBar();
    }

    private boolean blockIfPlanningStagesCalendarNotReady(String stageLabel) {
        if (!attendanceReadinessResolved) {
            String msg = "勤怠・カレンダーの準備状態を確認中です。数秒待ってから再実行してください。";
            appendLog("[" + stageLabel + "] " + msg);
            refreshAttendanceReadiness();
            showErrorDialog(stageLabel, msg);
            return true;
        }
        if (attendanceStage2Ready) {
            return false;
        }
        String msg =
                attendanceReadinessTooltip.isBlank()
                        ? "カレンダー正本 JSON が未準備です。"
                            + " attendance-data.json（会社カレンダー・メンバー勤怠）と"
                            + " machine-calendar-data.json をセットアップしてください。"
                        : attendanceReadinessTooltip;
        appendLog("[" + stageLabel + "] " + msg.replace('\n', ' '));
        showErrorDialog(stageLabel, msg.replace('\n', '\n'));
        return true;
    }

    /** @deprecated 呼び出し元互換。段階2専用ブロック。 */
    private boolean blockIfAttendanceNotReadyForStage2() {
        return blockIfPlanningStagesCalendarNotReady("段階2");
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

    /** 配台計画_タスク入力タブの段階2用 Gemini スキップを子プロセス環境へ反映する（段階2/2.1/3 試行で段階1設定を上書き）。 */
    private void overlayPlanInputStage2SkipGeminiApiEnv(Map<String, String> ui) {
        if (planInputTabController == null) {
            return;
        }
        ui.put(
                AppPaths.KEY_PM_AI_SKIP_GEMINI_API,
                planInputTabController.snapshotStage2SkipGeminiApi() ? "1" : "0");
    }

    /** 実行・ログタブ「その他」のチェックを子プロセス環境へ反映する。 */
    private void overlayMainRunSkipGeminiApiEnv(Map<String, String> ui) {
        if (mainRunTabController == null) {
            return;
        }
        ui.put(
                AppPaths.KEY_PM_AI_SKIP_GEMINI_API,
                mainRunTabController.snapshotSkipGeminiApi() ? "1" : "0");
    }

    /** 配台計画_タスク入力タブの組み合わせ表 need 超過チェックを子プロセス環境へ反映する。 */
    private void overlayPlanInputComboSheetMayExceedNeedEnv(Map<String, String> ui) {
        if (planInputTabController == null) {
            return;
        }
        ui.put(
                AppPaths.KEY_TEAM_ASSIGN_COMBO_SHEET_MAY_EXCEED_NEED,
                planInputTabController.snapshotComboSheetMayExceedNeed() ? "1" : "0");
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
        overlayMainRunSkipGeminiApiEnv(m);
        Stage2PythonChildEnv.stripLegacyWorkbookKeys(m);
        Stage2PythonChildEnv.ensureSkipWorkbookEnvSheetDefault(m);
        overlayPlanInputTabPathsForStage2(m);
        lastNetworkSourceResolution =
                Stage2PythonChildEnv.applyNetworkSourceAndChildPause(
                        m,
                        startupSkipTaskInputSourceDirListing,
                        startupSkipActualDetailSourceDirListing);
        AgentDebugLog.overlayPythonChildDebugEnv(m);
        overlayAttendancePathsForPython(m);
        return m;
    }

    /** 勤怠 JSON を工場既定サマリ同階層へ補完（空文字の env 行は上書き）。 */
    private void overlayAttendancePathsForPython(Map<String, String> m) {
        if (nz(m.get(AppPaths.KEY_PM_AI_ATTENDANCE_JSON)).isBlank()) {
            m.put(
                    AppPaths.KEY_PM_AI_ATTENDANCE_JSON,
                    AppPaths.attendanceDataJsonPath(m).toString());
        }
        if (nz(m.get(AppPaths.KEY_PM_AI_ATTENDANCE_CALENDAR_XLSX)).isBlank()) {
            m.put(
                    AppPaths.KEY_PM_AI_ATTENDANCE_CALENDAR_XLSX,
                    AppPaths.attendanceCalendarXlsxPath(m).toString());
        }
        if (nz(m.get(AppPaths.KEY_PM_AI_MACHINE_CALENDAR_JSON)).isBlank()) {
            m.put(
                    AppPaths.KEY_PM_AI_MACHINE_CALENDAR_JSON,
                    AppPaths.machineCalendarDataJsonPath(m).toString());
        }
    }

    /**
     * Fills {@link PlanInputTabController#ENV_PM_AI_PLAN_INPUT_PATH} and {@link
     * PlanInputTabController#ENV_TASK_PLAN_SHEET} from the dedicated plan-input tab when the env tab
     * leaves them blank. When the plan path is {@code plan_input_tasks.xlsx}, the tab sheet name
     * always wins over env {@code TASK_PLAN_SHEET} (stage-1 output uses {@link AppPaths#STAGE1_PLAN_OUTPUT_SHEET}).
     */
    private void overlayPlanInputTabPathsForStage2(Map<String, String> m) {
        overlayPlanInputTabPathsIfEnvBlank(m);
        String pipKey = PlanInputTabController.ENV_PM_AI_PLAN_INPUT_PATH;
        String pip = m.get(pipKey);
        if (!isDedicatedStage1PlanTasksWorkbook(pip)) {
            return;
        }
        String tpsKey = PlanInputTabController.ENV_TASK_PLAN_SHEET;
        String tabSheet = planInputTabController.snapshotPlanInputSheet();
        if (tabSheet != null && !tabSheet.isBlank()) {
            m.put(tpsKey, tabSheet.trim());
            return;
        }
        m.put(tpsKey, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
    }

    private static boolean isDedicatedStage1PlanTasksWorkbook(String path) {
        if (path == null || path.isBlank()) {
            return false;
        }
        try {
            return Path.of(path.trim())
                    .getFileName()
                    .toString()
                    .equalsIgnoreCase(AppPaths.STAGE1_PLAN_TASKS_FILENAME);
        } catch (Exception ignored) {
            return false;
        }
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
    private void applyFactorySitePortableAndNetworkDefaults(FactorySite site) {
        if (envRows == null || site == null) {
            return;
        }
        String task = site.taskInputSourceDir();
        String actual = site.actualDetailSourceDir();
        String portable = site.portableBundleSourceDir();
        String pmAiMaster = site.pmAiMasterWorkbookEnvValue(collectUiEnv());
        String pmAiSummary = site.pmAiSummaryAiDispatchWorkbookEnvValue(collectUiEnv());
        for (EnvVarRow r : envRows) {
            String name = r.getName() != null ? r.getName().trim() : "";
            if (AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR.equals(name)) {
                r.setValue(task);
            } else if (AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR.equals(name)) {
                r.setValue(actual);
            } else if (AppPaths.KEY_PM_AI_DAILY_REPORT_SOURCE_DIR.equals(name)) {
                r.setValue(site.dailyReportSourceDir());
            } else if (AppPaths.KEY_PM_AI_ORDER_DETAIL_SOURCE_DIR.equals(name)) {
                r.setValue(site.orderDetailSourceDir());
            } else if (AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR.equals(name)) {
                r.setValue(portable);
            } else if (AppPaths.KEY_PM_AI_MASTER_WORKBOOK.equals(name)) {
                r.setValue(pmAiMaster != null ? pmAiMaster : "");
            } else if (AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK.equals(name)) {
                r.setValue(pmAiSummary != null ? pmAiSummary : "");
            } else if (AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR.equals(name)) {
                r.setValue(site.aladdinMasterDir());
            } else if (AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE.equals(name)) {
                r.setValue(site.requestFormJuchuFile());
            } else if (AppPaths.KEY_PM_AI_MACHINE_DELIVERY_MANAGEMENT_XLSM.equals(name)) {
                r.setValue(site.machineDeliveryManagementXlsm());
            } else if (AppPaths.KEY_PM_AI_REQUEST_FORM_TPI_PDF_DIR.equals(name)) {
                r.setValue(site.requestFormTpiPdfDir());
            } else if (AppPaths.KEY_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR.equals(name)) {
                r.setValue(site.rdpPortableBundleSourceDir());
            } else if (AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR.equals(name)) {
                r.setValue(site.rdpLauncherDeployDir());
            } else if (AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR.equals(name)) {
                r.setValue(site.rdpOperatorUsersStoreDir());
            } else if (AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR.equals(name)) {
                r.setValue(site.rpaLauncherDeployDir());
            } else if (AppPaths.KEY_PM_AI_RPA_LAUNCHER_OPERATOR_USERS_STORE_DIR.equals(name)) {
                r.setValue(site.rpaLauncherOperatorUsersStoreDir());
            } else if (AppPaths.KEY_PM_AI_FACTORY_SITE.equals(name)) {
                r.setValue(site.name());
            }
        }
        GlobalInitSettingTarget.save(site);
        syncFactorySiteAttendanceEnvRows(site);
        syncFactorySiteRequestFormEnvRows(site);
        if (globalSettingsTabController != null) {
            globalSettingsTabController.refreshInitSettingTargetComboFromStore();
        }
        if (mainRunTabController != null) {
            factoryOperatorToolbar.refreshFactorySiteLogo();
            mainRunTabController.refreshFactorySiteComboFromStore();
        }
    }

    /** 工場別マスタ／サマリ更新後に勤怠 JSON・xlsx の env 行を同期する。 */
    private void syncFactorySiteAttendanceEnvRows(FactorySite site) {
        if (envRows == null || site == null) {
            return;
        }
        Map<String, String> ui = new LinkedHashMap<>(collectUiEnv());
        AppPaths.overlayFactorySiteAttendancePaths(ui, site);
        for (EnvVarRow row : envRows) {
            String name = row.getName() != null ? row.getName().trim() : "";
            if (ui.containsKey(name)) {
                row.setValue(ui.get(name));
            }
        }
    }

    /** 工場切替後に依頼書原本・受注ファイル等の env 行を当該工場へ揃える。 */
    private void syncFactorySiteRequestFormEnvRows(FactorySite site) {
        if (envRows == null || site == null) {
            return;
        }
        Map<String, String> ui = new LinkedHashMap<>(collectUiEnv());
        AppPaths.overlayFactorySiteRequestFormPaths(ui, site);
        for (EnvVarRow row : envRows) {
            String name = row.getName() != null ? row.getName().trim() : "";
            if (ui.containsKey(name)) {
                row.setValue(ui.get(name));
            }
        }
    }

    /** デスクトップ本体の終了後更新の直前: ユーザーへ再起動を明示する。 */
    private void showPortableUpgradeDeferredRestartDialog(String versionLabel) {
        Alert a = new Alert(AlertType.INFORMATION);
        initDialogOwnerIfSceneReady(a);
        applyAlertStylesheetsFromOwner(a);
        a.setTitle("自動バージョンアップ");
        a.setHeaderText("再起動します");
        String ver = versionLabel != null && !versionLabel.isBlank() ? versionLabel : "（新版）";
        a.setContentText(
                "pm-ai-data の更新が完了しました（版 "
                        + ver
                        + "）。\n\n"
                        + "PMD.exe とアプリケーション本体（app・runtime）を反映するため、"
                        + "このウィンドウを終了し、自動的に再起動します。\n\n"
                        + "再起動後、利用工場の設定を維持したまま環境を初期化します。"
                        + "操作者選択は表示しません（前回選択を復元します）。\n"
                        + "OK を押して続行してください。");
        a.showAndWait();
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
     * 納期管理ビューのデータ再読み込み中は段階1～段階3.5 の実行ボタンを無効化する。
     *
     * @param blocking {@code true} で無効化、{@code false} で通常の可否判定へ戻す
     */
    void setDeliveryCalendarReloadBlockingStageRuns(boolean blocking) {
        deliveryCalendarReloadBlockingStageRuns.set(blocking);
        if (mainRunTabController != null) {
            mainRunTabController.setDeliveryCalendarReloadBlocking(blocking);
        }
        if (planInputTabController != null) {
            planInputTabController.setDeliveryCalendarReloadBlocking(blocking);
        }
        if (dispatchInteractiveTabController != null) {
            dispatchInteractiveTabController.setDeliveryCalendarReloadBlocking(blocking);
        }
    }

    boolean isDeliveryCalendarReloadBlockingStageRuns() {
        return deliveryCalendarReloadBlockingStageRuns.get();
    }

    /**
     * メインシェルのタブを ID で選択する（配台試行ウィザードなどから）。
     */
    public void selectMainShellTab(MainShellTabId id) {
        if (tabPane == null || id == null) {
            return;
        }
        if (isEnvVarsInitializationPending() && id != MainShellTabId.ENV) {
            ensureMainShellEnvTabSelected();
            return;
        }
        if (FactoryOperatorUserStore.isGuestSession() && id != MainShellTabId.RUN) {
            ensureMainShellRunTabSelected();
            return;
        }
        if (id == MainShellTabId.RUN) {
            ensureMainShellRunTabSelected();
            return;
        }
        selectMainShellTabRecursive(tabPane, id);
    }

    /**
     * 「実行・ログ」葉タブを確実に選択する。タブ整理のグループ化後や {@link #rebuildMainShellTabsFromLayout}
     * 直後は 1 フレーム遅れて再試行する（無効化時の自動遷移でリモート等へ寄った場合も戻す）。
     */
    private void ensureMainShellRunTabSelected() {
        if (tabPane == null || mainShellTabRun == null) {
            return;
        }
        Runnable select =
                () -> {
                    if (!selectShellTabLeaf(mainShellTabRun)) {
                        selectMainShellTabRecursive(tabPane, MainShellTabId.RUN);
                    }
                    Tab effective =
                            resolveEffectiveLeafTab(
                                    tabPane.getSelectionModel().getSelectedItem());
                    if (effective != mainShellTabRun) {
                        selectShellTabLeaf(mainShellTabRun);
                    }
                    if (!suppressLazyMainShellTabContentSwap.get()) {
                        activateMainShellTabHeavyContentRecursive(
                                tabPane.getSelectionModel().getSelectedItem());
                    }
                    lastEffectiveShellLeaf = mainShellTabRun;
                };
        if (Platform.isFxApplicationThread()) {
            select.run();
            Platform.runLater(select);
        } else {
            Platform.runLater(select);
        }
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
    @Override
    public void appendLog(String line) {
        if (line != null && !line.isBlank()) {
            lastGlobalLogLine = line;
        }
        mainRunTabController.appendLog(line);
        refreshGlobalStatusBar();
    }

    void beginPipelineExecutionTiming(PipelineExecutionTimingKind kind) {
        pipelineExecutionTimingHistory.begin(kind);
    }

    void endPipelineExecutionTiming(PipelineExecutionTimingKind kind) {
        pipelineExecutionTimingHistory.end(kind);
    }

    PipelineExecutionTimingHistoryStore pipelineExecutionTimingHistory() {
        return pipelineExecutionTimingHistory;
    }

    /** グローバル設定の工場切替などで実行・ログタブ上部ロゴを更新する。 */
    void refreshMainRunTabFactoryLogo() {
        if (mainRunTabController != null) {
            factoryOperatorToolbar.refreshFactorySiteLogo();
        }
    }

    /**
     * 配台デスクトップの利用工場を切り替える。工場別ワークスペース（env + session 断片）を save→restore する。
     * 保存 env 行が無い場合は ui_ref 既定へ初期化する。
     */
    public void switchActiveFactorySite(FactorySite newSite) {
        switchActiveFactorySite(newSite, false);
    }

    private void switchActiveFactorySite(FactorySite newSite, boolean startup) {
        if (newSite == null || newSite == FactorySite.RDP_LAUNCHER) {
            return;
        }
        FactorySite oldSite = GlobalInitSettingTarget.load();
        if (oldSite == newSite) {
            return;
        }
        if (factorySiteSwitchInProgress) {
            return;
        }
        Map<String, String> ui = collectUiEnv();
        FactoryOperatorUserStore.configureForCurrentApp(ui, newSite);
        if (!FactorySiteOperatorAccess.isSessionOperatorAllowedForFactory(ui, newSite)) {
            String reason = FactorySiteOperatorAccess.comboBlockReasonJa(ui, newSite);
            if (!reason.isBlank()) {
                appendLog("[factory] 切替不可: " + reason);
            }
            refreshFactorySiteComboPresentation();
            if (mainRunTabController != null) {
                factoryOperatorToolbar.refreshFactorySiteComboFromStore();
            }
            if (globalSettingsTabController != null) {
                globalSettingsTabController.refreshInitSettingTargetComboFromStore();
            }
            return;
        }
        factorySiteSwitchInProgress = true;
        factorySwitchAwaitingBackgroundLoadBeforeModalClose = false;
        setFactorySiteCombosDisabled(true);
        if (startupTabBackgroundLoad != null) {
            startupTabBackgroundLoad.cancelForFactorySwitch();
        }
        beginFactorySiteSwitchBusy(oldSite, newSite);
        FactorySiteSwitchContext ctx = new FactorySiteSwitchContext(oldSite, newSite, startup);
        runAfterUiPulse(() -> runFactorySiteSwitchStep(ctx, 0));
    }

    private static final class FactorySiteSwitchContext {
        private final FactorySite oldSite;
        private final FactorySite newSite;
        private final boolean startup;
        private final long t0Nanos;
        private Optional<FactorySiteWorkspaceSnapshot> loaded = Optional.empty();

        private FactorySiteSwitchContext(FactorySite oldSite, FactorySite newSite, boolean startup) {
            this.oldSite = oldSite;
            this.newSite = newSite;
            this.startup = startup;
            this.t0Nanos = System.nanoTime();
        }

        private FactorySite oldSite() {
            return oldSite;
        }

        private FactorySite newSite() {
            return newSite;
        }

        private boolean startup() {
            return startup;
        }

        private long t0Nanos() {
            return t0Nanos;
        }

        private Optional<FactorySiteWorkspaceSnapshot> loaded() {
            return loaded;
        }

        private void setLoaded(Optional<FactorySiteWorkspaceSnapshot> value) {
            loaded = value != null ? value : Optional.empty();
        }
    }

    private void runFactorySiteSwitchStep(FactorySiteSwitchContext ctx, int step) {
        try {
            switch (step) {
                case 0 -> {
                    GlobalInitSettingTarget.setSuppressUiEnvInferencePersist(true);
                    updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_SAVING);
                    String operator = FactoryOperatorUserStore.sessionOperatorName();
                    if (!operator.isBlank()
                            && ctx.oldSite() != null
                            && ctx.oldSite() != FactorySite.RDP_LAUNCHER) {
                        FactorySiteWorkspaceStore.save(
                                operator, ctx.oldSite(), buildFactorySiteWorkspaceSnapshot());
                        FactorySiteWorkspaceStore.flushMemoryCacheToDisk(operator);
                    }
                    GlobalInitSettingTarget.save(ctx.newSite());
                    LastLaunchedFactorySiteStore.save(ctx.newSite());
                    if (!operator.isBlank()) {
                        FactorySiteWorkspaceStore.saveLastFactorySite(operator, ctx.newSite());
                    }
                    runAfterUiPulse(() -> runFactorySiteSwitchStep(ctx, 1));
                }
                case 1 -> {
                    updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_LOADING);
                    String operator = FactoryOperatorUserStore.sessionOperatorName();
                    Optional<FactorySiteWorkspaceSnapshot> loaded =
                            operator.isBlank()
                                    ? Optional.empty()
                                    : FactorySiteWorkspaceStore.load(operator, ctx.newSite());
                    ctx.setLoaded(loaded);
                    runAfterUiPulse(() -> runFactorySiteSwitchStep(ctx, 2));
                }
                case 2 -> {
                    updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_ENV);
                    applyFactorySiteWorkspaceRestore(ctx.newSite(), ctx.loaded());
                    runAfterUiPulse(() -> runFactorySiteSwitchStep(ctx, 3));
                }
                case 3 -> {
                    updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_REFRESH);
                    refreshFactoryDependentTabs(ctx.newSite(), true);
                    schedulePersistSessionDebounced();
                    runAfterUiPulse(() -> runFactorySiteSwitchStep(ctx, 4));
                }
                case 4 -> {
                    updateFactorySiteSwitchBusy(FactorySiteSwitchBusyDialog.STATUS_OPERATOR);
                    runAfterUiPulse(() -> runFactorySiteSwitchStep(ctx, 5));
                }
                case 5 -> {
                    completeEnvVarsStartupCheck(!(ctx.startup() && startupSequenceActive));
                    long ms = (System.nanoTime() - ctx.t0Nanos()) / 1_000_000L;
                    appendLog(
                            "[factory] 切替完了 "
                                    + ctx.oldSite().displayLabelJa()
                                    + "→"
                                    + ctx.newSite().displayLabelJa()
                                    + " ms="
                                    + ms);
                    runAfterUiPulse(() -> finishFactorySiteSwitch(ctx.newSite(), ctx.startup()));
                }
                default -> finishFactorySiteSwitch(null, false);
            }
        } catch (RuntimeException ex) {
            if (startupSequenceActive) {
                finishStartupSequenceProgressAndPrompt();
            } else {
                finishFactorySiteSwitch(null, false);
            }
            throw ex;
        }
    }

    /** 工場切替の進行中（モーダル表示前後を含む）。タブ再走査の抑止に使う。 */
    boolean isFactorySiteSwitchInProgress() {
        return factorySiteSwitchInProgress;
    }

    /**
     * 工場切替モーダルをいったん閉じ、操作者確認のあとタブ再読込中に再度表示する。
     *
     * <p>操作者ダイアログは進捗モーダルと重ねると FX スレッドが詰まるため、確認の直前だけ閉じる。
     * タブ再読込（リモート・カレンダー等）は起動時と同様に進捗モーダルを維持する。
     */
    private void finishFactorySiteSwitch(FactorySite newSite, boolean startup) {
        GlobalInitSettingTarget.setSuppressUiEnvInferencePersist(false);
        endFactorySiteSwitchBusy();
        factorySiteSwitchInProgress = false;
        setFactorySiteCombosDisabled(false);
        FactorySite site = newSite != null ? newSite : GlobalInitSettingTarget.load();
        if (site != null) {
            factorySwitchBusyTo = site;
        }
        if (!FactoryOperatorUserStore.isGuestSession() && !(startup && startupSequenceActive)) {
            requireOperatorSelectionForFactory(site, startup);
        }
        reloadAttendanceTabsFromJson(true);
        runAfterUiPulse(
                () -> {
                    notifyActiveMainShellTabAfterWorkspaceChange();
                    if (startupSequenceActive) {
                        finishStartupSequenceAfterEnvCheck();
                    } else if (startupTabBackgroundLoad != null) {
                        beginFactorySiteSwitchTabLoadBusy();
                        factorySwitchAwaitingBackgroundLoadBeforeModalClose = true;
                        startupTabBackgroundLoad.resetAndScheduleAfterFactorySwitch();
                        if (!FactorySiteSwitchBusySupport.keepBusyDialogForPostSwitchTabLoad(
                                startupSequenceActive, startupTabBackgroundLoadActive)) {
                            endFactorySiteSwitchBusy();
                        }
                    }
                });
    }

    @Override
    public Map<String, String> snapshotUiEnv() {
        Map<String, String> base = collectUiEnv();
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        if (operator.isBlank()) {
            return base;
        }
        Map<String, String> merged = new HashMap<>(base);
        merged.put(AppPaths.KEY_PM_AI_OPERATOR_USER, operator);
        return Map.copyOf(merged);
    }

    @Override
    public void refreshOperatorUserPresentation() {
        refreshMainRunTabOperatorLabel();
        refreshRemoteDesktopOperatorContext();
        refreshGlobalStatusBar();
    }

    @Override
    public void refreshRemoteDesktopOperatorContext() {
        if (remoteDesktopTabController != null) {
            remoteDesktopTabController.refreshForSessionOperatorChange();
        }
    }

    /** 実行・ログタブの操作者表示を更新する。 */
    void refreshMainRunTabOperatorLabel() {
        if (mainRunTabController != null) {
            factoryOperatorToolbar.refreshOperatorUserLabel();
        }
        if (operatorUserManagementTabController != null) {
            operatorUserManagementTabController.refreshPresentationQuietly();
        }
        if (requestFormInputTabController != null) {
            requestFormInputTabController.refreshSessionInputTantoLabel();
        }
    }

    /**
     * 工場別の操作者名を選択させる。キャンセル不可（起動時・工場切替時）。
     *
     * @param startup true のとき起動直後の案内文
     */
    @Override
    public void requireOperatorSelectionForFactory(FactorySite site, boolean startup) {
        OperatorUserSelectionSupport.requireOperatorSelectionForFactory(this, site, startup);
    }

    @Override
    public void changeSessionOperator(FactorySite site) {
        if (FactoryOperatorUserStore.isGuestSession()) {
            showWarningDialog(
                    "操作不可",
                    "ゲスト操作者は工場切替のみ利用できます。\n"
                            + "登録操作者で利用するには、アプリを再起動して操作者を選び直してください。");
            return;
        }
        OperatorUserSelectionSupport.changeSessionOperator(this, site);
        scheduleDesktopSessionSave();
        applyRunTabGating();
    }

    /** 実行・ログタブなどから、現在の工場向けに操作者を変更する。 */
    public void changeSessionOperator() {
        changeSessionOperator(GlobalInitSettingTarget.load());
    }

    @FXML
    private void onShellChangeSessionOperatorAction() {
        changeSessionOperator();
    }

    @FXML
    private void onShellChangeOperatorPinAction() {
        promptChangeSessionOperatorPin();
    }

    private Optional<String> promptAndVerifyOperatorPin(FactorySite factory, String operatorName) {
        if (primaryStage == null) {
            return Optional.empty();
        }
        try {
            if (FactoryOperatorUserStore.isPinLocked(factory, operatorName)) {
                showWarningDialog(
                        "PIN ロック",
                        "操作者「"
                                + operatorName
                                + "」は PIN を "
                                + FactoryOperatorUserStore.MAX_CONSECUTIVE_PIN_FAILURES
                                + " 回連続で間違えたためロックされています。\n"
                                + "ユーザー管理者タブでロック解除または PIN 再発行してください。");
                return Optional.empty();
            }
        } catch (IOException ex) {
            showWarningDialog("PIN", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return Optional.empty();
        }
        Dialog<String> dialog = new Dialog<>();
        prepareDialogForMainTheme(dialog);
        dialog.setTitle("PIN 認証");
        dialog.setHeaderText(null);
        Label hint =
                new Label(
                        "操作者「"
                                + operatorName
                                + "」の PIN（"
                                + FactoryOperatorUserStore.pinLengthRangeDescriptionJa()
                                + "）を入力してください。");
        hint.setWrapText(true);
        PasswordField pf = new PasswordField();
        pf.setPromptText("PIN");
        VBox box = new VBox(8, hint, new Label("PIN:"), pf);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        focusInputWhenDialogShown(dialog, pf);
        dialog.setResultConverter(
                bt -> {
                    if (bt != ButtonType.OK) {
                        return null;
                    }
                    String t = pf.getText();
                    return t != null ? t.strip() : "";
                });
        while (true) {
            Optional<String> ans = dialog.showAndWait();
            if (ans.isEmpty()) {
                return Optional.empty();
            }
            String pin = ans.get();
            if (FactoryOperatorUserStore.normalizePin(pin) == null) {
                showWarningDialog(
                        "PIN", FactoryOperatorUserStore.pinLengthRangeDescriptionJa() + "を入力してください。");
                continue;
            }
            try {
                FactoryOperatorUserStore.PinVerificationResult result =
                        FactoryOperatorUserStore.verifyPinAttempt(factory, operatorName, pin);
                switch (result) {
                    case SUCCESS -> {
                        return Optional.of(pin);
                    }
                    case LOCKED -> {
                        showWarningDialog(
                                "PIN ロック",
                                "操作者「"
                                        + operatorName
                                        + "」は PIN を "
                                        + FactoryOperatorUserStore.MAX_CONSECUTIVE_PIN_FAILURES
                                        + " 回連続で間違えたためロックされました。\n"
                                        + "ユーザー管理者タブでロック解除または PIN 再発行してください。");
                        return Optional.empty();
                    }
                    case WRONG_PIN -> {
                        int remaining = FactoryOperatorUserStore.remainingPinAttempts(factory, operatorName);
                        showWarningDialog(
                                "PIN",
                                remaining > 0
                                        ? "PIN が正しくありません。残り "
                                                + remaining
                                                + " 回でロックされます。"
                                        : "PIN が正しくありません。");
                    }
                    default -> showWarningDialog("PIN", "PIN が正しくありません。");
                }
            } catch (IOException ex) {
                showWarningDialog("PIN", ex.getMessage() != null ? ex.getMessage() : ex.toString());
                return Optional.empty();
            }
        }
    }

    /** ランダム初期 PIN でログイン後、初回のみ PIN 変更を強制する。 */
    private boolean promptRequiredInitialPinChange(
            FactorySite factory, String operatorName, String currentPin) {
        if (primaryStage == null) {
            return false;
        }
        Dialog<ButtonType> dialog = new Dialog<>();
        prepareDialogForMainTheme(dialog);
        dialog.setTitle("初回 PIN 変更（必須）");
        dialog.setHeaderText(null);
        Label hint =
                new Label(
                        "操作者「"
                                + operatorName
                                + "」は初回ログインのため、新しい PIN（"
                                + FactoryOperatorUserStore.pinLengthRangeDescriptionJa()
                                + "）を設定してください。");
        hint.setWrapText(true);
        PasswordField newPf = new PasswordField();
        newPf.setPromptText("新しい PIN");
        PasswordField confirmPf = new PasswordField();
        confirmPf.setPromptText("新しい PIN（確認）");
        VBox box = new VBox(8, hint, new Label("新しい PIN:"), newPf, new Label("確認:"), confirmPf);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK);
        focusInputWhenDialogShown(dialog, newPf);
        while (true) {
            Optional<ButtonType> ans = dialog.showAndWait();
            if (ans.isEmpty() || ans.get() != ButtonType.OK) {
                return false;
            }
            String newPin = newPf.getText() != null ? newPf.getText().strip() : "";
            String confirmPin = confirmPf.getText() != null ? confirmPf.getText().strip() : "";
            if (!newPin.equals(confirmPin)) {
                showWarningDialog("初回 PIN 変更", "新しい PIN と確認入力が一致しません。");
                continue;
            }
            if (FactoryOperatorUserStore.normalizePin(newPin) == null) {
                showWarningDialog(
                        "初回 PIN 変更",
                        "新しい PIN は "
                                + FactoryOperatorUserStore.pinLengthRangeDescriptionJa()
                                + " です。");
                continue;
            }
            try {
                FactoryOperatorUserStore.changePinOnFirstLogin(factory, operatorName, currentPin, newPin);
                appendLog("[operator-user] 初回 PIN を変更しました: " + operatorName);
                showInformationDialog("初回 PIN 変更", "PIN を変更しました。ログインを続行します。");
                return true;
            } catch (Exception ex) {
                showWarningDialog("初回 PIN 変更", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
        }
    }

    /** 実行・ログタブなどから、ログイン中操作者の PIN 変更ダイアログを開く。 */
    void promptChangeSessionOperatorPin() {
        if (primaryStage == null) {
            return;
        }
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        if (operator.isBlank()) {
            showWarningDialog("PIN 変更", "操作者が未選択です。先に操作者名を選んでください。");
            return;
        }
        if (FactoryOperatorUserStore.isGuestSession()) {
            showWarningDialog("PIN 変更", "ゲストユーザーは PIN を設定できません。");
            return;
        }
        FactorySite factory = GlobalInitSettingTarget.load();
        boolean hasPin;
        try {
            hasPin = FactoryOperatorUserStore.hasPin(factory, operator);
            if (FactoryOperatorUserStore.isPinLocked(factory, operator)) {
                showWarningDialog(
                        "PIN ロック",
                        "PIN がロックされています。ユーザー管理者タブでロック解除してください。");
                return;
            }
        } catch (IOException ex) {
            showWarningDialog("PIN 変更", ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return;
        }

        Dialog<ButtonType> dialog = new Dialog<>();
        prepareDialogForMainTheme(dialog);
        dialog.setTitle(hasPin ? "PIN 変更" : "PIN 設定");
        dialog.setHeaderText(null);
        Label hint =
                new Label(
                        "操作者「"
                                + operator
                                + "」の PIN（"
                                + FactoryOperatorUserStore.pinLengthRangeDescriptionJa()
                                + "）を"
                                + (hasPin ? "変更" : "設定")
                                + "します。");
        hint.setWrapText(true);
        PasswordField currentPf = new PasswordField();
        currentPf.setPromptText("現在の PIN");
        PasswordField newPf = new PasswordField();
        newPf.setPromptText("新しい PIN");
        PasswordField confirmPf = new PasswordField();
        confirmPf.setPromptText("新しい PIN（確認）");
        VBox box;
        if (hasPin) {
            box = new VBox(8, hint, new Label("現在の PIN:"), currentPf, new Label("新しい PIN:"), newPf, new Label("確認:"), confirmPf);
        } else {
            box = new VBox(8, hint, new Label("新しい PIN:"), newPf, new Label("確認:"), confirmPf);
        }
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        focusInputWhenDialogShown(dialog, hasPin ? currentPf : newPf);
        Optional<ButtonType> ans = dialog.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        String newPin = newPf.getText() != null ? newPf.getText().strip() : "";
        String confirmPin = confirmPf.getText() != null ? confirmPf.getText().strip() : "";
        if (!newPin.equals(confirmPin)) {
            showWarningDialog("PIN 変更", "新しい PIN と確認入力が一致しません。");
            return;
        }
        if (FactoryOperatorUserStore.normalizePin(newPin) == null) {
            showWarningDialog(
                    "PIN 変更",
                    "新しい PIN は " + FactoryOperatorUserStore.pinLengthRangeDescriptionJa() + " です。");
            return;
        }
        String currentPin = hasPin && currentPf.getText() != null ? currentPf.getText().strip() : "";
        try {
            FactoryOperatorUserStore.changePinByUser(factory, operator, currentPin, newPin);
            appendLog("[operator-user] PIN を" + (hasPin ? "変更" : "設定") + "しました: " + operator);
            showInformationDialog("PIN 変更", "PIN を" + (hasPin ? "変更" : "設定") + "しました。");
        } catch (Exception ex) {
            showWarningDialog("PIN 変更", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    private void maybePromptOperatorUserAtStartup() {
        if (skipOperatorPromptAfterPortableUpgrade.compareAndSet(true, false)) {
            startupSequenceActive = true;
            beginEnvVarsStartupCheckBusy(EnvVarsStartupCheckBusyDialog.STATUS_STABILIZE);
            runAfterUiPulse(
                    () -> {
                        try {
                            completeEnvVarsStartupCheck(false);
                            finishStartupSequenceAfterEnvCheck();
                        } catch (RuntimeException ex) {
                            finishStartupSequenceProgressAndPrompt();
                            throw ex;
                        }
                    });
            return;
        }
        if (deferOperatorPromptForPortableUpgrade.get()) {
            return;
        }
        FactorySite factory = StartupFactorySiteResolver.resolve();
        GlobalInitSettingTarget.save(factory);
        FactoryOperatorUserStore.configureFromUi(collectUiEnv(), factory);
        if (FactoryOperatorUserStore.usingLocalStoreFallback()) {
            appendLog(
                    "[startup] 操作者設定: 共有フォルダに書き込めないためローカルに退避しています（"
                            + FactoryOperatorUserStore.storePath()
                            + "）。グローバル設定の工場と環境変数のサマリパスを確認してください。");
        }
        requireOperatorSelectionForFactory(factory, true);
        runOperatorStartupWorkspaceAndEnvCheckWithProgress();
    }

    private Optional<String> promptOperatorUserChoice(FactorySite site, boolean startup) {
        if (primaryStage == null) {
            return Optional.empty();
        }
        List<String> names;
        try {
            names = FactoryOperatorUserStore.loginChoicesForFactory(site);
        } catch (IOException ex) {
            names = new ArrayList<>(FactoryOperatorUserStore.DEFAULT_NAMES);
            if (!names.contains(FactoryOperatorUserStore.GUEST_OPERATOR_NAME)) {
                names.add(FactoryOperatorUserStore.GUEST_OPERATOR_NAME);
            }
        }
        String pref;
        try {
            pref = FactoryOperatorUserStore.lastSelectedForFactory(site);
        } catch (IOException ex) {
            pref = "";
        }
        if (pref.isBlank() || !names.contains(pref)) {
            pref = names.get(0);
        }
        ChoiceDialog<String> d = new ChoiceDialog<>(pref, names);
        prepareDialogForMainTheme(d);
        d.setTitle("操作者名の選択");
        d.setHeaderText(null);
        d.setContentText(
                (startup ? "配台システムを利用する操作者名を選んでください。\n" : "")
                        + "工場: "
                        + site.displayLabelJa()
                        + "\n（作成者表示に使用します。一覧の編集はユーザー管理者タブから行えます。）\n"
                        + "「"
                        + FactoryOperatorUserStore.GUEST_OPERATOR_NAME
                        + "」は PIN 不要ですが、サマリ Excel は生成できません。");
        return d.showAndWait();
    }

    /** {@code jp.co.pm.ai.desktop.dispatch.rules} 子タブ向け。 */
    public Map<String, String> dispatchRulesUiEnv() {
        return snapshotUiEnv();
    }

    /** {@code jp.co.pm.ai.desktop.dispatch.rules} 子タブ向け。 */
    public void dispatchRulesAppendLog(String line) {
        appendLog(line);
    }

    /** 環境変数タブ「配台 Gemini モデル優先」から無料枠 Flash-Lite 一覧を即時再取得する。 */
    public void requestGeminiFreeTierModelsForceRefresh() {
        if (geminiFreeTierModelsRefreshService != null) {
            geminiFreeTierModelsRefreshService.refreshNow(true);
        }
    }

    /** 環境変数の GEMINI_MODEL / TRY_ORDER 変更を API モデルベンチマークタブへ反映する。 */
    public void refreshApiModelBenchmarkDerivedLabels() {
        if (apiModelBenchmarkTabController != null) {
            apiModelBenchmarkTabController.refreshShellDerivedLabels();
        }
    }

    private void onGeminiFreeTierModelsRefreshFinished(
            GeminiFreeTierModelsRefreshService.RefreshResult result) {
        javafx.application.Platform.runLater(
                () -> {
                    if (result != null && result.success() && !result.modelIds().isEmpty()) {
                        applyGeminiFreeTierModelsToTryOrderEnvIfUnpinned(result.modelIds());
                    }
                    if (result != null && result.message() != null && !result.message().isBlank()) {
                        appendLog("[gemini-free-tier] " + result.message());
                    }
                    if (envTabController != null && result != null) {
                        envTabController.onGeminiFreeTierRefreshCompleted(result);
                    }
                });
    }

    /**
     * {@code GEMINI_MODEL} 未設定時のみ、日次／手動更新で得た試行列を環境変数表へ反映する。
     */
    private void applyGeminiFreeTierModelsToTryOrderEnvIfUnpinned(List<String> modelIds) {
        if (envRows == null || modelIds == null || modelIds.isEmpty()) {
            return;
        }
        boolean pinned = false;
        for (EnvVarRow r : envRows) {
            String n = r.getName() != null ? r.getName().strip() : "";
            if ("GEMINI_MODEL".equals(n)) {
                String v = r.getValue();
                pinned = v != null && !v.isBlank();
                break;
            }
        }
        if (pinned) {
            return;
        }
        EnvVarRow tryRow = null;
        for (EnvVarRow r : envRows) {
            String n = r.getName() != null ? r.getName().strip() : "";
            if ("GEMINI_MODEL_TRY_ORDER".equals(n)) {
                tryRow = r;
                break;
            }
        }
        if (tryRow == null) {
            tryRow = new EnvVarRow();
            tryRow.setName("GEMINI_MODEL_TRY_ORDER");
            tryRow.setDescription(
                    "カンマ区切りで試行順。GEMINI_MODEL 未設定時のみ有効。日次更新で Flash-Lite 無料枠候補を自動反映。");
            envRows.add(tryRow);
        }
        tryRow.setValue(
                String.join(
                        ",",
                        GeminiDispatchModelTryOrderDefaults.withPlanningCorePriorityFirst(
                                modelIds)));
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
        Map<String, String> ui = new HashMap<>(collectUiEnv());
        overlayMainRunSkipGeminiApiEnv(ui);
        return childEnvForPython(ui);
    }

    /**
     * Environment for {@code dispatch_interactive_trial.py}: same stage-2 overrides as {@link #runStage}:
     * {@link AppPaths#KEY_PM_AI_STAGE2_WRITE_EXCEL}（常に 1）と {@link AppPaths#KEY_PM_AI_RESULT_BOOK_FONT} from the run tab,
     * {@link AppPaths#KEY_PM_AI_STAGE2_SKIP_TODAY_DISPATCH} from the plan-input tab, and
     * {@link AppPaths#KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH} は常に無効（加工途中はタスク入力タブで翌日配台量を設定）。
     */
    public Map<String, String> snapshotDispatchTrialPythonEnv() {
        return snapshotDispatchTrialPythonEnv(null, null);
    }

    /**
     * {@link #snapshotDispatchTrialPythonEnv()} に加え、段階3.5 残業シミュレーション JSON パスを載せる。
     */
    public Map<String, String> snapshotDispatchTrialPythonEnv(java.nio.file.Path overtimeSimulationJson) {
        return snapshotDispatchTrialPythonEnv(overtimeSimulationJson, null);
    }

    /** 段階3 配台試行用の子プロセス環境（残業 JSON は段階2.1 では使わない）。 */
    public Map<String, String> snapshotDispatchTrialPythonEnv(
            java.nio.file.Path overtimeSimulationJson, java.nio.file.Path unusedFloorJson) {
        Map<String, String> ui = new HashMap<>(collectUiEnv());
        ui.put(AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL, "1");
        ui.put(
                AppPaths.KEY_PM_AI_STAGE2_SKIP_TODAY_DISPATCH,
                mainRunTabController.snapshotStage2SkipTodayDispatch() ? "1" : "0");
        ui.put(AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH, "0");
        applyStage2NextDayDialogEnvs(ui);
        overlayPlanInputComboSheetMayExceedNeedEnv(ui);
        overlayMainRunSkipGeminiApiEnv(ui);
        overlayPlanInputStage2SkipGeminiApiEnv(ui);
        String resultFont = mainRunTabController.snapshotStage2ResultBookFont();
        if (resultFont != null && !resultFont.isBlank()) {
            ui.put(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT, resultFont.trim());
        } else {
            ui.remove(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT);
        }
        ui.remove(AppPaths.KEY_PM_AI_OVERTIME_SIMULATION_JSON);
        ui.remove(AppPaths.KEY_PM_AI_STAGE2_1_OVERTIME);
        return childEnvForPython(ui);
    }

    /** 段階2.1: 残業シミュ付きフル再配台。成果物は {@code output/stage21/} へ。 */
    public Map<String, String> snapshotStage21PythonEnv(java.nio.file.Path overtimeSimulationJson) {
        Map<String, String> ui = new HashMap<>(collectUiEnv());
        ui.put(AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL, "1");
        ui.put(
                AppPaths.KEY_PM_AI_STAGE2_SKIP_TODAY_DISPATCH,
                mainRunTabController.snapshotStage2SkipTodayDispatch() ? "1" : "0");
        ui.put(AppPaths.KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH, "0");
        applyStage2NextDayDialogEnvs(ui);
        overlayPlanInputComboSheetMayExceedNeedEnv(ui);
        overlayMainRunSkipGeminiApiEnv(ui);
        overlayPlanInputStage2SkipGeminiApiEnv(ui);
        String resultFont = mainRunTabController.snapshotStage2ResultBookFont();
        if (resultFont != null && !resultFont.isBlank()) {
            ui.put(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT, resultFont.trim());
        } else {
            ui.remove(AppPaths.KEY_PM_AI_RESULT_BOOK_FONT);
        }
        java.nio.file.Path stage21Dir =
                AppPaths.resolveStage21OutputDir(ui).toAbsolutePath().normalize();
        ui.put(AppPaths.KEY_PM_AI_OUTPUT_DIR, stage21Dir.toString());
        ui.put(AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR, stage21Dir.toString());
        ui.put(AppPaths.KEY_PM_AI_STAGE2_1_OVERTIME, "1");
        if (overtimeSimulationJson != null) {
            ui.put(
                    AppPaths.KEY_PM_AI_OVERTIME_SIMULATION_JSON,
                    overtimeSimulationJson.toAbsolutePath().normalize().toString());
        } else {
            ui.remove(AppPaths.KEY_PM_AI_OVERTIME_SIMULATION_JSON);
        }
        return childEnvForPython(ui);
    }

    /** 段階2.1: 残業シミュレーション上書き JSON を stage21 出力ディレクトリへ書く。 */
    public java.nio.file.Path writeStage21OvertimeSimulationOverridesJson(
            jp.co.pm.ai.desktop.dispatch.OvertimeSimulationOverridesWriter.OverridesPayload payload)
            throws Exception {
        java.nio.file.Path dir = AppPaths.resolveStage21OutputDir(collectUiEnv());
        java.nio.file.Files.createDirectories(dir);
        java.nio.file.Path target = dir.resolve("overtime_simulation_overrides.json");
        jp.co.pm.ai.desktop.dispatch.OvertimeSimulationOverridesWriter.write(target, payload);
        appendLog("[stage2.1] 残業シミュレーション JSON: " + target.toAbsolutePath().normalize());
        return target;
    }

    /** 段階2.1 正常終了後: 勤怠適用サマリ付き完了通知。 */
    void notifyStage21OvertimeSimulationSuccess() {
        notifyStage21OvertimeSimulationSuccess(null);
    }

    /** 段階2.1 正常終了後: 勤怠適用サマリ付き完了通知。 */
    void notifyStage21OvertimeSimulationSuccess(
            jp.co.pm.ai.desktop.dispatch.Stage21OutputPromoter.Result promoted) {
        appendLog("[end] 段階2.1（残業/休出シミュ）正常終了");
        MacroCompleteChime.playIfAvailable(collectUiEnv());
        selectMainShellTab(MainShellTabId.DISPATCH_INTERACTIVE);
        java.nio.file.Path jsonPath =
                AppPaths.resolveResultDispatchTableJsonPath(collectUiEnv());
        jp.co.pm.ai.desktop.dispatch.Stage21TrialSnapshotStore.Stage21TrialMeta meta =
                jp.co.pm.ai.desktop.dispatch.Stage21TrialSnapshotStore.tryLoadMeta(jsonPath);
        StringBuilder body = new StringBuilder();
        body.append("段階2.1（残業/休出シミュ）の処理が正常終了しました。\n\n");
        if (meta.hasAttendanceMeta() && meta.overrideSummary() != null) {
            body.append(meta.overrideSummary().formatSummaryLine()).append('\n');
        }
        if (promoted != null && promoted.filesCopied() > 0) {
            body.append("\nメイン output へ ")
                    .append(promoted.filesCopied())
                    .append(" 件の成果物を正本反映しました。");
            if (promoted.mainDispatchJson() != null) {
                body.append("\n配台表: ").append(promoted.mainDispatchJson());
            }
            body.append(
                    "\n配台計画手動修正タブで (段階2後) と (段階2.1後) を比較できます。");
        } else {
            body.append("\nメイン output（結果_配台表.json・計画 JSON 等）へ正本反映済みです。");
        }
        showStageCompletionDialog("段階2.1 完了", body.toString());
    }

    /** 段階2.1 異常終了後。 */
    void notifyStage21OvertimeSimulationFailure(String detailMessage) {
        selectMainShellTab(MainShellTabId.RUN);
        Alert alert = new Alert(AlertType.ERROR);
        initDialogOwnerIfSceneReady(alert);
        applyAlertStylesheetsFromOwner(alert);
        alert.setTitle("段階2.1 失敗");
        alert.setHeaderText(null);
        StringBuilder body = new StringBuilder();
        body.append("段階2.1（残業/休出シミュ）が異常終了しました。\n");
        if (detailMessage != null && !detailMessage.isBlank()) {
            body.append(detailMessage.trim()).append('\n');
        }
        body.append("\n詳細は「実行・ログ」タブのログを確認してください。");
        applyScrollableAlertBody(alert, body.toString());
        alert.showAndWait();
    }

    /** 実行・ログタブ「その他」の Gemini スキップチェック状態（配台不要ルールタブの Java 側 AI 用）。 */
    public boolean mainRunSkipGeminiApi() {
        return mainRunTabController != null && mainRunTabController.snapshotSkipGeminiApi();
    }

    void acceptReloadAfterStage1PlanInput(Runnable r) {
        this.reloadAfterStage1PlanInput = r;
    }

    /** 段階1キャッシュクリア時: 配台計画_タスク入力タブの表を空にする。 */
    void clearPlanInputTableForStage1CacheClear() {
        if (planInputTabController != null) {
            planInputTabController.clearTableForStage1CacheClear();
        }
    }

    /** 段階2.0 実行直前: 段階2〜段階2.1 の成果物を削除し関連タブを初期化する（確認ダイアログなし）。 */
    private void clearStage2CachesBeforeStage2Run() {
        Map<String, String> ui = collectUiEnv();
        PipelineDownstreamResultsClearer.ClearResult cleared =
                PipelineDownstreamResultsClearer.clearStage2Downstream(
                        ui, true /* 当日配台ソース束は段階2で再利用する */);
        for (String line : cleared.detailLines()) {
            appendLog(line);
        }
        if (cleared.anyFailed()) {
            appendLog("[stage2] 実行前: 段階2〜段階2.1 成果物の一部を削除できませんでした。");
        } else {
            appendLog("[stage2] 実行前に段階2〜段階2.1 キャッシュをクリアしました。");
        }
        pendingStage21OvertimeJsonPath = null;
        pendingStage2InProgressNextDayJsonPath = null;
        pendingStage2AladdinTodayExcludeJsonPath = null;
        syncUiAfterDownstreamPipelineResultsCleared();
    }

    /** 段階1開始時: 段階2〜段階2.1 成果物削除後に関連タブの表示を初期化する。 */
    private void syncUiAfterDownstreamPipelineResultsCleared() {
        if (dispatchInteractiveTabController != null) {
            dispatchInteractiveTabController.resetTableDisplayForStage2Run();
        }
        if (mainRunTabController != null) {
            mainRunTabController.setStage2ArtifactPaths("", "");
        }
        if (specialRulesTabController != null) {
            specialRulesTabController.reloadTraceFromDisk();
        }
        invalidateDeliveryCalendarAfterPipelineRun();
        refreshEquipmentGanttGraphicAfterPipelineRun();
        refreshOperatorCardAfterPipelineRun();
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
        if (lastDispatchTableDirty != dispatchTableDirty) {
            dispatchTableDirtyGeneration++;
            lastDispatchTableDirty = dispatchTableDirty;
        }
        if (planInputTabController != null) {
            planInputTabController.setStage2BlockedByUnsavedDispatchEdit(dispatchTableDirty);
        }
    }

    void triggerStage1() {
        if (blockIfStage2SourceGuardBusy("段階1")) {
            return;
        }
        if (blockIfPlanningStagesCalendarNotReady("段階1")) {
            return;
        }
        if (blockIfPipelineCheckBlocksStage1()) {
            return;
        }
        if (mainRunTabController == null || !mainRunTabController.snapshotTodayDispatch()) {
            stage1StartedWithTodayDispatch = false;
            pendingTodayDispatchSourcePair = null;
            startStage1AfterStrictBundleInvalidation();
            return;
        }
        loadTodayDispatchSourceCandidatesForStage1();
    }

    private void loadTodayDispatchSourceCandidatesForStage1() {
        Map<String, String> ui = new HashMap<>(collectUiEnv());
        Dialog<Void> wait = new Dialog<>();
        wait.initOwner(primaryStage);
        wait.setTitle("段階1");
        wait.setHeaderText("当日配台ソース候補を取得中…");
        wait.getDialogPane().getButtonTypes().add(ButtonType.CANCEL);
        wait.getDialogPane().setContent(new ProgressIndicator());

        Task<List<Stage1SourcePairMatcher.MatchedPair>> task =
                new Task<>() {
                    @Override
                    protected List<Stage1SourcePairMatcher.MatchedPair> call() {
                        return Stage1SourcePairMatcher.buildSelectableRows(ui);
                    }
                };
        wait.setOnCloseRequest(event -> task.cancel());
        task.setOnSucceeded(
                event -> {
                    wait.close();
                    if (prepareTodayDispatchSourceSelectionForStage1(task.getValue())) {
                        startStage1AfterStrictBundleInvalidation();
                    }
                });
        task.setOnCancelled(
                event -> {
                    wait.close();
                    appendLog("[stage1] 当日配台: ソース候補取得をキャンセルしました。");
                });
        task.setOnFailed(
                event -> {
                    wait.close();
                    Throwable failure = task.getException();
                    String detail =
                            failure != null && failure.getMessage() != null
                                    ? failure.getMessage()
                                    : String.valueOf(failure);
                    appendLog("[stage1] 当日配台: ソース候補取得に失敗: " + detail);
                    showErrorDialog("段階1", "当日配台ソース候補の取得に失敗しました。\n" + detail);
                });
        Thread worker = new Thread(task, "today-dispatch-source-scan");
        worker.setDaemon(true);
        worker.start();
        wait.show();
    }

    private void startStage1AfterStrictBundleInvalidation() {
        Map<String, String> ui = new HashMap<>(collectUiEnv());
        Stage1SourceBundleCompletionGate.Result invalidation =
                Stage1SourceBundleCompletionGate.invalidateBeforeStage1(
                        () -> Stage1SourceBundleIo.deleteDefaultIfExists(ui));
        if (!invalidation.completionAllowed()) {
            appendLog("[stage1] " + invalidation.message());
            showErrorDialog(
                    "段階1",
                    invalidation.message() + "\n旧bundleを無効化できないため段階1を開始しません。");
            return;
        }
        runStage(STAGE1);
    }

    /**
     * 当日配台 ON のとき段階1直前に加工計画取得時刻を選び、skip_today を自動設定する。
     *
     * @return false なら段階1を中止
     */
    private boolean prepareTodayDispatchSourceSelectionForStage1(
            List<Stage1SourcePairMatcher.MatchedPair> rows) {
        if (rows.isEmpty()) {
            appendLog("[stage1] 当日配台: 加工計画の候補が見つかりません。");
            showErrorDialog(
                    "段階1",
                    "「当日配台する」ですが、加工計画フォルダに候補ファイルがありません。\n"
                            + "PM_AI_TASK_INPUT_SOURCE_DIR を確認するか、「当日は配台しない」を選んでください。");
            return false;
        }
        Optional<Stage1SourcePairMatcher.MatchedPair> chosen =
                TodayDispatchSourceSelectionDialog.prompt(primaryStage, rows);
        if (chosen.isEmpty()) {
            appendLog("[stage1] 当日配台: ソース選択をキャンセルしました。");
            return false;
        }
        Stage1SourcePairMatcher.MatchedPair pair = chosen.get();
        if (pair.dailyReport() == null) {
            appendLog("[stage1] 当日配台: 同日の加工日報が無いため実行できません。");
            showErrorDialog(
                    "段階1",
                    "選択した加工計画に対応する同日の加工日報がありません。\n"
                            + "日報を取得するか、別の計画取得時刻を選んでください。");
            return false;
        }
        stage1StartedWithTodayDispatch = true;
        pendingTodayDispatchSourcePair = pair;
        if (pair.largeDeltaWarning()) {
            appendLog(
                    "[stage1] 当日配台: 計画と日報の取得時刻差が "
                            + pair.deltaMinutes()
                            + " 分です（警告のみ・実行は続行します）。");
        }
        boolean skipToday =
                Stage2SkipTodayDispatchPolicy.shouldSkipTodayDispatch(
                        pair.plan().extractionTime());
        mainRunTabController.applyStage2SkipTodayDispatchFromSession(skipToday);
        appendLog(
                "[stage1] 当日配台: 計画="
                        + pair.plan().fileName()
                        + " 日報="
                        + pair.dailyReport().fileName()
                        + " skip_today="
                        + (skipToday ? "1" : "0"));
        return true;
    }

    private boolean guardTodayDispatchSourceBundleBeforeStageRun(
            String stageLabel, Runnable continuation) {
        Map<String, String> guardEnvironment = new HashMap<>(collectUiEnv());
        Stage2SourceGuardSnapshot startedSnapshot =
                captureStage2SourceGuardSnapshot(guardEnvironment);
        boolean submitted =
                stage2SourceGuardCoordinator.submit(
                        () -> evaluateTodayDispatchSourceGuard(startedSnapshot, guardEnvironment),
                        outcome -> {
                            Stage2SourceGuardSnapshot currentSnapshot =
                                    captureStage2SourceGuardSnapshot(
                                            new HashMap<>(collectUiEnv()));
                            if (!startedSnapshot.matches(currentSnapshot)) {
                                pendingTodayDispatchStageBundle = null;
                                String message =
                                        "ガード中に実行条件が変更されました。"
                                                + startedSnapshot.mismatchMessage(currentSnapshot)
                                                + "\n状態を確認して段階処理を再実行してください。";
                                appendLog("[" + stageLabel + "] 実行条件変更のため中止しました。");
                                showErrorDialog(stageLabel, message);
                                Platform.runLater(this::applyRunTabGating);
                                return;
                            }
                            if (!outcome.guard().allowed()) {
                                pendingTodayDispatchStageBundle = null;
                                appendLog(
                                        "["
                                                + stageLabel
                                                + "] "
                                                + outcome.guard().message().replace('\n', ' '));
                                showErrorDialog(stageLabel, outcome.guard().message());
                                Platform.runLater(this::applyRunTabGating);
                                return;
                            }
                            pendingTodayDispatchStageBundle = outcome.bundle();
                            if (outcome.bundle() != null && mainRunTabController != null) {
                                mainRunTabController.applyStage2SkipTodayDispatchFromSession(
                                        Stage2SkipTodayDispatchPolicy.shouldSkipTodayDispatch(
                                                outcome.bundle().planExtractionTime()));
                            }
                            continuation.run();
                            Platform.runLater(this::applyRunTabGating);
                        },
                        failure -> {
                            pendingTodayDispatchStageBundle = null;
                            appendLog(
                                    "["
                                            + stageLabel
                                            + "] 固定ソース確認に失敗: "
                                            + (failure.getMessage() != null
                                                    ? failure.getMessage()
                                                    : failure.getClass().getSimpleName()));
                            showErrorDialog(stageLabel, "固定ソース確認に失敗しました。段階処理を再実行してください。");
                            Platform.runLater(this::applyRunTabGating);
                        });
        if (!submitted) {
            appendLog("[" + stageLabel + "] 固定ソース確認中のため重複起動を拒否しました。");
            return false;
        }
        applyRunTabGating();
        return true;
    }

    private StageSourceGuardOutcome evaluateTodayDispatchSourceGuard(
            Stage2SourceGuardSnapshot snapshot, Map<String, String> guardEnvironment)
            throws IOException {
        if (!snapshot.todayDispatch()) {
            return new StageSourceGuardOutcome(Stage2SourceConsistencyGuard.Result.ok(), null);
        }
        Optional<Stage1SourceBundle> bundle =
                Stage1SourceBundleIo.readIfPresentStrict(guardEnvironment);
        if (bundle.isEmpty()) {
            bundle = recoverTodayDispatchSourceBundleFromPendingPair(guardEnvironment);
        }
        Stage2SourceConsistencyGuard.Result guard =
                Stage2SourceConsistencyGuard.verify(
                        guardEnvironment, bundle.orElse(null));
        return new StageSourceGuardOutcome(guard, guard.allowed() ? bundle.orElse(null) : null);
    }

    /**
     * 段階1成功時にディスクへ書けなかった／消えた場合でも、同一セッションで選んだペアが残っていれば
     * ソース束を再保存して段階2を継続できるようにする。
     */
    private Optional<Stage1SourceBundle> recoverTodayDispatchSourceBundleFromPendingPair(
            Map<String, String> ui) throws IOException {
        if (pendingTodayDispatchSourcePair == null
                || pendingTodayDispatchSourcePair.dailyReport() == null) {
            return Optional.empty();
        }
        Stage1SourceBundle recovered =
                Stage1SourceBundle.fromMatchedPair(
                        pendingTodayDispatchSourcePair, System.currentTimeMillis());
        Stage1SourceBundleIo.writeDefault(ui, recovered);
        return Optional.of(recovered);
    }

    private Stage2SourceGuardSnapshot captureStage2SourceGuardSnapshot(
            Map<String, String> environment) {
        return new Stage2SourceGuardSnapshot(
                mainRunTabController != null && mainRunTabController.snapshotTodayDispatch(),
                planInputTabController != null
                        && planInputTabController.isPlanInputTableDirtySinceSave(),
                planInputTabController != null
                        ? planInputTabController.snapshotPlanInputDirtyGeneration()
                        : 0L,
                dispatchInteractiveTabController != null
                        && dispatchInteractiveTabController.isDispatchDocDirtySinceSave(),
                dispatchTableDirtyGeneration,
                runLock.get(),
                environment);
    }

    private record StageSourceGuardOutcome(
            Stage2SourceConsistencyGuard.Result guard, Stage1SourceBundle bundle) {}

    private void runStageAfterStage2SourceGuard(String script) {
        stage2SourceGuardRunHandoff = true;
        try {
            runStage(script);
        } finally {
            stage2SourceGuardRunHandoff = false;
        }
    }

    private void overlayTodayDispatchSourcesForStageRun(
            Map<String, String> uiRun, String script) {
        if (mainRunTabController == null || !mainRunTabController.snapshotTodayDispatch()) {
            return;
        }
        if (STAGE1.equals(script) && pendingTodayDispatchSourcePair != null) {
            overlayTodayDispatchPairToEnv(uiRun, pendingTodayDispatchSourcePair);
            return;
        }
        if (STAGE2.equals(script) || STAGE2_1.equals(script)) {
            if (pendingTodayDispatchStageBundle != null) {
                Stage2SourceConsistencyGuard.overlayBundlePaths(
                        uiRun, pendingTodayDispatchStageBundle);
            }
        }
    }

    private static void overlayTodayDispatchPairToEnv(
            Map<String, String> env, Stage1SourcePairMatcher.MatchedPair pair) {
        if (env == null || pair == null || pair.plan() == null || pair.dailyReport() == null) {
            return;
        }
        String planPath = pair.plan().path().toString();
        env.put(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH, planPath);
        env.put(AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK, planPath);
        env.put(
                jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup.KEY_DAILY_REPORT_CSV_PATH,
                pair.dailyReport().path().toString());
    }

    private Stage1SourceBundleCompletionGate.Result persistStage1SourceBundleAfterSuccess() {
        /*
         * 完了時点のチェックボックスではなく、段階1開始時の当日配台意図を使う。
         * 実行中のセッション再適用で OFF になると、旧束削除だけして保存せず完了扱いになるため。
         */
        boolean todayDispatch = stage1StartedWithTodayDispatch;
        boolean bundleReady =
                pendingTodayDispatchSourcePair != null
                        && pendingTodayDispatchSourcePair.dailyReport() != null;
        Map<String, String> ui = new HashMap<>(collectUiEnv());
        Path bundlePath = Stage1SourceBundleIo.defaultCachePath(ui);
        Stage1SourceBundleCompletionGate.Result result =
                Stage1SourceBundleCompletionGate.persist(
                        todayDispatch,
                        bundleReady,
                        () -> Stage1SourceBundleIo.deleteDefaultIfExists(ui),
                        () -> {
                            Stage1SourceBundle bundle =
                                    Stage1SourceBundle.fromMatchedPair(
                                            pendingTodayDispatchSourcePair,
                                            System.currentTimeMillis());
                            Stage1SourceBundleIo.writeDefault(ui, bundle);
                            if (!java.nio.file.Files.isRegularFile(bundlePath)) {
                                throw new java.io.IOException(
                                        "ソース束ファイルが作成されませんでした: " + bundlePath);
                            }
                        });
        if (result.completionAllowed() && todayDispatch) {
            appendLog(
                    "[stage1] 当日配台: ソース束を固定しました（"
                            + bundlePath.toAbsolutePath().normalize()
                            + "）。");
        } else if (result.completionAllowed() && !todayDispatch) {
            appendLog(
                    "[stage1] 当日配台 OFF のためソース束は保存していません。"
                            + " 段階2で当日配台を使う場合は ON にして段階1を再実行してください。");
        }
        return result;
    }

    void triggerStage2() {
        if (blockIfStage2SourceGuardBusy("段階2")) {
            return;
        }
        if (blockIfMaterialLookupTablesHaveBlankValues("段階2")) {
            return;
        }
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
        if (blockIfAttendanceNotReadyForStage2()) {
            return;
        }
        if (!guardTodayDispatchSourceBundleBeforeStageRun(
                "段階2", this::continueStage2AfterSourceGuard)) {
            return;
        }
    }

    private void continueStage2AfterSourceGuard() {
        boolean skipCacheClear =
                planInputTabController != null
                        && planInputTabController.snapshotStage2SkipCacheClearBeforeRun();
        if (!skipCacheClear) {
            clearStage2CachesBeforeStage2Run();
        }
        if (planInputTabController != null) {
            if (!prepareStage2NextDayDialogJsonPaths(collectUiEnv())) {
                return;
            }
        } else {
            pendingStage2InProgressNextDayJsonPath = null;
            pendingStage2AladdinTodayExcludeJsonPath = null;
        }
        runStageAfterStage2SourceGuard(STAGE2);
    }

    /**
     * 当日配台（skip_today OFF）のとき加工途中ダイアログ ①/③ を実質省略する。
     */
    private jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode effectiveStage2NextDayDialogMode() {
        if (planInputTabController == null) {
            return jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode.defaultMode();
        }
        jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode mode =
                planInputTabController.snapshotStage2NextDayDialogMode();
        if (snapshotStage2SkipTodayDispatch()) {
            return mode;
        }
        return switch (mode) {
            case IN_PROGRESS -> jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode.NONE;
            case BOTH -> jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode.ALADDIN_TODAY_EXCLUDE;
            default -> mode;
        };
    }

    /**
     * 段階2直前ダイアログ（①加工途中 / ②アラジン除外 / ③両方）の確定と JSON 書込。
     *
     * @return false なら段階2を中止（キャンセルまたは保存失敗）
     */
    private boolean prepareStage2NextDayDialogJsonPaths(Map<String, String> ui) {
        jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode requestedMode =
                planInputTabController != null
                        ? planInputTabController.snapshotStage2NextDayDialogMode()
                        : jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode.defaultMode();
        jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode mode =
                effectiveStage2NextDayDialogMode();
        if (snapshotTodayDispatch()
                && !snapshotStage2SkipTodayDispatch()
                && mode != requestedMode
                && requestedMode.runsInProgressDialog()) {
            appendLog(
                    "[stage2] 当日配台する のため加工途中の翌日配台ダイアログ(①)を省略します。"
                            + " 設定する場合は「当日は配台しない」を選んでください。");
        }
        pendingStage2InProgressNextDayJsonPath = null;
        pendingStage2AladdinTodayExcludeJsonPath = null;

        java.nio.file.Path inProgressCache =
                jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo.defaultCachePath(ui);
        java.nio.file.Path aladdinCache =
                jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo.defaultCachePath(
                        ui);

        if (mode == jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode.NONE) {
            jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo.deleteIfExists(inProgressCache);
            jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo.deleteIfExists(
                    aladdinCache);
            return true;
        }

        List<jp.co.pm.ai.desktop.ui.Stage2InProgressNextDayDispatchDialog.Row> inProgressRows =
                mode.runsInProgressDialog()
                        ? planInputTabController.collectInProgressRowsForNextDayDialog(ui)
                        : List.of();
        List<jp.co.pm.ai.desktop.ui.Stage2AladdinTodayExcludeNextDayDispatchDialog.Row> aladdinRows =
                mode.runsAladdinExcludeDialog()
                        ? planInputTabController.collectAladdinTodayExcludeRowsForNextDayDialog(ui)
                        : List.of();

        if (mode == jp.co.pm.ai.planning.stage2.Stage2NextDayDialogMode.ALADDIN_TODAY_EXCLUDE) {
            var inProgressForHint =
                    planInputTabController.collectInProgressRowsForNextDayDialog(ui);
            if (!inProgressForHint.isEmpty()) {
                appendLog(
                        "[stage2] 加工途中 "
                                + inProgressForHint.size()
                                + " 件は②には表示しません（実加工数>0）。"
                                + " 設定する場合はラジオ「加工途中…」または「①と②の両方」を選んでください。");
            }
        }
        if (mode.runsAladdinExcludeDialog() && aladdinRows.isEmpty()) {
            appendLog(
                    "[stage2] アラジン当日対象: 0件のため表示を省略"
                            + "（shaped_aladdin_plan 未読込・当日列なし・実加工>0 の行のみ等）。");
        }

        if (inProgressRows.isEmpty() && aladdinRows.isEmpty()) {
            jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo.deleteIfExists(
                    inProgressCache);
            jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo.deleteIfExists(
                    aladdinCache);
            return true;
        }

        var result =
                jp.co.pm.ai.desktop.ui.Stage2NextDayDispatchDialog.prompt(
                        primaryStage, inProgressRows, aladdinRows);
        if (result.isEmpty()) {
            appendLog("[stage2] 翌日配台量設定をキャンセルしました。");
            return false;
        }

        if (mode.runsInProgressDialog() && !result.get().inProgressEntries().isEmpty()) {
            try {
                jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo.write(
                        inProgressCache, result.get().inProgressEntries());
                pendingStage2InProgressNextDayJsonPath =
                        inProgressCache.toAbsolutePath().normalize();
            } catch (Exception ex) {
                appendLog("[stage2] 翌日配台量 JSON の保存に失敗: " + ex.getMessage());
                return false;
            }
        } else {
            jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo.deleteIfExists(
                    inProgressCache);
        }

        if (mode.runsAladdinExcludeDialog() && !result.get().aladdinExcludeEntries().isEmpty()) {
            try {
                jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo.write(
                        aladdinCache, result.get().aladdinExcludeEntries());
                pendingStage2AladdinTodayExcludeJsonPath =
                        aladdinCache.toAbsolutePath().normalize();
            } catch (Exception ex) {
                appendLog("[stage2] アラジン翌日除外 JSON の保存に失敗: " + ex.getMessage());
                return false;
            }
        } else {
            jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo.deleteIfExists(
                    aladdinCache);
        }

        return true;
    }

    /** 段階2.1（残業/休出シミュ）: 時間外ウィザードを起動し、確定後に {@link #triggerStage21} へ進む。 */
    void launchStage21OvertimeSimulationWizard() {
        if (blockIfStage2SourceGuardBusy("段階2.1")) {
            return;
        }
        if (blockIfMaterialLookupTablesHaveBlankValues("段階2.1")) {
            return;
        }
        if (planInputTabController != null
                && planInputTabController.isPlanInputTableDirtySinceSave()) {
            appendLog(
                    "[stage2.1] 配台計画_タスク入力タブの表に未保存の変更があります。「保存」または「再読み」で確定してから実行してください。");
            showErrorDialog(
                    "段階2.1",
                    "配台計画_タスク入力タブの変更を「保存」または「再読み」で確定してから段階2.1 を実行してください。");
            return;
        }
        if (dispatchInteractiveTabController != null
                && dispatchInteractiveTabController.isDispatchDocDirtySinceSave()) {
            appendLog(
                    "[stage2.1] 配台計画手動修正に未保存の変更があります。「保存」してから実行してください。");
            showErrorDialog(
                    "段階2.1",
                    "配台計画手動修正タブの変更を「保存 (JSON+xlsx)」で確定してから段階2.1 を実行してください。");
            return;
        }
        if (blockIfAttendanceNotReadyForStage2()) {
            showErrorDialog(
                    "段階2.1",
                    attendanceReadinessTooltip.isBlank()
                            ? "勤怠正本（attendance-data.json）が未準備です。会社カレンダー／メンバー勤怠タブでセットアップしてください。"
                            : attendanceReadinessTooltip);
            return;
        }
        java.nio.file.Path mainJson = AppPaths.resolveResultDispatchTableJsonPath(collectUiEnv());
        if (!java.nio.file.Files.isRegularFile(mainJson)) {
            appendLog(
                    "[stage2.1] 段階2 の成果物（結果_配台表.json）がありません。先に段階2を実行してください。");
            showErrorDialog(
                    "段階2.1",
                    "段階2.1 を実行する前に段階2を実行し、結果_配台表.json を生成してください。");
            return;
        }
        final java.nio.file.Path pyExe = resolveStagePythonExecutablePath(collectUiEnv());
        final java.nio.file.Path pyDir = AppPaths.resolvePythonScriptDir(collectUiEnv());
        final Map<String, String> pyEnv = snapshotDispatchTrialPythonEnv();
        final Stage owner = primaryStage;
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                AttendanceOvertimePreview.Preview preview =
                                        AttendanceOvertimePreviewPython.load(
                                                pyExe, pyDir, pyEnv, this::appendLog);
                                Platform.runLater(
                                        () ->
                                                OvertimeSimulationWizard.show(
                                                        owner,
                                                        this,
                                                        preview,
                                                        OvertimeSimulationWizard.Target.STAGE21,
                                                        this::triggerStage21));
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () ->
                                                showErrorDialog(
                                                        "段階2.1",
                                                        "勤怠プレビューの取得に失敗しました。\n"
                                                                + (ex.getMessage() != null
                                                                        ? ex.getMessage()
                                                                        : ex)));
                            }
                        },
                        "dispatch-stage21-preview");
        worker.setDaemon(true);
        worker.start();
    }

    /** 段階2.1（残業/休出シミュ）: ウィザード確定後にフル再配台（output/stage21/）。 */
    void triggerStage21(java.nio.file.Path overtimeSimulationJson) {
        if (blockIfStage2SourceGuardBusy("段階2.1")) {
            return;
        }
        if (blockIfMaterialLookupTablesHaveBlankValues("段階2.1")) {
            return;
        }
        java.nio.file.Path mainJson =
                AppPaths.resolveResultDispatchTableJsonPath(collectUiEnv());
        if (!java.nio.file.Files.isRegularFile(mainJson)) {
            appendLog(
                    "[stage2.1] 段階2 の成果物（結果_配台表.json）がありません。先に段階2を実行してください。");
            showErrorDialog(
                    "段階2.1",
                    "段階2.1 を実行する前に段階2を実行し、結果_配台表.json を生成してください。");
            return;
        }
        if (dispatchInteractiveTabController != null
                && dispatchInteractiveTabController.isDispatchDocDirtySinceSave()) {
            appendLog(
                    "[stage2.1] 配台計画手動修正に未保存の変更があります。JSON を「保存」するか「再読み」後に実行してください。");
            return;
        }
        if (planInputTabController != null
                && planInputTabController.isPlanInputTableDirtySinceSave()) {
            appendLog(
                    "[stage2.1] 配台計画_タスク入力タブの表に未保存の変更があります。「保存」または「再読み」で確定してから実行してください。");
            return;
        }
        if (!guardTodayDispatchSourceBundleBeforeStageRun(
                "段階2.1",
                () -> continueStage21AfterSourceGuard(overtimeSimulationJson, mainJson))) {
            return;
        }
    }

    private void continueStage21AfterSourceGuard(
            java.nio.file.Path overtimeSimulationJson, java.nio.file.Path mainJson) {
        if (overtimeSimulationJson == null
                || !java.nio.file.Files.isRegularFile(overtimeSimulationJson)) {
            appendLog("[stage2.1] 残業シミュレーション JSON が無効です。");
            return;
        }
        pendingStage21OvertimeJsonPath =
                overtimeSimulationJson.toAbsolutePath().normalize();
        if (dispatchInteractiveTabController != null) {
            dispatchInteractiveTabController.captureStage21BaselineBeforeRun(
                    mainJson, pendingStage21OvertimeJsonPath);
        }
        runStageAfterStage2SourceGuard(STAGE2_1);
    }

    /** 段階2.1 実行時のみ子プロセスへ渡す残業 JSON。 */
    private java.nio.file.Path pendingStage21OvertimeJsonPath;

    /** 直前の段階2ダイアログで確定した JSON。{@link #runStage} の段階2のみで子プロセスへ渡す。 */
    private java.nio.file.Path pendingStage2InProgressNextDayJsonPath;

    private java.nio.file.Path pendingStage2AladdinTodayExcludeJsonPath;

    /** 当日配台 ON 時、段階1直前にユーザーが選んだ plan+daily ペア。 */
    private Stage1SourcePairMatcher.MatchedPair pendingTodayDispatchSourcePair;

    /**
     * 段階1開始時点の当日配台 ON/OFF。完了時のチェックボックスやセッション再適用で変わっても、
     * ソース束の保存要否は開始時の意図に従う。
     */
    private boolean stage1StartedWithTodayDispatch;

    private boolean lastDispatchTableDirty;
    private long dispatchTableDirtyGeneration;

    private Stage1SourceBundle pendingTodayDispatchStageBundle;

    private final Stage2SourceGuardCoordinator stage2SourceGuardCoordinator =
            new Stage2SourceGuardCoordinator(
                    command -> {
                        Thread worker =
                                new Thread(command, "today-dispatch-stage-source-guard");
                        worker.setDaemon(true);
                        worker.start();
                    },
                    Platform::runLater);

    /** 固定ソースガードの成功コールバックから runLock 取得へ渡すFXスレッド内限定フラグ。 */
    private boolean stage2SourceGuardRunHandoff;

    private void applyStage2InProgressNextDayDispatchEnv(Map<String, String> ui) {
        java.nio.file.Path p = pendingStage2InProgressNextDayJsonPath;
        if (p != null && java.nio.file.Files.isRegularFile(p)) {
            ui.put(AppPaths.KEY_PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON, p.toString());
        } else {
            ui.remove(AppPaths.KEY_PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON);
            jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo.deleteIfExists(
                    jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo.defaultCachePath(ui));
        }
    }

    private void applyStage2AladdinTodayExcludeNextDayEnv(Map<String, String> ui) {
        java.nio.file.Path p = pendingStage2AladdinTodayExcludeJsonPath;
        if (p != null && java.nio.file.Files.isRegularFile(p)) {
            ui.put(AppPaths.KEY_PM_AI_STAGE2_ALADDIN_TODAY_EXCLUDE_NEXT_DAY_JSON, p.toString());
        } else {
            ui.remove(AppPaths.KEY_PM_AI_STAGE2_ALADDIN_TODAY_EXCLUDE_NEXT_DAY_JSON);
            jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo.deleteIfExists(
                    jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo
                            .defaultCachePath(ui));
        }
    }

    private void applyStage2NextDayDialogEnvs(Map<String, String> ui) {
        applyStage2InProgressNextDayDispatchEnv(ui);
        applyStage2AladdinTodayExcludeNextDayEnv(ui);
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

    /** 配台計画手動修正タブの配台ロール単位 (m) 解決用。未初期化時は {@code null}。 */
    /** Package / dispatch-rules access to plan input tab. */
public PlanInputTabController planInputTabControllerForDispatchRollUnit() {
        return planInputTabController;
    }

    /**
     * 配台試行（段階3）正常終了後: 納期管理ビューは段階3前・段階3後（配台結果）のみ反映し、サマリ xlsx を更新する。
     */
    void reloadDeliveryCalendarInBackgroundAfterDispatchTrialSuccess() {
        if (deliveryCalendarViewTabController != null) {
            deliveryCalendarViewTabController.reloadInBackgroundAfterDispatchTrialSuccess();
        }
    }

    /** 納期管理ビュー内アラジン加工計画タブをソースから再読込する。JavaFX スレッドから呼ぶこと。 */
    void refreshAladdinProcessingPlanTabFromDisk() {
        if (deliveryCalendarViewTabController != null) {
            deliveryCalendarViewTabController.refreshAladdinProcessingPlanTabFromDisk();
        }
    }

    /** アラジン加工計画ソース再読込後、原本転記・計画確認の確認チェックを必要に応じてリセットする。 */
    void notifyAladdinProcessingPlanSourceReloaded(java.nio.file.Path sourceFile) {
        if (requestFormPipelineCheckTabController != null) {
            requestFormPipelineCheckTabController.onAladdinProcessingPlanSourceUpdated(sourceFile);
        }
    }

    /**
     * 段階2直前ダイアログ②: アラジン shaped 表（JSON → 納期管理ビュー内タブ → ソース再読込）。
     */
    AladdinShapedPlanQtyLookup.ShapedTable snapshotShapedAladdinPlanTable(Map<String, String> ui) {
        if (deliveryCalendarViewTabController != null) {
            return deliveryCalendarViewTabController.snapshotShapedAladdinPlanTable(ui);
        }
        return AladdinShapedPlanQtyLookup.loadShapedTable(
                AppPaths.resolveShapedAladdinPlanJsonPath(ui != null ? ui : Map.of()));
    }

    /**
     * 段階2実行前: master「組み合わせ表」に無い工程+機械があればダイアログで配台不要を確認する。
     *
     * @return {@code false} のとき段階2実行を中止（キャンセル）
     */
    private boolean confirmStage2UnknownMasterCombinationsBeforeRun() {
        try {
            Stage2UnknownMasterCombinationPrompt.PromptBundle bundle =
                    Stage2UnknownMasterCombinationPrompt.collectUnknownPairs(collectUiEnv());
            if (bundle.empty()) {
                return true;
            }
            appendLog(
                    "[stage2] 組み合わせ表に無い工程+機械 "
                            + bundle.pairs().size()
                            + " 件 — 配台不要確認ダイアログを表示します。");
            Optional<Stage2UnknownMasterCombinationDialogResult> entered =
                    Stage2UnknownMasterCombinationDialog.prompt(primaryStageForDialogs(), bundle);
            if (entered.isEmpty()) {
                appendLog("[stage2] マスタ未登録工程+機械の確認をキャンセルしたため段階2を中止します。");
                showWarningDialog(
                        "段階2 中止",
                        "マスタ未登録の工程+機械の確認をキャンセルしたため、段階2は実行しませんでした。");
                return false;
            }
            applyStage2UnknownMasterCombinationSelections(entered.get());
            return true;
        } catch (IOException ex) {
            appendLog(
                    "[stage2] マスタ未登録工程+機械の確認に失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            showWarningDialog(
                    "段階2 確認失敗",
                    "マスタ未登録の工程+機械を確認できませんでした。\n"
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString())
                            + "\n\n段階2を続行します。");
            return true;
        }
    }

    /**
     * 段階2実行前: 計画タスクの工程+機械が skills シートに列として無ければ確認ダイアログを出す。
     *
     * @return {@code false} のとき段階2実行を中止（キャンセル）
     */
    private boolean confirmStage2MissingSkillsColumnsBeforeRun() {
        try {
            PlanTasksMissingSkillsColumnPrompt.PromptBundle bundle =
                    PlanTasksMissingSkillsColumnPrompt.collectMissingPairs(collectUiEnv());
            if (bundle.empty()) {
                return true;
            }
            appendLog(
                    "[stage2] skills シートに無い工程+機械 "
                            + bundle.pairs().size()
                            + " 件 — 配台阻害の確認ダイアログを表示します。");
            Optional<Boolean> entered =
                    MissingSkillsSheetColumnDialog.prompt(
                            primaryStageForDialogs(), bundle, true);
            if (entered.isEmpty() || !entered.get()) {
                appendLog(
                        "[stage2] skills シート未登録の確認をキャンセルしたため段階2を中止します。");
                showWarningDialog(
                        "段階2 中止",
                        "skills シートに未登録の工程+機械があるため、段階2は実行しませんでした。\n\n"
                                + bundle.summaryJa(12)
                                + "\n\nmaster の skills シートへ列を追加してください。");
                return false;
            }
            appendLog("[stage2] skills シート未登録あり — ユーザーが続行を選択しました。");
            return true;
        } catch (IOException ex) {
            appendLog(
                    "[stage2] skills シート未登録の確認に失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            showWarningDialog(
                    "段階2 確認失敗",
                    "skills シート未登録の工程+機械を確認できませんでした。\n"
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString())
                            + "\n\n段階2を続行します。");
            return true;
        }
    }

    /**
     * 段階1正常終了後: 計画タスクの工程+機械が skills シートに無ければ警告ダイアログを出す。
     */
    private void warnStage1MissingSkillsColumnsAfterSuccess() {
        try {
            PlanTasksMissingSkillsColumnPrompt.PromptBundle bundle =
                    PlanTasksMissingSkillsColumnPrompt.collectMissingPairs(collectUiEnv());
            if (bundle.empty()) {
                return;
            }
            appendLog(
                    "[stage1] skills シートに無い工程+機械 "
                            + bundle.pairs().size()
                            + " 件 — 段階2配台阻害の警告を表示します。");
            MissingSkillsSheetColumnDialog.prompt(primaryStageForDialogs(), bundle, false);
        } catch (IOException ex) {
            appendLog(
                    "[stage1] skills シート未登録の確認に失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
    }

    private void applyStage2UnknownMasterCombinationSelections(
            Stage2UnknownMasterCombinationDialogResult result) throws IOException {
        if (result == null || result.markExclude() == null || result.markExclude().isEmpty()) {
            appendLog("[stage2] マスタ未登録工程+機械: 配台不要への更新はありません。");
            return;
        }
        Stage2UnknownMasterCombinationPrompt.ApplySummary applied =
                Stage2UnknownMasterCombinationPrompt.applyExcludeSelections(
                        collectUiEnv(), result.markExclude());
        appendLog(
                "[stage2] マスタ未登録工程+機械を配台不要へ反映しました（JSON "
                        + applied.excludeRulesUpdated()
                        + " 件、計画行 "
                        + applied.planRowsUpdated()
                        + " 行）。");
        syncExcludeRulesJsonPathToEnvTab(
                Stage2UnknownMasterCombinationPrompt.resolveExcludeRulesJsonPath(collectUiEnv())
                        .map(Path::toString)
                        .orElse(null));
        if (excludeRulesTabController != null) {
            excludeRulesTabController.tryStartupLoadFromPathField();
        }
        if (planInputTabController != null) {
            planInputTabController.reloadQuietlyFromDisk();
        }
    }

    /**
     * 段階1／2 実行前: 材料・製品種類ルックアップ表をサマリ Excel 同フォルダへ確保し、
     * 子プロセス向け環境変数を揃える。
     */
    private void overlayDispatchLookupTablePathsForStageRun(Map<String, String> uiRun) {
        if (uiRun == null) {
            return;
        }
        AppPaths.ensureAllDispatchLookupTablesFromRepoIfMissing(uiRun);
        putDispatchLookupTableEnvIfPresent(
                uiRun, "RAW_FABRIC_WIDTH_TABLE_PATH", AppPaths.DISPATCH_LOOKUP_USED_RAW_WIDTH);
        putDispatchLookupTableEnvIfPresent(
                uiRun, "ROLL_UNIT_BY_USED_RAW_TABLE_PATH", AppPaths.DISPATCH_LOOKUP_USED_RAW_ROLL);
        putDispatchLookupTableEnvIfPresent(
                uiRun, "PRODUCT_WIDTH_TABLE_PATH", AppPaths.DISPATCH_LOOKUP_PRODUCT_WIDTH);
        putDispatchLookupTableEnvIfPresent(
                uiRun, "PRODUCT_LENGTH_TABLE_PATH", AppPaths.DISPATCH_LOOKUP_PRODUCT_LENGTH);
        putDispatchLookupTableEnvIfPresent(
                uiRun, "PRODUCT_THICKNESS_TABLE_PATH", AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK);
        putDispatchLookupTableEnvIfPresent(
                uiRun, "ROLL_UNIT_LENGTH_TABLE_PATH", AppPaths.DISPATCH_LOOKUP_PRODUCT_ROLL);
    }

    private void putDispatchLookupTableEnvIfPresent(
            Map<String, String> uiRun, String envKey, String filename) {
        Path p = AppPaths.dispatchLookupTablePath(uiRun, filename);
        if (!Files.isRegularFile(p)) {
            return;
        }
        uiRun.put(envKey, p.toString());
        syncEnvTabValue(envKey, p.toString());
    }

    private void syncEnvTabValue(String envKey, String pathStr) {
        if (envRows == null || envKey == null || envKey.isBlank() || pathStr == null || pathStr.isBlank()) {
            return;
        }
        try {
            for (EnvVarRow row : envRows) {
                String k = row.getName() != null ? row.getName().trim() : "";
                if (envKey.equals(k)) {
                    row.setValue(pathStr);
                    appendLog("[env] " + envKey + "=" + pathStr);
                    return;
                }
            }
        } catch (Exception ex) {
            appendLog("[env] " + envKey + " 更新に失敗: " + ex.getMessage());
        }
    }

    /** 環境変数タブの行を更新する（session-state 保存は行リスナーで debounce される）。 */
    @Override
    public void updateEnvTabValue(String envKey, String value) {
        syncEnvTabValue(envKey, value);
        if (AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR.equals(envKey)
                && AppPaths.isRequestFormOriginalDirEnvConfigured(collectUiEnv())
                && requestFormPipelineCheckTabController != null) {
            requestFormPipelineCheckTabController.onRequestFormOriginalDirEnvConfigured();
        }
    }

    /**
     * 段階1／2 実行前: サマリ Excel と同一フォルダの {@code stage1_exclude_rules.json} を確保し、
     * 子プロセスへ渡す {@code PM_AI_EXCLUDE_RULES_JSON} を揃える。
     */
    private void overlayWorkingExcludeRulesJsonPathForStageRun(Map<String, String> uiRun) {
        if (uiRun == null) {
            return;
        }
        AppPaths.resolveDefaultExcludeRulesJsonPath(uiRun)
                .ifPresent(
                        p -> {
                            uiRun.put(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON, p.toString());
                            syncEnvTabValue(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON, p.toString());
                        });
    }

    /** 段階1／2 実行前: 特別ルール JSON を run_snapshots へ凍結し env を上書き。 */
    private void overlayDispatchSpecialRulesForStageRun(Map<String, String> uiRun, String script) {
        if (uiRun == null) {
            return;
        }
        String stage =
                STAGE1.equals(script)
                        ? "stage1"
                        : STAGE2.equals(script) ? "stage2" : "stage";
        try {
            AppPaths.ensureDispatchSpecialRulesJsonFromRepoIfMissing(uiRun);
            var capture = DispatchRuleStageRunOverlay.captureForStage(stage, uiRun);
            if (capture.snapshotPath() != null && Files.isRegularFile(capture.snapshotPath())) {
                uiRun.put(
                        DispatchRulePaths.KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON,
                        capture.snapshotPath().toString());
                syncEnvTabValue(
                        DispatchRulePaths.KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON,
                        capture.snapshotPath().toString());
                appendLog("[dispatch-rules] run snapshot: " + capture.runId());
            } else {
                AppPaths.resolveDefaultDispatchSpecialRulesJsonPath(uiRun)
                        .ifPresent(
                                p -> {
                                    uiRun.put(DispatchRulePaths.KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON, p.toString());
                                    syncEnvTabValue(DispatchRulePaths.KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON, p.toString());
                                });
            }
        } catch (IOException ex) {
            appendLog("[dispatch-rules] snapshot failed: " + ex.getMessage());
        }
    }

    private void syncExcludeRulesJsonPathToEnvTab(String pathStr) {
        syncEnvTabValue(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON, pathStr);
    }

    private void showStage2FailureWithUnknownMasterComboRetry(Integer code, List<String> tailSnap) {
        try {
            Stage2UnknownMasterCombinationPrompt.PromptBundle bundle =
                    Stage2UnknownMasterCombinationPrompt.collectUnknownPairs(collectUiEnv());
            if (!bundle.empty()) {
                Alert alert = new Alert(AlertType.ERROR);
                initDialogOwnerIfSceneReady(alert);
                applyAlertStylesheetsFromOwner(alert);
                alert.setTitle("段階2 失敗");
                alert.setHeaderText("計画データの検証エラーです。マスタ未登録の工程+機械が残っています。");
                StringBuilder body = new StringBuilder();
                body.append(exitCodeLegend(code != null ? code : -1)).append('\n');
                body.append(exitHintJa(code != null ? code : -1)).append("\n\n");
                body.append("組み合わせ表に無い工程+機械: ").append(bundle.pairs().size()).append(" 件\n");
                body.append("「配台不要を設定して再実行」で確認ダイアログを開き、段階2を再実行できます。");
                appendTailLinesToFailureBody(body, tailSnap);
                applyScrollableAlertBody(alert, body.toString());
                ButtonType retry = new ButtonType("配台不要を設定して再実行");
                alert.getButtonTypes().setAll(retry, ButtonType.OK);
                Optional<ButtonType> ans = alert.showAndWait();
                if (ans.isPresent() && ans.get() == retry) {
                    Optional<Stage2UnknownMasterCombinationDialogResult> entered =
                            Stage2UnknownMasterCombinationDialog.prompt(
                                    primaryStageForDialogs(), bundle);
                    if (entered.isPresent()) {
                        applyStage2UnknownMasterCombinationSelections(entered.get());
                        Platform.runLater(() -> runStage(STAGE2));
                    }
                }
                return;
            }
        } catch (IOException ex) {
            appendLog(
                    "[stage2] 失敗後のマスタ未登録確認に失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
        showStageFailureDialog(STAGE2, code, null, tailSnap);
    }

    private void appendTailLinesToFailureBody(StringBuilder body, List<String> tailLines) {
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
    }

    /** 段階1完了後: EC面区分が「不明」の依頼NOがあれば選択ダイアログを表示し plan_input を更新する。 */
    private void promptStage1EcSideUnknownAfterSuccess() {
        try {
            Stage1EcSideUnknownPrompt.PromptBundle bundle =
                    Stage1EcSideUnknownPrompt.collectUnknownIraiNos(collectUiEnv());
            if (bundle.empty()) {
                return;
            }
            appendLog(
                    "[stage1] EC面区分が不明の依頼 "
                            + bundle.items().size()
                            + " 件 — 選択ダイアログを表示します。");
            Optional<Stage1EcSideUnknownDialogResult> entered =
                    Stage1EcSideUnknownDialog.prompt(primaryStageForDialogs(), bundle);
            if (entered.isEmpty()) {
                appendLog(
                        "[stage1] EC面区分の選択ダイアログをキャンセルしました（不明のまま）。"
                                + " 配台計画タスク入力タブで手動修正できます。");
                return;
            }
            Stage1EcSideUnknownPrompt.ApplySummary applied =
                    Stage1EcSideUnknownPrompt.applySelections(
                            collectUiEnv(), entered.get().selections());
            appendLog("[stage1] EC面区分を " + applied.rowsUpdated() + " 行更新しました。");
            if (reloadAfterStage1PlanInput != null) {
                reloadAfterStage1PlanInput.run();
            }
        } catch (IOException ex) {
            appendLog(
                    "[stage1] EC面区分選択ダイアログ失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            showWarningDialog(
                    "EC面区分",
                    "EC面区分の選択ダイアログを表示できませんでした。\n"
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString())
                            + "\n\n配台計画タスク入力タブで手動修正してください。");
        }
    }

    /**
     * 段階1マージ後: 新規追記された空欄キーがあれば入力ダイアログを出し、OK 時に {@code code/} へ書き戻す。
     */
    private void promptStage1NewMaterialLookupsAfterMerge() {
        try {
            CodeDispatchLookupTablesValidator.ValidationResult vr =
                    CodeDispatchLookupTablesValidator.validateNoBlankValues(collectUiEnv());
            if (vr.ok()) {
                return;
            }
            CodeDispatchLookupTablesBlankPrompt.PromptBundle bundle =
                    CodeDispatchLookupTablesBlankPrompt.collectPrompt(collectUiEnv(), vr);
            if (bundle.empty()) {
                return;
            }
            appendLog(
                    "[stage1] 新規材料・製品種類 "
                            + (bundle.products() != null ? bundle.products().size() : 0)
                            + " 製品 / "
                            + (bundle.usedRaws() != null ? bundle.usedRaws().size() : 0)
                            + " 原反 — 入力ダイアログを表示します。");
            Optional<Stage1NewMaterialLookupDialog.Result> entered =
                    Stage1NewMaterialLookupDialog.prompt(primaryStageForDialogs(), bundle);
            if (entered.isEmpty()) {
                appendLog(
                        "[stage1] 新規材料・製品種類の入力ダイアログをキャンセルしました（空欄のまま）。"
                                + " 段階2・段階3実行前に「材料・製品種類情報」タブで入力してください。");
                return;
            }
            CodeDispatchLookupTablesBlankPrompt.ApplySummary applied =
                    CodeDispatchLookupTablesBlankPrompt.applyInputs(
                            collectUiEnv(),
                            entered.get().products(),
                            entered.get().usedRaws());
            appendLog(
                    "[stage1] 新規材料・製品種類をダイアログ入力で登録しました（"
                            + applied.updatedFields()
                            + " フィールド更新）。");
            if (codeDispatchLookupTablesTabController != null) {
                codeDispatchLookupTablesTabController.reloadAllFromDisk();
            }
        } catch (IOException ex) {
            appendLog(
                    "[stage1] 新規材料・製品種類ダイアログ失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            showWarningDialog(
                    "材料テーブル入力",
                    "新規材料・製品種類の入力ダイアログを表示できませんでした。\n"
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString())
                            + "\n\n「材料・製品種類情報」タブで手動入力してください。");
        }
    }

    private void syncRequestFormFeedLocFromStage1Plan() {
        if (requestFormInputTabController == null) {
            return;
        }
        int added = requestFormInputTabController.mergeFeedLocFromStage1Plan(collectUiEnv());
        if (added > 0) {
            appendLog("[stage1] 依頼書入力の投入場所候補を計画データから " + added + " 件追加しました。");
            scheduleDesktopSessionSave();
        } else if (added < 0) {
            appendLog("[stage1] 依頼書入力の投入場所候補の計画データ取込に失敗しました。");
        }
    }

    private String buildStage1CompletionMessage() {
        String skillsNote = "";
        try {
            PlanTasksMissingSkillsColumnPrompt.PromptBundle skillsBundle =
                    PlanTasksMissingSkillsColumnPrompt.collectMissingPairs(collectUiEnv());
            if (!skillsBundle.empty()) {
                skillsNote =
                        "\n\nmaster「skills」シートに未登録の工程+機械が "
                                + skillsBundle.pairs().size()
                                + " 件あります（段階2で配台されません）。\n"
                                + skillsBundle.summaryJa(8);
            }
        } catch (IOException ignored) {
            // 完了メッセージはベストエフォート
        }
        try {
            CodeDispatchLookupTablesValidator.ValidationResult vr =
                    CodeDispatchLookupTablesValidator.validateNoBlankValues(collectUiEnv());
            if (vr.ok()) {
                return "段階1 の処理が正常終了しました。" + skillsNote;
            }
            return "段階1 の処理が正常終了しました。\n\n"
                    + "材料・製品種類情報（code/）に値が空欄の行が残っています。"
                    + " 段階2・段階3の前に「材料・製品種類情報」タブで入力してください。"
                    + skillsNote;
        } catch (IOException ex) {
            return "段階1 の処理が正常終了しました。" + skillsNote;
        }
    }

    /**
     * 材料・製品種類情報（{@code code/}）に値が空欄の行があればログと警告を出し true（呼び出し側は処理を中止）。
     *
     * @param operationLabelJa ユーザー向けの処理名（例: 段階2、配台試行（段階3））
     */
    boolean blockIfMaterialLookupTablesHaveBlankValues(String operationLabelJa) {
        try {
            CodeDispatchLookupTablesValidator.ValidationResult vr =
                    CodeDispatchLookupTablesValidator.validateNoBlankValues(collectUiEnv());
            if (vr.ok()) {
                return false;
            }
            appendLog(
                    "[材料テーブル] "
                            + operationLabelJa
                            + " を中止（材料・製品種類情報に値が空欄の行があります）");
            for (String line : vr.logLines()) {
                appendLog(line);
            }
            showWarningDialog(
                    "材料テーブル未入力",
                    operationLabelJa
                            + " は実行できません。\n\n"
                            + vr.messageJa(12));
            return true;
        } catch (IOException ex) {
            appendLog(
                    "[材料テーブル] "
                            + operationLabelJa
                            + " を中止（材料テーブルの検証に失敗: "
                            + ex.getMessage()
                            + "）");
            showErrorDialog(
                    "材料テーブル検証失敗",
                    operationLabelJa
                            + " は実行できません。\n材料・製品種類情報（code/）の読み込みに失敗しました。\n"
                            + ex.getMessage());
            return true;
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

    /**
     * 段階2／段階3 完了後: 最新 member_schedule と結果_配台表 JSON でオペレーターカードを再読込・プレビュー更新する。
     */
    void refreshOperatorCardAfterPipelineRun() {
        if (operatorCardTabController == null) {
            return;
        }
        try {
            Map<String, String> ui = collectUiEnv();
            Path dir = AppPaths.defaultPlanningOutputDir(ui);
            Path newestMember = Stage2OutputNaming.newestPrimaryMemberXlsx(dir);
            if (newestMember == null) {
                newestMember = Stage2OutputNaming.newestPrimaryMemberJson(dir);
            }
            String memStr = newestMember != null ? newestMember.toString() : "";
            operatorCardTabController.tryAutofillMemberJsonFromStage2(memStr);
            operatorCardTabController.syncAfterPipelineArtifactRefresh();
        } catch (Exception ex) {
            appendLog("[operator-card] 更新エラー: " + ex.getMessage());
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
            refreshOperatorCardAfterPipelineRun();
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
                    "[startup] 初回起動: "
                            + AppPaths.PORTABLE_FIRST_LAUNCH_MARKER_FILE
                            + " を削除しました（工場既定="
                            + firstLaunchSite.displayLabelJa()
                            + "）。");
        } catch (Exception ex) {
            appendLog(
                    "[startup] 初回起動の環境変数初期化に失敗（"
                            + AppPaths.PORTABLE_FIRST_LAUNCH_MARKER_FILE
                            + " は残します）: "
                            + ex.getMessage());
        }
    }

    /**
     * バージョンアップ後の再起動（{@link PortableBundleUpgradeFollowUp} 待ち）では、セッション復元直後に
     * 新 JAR の ui_ref 既定へ環境変数初期化を強制する。スプラッシュ後の {@link
     * #finishPortableUpgradeWithFactorySitePrompt} でも再度実行する。
     */
    private void maybeForceEnvInitAfterPortableUpgradeRestart() {
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        if (!PortableBundleSelfUpdater.isPortableBundleLayout(cwd)) {
            return;
        }
        if (!PortableBundleUpgradeFollowUp.isPendingFor(cwd)) {
            return;
        }
        FactorySite site = resolveFactorySiteForPortableUpgrade(Optional.empty());
        appendLog(
                "[startup] バージョンアップ後の再起動: 環境変数初期化を強制実行します。工場="
                        + site.displayLabelJa());
        applyFactoryScopedGlobalAndEnvReset(site, true);
        persistOperatorWorkspaceForEnvInitBaseline(site);
        recordEnvInitializationBaseline();
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
        if (PortableBundleUpdateLauncher.tryApplyStagedBundleOnColdStart(cwd, this::appendLog)) {
            appendLog("[startup] 前回保留のデスクトップ本体更新を適用しました。");
        }
        Path localData = cwd.resolve("pm-ai-data").normalize();
        if (PortableBundleUpgradeFollowUp.isPendingFor(cwd)) {
            appendLog(
                    "[startup] バージョンアップ後の再起動を検出: 環境変数初期化を強制実行します。");
            deferOperatorPromptForPortableUpgrade.set(true);
            finishPortableUpgradeWithFactorySitePrompt(
                    cwd,
                    localData,
                    0,
                    null,
                    "（デスクトップ本体の再起動後）",
                    Optional.empty());
            return;
        }
        Map<String, String> ui = collectUiEnv();
        String raw = ui.get(AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR);
        if (raw == null || raw.isBlank()) {
            Alert a = new Alert(AlertType.INFORMATION);
            initDialogOwnerIfSceneReady(a);
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
        if (!PortableBundleSelfUpdater.isValidPortableBundleCanonical(canonical)) {
            appendLog(
                    "[startup] 正本パスにアクセスできません: "
                            + PortableBundleSelfUpdater.safePathForLog(canonical));
            Alert w = new Alert(AlertType.WARNING);
            initDialogOwnerIfSceneReady(w);
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
        Optional<PortableBundleBuildManifest> buildManifest =
                PortableBundleBuildManifest.readBesideCanonical(canonical);
        if (!PortableBundleSelfUpdater.shouldUpdateBundle(canonical, cwd, localData)) {
            String reason =
                    cv.isEmpty()
                            ? "正本の version.txt が読めません（ZIP の隣、または pm-ai-package-release 直下）。"
                            : "ローカル版が正本以上で、デスクトップ JAR も一致（更新不要）。";
            appendLog(
                    "[startup] 自動バージョンアップはスキップ: "
                            + reason
                            + " 正本="
                            + cv.map(BigDecimal::toPlainString).orElse("（なし）")
                            + " ローカル="
                            + lv.map(BigDecimal::toPlainString).orElse("（なし・0扱い）")
                            + (buildManifest.map(m -> " " + m.summaryForLog()).orElse(""))
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
                        ? "ZIP を展開し、pm-ai-data とデスクトップ本体（PMD.exe・app・runtime）を更新します。\n"
                                + "本体更新後は自動的にアプリを再起動します。"
                        : "正本から pm-ai-data を同期します（ZIP が無い場合はデスクトップ本体は更新しません）。";
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        initDialogOwnerIfSceneReady(confirm);
        applyAlertStylesheetsFromOwner(confirm);
        confirm.setTitle("自動バージョンアップ");
        confirm.setHeaderText(null);
        confirm.setContentText(
                "正本の更新があります（版 "
                        + canonVerStr
                        + "、ローカル "
                        + localVerStr
                        + "）。\n"
                        + syncHint
                        + "\n実行してよいですか？");
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            appendLog("[startup] ポータル同期はユーザー操作によりスキップしました（版 " + canonVerStr + " → 保留）。");
            return;
        }

        PortableBundleUpgradeUiSnapshot.capture(collectDesktopSession());

        final FactorySite upgradeFactorySite = resolveFactorySiteForPortableUpgrade(Optional.of(canonical));
        deferOperatorPromptForPortableUpgrade.set(true);

        selectMainShellTab(MainShellTabId.RUN);
        mainRunTabController.prepareRunTabForPortableBundleSync();

        PortableBundleUpgradeLog upgradeFileLog = null;
        try {
            upgradeFileLog = PortableBundleUpgradeLog.open(cwd, localData);
        } catch (IOException logOpenEx) {
            appendLog(
                    "[startup] バージョンアップログファイルを作成できません: "
                            + logOpenEx.getMessage());
        }
        final PortableBundleUpgradeLog fileLog = upgradeFileLog;

        String startBanner =
                "[startup] --- ポータルバージョンアップ開始（正本 "
                        + canonVerStr
                        + " → ローカル "
                        + localVerStr
                        + "）---";
        appendLog(startBanner);
        fileLogLine(fileLog, startBanner);
        if (fileLog != null) {
            String logPathMsg =
                    "[startup] バージョンアップログファイル: "
                            + PortableBundleSelfUpdater.safePathForLog(fileLog.logFile());
            appendLog(logPathMsg);
            fileLogLine(fileLog, logPathMsg);
        }
        String sourceMsg =
                "[startup] 正本: "
                        + PortableBundleSelfUpdater.safePathForLog(canonical)
                        + (upgradeZip.isPresent()
                                ? " / ZIP: "
                                        + PortableBundleSelfUpdater.safePathForLog(upgradeZip.get())
                                : " / フォルダ同期");
        appendLog(sourceMsg);
        fileLogLine(fileLog, sourceMsg);

        final boolean zipUpgradeMode = upgradeZip.isPresent();
        Stage wait = new Stage();
        wait.initModality(Modality.APPLICATION_MODAL);
        if (primaryStage != null && primaryStage.getScene() != null) {
            wait.initOwner(primaryStage);
        }
        wait.setTitle("自動バージョンアップ");
        wait.setMinWidth(580);
        wait.setMinHeight(zipUpgradeMode ? 420 : 320);
        VBox root = new VBox(14);
        root.setAlignment(Pos.CENTER_LEFT);
        root.setStyle("-fx-padding: 24;");
        String waitHint =
                fileLog != null
                        ? "正本から pm-ai-data とデスクトップ本体を更新しています…\n詳細は「実行・ログ」タブとログファイルに記録します。"
                        : "正本から pm-ai-data とデスクトップ本体を更新しています…\n詳細は「実行・ログ」タブに追記されます。";
        Label msg = new Label(waitHint);
        msg.setWrapText(true);
        msg.setMaxWidth(520);

        Label downloadCaption = new Label("① 正本ZIPの取得（共有フォルダ → ローカル）");
        ProgressBar downloadBar = new ProgressBar(0);
        downloadBar.setMaxWidth(Double.MAX_VALUE);
        downloadBar.setPrefWidth(520);
        Label extractCaption = new Label("② ZIPの展開");
        ProgressBar extractBar = new ProgressBar(0);
        extractBar.setMaxWidth(Double.MAX_VALUE);
        extractBar.setPrefWidth(520);
        Label syncCaption = new Label("③ pm-ai-data へのファイル同期");
        ProgressBar syncBar = new ProgressBar(0);
        syncBar.setMaxWidth(Double.MAX_VALUE);
        syncBar.setPrefWidth(520);
        Label desktopCaption = new Label("④ デスクトップ本体のステージング（PMD.exe・app・runtime）");
        ProgressBar desktopBar = new ProgressBar(0);
        desktopBar.setMaxWidth(Double.MAX_VALUE);
        desktopBar.setPrefWidth(520);

        VBox downloadBox = new VBox(6, downloadCaption, downloadBar);
        VBox extractBox = new VBox(6, extractCaption, extractBar);
        VBox syncBox = new VBox(6, syncCaption, syncBar);
        VBox desktopBox = new VBox(6, desktopCaption, desktopBar);
        if (!zipUpgradeMode) {
            downloadBox.setManaged(false);
            downloadBox.setVisible(false);
            extractBox.setManaged(false);
            extractBox.setVisible(false);
            desktopBox.setManaged(false);
            desktopBox.setVisible(false);
            downloadBar.setProgress(1);
            extractBar.setProgress(1);
            desktopBar.setProgress(1);
        } else {
            downloadBar.setProgress(-1);
            extractBar.setProgress(0);
            desktopBar.setProgress(0);
        }
        syncBar.setProgress(-1);

        Label detail = new Label("準備中…");
        detail.setWrapText(true);
        detail.setMaxWidth(520);
        detail.getStyleClass().add("pm-portable-sync-detail");

        root.getChildren().addAll(msg, downloadBox, extractBox, syncBox, desktopBox, detail);
        VBox.setVgrow(downloadBar, Priority.NEVER);
        Scene waitScene = new Scene(root, 580, zipUpgradeMode ? 420 : 320);
        wait.setScene(waitScene);
        if (primaryStage != null && primaryStage.getScene() != null) {
            waitScene.getStylesheets().setAll(primaryStage.getScene().getStylesheets());
        }
        wait.show();

        final AtomicLong lastProgressUiNanos = new AtomicLong(0);
        PortableBundleUpgradeProgress.Listener upgradeProgress =
                (phase, done, total, detailLine) -> {
                    boolean phaseEdge = done <= 0 || (total > 0 && done >= total);
                    if (phaseEdge && fileLog != null) {
                        fileLog.appendLine(
                                "[progress] "
                                        + phase
                                        + " "
                                        + done
                                        + "/"
                                        + (total > 0 ? total : "?"));
                    }
                    long now = System.nanoTime();
                    boolean force =
                            phaseEdge
                                    || detailLine != null && !detailLine.isBlank();
                    if (!force && now - lastProgressUiNanos.get() < 50_000_000L) {
                        return;
                    }
                    lastProgressUiNanos.set(now);
                    Platform.runLater(
                            () ->
                                    applyPortableUpgradeProgressToBars(
                                            zipUpgradeMode,
                                            phase,
                                            done,
                                            total,
                                            detailLine,
                                            downloadCaption,
                                            downloadBar,
                                            extractCaption,
                                            extractBar,
                                            syncCaption,
                                            syncBar,
                                            desktopCaption,
                                            desktopBar,
                                            detail));
                };

        final Path[] localZipHolder = new Path[1];
        final Path[] extractedHolder = new Path[1];
        final AtomicBoolean deferredDesktopRelaunch = new AtomicBoolean();
        final AtomicInteger filesSynced = new AtomicInteger();
        Consumer<String> portableSyncLog =
                line -> {
                    if (line != null && line.contains("同期: ")) {
                        filesSynced.incrementAndGet();
                    }
                    fileLogLine(fileLog, line);
                    mainRunTabController.appendPortableBundleSyncLog(line);
                };
        Task<Void> task =
                new Task<>() {
                    @Override
                    protected Void call() throws Exception {
                        Path syncSource;
                        Optional<Path> zipForSync =
                                PortableBundleSelfUpdater.resolveEffectiveUpgradeZip(canonical);
                        if (zipForSync.isPresent()) {
                            Path remoteZip = zipForSync.get();
                            localZipHolder[0] =
                                    PortableBundleSelfUpdater.copyUpgradeZipToLocal(
                                            remoteZip, portableSyncLog, upgradeProgress);
                            Path tmp =
                                    PortableBundleSelfUpdater.extractUpgradeZipToTempDirectory(
                                            localZipHolder[0], portableSyncLog, upgradeProgress);
                            extractedHolder[0] = tmp;
                            syncSource = tmp.resolve("pm-ai-data");
                            if (!Files.isDirectory(syncSource)) {
                                throw new IOException(
                                        "ZIP 内に pm-ai-data フォルダがありません: " + remoteZip);
                            }
                        } else {
                            portableSyncLog.accept(
                                    PortableBundleSelfUpdater.PORTABLE_SYNC_LOG_PREFIX
                                            + "正本フォルダから同期: "
                                            + PortableBundleSelfUpdater.safePathForLog(
                                                    PortableBundleSelfUpdater.resolveSyncSourceRoot(
                                                            canonical)));
                            syncSource = PortableBundleSelfUpdater.resolveSyncSourceRoot(canonical);
                        }
                        PortableBundleSelfUpdater.syncFromCanonical(
                                syncSource, localData, portableSyncLog, upgradeProgress);
                        PortableBundleSelfUpdater.copyOuterVersionTxtToLocal(canonical, cwd, localData);

                        Optional<Path> desktopBundleRoot = Optional.empty();
                        if (extractedHolder[0] != null) {
                            desktopBundleRoot =
                                    PortableBundleSelfUpdater.resolveDesktopBundleRoot(
                                            extractedHolder[0]);
                        } else if (PortableBundleSelfUpdater.hasDesktopInstallLayout(canonical)) {
                            desktopBundleRoot = Optional.of(canonical);
                        }
                        if (desktopBundleRoot.isPresent()) {
                            Path staging = PortableBundlePendingUpdate.defaultStagingDirectory();
                            PortableBundleSelfUpdater.stageDesktopBundleForRelaunch(
                                    desktopBundleRoot.get(), staging, portableSyncLog);
                            deferredDesktopRelaunch.set(true);
                        }
                        return null;
                    }
                };
        task.setOnSucceeded(
                e -> {
                    mainRunTabController.flushPortableBundleSyncLog();
                    if (localZipHolder[0] != null) {
                        try {
                            Files.deleteIfExists(localZipHolder[0]);
                        } catch (IOException ignored) {
                            /* best-effort */
                        }
                    }
                    if (extractedHolder[0] != null) {
                        PortableBundleSelfUpdater.deleteDirectoryRecursive(
                                extractedHolder[0], portableSyncLog);
                    }
                    wait.close();
                    if (deferredDesktopRelaunch.get()) {
                        try {
                            /* 環境変数初期化は再起動後（新 JAR の ui_ref）で強制実行する */
                            PortableBundleUpgradeFollowUp.writePending(
                                    cwd, canonVerStr, upgradeFactorySite);
                            appendLog(
                                    "[startup] バージョンアップ後の再起動待ち: 環境変数初期化記録を無効化しました。再起動後に強制初期化します。");
                            showPortableUpgradeDeferredRestartDialog(canonVerStr);
                            long pid = ProcessHandle.current().pid();
                            Path staging = PortableBundlePendingUpdate.defaultStagingDirectory();
                            PortableBundlePendingUpdate.write(
                                    cwd, staging, canonVerStr, pid, canonical);
                            PortableBundleUpdateLauncher.launchDeferredDesktopApply(
                                    cwd,
                                    staging,
                                    pid,
                                    canonVerStr,
                                    canonical,
                                    this::appendLog);
                            appendLog(
                                    "[startup] デスクトップ本体を適用するため終了します（pmd-apply-portable-update.ps1 が再起動します）。"
                                            + " 環境変数・グローバル設定は反映済みです。");
                            fileLogLine(fileLog, "[startup] deferred desktop apply launched");
                            if (fileLog != null) {
                                fileLog.close(true, "deferred desktop apply");
                            }
                            suppressCloseConfirmation = true;
                            Platform.exit();
                        } catch (IOException ex) {
                            appendLog(
                                    "[startup] デスクトップ本体の終了後更新の起動に失敗: "
                                            + ex.getMessage());
                            fileLogLine(fileLog, "[startup] deferred launch failed: " + ex.getMessage());
                            finishPortableUpgradeWithFactorySitePrompt(
                                    cwd,
                                    localData,
                                    filesSynced.get(),
                                    fileLog,
                                    "（本体再起動の起動に失敗したため、同期直後に環境を反映）",
                                    Optional.of(canonical));
                        }
                        return;
                    }
                    applyPortableUpgradeBundledPolicyFromPmAiData(localData);
                    applyBundledPortableDefaultsIfPresent();
                    String doneBanner =
                            "[startup] --- ポータルバージョンアップ完了（同期ファイル約 "
                                    + filesSynced.get()
                                    + " 件。上記 [portable-sync] 行を参照）---";
                    appendLog(doneBanner);
                    fileLogLine(fileLog, doneBanner);
                    if (fileLog != null) {
                        fileLog.close(
                                true, "同期ファイル約 " + filesSynced.get() + " 件");
                    }
                    applyRepoFolderPathNormalization();
                    finishPortableUpgradeWithFactorySitePrompt(
                            cwd,
                            localData,
                            filesSynced.get(),
                            fileLog,
                            "（デスクトップ本体の変更が無いため再起動は不要です）",
                            Optional.of(canonical));
                });
        task.setOnFailed(
                e -> {
                    deferOperatorPromptForPortableUpgrade.set(false);
                    mainRunTabController.flushPortableBundleSyncLog();
                    if (localZipHolder[0] != null) {
                        try {
                            Files.deleteIfExists(localZipHolder[0]);
                        } catch (IOException ignored) {
                            /* best-effort */
                        }
                    }
                    if (extractedHolder[0] != null) {
                        PortableBundleSelfUpdater.deleteDirectoryRecursive(
                                extractedHolder[0], portableSyncLog);
                    }
                    wait.close();
                    Throwable ex = task.getException();
                    String errorDetail = ex != null ? ex.getMessage() : "不明なエラー";
                    appendLog("[startup] ポータル同期に失敗: " + errorDetail);
                    fileLogLine(fileLog, "[startup] ポータル同期に失敗: " + errorDetail);
                    if (fileLog != null) {
                        if (ex != null) {
                            fileLog.appendThrowable("portable-sync task", ex);
                        }
                        fileLog.close(false, errorDetail);
                    }
                    Alert er = new Alert(AlertType.WARNING);
                    initDialogOwnerIfSceneReady(er);
                    applyAlertStylesheetsFromOwner(er);
                    er.setTitle("自動バージョンアップ");
                    er.setHeaderText(null);
                    er.setContentText("正本からの同期に失敗しました。\n" + errorDetail);
                    er.showAndWait();
                    /* VU 確認後にスキップしていた起動時案内を、失敗時のみ通常どおり出す */
                    maybePromptRequestFormOriginalDirAtStartup();
                    maybePromptOperatorUserAtStartup();
                });
        Thread t = new Thread(task, "pm-ai-portable-sync");
        t.setDaemon(true);
        t.start();
    }

    private int refreshPersonBadgeSkillsMembersFromMaster() {
        if (ganttPersonBadgeDesignTabController != null) {
            return ganttPersonBadgeDesignTabController.reloadSkillsMembersAfterMasterEnvChange();
        }
        return 0;
    }

    private void applyPortableUpgradeBundledPolicyFromPmAiData(Path localData) {
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
    }

    /**
     * ポータル同期後の工場既定反映と環境変数初期化の強制実行。デスクトップ本体再起動後は {@link
     * PortableBundleUpgradeFollowUp} 経由でここだけ再実行する。
     *
     * <p>工場選択ダイアログは出さず、アップデート前の利用工場（または正本 UNC からの推定）を維持する。操作者選択もスキップし、
     * 前回選択の復元のみ試みる。{@code init_setting} のグローバル設定を適用したうえで、環境変数タブを
     * {@link #applyEnvRowsFullBundledResetAndPersist} で ui_ref 既定へ強制初期化する（工場ワークスペースに保存されていた環境変数行は復元しない）。
     */
    private void finishPortableUpgradeWithFactorySitePrompt(
            Path cwd,
            Path localData,
            int filesSyncedApprox,
            PortableBundleUpgradeLog fileLog,
            String completionNoteSuffix,
            Optional<Path> canonicalOpt) {
        applyPortableUpgradeBundledPolicyFromPmAiData(localData);
        FactorySite siteAfterUpgrade = resolveFactorySiteForPortableUpgrade(canonicalOpt);
        GlobalInitSettingTarget.save(siteAfterUpgrade);
        restoreOperatorAfterPortableUpgrade(siteAfterUpgrade);
        skipOperatorPromptAfterPortableUpgrade.set(true);
        deferOperatorPromptForPortableUpgrade.set(false);
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        applyFactoryScopedGlobalAndEnvReset(siteAfterUpgrade, true);
        applyBundledPortableDefaultsIfPresent();
        Optional<FactorySiteWorkspaceSnapshot> workspace = Optional.empty();
        if (!operator.isBlank()) {
            FactorySiteWorkspaceMigrator.migrateIfNeeded(
                    operator,
                    siteAfterUpgrade,
                    snapshotUiEnvRows(),
                    DesktopSessionStateStore.load(),
                    collectUiEnv());
            FactorySiteWorkspaceStore.warmMemoryCacheFromDisk(operator);
            workspace = FactorySiteWorkspaceStore.load(operator, siteAfterUpgrade);
        }
        if (workspace.isPresent()) {
            applyFactoryWorkspaceSessionFragment(workspace.get(), true);
        }
        applyPortableUpgradeShellUiSnapshotIfPresent();
        persistOperatorWorkspaceForEnvInitBaseline(siteAfterUpgrade);
        recordEnvInitializationBaseline();
        int badgeMembers = refreshPersonBadgeSkillsMembersFromMaster();
        DesktopSessionStateStore.save(collectDesktopSessionForGlobalPersistence());
        mainRunTabController.refreshAppVersionLabel();
        mainRunTabController.refreshOpenWorkbookHintLabels();
        factoryOperatorToolbar.refreshFactorySiteLogo();
        refreshFactorySiteComboPresentation();
        PortableBundleUpgradeFollowUp.clear();
        refreshEnvVarsInitializedAtToolbarLabel();
        applyRunTabGating();
        String completion =
                "[startup] ポータル同期が完了しました（version.txt・pm-ai-data／init_setting をリポジトリへ反映）。"
                        + "環境変数を強制初期化し、タブ配置を維持して反映しました。"
                        + " 工場既定: "
                        + siteAfterUpgrade.displayLabelJa()
                        + "（アップデート前の利用工場を維持）。"
                        + (completionNoteSuffix != null ? completionNoteSuffix : "");
        appendLog(completion);
        fileLogLine(fileLog, completion);
        if (badgeMembers > 0) {
            appendLog(
                    "[startup] 担当バッジ: skills メンバーをマスタから再読込しました（"
                            + badgeMembers
                            + " 名）。");
        }
        if (filesSyncedApprox > 0 && fileLog != null) {
            fileLog.appendLine(
                    "[startup] finishPortableUpgrade filesSyncedApprox=" + filesSyncedApprox);
        }
    }

    /**
     * ポータル自動バージョンアップ時に利用工場を決める。環境タブの UNC 推定 → 永続ファイル → 正本パス の順。
     * いずれも判定不能のときのみ {@link FactorySite#KONAN}。
     */
    private FactorySite resolveFactorySiteForPortableUpgrade(Optional<Path> canonicalOpt) {
        Optional<FactorySite> fromFollowUp =
                PortableBundleUpgradeFollowUp.readIfPresent()
                        .flatMap(PortableBundleUpgradeFollowUp::factorySiteOrEmpty);
        if (fromFollowUp.isPresent()) {
            return fromFollowUp.get();
        }
        Optional<FactorySite> fromEnv = FactorySite.inferFromUiEnv(collectUiEnv());
        if (fromEnv.isPresent()) {
            return fromEnv.get();
        }
        FactorySite stored = GlobalInitSettingTarget.load();
        if (stored != null) {
            return stored;
        }
        Optional<FactorySite> fromCanonical =
                canonicalOpt.flatMap(FactorySite::inferFromPortableBundleInitSetting);
        if (fromCanonical.isEmpty() && canonicalOpt.isPresent()) {
            fromCanonical =
                    FactorySite.inferFromPortableBundleSourceValue(canonicalOpt.get().toString());
        }
        return fromCanonical.orElse(FactorySite.KONAN);
    }

    /** バージョンアップ後: 操作者選択ダイアログは出さず、前回選択の復元のみ試みる。 */
    private void restoreOperatorAfterPortableUpgrade(FactorySite site) {
        FactoryOperatorUserStore.configureFromUi(collectUiEnv(), site);
        try {
            if (FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(site)) {
                appendLog(
                        "[startup] 操作者: "
                                + FactoryOperatorUserStore.sessionOperatorName()
                                + " （"
                                + site.displayLabelJa()
                                + "・バージョンアップ後に前回選択を復元）"
                                + (FactoryOperatorUserStore.isGuestOperator(
                                                FactoryOperatorUserStore.sessionOperatorName())
                                        ? " ※サマリ Excel 生成不可"
                                        : ""));
            } else {
                appendLog(
                        "[startup] 操作者: バージョンアップ後は選択を省略しました（後から実行・ログタブ等で選べます）。");
            }
        } catch (IOException ex) {
            appendLog(
                    "[startup] 操作者の復元をスキップ: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
        refreshMainRunTabOperatorLabel();
    }

    private static void fileLogLine(PortableBundleUpgradeLog fileLog, String line) {
        if (fileLog != null && line != null && !line.isBlank()) {
            fileLog.appendLine(line);
        }
    }

    private static void applyPortableUpgradeProgressToBars(
            boolean zipUpgradeMode,
            PortableBundleUpgradeProgress.Phase phase,
            long done,
            long total,
            String detailLine,
            Label downloadCaption,
            ProgressBar downloadBar,
            Label extractCaption,
            ProgressBar extractBar,
            Label syncCaption,
            ProgressBar syncBar,
            Label desktopCaption,
            ProgressBar desktopBar,
            Label detail) {
        if (zipUpgradeMode) {
            switch (phase) {
                case DOWNLOAD -> {
                    applyProgressBarValue(downloadBar, done, total);
                    downloadCaption.setText(
                            progressCaption(
                                    "① 正本ZIPの取得（共有フォルダ → ローカル）", done, total));
                }
                case EXTRACT -> {
                    downloadBar.setProgress(1.0);
                    downloadCaption.setText("① 正本ZIPの取得（完了）");
                    applyProgressBarValue(extractBar, done, total);
                    extractCaption.setText(progressCaption("② ZIPの展開", done, total));
                }
                case SYNC_PM_AI_DATA -> {
                    downloadBar.setProgress(1.0);
                    downloadCaption.setText("① 正本ZIPの取得（完了）");
                    extractBar.setProgress(1.0);
                    extractCaption.setText("② ZIPの展開（完了）");
                    applyProgressBarValue(syncBar, done, total);
                    syncCaption.setText(progressCaption("③ pm-ai-data へのファイル同期", done, total));
                }
                case SYNC_DESKTOP -> {
                    downloadBar.setProgress(1.0);
                    downloadCaption.setText("① 正本ZIPの取得（完了）");
                    extractBar.setProgress(1.0);
                    extractCaption.setText("② ZIPの展開（完了）");
                    syncBar.setProgress(1.0);
                    syncCaption.setText("③ pm-ai-data へのファイル同期（完了）");
                    applyProgressBarValue(desktopBar, done, total);
                    desktopCaption.setText(
                            progressCaption("④ デスクトップ本体のステージング", done, total));
                }
                default -> {
                    /* not reached */
                }
            }
        } else {
            if (phase == PortableBundleUpgradeProgress.Phase.SYNC_DESKTOP) {
                applyProgressBarValue(desktopBar, done, total);
                desktopCaption.setText(progressCaption("④ デスクトップ本体のステージング", done, total));
            } else {
                applyProgressBarValue(syncBar, done, total);
                syncCaption.setText(progressCaption("③ pm-ai-data へのファイル同期", done, total));
            }
        }
        if (detailLine != null && !detailLine.isBlank()) {
            String shortLine =
                    detailLine.length() > 160 ? detailLine.substring(0, 157) + "…" : detailLine;
            detail.setText(shortLine);
        } else if (phase == PortableBundleUpgradeProgress.Phase.DOWNLOAD && total > 0) {
            detail.setText(
                    "取得: "
                            + PortableBundleSelfUpdater.formatByteSize(done)
                            + " / "
                            + PortableBundleSelfUpdater.formatByteSize(total));
        } else if (total > 0) {
            detail.setText(done + " / " + total);
        }
    }

    private static void applyProgressBarValue(ProgressBar bar, long done, long total) {
        if (total > 0) {
            bar.setProgress(Math.min(1.0, done / (double) total));
        } else {
            bar.setProgress(-1);
        }
    }

    private static String progressCaption(String base, long done, long total) {
        if (total <= 0) {
            return base + " …";
        }
        int pct = (int) Math.min(100, (done * 100) / total);
        return base + " (" + pct + "%)";
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
            case AppPaths.KEY_PM_AI_DAILY_REPORT_SOURCE_DIR -> {
                return AppPaths.defaultDailyReportSourceDirForFactory(GlobalInitSettingTarget.load());
            }
            case AppPaths.KEY_PM_AI_ORDER_DETAIL_SOURCE_DIR -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR -> {
                return AppPaths.resolveAladdinMasterDir(u).toString();
            }
            case AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE -> {
                return AppPaths.defaultRequestFormJuchuFileForFactory(GlobalInitSettingTarget.load());
            }
            case AppPaths.KEY_PM_AI_MACHINE_DELIVERY_MANAGEMENT_XLSM -> {
                return AppPaths.defaultMachineDeliveryManagementXlsmForFactory(
                        GlobalInitSettingTarget.load());
            }
            case AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_REQUEST_FORM_TPI_PDF_DIR -> {
                return AppPaths.defaultRequestFormTpiPdfDirForFactory(GlobalInitSettingTarget.load());
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

    /** 廃止した {@link AppPaths#KEY_MASTER_WORKBOOK_FILE} を {@link AppPaths#KEY_PM_AI_MASTER_WORKBOOK} へ移行する。 */
    private void migrateLegacyMasterWorkbookFileEnvRows() {
        if (envRows == null) {
            return;
        }
        String legacy = envTabValueTrimmed(AppPaths.KEY_MASTER_WORKBOOK_FILE);
        if (legacy.isEmpty()) {
            return;
        }
        Optional<String> migrated =
                AppPaths.migrateLegacyMasterWorkbookFileToPmAi(collectUiEnv(), legacy);
        if (migrated.isEmpty()) {
            return;
        }
        for (EnvVarRow row : envRows) {
            String name = row.getName() != null ? row.getName().trim() : "";
            if (!AppPaths.KEY_PM_AI_MASTER_WORKBOOK.equals(name)) {
                continue;
            }
            if (envTabValueTrimmed(AppPaths.KEY_PM_AI_MASTER_WORKBOOK).isEmpty()) {
                row.setValue(migrated.get());
            }
            break;
        }
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
        AppPaths.ensureAllDispatchLookupTablesFromRepoIfMissing(ctx);
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
        return switch (k) {
            case PlanInputTabController.ENV_TASK_PLAN_SHEET ->
                    PlanInputTabController.DEFAULT_PLAN_INPUT_SHEET_NAME;
            case "MASTER_SPEED_SHEET_NAME" -> "speed";
            case "MASTER_SPEED_FIRST_EXCEL_COL" -> "4";
            case AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK ->
                    AppPaths.summarySharedDataDir(u).toString();
            case "RAW_FABRIC_WIDTH_TABLE_PATH" -> {
                Path p = AppPaths.dispatchLookupTablePath(u, AppPaths.DISPATCH_LOOKUP_USED_RAW_WIDTH);
                yield Files.isRegularFile(p)
                        ? p.toAbsolutePath().normalize().toString()
                        : "";
            }
            case "ROLL_UNIT_BY_USED_RAW_TABLE_PATH" -> {
                Path p = AppPaths.dispatchLookupTablePath(u, AppPaths.DISPATCH_LOOKUP_USED_RAW_ROLL);
                yield Files.isRegularFile(p)
                        ? p.toAbsolutePath().normalize().toString()
                        : "";
            }
            case "PRODUCT_WIDTH_TABLE_PATH" -> {
                Path p = AppPaths.dispatchLookupTablePath(u, AppPaths.DISPATCH_LOOKUP_PRODUCT_WIDTH);
                yield Files.isRegularFile(p)
                        ? p.toAbsolutePath().normalize().toString()
                        : "";
            }
            case "PRODUCT_LENGTH_TABLE_PATH" -> {
                Path p = AppPaths.dispatchLookupTablePath(u, AppPaths.DISPATCH_LOOKUP_PRODUCT_LENGTH);
                yield Files.isRegularFile(p)
                        ? p.toAbsolutePath().normalize().toString()
                        : "";
            }
            case "PRODUCT_THICKNESS_TABLE_PATH" -> {
                Path p = AppPaths.dispatchLookupTablePath(u, AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK);
                yield Files.isRegularFile(p)
                        ? p.toAbsolutePath().normalize().toString()
                        : "";
            }
            case "ROLL_UNIT_LENGTH_TABLE_PATH" -> {
                Path p = AppPaths.dispatchLookupTablePath(u, AppPaths.DISPATCH_LOOKUP_PRODUCT_ROLL);
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
