package jp.co.pm.ai.desktop;

import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.EnumSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Function;

import javafx.animation.KeyFrame;
import javafx.animation.PauseTransition;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextField;
import javafx.scene.control.TextInputControl;
import javafx.scene.control.TitledPane;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.Tooltip;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.actuals.DashboardLoadErrorFormatter;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardBuilder;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.LoadedSources;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.ReloadDecision;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.SourceFingerprint;
import jp.co.pm.ai.desktop.ui.EquipmentStatusCardFactory;
import jp.co.pm.ai.desktop.ui.EquipmentStatusCardFactory.DisplayOptions;
import jp.co.pm.ai.desktop.ui.EquipmentStatusDashboardAppearanceApplier;
import jp.co.pm.ai.desktop.ui.EquipmentStatusDashboardAppearancePanel;
import jp.co.pm.ai.desktop.ui.EquipmentStatusDashboardFilter;
import jp.co.pm.ai.desktop.ui.EquipmentStatusDashboardFilter.SortOrder;
import jp.co.pm.ai.desktop.ui.EquipmentStatusDashboardFilter.StatusCounts;
import jp.co.pm.ai.desktop.ui.EquipmentStatusFullscreenStage;

/** メインシェル「ダッシュボード」タブ。 */
public final class EquipmentStatusDashboardTabController {

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");
    private static final DateTimeFormatter TIME_FMT = DateTimeFormatter.ofPattern("HH:mm:ss");
    private static final String ERROR_STYLE_CLASS = "pm-equipment-status-error";
    private static final String STALE_CARDS_STYLE_CLASS = "pm-equipment-status-cards-stale";

    /** 見た目設定の連続操作でカード全再生成が何度も走らないよう待つ時間。 */
    private static final Duration RENDER_DEBOUNCE = Duration.millis(180);

    /** セッション JSON への同期書込は描画よりさらに緩く間引く。 */
    private static final Duration PERSIST_DEBOUNCE = Duration.millis(600);

    private MainShellController shell;

    @FXML private Button reloadButton;
    @FXML private Button actualTodayButton;
    @FXML private Button planTodayButton;
    @FXML private ToggleButton fullscreenToggle;
    @FXML private CheckBox autoRefreshCheckBox;
    @FXML private Spinner<Integer> autoRefreshIntervalSpinner;
    @FXML private Label nextRefreshLabel;
    @FXML private DatePicker actualDatePicker;
    @FXML private DatePicker planDatePicker;
    @FXML private Label lastUpdatedLabel;
    @FXML private ProgressIndicator loadingIndicator;
    @FXML private Label loadingStatusLabel;
    @FXML private CheckBox showAladdinCheckBox;
    @FXML private CheckBox showDispatchCheckBox;
    @FXML private ToggleButton filterStoppedToggle;
    @FXML private ToggleButton filterRunningToggle;
    @FXML private ToggleButton filterCompletedToggle;
    @FXML private TextField machineFilterField;
    @FXML private ComboBox<String> sortOrderCombo;
    @FXML private Label machineCountLabel;
    @FXML private Label loadStatsLabel;
    @FXML private Label sourceSummaryLabel;
    @FXML private HBox pastDateBanner;
    @FXML private Label pastDateBannerLabel;
    @FXML private HBox staleBanner;
    @FXML private Label staleBannerLabel;
    @FXML private TitledPane appearancePane;
    @FXML private VBox appearanceControlsHost;
    @FXML private BorderPane tabRoot;
    @FXML private ScrollPane cardScrollPane;
    @FXML private FlowPane cardFlowPane;
    private HBox cardFlowHost;
    @FXML private VBox emptyStatePane;

    private final EquipmentStatusFullscreenStage fullscreenStage = new EquipmentStatusFullscreenStage();

    private final ExecutorService reloadPool =
            Executors.newSingleThreadExecutor(
                    r -> {
                        Thread t = new Thread(r, "equipment-status-dashboard-reload");
                        t.setDaemon(true);
                        return t;
                    });

    private Timeline autoRefreshTimeline;
    private int autoRefreshRemainingSec;
    private PauseTransition renderDebounce;
    private PauseTransition persistDebounce;
    private Task<ReloadDecision> activeReloadTask;
    private final AtomicBoolean tabActive = new AtomicBoolean(false);
    private final AtomicBoolean suppressUiEvents = new AtomicBoolean(false);
    private final AtomicBoolean loading = new AtomicBoolean(false);

    private LoadedSources cachedSources;
    private SourceFingerprint loadedFingerprint;
    private List<EquipmentMachineStatus> currentStatuses = List.of();
    private List<EquipmentMachineStatus> visibleStatuses = List.of();
    private String lastLoadErrorDetail = "";
    private LocalTime lastSuccessAt;
    private boolean tabCardsStale;

    private LocalDate actualDate = LocalDate.now();
    private LocalDate planDate = LocalDate.now();

    /** 実績日を手動で選ぶまでは当日に追従する（壁掛け運用で日付が変わっても古い日を出し続けないため）。 */
    private boolean followToday = true;

    private EquipmentStatusDashboardAppearancePrefs appearancePrefs =
            EquipmentStatusDashboardAppearancePrefs.defaults();
    private EquipmentStatusDashboardAppearancePanel appearancePanel;

    @FXML
    private void initialize() {
        LocalDate today = LocalDate.now();
        actualDate = today;
        planDate = today;
        if (actualDatePicker != null) {
            actualDatePicker.setValue(today);
            actualDatePicker
                    .valueProperty()
                    .addListener(
                            (obs, prev, cur) -> {
                                if (suppressUiEvents.get() || cur == null) {
                                    return;
                                }
                                followToday = false;
                                setActualDate(cur);
                            });
        }
        if (planDatePicker != null) {
            planDatePicker.setValue(today);
            planDatePicker
                    .valueProperty()
                    .addListener(
                            (obs, prev, cur) -> {
                                if (suppressUiEvents.get() || cur == null) {
                                    return;
                                }
                                planDate = cur;
                                rebuildFromCache();
                            });
        }

        appearancePanel =
                new EquipmentStatusDashboardAppearancePanel(
                        appearancePrefs, p -> onAppearanceChanged(p));
        if (appearanceControlsHost != null) {
            appearanceControlsHost.getChildren().setAll(appearancePanel.buildRoot());
        }

        if (cardScrollPane != null && cardFlowPane != null) {
            cardFlowHost = new HBox(cardFlowPane);
            cardScrollPane.setContent(cardFlowHost);
            cardScrollPane.viewportBoundsProperty()
                    .addListener((o, a, b) -> applyFlowLayout(appearancePrefs, b.getWidth()));
        }

        fullscreenStage.setOnClose(
                () ->
                        Platform.runLater(
                                () -> {
                                    if (fullscreenToggle != null) {
                                        fullscreenToggle.setSelected(false);
                                    }
                                    if (tabCardsStale) {
                                        refreshView();
                                    }
                                    updateAutoRefreshTimer();
                                }));
        fullscreenStage.setOnReloadRequest(() -> reloadFromSources(true));
        fullscreenStage.setOnAdjustActualDateDays(
                days -> {
                    if (days == 0) {
                        return;
                    }
                    LocalDate base = actualDate != null ? actualDate : LocalDate.now();
                    followToday = false;
                    setActualDate(base.plusDays(days));
                });
        fullscreenStage.setOnAdjustPlanDateDays(
                days -> {
                    if (days == 0) {
                        return;
                    }
                    adjustPlanDateByDays(days);
                });

        if (tabRoot != null) {
            tabRoot.setFocusTraversable(true);
            tabRoot.addEventFilter(KeyEvent.KEY_PRESSED, this::onDashboardKeyPressed);
        }

        initializeViewControls();
        initializeTooltips();

        if (autoRefreshCheckBox != null) {
            autoRefreshCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, selected) -> {
                                updateAutoRefreshTimer();
                                if (selected && tabActive.get()) {
                                    reloadFromSources();
                                }
                            });
        }
        if (autoRefreshIntervalSpinner != null) {
            autoRefreshIntervalSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(
                            DesktopSessionState.MIN_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                            DesktopSessionState.MAX_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                            DesktopSessionState.DEFAULT_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC));
            EquipmentStatusDashboardAppearancePanel.commitEditorTextOnFocusLost(
                    autoRefreshIntervalSpinner);
            autoRefreshIntervalSpinner
                    .valueProperty()
                    .addListener((o, a, b) -> updateAutoRefreshTimer());
        }

        refreshView();
    }

    private void initializeViewControls() {
        Runnable displayRefresh = this::refreshView;
        if (showAladdinCheckBox != null) {
            showAladdinCheckBox.selectedProperty().addListener((o, a, b) -> displayRefresh.run());
        }
        if (showDispatchCheckBox != null) {
            showDispatchCheckBox.selectedProperty().addListener((o, a, b) -> displayRefresh.run());
        }
        for (ToggleButton t : statusFilterToggles()) {
            t.selectedProperty().addListener((o, a, b) -> displayRefresh.run());
        }
        if (machineFilterField != null) {
            machineFilterField.textProperty().addListener((o, a, b) -> displayRefresh.run());
        }
        if (sortOrderCombo != null) {
            for (SortOrder order : SortOrder.values()) {
                sortOrderCombo.getItems().add(order.label());
            }
            sortOrderCombo.setValue(SortOrder.MACHINE_NAME.label());
            sortOrderCombo.valueProperty().addListener((o, a, b) -> displayRefresh.run());
        }
    }

    private void initializeTooltips() {
        installTooltip(reloadButton, "実績・アラジン・配台のソースを読み直す");
        installTooltip(fullscreenToggle, "壁掛けディスプレイ向けの全画面表示に切り替える");
        installTooltip(actualTodayButton, "実績日と予定日をまとめて当日に戻す");
        installTooltip(planTodayButton, "予定日だけを当日に戻す（Shift + ← / → でも前後できる）");
        installTooltip(machineFilterField, "機械名の部分一致で絞り込む（全角・半角は区別しない）");
        installTooltip(sortOrderCombo, "カードの並び順");
        installTooltip(filterStoppedToggle, "停機の機械だけを表示（もう一度押すと解除）");
        installTooltip(filterRunningToggle, "稼働中の機械だけを表示（もう一度押すと解除）");
        installTooltip(filterCompletedToggle, "予定達成の機械だけを表示（もう一度押すと解除）");
        installTooltip(nextRefreshLabel, "自動更新の次回実行までの残り時間");
    }

    private static void installTooltip(javafx.scene.control.Control control, String text) {
        if (control != null) {
            control.setTooltip(new Tooltip(text));
        }
    }

    private List<ToggleButton> statusFilterToggles() {
        List<ToggleButton> toggles = new java.util.ArrayList<>(3);
        if (filterStoppedToggle != null) {
            toggles.add(filterStoppedToggle);
        }
        if (filterRunningToggle != null) {
            toggles.add(filterRunningToggle);
        }
        if (filterCompletedToggle != null) {
            toggles.add(filterCompletedToggle);
        }
        return toggles;
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    public void applyDashboardSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        suppressUiEvents.set(true);
        try {
            LocalDate today = LocalDate.now();
            actualDate = s.resolveEquipmentStatusDashboardActualDate(today);
            planDate = s.resolveEquipmentStatusDashboardPlanDate(today);
            if (actualDatePicker != null) {
                actualDatePicker.setValue(actualDate);
            }
            if (planDatePicker != null) {
                planDatePicker.setValue(planDate);
            }
            if (autoRefreshCheckBox != null) {
                autoRefreshCheckBox.setSelected(s.equipmentStatusDashboardAutoRefreshEnabled());
            }
            if (autoRefreshIntervalSpinner != null) {
                int iv =
                        Math.max(
                                DesktopSessionState.MIN_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                                Math.min(
                                        DesktopSessionState.MAX_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                                        s.equipmentStatusDashboardAutoRefreshIntervalSec()));
                autoRefreshIntervalSpinner.getValueFactory().setValue(iv);
            }
            if (showAladdinCheckBox != null) {
                showAladdinCheckBox.setSelected(s.equipmentStatusDashboardShowAladdinPlans());
            }
            if (showDispatchCheckBox != null) {
                showDispatchCheckBox.setSelected(s.equipmentStatusDashboardShowDispatchPlans());
            }
            appearancePrefs = s.equipmentStatusDashboardAppearance();
            if (appearancePanel != null) {
                appearancePanel.applyPrefs(appearancePrefs);
            }
        } finally {
            suppressUiEvents.set(false);
        }
        refreshView();
    }

    /** アプリ起動時: 実績日・予定日を当日に揃える（前回セッションの日付は復元しない）。 */
    public void resetDashboardDatesToToday() {
        followToday = true;
        setActualDate(LocalDate.now());
    }

    public String snapshotActualDateIso() {
        return actualDate != null ? actualDate.toString() : "";
    }

    public String snapshotPlanDateIso() {
        return planDate != null ? planDate.toString() : "";
    }

    public boolean snapshotAutoRefreshEnabled() {
        return autoRefreshCheckBox == null || autoRefreshCheckBox.isSelected();
    }

    public int snapshotAutoRefreshIntervalSec() {
        if (autoRefreshIntervalSpinner == null || autoRefreshIntervalSpinner.getValue() == null) {
            return DesktopSessionState.DEFAULT_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC;
        }
        return Math.max(
                DesktopSessionState.MIN_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                Math.min(
                        DesktopSessionState.MAX_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                        autoRefreshIntervalSpinner.getValue()));
    }

    public boolean snapshotShowAladdinPlans() {
        return showAladdinCheckBox == null || showAladdinCheckBox.isSelected();
    }

    public boolean snapshotShowDispatchPlans() {
        return showDispatchCheckBox == null || showDispatchCheckBox.isSelected();
    }

    public EquipmentStatusDashboardAppearancePrefs snapshotAppearancePrefs() {
        if (appearancePanel != null) {
            appearancePrefs = appearancePanel.snapshot();
        }
        return appearancePrefs != null
                ? appearancePrefs
                : EquipmentStatusDashboardAppearancePrefs.defaults();
    }

    public void onMainShellTabSelected() {
        tabActive.set(true);
        if (tabRoot != null) {
            tabRoot.requestFocus();
        }
        rolloverDatesIfFollowingToday();
        reloadFromSources();
        updateAutoRefreshTimer();
    }

    public void onMainShellTabDeselected() {
        tabActive.set(false);
        updateAutoRefreshTimer();
    }

    @FXML
    private void onReloadAction() {
        reloadFromSources(true);
    }

    @FXML
    private void onActualTodayAction() {
        followToday = true;
        setActualDate(LocalDate.now());
    }

    @FXML
    private void onPlanTodayAction() {
        setPlanDate(LocalDate.now());
    }

    @FXML
    private void onFullscreenToggleAction() {
        boolean wantFullscreen = fullscreenToggle != null && fullscreenToggle.isSelected();
        if (!wantFullscreen) {
            fullscreenStage.hide();
            updateAutoRefreshTimer();
            return;
        }
        if (shell == null || shell.getPrimaryStage() == null) {
            fullscreenToggle.setSelected(false);
            if (shell != null) {
                shell.appendLog("[dashboard] 全画面表示に失敗: メインウィンドウが未初期化");
            }
            return;
        }
        DisplayOptions opts = currentDisplayOptions();
        EquipmentStatusDashboardAppearancePrefs ap = snapshotAppearancePrefs();
        fullscreenStage.show(
                shell.getPrimaryStage(),
                visibleStatuses,
                opts,
                ap,
                badgeStyleResolver(),
                buildMetaSummary(),
                opts.actualDateLabel(),
                opts.planDateLabel(),
                cachedSources != null,
                lastLoadErrorDetail);
        updateAutoRefreshTimer();
    }

    private void onAppearanceChanged(EquipmentStatusDashboardAppearancePrefs prefs) {
        appearancePrefs =
                prefs != null ? prefs : EquipmentStatusDashboardAppearancePrefs.defaults();
        if (renderDebounce == null) {
            renderDebounce = new PauseTransition(RENDER_DEBOUNCE);
            renderDebounce.setOnFinished(e -> refreshView());
        }
        renderDebounce.playFromStart();
        if (persistDebounce == null) {
            persistDebounce = new PauseTransition(PERSIST_DEBOUNCE);
            persistDebounce.setOnFinished(
                    e -> {
                        if (shell != null) {
                            shell.persistDesktopSessionNow();
                        }
                    });
        }
        persistDebounce.playFromStart();
    }

    private void setActualDate(LocalDate date) {
        if (date == null) {
            return;
        }
        actualDate = date;
        planDate = date;
        suppressUiEvents.set(true);
        try {
            if (actualDatePicker != null) {
                actualDatePicker.setValue(date);
            }
            if (planDatePicker != null) {
                planDatePicker.setValue(date);
            }
        } finally {
            suppressUiEvents.set(false);
        }
        rebuildFromCache();
    }

    private void setPlanDate(LocalDate date) {
        if (date == null) {
            return;
        }
        planDate = date;
        suppressUiEvents.set(true);
        try {
            if (planDatePicker != null) {
                planDatePicker.setValue(date);
            }
        } finally {
            suppressUiEvents.set(false);
        }
        rebuildFromCache();
    }

    private void adjustPlanDateByDays(int days) {
        if (days == 0) {
            return;
        }
        LocalDate base = planDate != null ? planDate : LocalDate.now();
        setPlanDate(base.plusDays(days));
    }

    /** 当日追従中に日付が変わっていたら実績日・予定日を進める。 */
    private void rolloverDatesIfFollowingToday() {
        if (!followToday) {
            return;
        }
        LocalDate today = LocalDate.now();
        if (!today.equals(actualDate)) {
            setActualDate(today);
        }
    }

    private void onDashboardKeyPressed(KeyEvent e) {
        if (!tabActive.get() || fullscreenStage.isShowing()) {
            return;
        }
        if (!e.isShiftDown() || e.isControlDown() || e.isAltDown() || e.isMetaDown()) {
            return;
        }
        int days = arrowDayShift(e.getCode());
        if (days == 0 || skipDashboardShortcutTarget(e.getTarget())) {
            return;
        }
        adjustPlanDateByDays(days);
        e.consume();
    }

    static int arrowDayShift(KeyCode code) {
        return switch (code) {
            case LEFT -> -1;
            case RIGHT -> 1;
            default -> 0;
        };
    }

    static boolean skipDashboardShortcutTarget(Object target) {
        return target instanceof TextInputControl t && t.isEditable();
    }

    private void reloadFromSources() {
        reloadFromSources(false);
    }

    private void reloadFromSources(boolean userInitiated) {
        if (shell == null) {
            return;
        }
        if (activeReloadTask != null && activeReloadTask.isRunning()) {
            return;
        }
        setLoading(true);
        final SourceFingerprint previousFingerprint = loadedFingerprint;
        final boolean haveCache = cachedSources != null;
        Task<ReloadDecision> task =
                new Task<>() {
                    @Override
                    protected ReloadDecision call() {
                        return EquipmentStatusDashboardSourceLoader.loadIfChanged(
                                shell.snapshotUiEnv(), previousFingerprint, haveCache);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    ReloadDecision decision = task.getValue();
                    activeReloadTask = null;
                    setLoading(false);
                    if (decision == null || decision.sourcesUnchanged()) {
                        onSourcesUnchanged(userInitiated);
                        return;
                    }
                    lastLoadErrorDetail = "";
                    lastSuccessAt = LocalTime.now();
                    clearLoadErrorPresentation();
                    loadedFingerprint = decision.fingerprint();
                    cachedSources = decision.sources();
                    updateLastUpdatedLabel("最終更新");
                    rebuildFromCache();
                });
        task.setOnFailed(
                e -> {
                    activeReloadTask = null;
                    setLoading(false);
                    Throwable ex = task.getException();
                    Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
                    String sourceContext = EquipmentStatusDashboardSourceLoader.formatSourceContext(ui);
                    lastLoadErrorDetail = DashboardLoadErrorFormatter.formatDetail(ex);
                    presentLoadError(sourceContext, lastLoadErrorDetail, ex);
                    logLoadError(sourceContext, ex);
                    if (userInitiated && shell != null) {
                        shell.showErrorDialog(
                                "ダッシュボード読込エラー",
                                sourceContext
                                        + "\n\n"
                                        + lastLoadErrorDetail
                                        + "\n\n詳細スタックは「実行・ログ」タブを参照してください。");
                    }
                });
        activeReloadTask = task;
        reloadPool.execute(task);
    }

    /** ソースが変わっていなかったとき。画面に「読込中…」が残らないよう表示を必ず戻す。 */
    private void onSourcesUnchanged(boolean userInitiated) {
        lastSuccessAt = LocalTime.now();
        updateLastUpdatedLabel("最終確認");
        refreshView();
        if (shell != null) {
            shell.appendLog("[dashboard] ソース未変更のため読込を省略");
        }
        if (userInitiated && machineCountLabel != null) {
            machineCountLabel.setText(
                    machineCountLabel.getText() + "（変更なし）");
        }
    }

    private void setLoading(boolean on) {
        loading.set(on);
        if (reloadButton != null) {
            reloadButton.setDisable(on);
        }
        if (loadingIndicator != null) {
            loadingIndicator.setVisible(on);
            loadingIndicator.setManaged(on);
        }
        if (loadingStatusLabel != null) {
            loadingStatusLabel.setVisible(on);
            loadingStatusLabel.setManaged(on);
        }
        if (!on && fullscreenStage.isShowing()) {
            fullscreenStage.setMetaText(buildMetaSummary());
        }
        fullscreenStage.setLoadingVisible(on);
    }

    private void rebuildFromCache() {
        LocalDate actual = actualDate != null ? actualDate : LocalDate.now();
        LocalDate plan = planDate != null ? planDate : LocalDate.now();
        if (cachedSources != null) {
            currentStatuses =
                    EquipmentStatusDashboardBuilder.build(
                            cachedSources.actuals(),
                            cachedSources.aladdin(),
                            cachedSources.dispatch(),
                            actual,
                            plan);
        } else {
            currentStatuses = List.of();
        }
        refreshView();
    }

    /** 絞込・並べ替えを当て直し、ラベル・カード・全画面を現在の状態に揃える。 */
    private void refreshView() {
        EquipmentStatusDashboardAppearancePrefs ap = snapshotAppearancePrefs();
        visibleStatuses =
                EquipmentStatusDashboardFilter.apply(
                        currentStatuses,
                        selectedStatusFilter(),
                        machineFilterField != null ? machineFilterField.getText() : "",
                        selectedSortOrder());
        StatusCounts counts = EquipmentStatusDashboardFilter.countByStatus(currentStatuses);
        updateStatusFilterLabels(counts);
        updateSummaryLabels(
                actualDate != null ? actualDate : LocalDate.now(),
                planDate != null ? planDate : LocalDate.now());
        updateBanners();
        if (fullscreenStage.isShowing()) {
            tabCardsStale = true;
            syncFullscreen(ap);
            return;
        }
        tabCardsStale = false;
        applyFlowLayout(ap, viewportWidth());
        renderCards(ap);
    }

    private void syncFullscreen(EquipmentStatusDashboardAppearancePrefs ap) {
        DisplayOptions opts = currentDisplayOptions();
        fullscreenStage.setHeaderDates(opts.actualDateLabel(), opts.planDateLabel());
        fullscreenStage.rebuildCards(
                visibleStatuses,
                opts,
                ap,
                badgeStyleResolver(),
                opts.actualDateLabel(),
                opts.planDateLabel(),
                cachedSources != null,
                lastLoadErrorDetail);
        fullscreenStage.setMetaText(buildMetaSummary());
    }

    private Set<EquipmentMachineStatus.Status> selectedStatusFilter() {
        Set<EquipmentMachineStatus.Status> selected =
                EnumSet.noneOf(EquipmentMachineStatus.Status.class);
        if (filterStoppedToggle != null && filterStoppedToggle.isSelected()) {
            selected.add(EquipmentMachineStatus.Status.STOPPED);
        }
        if (filterRunningToggle != null && filterRunningToggle.isSelected()) {
            selected.add(EquipmentMachineStatus.Status.RUNNING);
        }
        if (filterCompletedToggle != null && filterCompletedToggle.isSelected()) {
            selected.add(EquipmentMachineStatus.Status.COMPLETED);
        }
        return selected;
    }

    private SortOrder selectedSortOrder() {
        return sortOrderCombo != null
                ? SortOrder.fromLabel(sortOrderCombo.getValue())
                : SortOrder.MACHINE_NAME;
    }

    private void updateStatusFilterLabels(StatusCounts counts) {
        if (filterStoppedToggle != null) {
            filterStoppedToggle.setText("停機 " + counts.stopped());
        }
        if (filterRunningToggle != null) {
            filterRunningToggle.setText("稼働中 " + counts.running());
        }
        if (filterCompletedToggle != null) {
            filterCompletedToggle.setText("予定達成 " + counts.completed());
        }
    }

    private double viewportWidth() {
        return cardScrollPane != null ? cardScrollPane.getViewportBounds().getWidth() : 0;
    }

    private void applyFlowLayout(
            EquipmentStatusDashboardAppearancePrefs ap, double viewportWidth) {
        if (cardFlowPane == null) {
            return;
        }
        boolean fillViewport =
                EquipmentStatusDashboardAppearanceApplier.configureFlowPane(
                        cardFlowPane, ap, false, viewportWidth);
        if (cardFlowHost != null) {
            EquipmentStatusDashboardAppearanceApplier.applyFlowHostLayout(
                    cardFlowHost, cardFlowPane, fillViewport);
        }
        if (cardScrollPane != null) {
            cardScrollPane.setFitToWidth(
                    EquipmentStatusDashboardAppearanceApplier.scrollShouldFitToWidth(ap));
        }
    }

    private void renderCards(EquipmentStatusDashboardAppearancePrefs ap) {
        if (cardFlowPane == null) {
            return;
        }
        double scrollBefore = cardScrollPane != null ? cardScrollPane.getVvalue() : 0;
        cardFlowPane.getChildren().clear();
        DisplayOptions opts = currentDisplayOptions();
        boolean sourcesLoaded = cachedSources != null;
        boolean empty = visibleStatuses == null || visibleStatuses.isEmpty();
        if (emptyStatePane != null) {
            emptyStatePane.getChildren().clear();
            boolean showEmpty = empty && !loading.get();
            emptyStatePane.setVisible(showEmpty);
            emptyStatePane.setManaged(showEmpty);
            if (showEmpty) {
                emptyStatePane
                        .getChildren()
                        .add(
                                EquipmentStatusCardFactory.createEmptyState(
                                        opts.actualDateLabel(),
                                        opts.planDateLabel(),
                                        sourcesLoaded,
                                        false,
                                        lastLoadErrorDetail,
                                        () -> reloadFromSources(true)));
            }
        }
        if (empty) {
            return;
        }
        Function<String, PersonBadgeStyle> resolver = badgeStyleResolver();
        for (EquipmentMachineStatus s : visibleStatuses) {
            cardFlowPane
                    .getChildren()
                    .add(EquipmentStatusCardFactory.createCard(s, opts, ap, resolver, false));
        }
        restoreScrollPosition(scrollBefore);
    }

    /** カードを作り直すと ScrollPane の位置が先頭に戻るため、自動更新で読んでいた場所を失わないよう復元する。 */
    private void restoreScrollPosition(double vvalue) {
        if (cardScrollPane == null || vvalue <= 0) {
            return;
        }
        Platform.runLater(() -> cardScrollPane.setVvalue(vvalue));
    }

    private DisplayOptions currentDisplayOptions() {
        LocalDate actual = actualDate != null ? actualDate : LocalDate.now();
        LocalDate plan = planDate != null ? planDate : LocalDate.now();
        LocalDate today = LocalDate.now();
        return new DisplayOptions(
                showAladdinCheckBox == null || showAladdinCheckBox.isSelected(),
                showDispatchCheckBox == null || showDispatchCheckBox.isSelected(),
                actual.format(DATE_FMT),
                plan.format(DATE_FMT),
                actual.equals(today));
    }

    private Function<String, PersonBadgeStyle> badgeStyleResolver() {
        return shell != null
                ? shell.personBadgeStyleResolverForGantt()
                : (String __) -> PersonBadgeStyle.defaultStyle();
    }

    private void updateSummaryLabels(LocalDate actual, LocalDate plan) {
        if (machineCountLabel != null) {
            machineCountLabel.setText(machineCountText());
        }
        if (sourceSummaryLabel != null && cachedSources != null && lastLoadErrorDetail.isBlank()) {
            String summary =
                    "実績="
                            + cachedSources.actualSourceLabel()
                            + "  アラジン="
                            + cachedSources.aladdinSourceLabel()
                            + "  配台="
                            + cachedSources.dispatchSourceLabel();
            if (cachedSources.loadNotice() != null && !cachedSources.loadNotice().isBlank()) {
                summary += "  ※ " + cachedSources.loadNotice();
            }
            sourceSummaryLabel.setText(summary);
        }
        if (loadStatsLabel != null && lastLoadErrorDetail.isBlank()) {
            loadStatsLabel.setText(
                    cachedSources == null
                            ? ""
                            : EquipmentStatusDashboardSourceLoader.formatLoadStatsSummary(
                                    cachedSources.loadStats()));
        }
    }

    private String machineCountText() {
        if (cachedSources == null) {
            return lastLoadErrorDetail.isBlank() ? "—" : "読込エラー";
        }
        int total = currentStatuses.size();
        int shown = visibleStatuses.size();
        if (total == 0) {
            return "該当 0台（非稼働日の可能性）";
        }
        if (shown != total) {
            return "表示 " + shown + " / " + total + "台";
        }
        return total + "台";
    }

    private void updateBanners() {
        LocalDate today = LocalDate.now();
        LocalDate actual = actualDate != null ? actualDate : today;
        boolean notToday = !actual.equals(today);
        if (pastDateBanner != null) {
            pastDateBanner.setVisible(notToday);
            pastDateBanner.setManaged(notToday);
        }
        if (pastDateBannerLabel != null && notToday) {
            pastDateBannerLabel.setText(
                    (actual.isBefore(today) ? "過去日" : "未来日")
                            + "を表示中: "
                            + actual.format(DATE_FMT)
                            + "（当日ではありません）");
        }
        boolean stale = !lastLoadErrorDetail.isBlank() && cachedSources != null;
        if (staleBanner != null) {
            staleBanner.setVisible(stale);
            staleBanner.setManaged(stale);
        }
        if (staleBannerLabel != null && stale) {
            staleBannerLabel.setText(
                    "更新に失敗しました。"
                            + (lastSuccessAt != null
                                    ? lastSuccessAt.format(TIME_FMT) + " 時点の"
                                    : "以前の")
                            + "データを表示しています — "
                            + lastLoadErrorDetail.replace('\n', ' '));
        }
        if (cardFlowPane != null) {
            boolean marked = cardFlowPane.getStyleClass().contains(STALE_CARDS_STYLE_CLASS);
            if (stale && !marked) {
                cardFlowPane.getStyleClass().add(STALE_CARDS_STYLE_CLASS);
            } else if (!stale && marked) {
                cardFlowPane.getStyleClass().remove(STALE_CARDS_STYLE_CLASS);
            }
        }
    }

    private String buildMetaSummary() {
        if (loading.get()) {
            return "データ読込中…";
        }
        StringBuilder sb = new StringBuilder();
        if (machineCountLabel != null) {
            sb.append(machineCountLabel.getText());
        }
        if (lastUpdatedLabel != null) {
            if (!sb.isEmpty()) {
                sb.append("  ");
            }
            sb.append(lastUpdatedLabel.getText());
        }
        if (cachedSources != null && cachedSources.loadStats() != null) {
            String stats =
                    EquipmentStatusDashboardSourceLoader.formatLoadStatsSummary(
                            cachedSources.loadStats());
            if (!stats.isBlank()) {
                if (!sb.isEmpty()) {
                    sb.append("  ");
                }
                sb.append(stats);
            }
        }
        return sb.toString();
    }

    private void updateLastUpdatedLabel(String prefix) {
        if (lastUpdatedLabel != null) {
            lastUpdatedLabel.setText(prefix + ": " + LocalTime.now().format(TIME_FMT));
        }
    }

    private void updateAutoRefreshTimer() {
        if (autoRefreshTimeline != null) {
            autoRefreshTimeline.stop();
            autoRefreshTimeline = null;
        }
        boolean want =
                (tabActive.get() || fullscreenStage.isShowing())
                        && (autoRefreshCheckBox == null || autoRefreshCheckBox.isSelected());
        if (!want) {
            if (nextRefreshLabel != null) {
                nextRefreshLabel.setText("自動更新 停止中");
            }
            return;
        }
        autoRefreshRemainingSec = snapshotAutoRefreshIntervalSec();
        updateNextRefreshLabel();
        autoRefreshTimeline =
                new Timeline(new KeyFrame(Duration.seconds(1), e -> onAutoRefreshTick()));
        autoRefreshTimeline.setCycleCount(Timeline.INDEFINITE);
        autoRefreshTimeline.play();
    }

    private void onAutoRefreshTick() {
        autoRefreshRemainingSec--;
        if (autoRefreshRemainingSec <= 0) {
            autoRefreshRemainingSec = snapshotAutoRefreshIntervalSec();
            rolloverDatesIfFollowingToday();
            reloadFromSources();
        }
        updateNextRefreshLabel();
    }

    private void updateNextRefreshLabel() {
        if (nextRefreshLabel == null) {
            return;
        }
        int sec = Math.max(0, autoRefreshRemainingSec);
        nextRefreshLabel.setText(
                String.format("次回更新まで %d:%02d", sec / 60, sec % 60));
    }

    private void presentLoadError(String sourceContext, String detail, Throwable ex) {
        if (machineCountLabel != null) {
            machineCountLabel.setText(cachedSources == null ? "読込エラー" : machineCountText());
            addErrorStyle(machineCountLabel);
        }
        if (sourceSummaryLabel != null) {
            sourceSummaryLabel.setText(sourceContext);
        }
        if (loadStatsLabel != null) {
            loadStatsLabel.setText(
                    "読込エラー — " + DashboardLoadErrorFormatter.formatShortDetail(ex));
            addErrorStyle(loadStatsLabel);
            Tooltip tooltip =
                    new Tooltip(
                            sourceContext
                                    + "\n\n"
                                    + detail
                                    + "\n\n"
                                    + DashboardLoadErrorFormatter.formatStackTrace(ex));
            tooltip.setWrapText(true);
            tooltip.setMaxWidth(720);
            loadStatsLabel.setTooltip(tooltip);
        }
        refreshView();
    }

    private void clearLoadErrorPresentation() {
        removeErrorStyle(machineCountLabel);
        removeErrorStyle(loadStatsLabel);
        if (loadStatsLabel != null) {
            loadStatsLabel.setTooltip(null);
        }
        if (sourceSummaryLabel != null) {
            sourceSummaryLabel.setTooltip(null);
        }
    }

    private static void addErrorStyle(Label label) {
        if (label != null && !label.getStyleClass().contains(ERROR_STYLE_CLASS)) {
            label.getStyleClass().add(ERROR_STYLE_CLASS);
        }
    }

    private static void removeErrorStyle(Label label) {
        if (label != null) {
            label.getStyleClass().remove(ERROR_STYLE_CLASS);
        }
    }

    private void logLoadError(String sourceContext, Throwable ex) {
        if (shell == null) {
            return;
        }
        shell.appendLog("[dashboard] 読込失敗");
        for (String line : sourceContext.split("\n")) {
            if (!line.isBlank()) {
                shell.appendLog("[dashboard] " + line.strip());
            }
        }
        shell.appendLog("[dashboard] " + DashboardLoadErrorFormatter.formatDetail(ex));
        for (String line : DashboardLoadErrorFormatter.formatStackTrace(ex).split("\n")) {
            if (!line.isBlank()) {
                shell.appendLog("[dashboard] " + line);
            }
        }
    }
}
