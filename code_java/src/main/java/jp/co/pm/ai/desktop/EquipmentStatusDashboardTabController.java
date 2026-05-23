package jp.co.pm.ai.desktop;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Function;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TitledPane;
import javafx.scene.control.ToggleButton;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.VBox;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
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
import jp.co.pm.ai.desktop.ui.EquipmentStatusFullscreenStage;

/** メインシェル「ダッシュボード」タブ。 */
public final class EquipmentStatusDashboardTabController {

    private static final int AUTO_REFRESH_SEC = 60;
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/M/d");

    private MainShellController shell;

    @FXML private Button reloadButton;
    @FXML private ToggleButton fullscreenToggle;
    @FXML private CheckBox autoRefreshCheckBox;
    @FXML private DatePicker actualDatePicker;
    @FXML private DatePicker planDatePicker;
    @FXML private Label lastUpdatedLabel;
    @FXML private ProgressIndicator loadingIndicator;
    @FXML private Label loadingStatusLabel;
    @FXML private CheckBox showAladdinCheckBox;
    @FXML private CheckBox showDispatchCheckBox;
    @FXML private Label planDateSummaryLabel;
    @FXML private Label actualDateSummaryLabel;
    @FXML private Label machineCountLabel;
    @FXML private Label sourceSummaryLabel;
    @FXML private TitledPane appearancePane;
    @FXML private VBox appearanceControlsHost;
    @FXML private ScrollPane cardScrollPane;
    @FXML private FlowPane cardFlowPane;
    @FXML private VBox emptyStatePane;
    @FXML private VBox loadingOverlayPane;

    private final EquipmentStatusFullscreenStage fullscreenStage = new EquipmentStatusFullscreenStage();

    private Timeline autoRefreshTimeline;
    private Task<ReloadDecision> activeReloadTask;
    private final AtomicBoolean tabActive = new AtomicBoolean(false);
    private final AtomicBoolean suppressUiEvents = new AtomicBoolean(false);
    private final AtomicBoolean loading = new AtomicBoolean(false);

    private LoadedSources cachedSources;
    private SourceFingerprint loadedFingerprint;
    private List<EquipmentMachineStatus> currentStatuses = List.of();

    private LocalDate actualDate = LocalDate.now();
    private LocalDate planDate = LocalDate.now();
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
                                actualDate = cur;
                                rebuildFromCache();
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
            cardScrollPane.viewportBoundsProperty()
                    .addListener(
                            (o, a, b) ->
                                    EquipmentStatusDashboardAppearanceApplier.configureFlowPane(
                                            cardFlowPane,
                                            appearancePrefs,
                                            false,
                                            b.getWidth()));
        }

        fullscreenStage.setOnClose(
                () ->
                        Platform.runLater(
                                () -> {
                                    if (fullscreenToggle != null) {
                                        fullscreenToggle.setSelected(false);
                                    }
                                    updateAutoRefreshTimer();
                                }));

        Runnable displayRefresh = this::rebuildFromCache;
        if (showAladdinCheckBox != null) {
            showAladdinCheckBox.selectedProperty().addListener((o, a, b) -> displayRefresh.run());
        }
        if (showDispatchCheckBox != null) {
            showDispatchCheckBox.selectedProperty().addListener((o, a, b) -> displayRefresh.run());
        }
        if (autoRefreshCheckBox != null) {
            autoRefreshCheckBox
                    .selectedProperty()
                    .addListener((o, a, b) -> updateAutoRefreshTimer());
        }
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
        reloadFromSources();
        updateAutoRefreshTimer();
    }

    public void onMainShellTabDeselected() {
        tabActive.set(false);
        updateAutoRefreshTimer();
    }

    @FXML
    private void onReloadAction() {
        reloadFromSources();
    }

    @FXML
    private void onActualTodayAction() {
        setActualDate(LocalDate.now());
    }

    @FXML
    private void onPlanTodayAction() {
        setPlanDate(LocalDate.now());
    }

    @FXML
    private void onFullscreenToggleAction() {
        if (shell == null || shell.getPrimaryStage() == null) {
            return;
        }
        if (fullscreenToggle != null && fullscreenToggle.isSelected()) {
            DisplayOptions opts = currentDisplayOptions();
            EquipmentStatusDashboardAppearancePrefs ap = snapshotAppearancePrefs();
            fullscreenStage.show(
                    shell.getPrimaryStage(),
                    currentStatuses,
                    opts,
                    ap,
                    badgeStyleResolver(),
                    buildMetaSummary(),
                    opts.actualDateLabel(),
                    opts.planDateLabel(),
                    cachedSources != null);
            updateAutoRefreshTimer();
        } else {
            fullscreenStage.hide();
            updateAutoRefreshTimer();
        }
    }

    private void onAppearanceChanged(EquipmentStatusDashboardAppearancePrefs prefs) {
        appearancePrefs =
                prefs != null ? prefs : EquipmentStatusDashboardAppearancePrefs.defaults();
        applyFlowLayout();
        renderCards();
        if (fullscreenStage.isShowing()) {
            DisplayOptions opts = currentDisplayOptions();
            fullscreenStage.rebuildCards(
                    currentStatuses,
                    opts,
                    appearancePrefs,
                    badgeStyleResolver(),
                    opts.actualDateLabel(),
                    opts.planDateLabel(),
                    cachedSources != null);
        }
        if (shell != null) {
            shell.persistDesktopSessionNow();
        }
    }

    private void setActualDate(LocalDate date) {
        if (date == null) {
            return;
        }
        actualDate = date;
        suppressUiEvents.set(true);
        try {
            if (actualDatePicker != null) {
                actualDatePicker.setValue(date);
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

    private void reloadFromSources() {
        if (shell == null) {
            return;
        }
        if (activeReloadTask != null && activeReloadTask.isRunning()) {
            return;
        }
        final SourceFingerprint previousFingerprint = loadedFingerprint;
        final boolean haveCache = cachedSources != null;
        activeReloadTask =
                new Task<>() {
                    @Override
                    protected ReloadDecision call() throws Exception {
                        ReloadDecision decision =
                                EquipmentStatusDashboardSourceLoader.loadIfChanged(
                                        shell.snapshotUiEnv(), previousFingerprint, haveCache);
                        if (!decision.sourcesUnchanged()) {
                            Platform.runLater(() -> setLoading(true));
                        }
                        return decision;
                    }
                };
        activeReloadTask.setOnSucceeded(
                e -> {
                    ReloadDecision decision = activeReloadTask.getValue();
                    if (decision == null || decision.sourcesUnchanged()) {
                        if (shell != null) {
                            shell.appendLog("[dashboard] ソース未変更のため読込を省略");
                        }
                        return;
                    }
                    setLoading(false);
                    loadedFingerprint = decision.fingerprint();
                    cachedSources = decision.sources();
                    rebuildFromCache();
                    updateLastUpdatedLabel();
                });
        activeReloadTask.setOnFailed(
                e -> {
                    setLoading(false);
                    if (machineCountLabel != null) {
                        machineCountLabel.setText("読込エラー");
                    }
                    if (sourceSummaryLabel != null) {
                        sourceSummaryLabel.setText("［再読込］でやり直してください");
                    }
                    if (shell != null && activeReloadTask.getException() != null) {
                        shell.appendLog(
                                "[dashboard] reload failed: "
                                        + activeReloadTask.getException().getMessage());
                    }
                });
        Thread t = new Thread(activeReloadTask, "equipment-status-dashboard-reload");
        t.setDaemon(true);
        t.start();
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
        if (loadingOverlayPane != null) {
            boolean showOverlay = on && (cachedSources == null || currentStatuses.isEmpty());
            loadingOverlayPane.setVisible(showOverlay);
            loadingOverlayPane.setManaged(showOverlay);
        }
        if (cardScrollPane != null) {
            cardScrollPane.setOpacity(on && cachedSources != null ? 0.5 : 1.0);
        }
        if (on) {
            if (machineCountLabel != null) {
                machineCountLabel.setText("読込中…");
            }
            if (sourceSummaryLabel != null) {
                sourceSummaryLabel.setText("実績・アラジン・配台を読み込んでいます");
            }
        } else if (fullscreenStage.isShowing()) {
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
        updateSummaryLabels(actual, plan);
        applyFlowLayout();
        renderCards();
        if (fullscreenStage.isShowing()) {
            DisplayOptions opts = currentDisplayOptions();
            fullscreenStage.rebuildCards(
                    currentStatuses,
                    opts,
                    snapshotAppearancePrefs(),
                    badgeStyleResolver(),
                    opts.actualDateLabel(),
                    opts.planDateLabel(),
                    cachedSources != null);
            fullscreenStage.setMetaText(buildMetaSummary());
        }
    }

    private void applyFlowLayout() {
        if (cardFlowPane == null) {
            return;
        }
        double viewport =
                cardScrollPane != null ? cardScrollPane.getViewportBounds().getWidth() : 0;
        EquipmentStatusDashboardAppearanceApplier.configureFlowPane(
                cardFlowPane, appearancePrefs, false, viewport);
    }

    private void renderCards() {
        if (cardFlowPane == null) {
            return;
        }
        cardFlowPane.getChildren().clear();
        DisplayOptions opts = currentDisplayOptions();
        EquipmentStatusDashboardAppearancePrefs ap = snapshotAppearancePrefs();
        boolean sourcesLoaded = cachedSources != null;
        boolean empty = currentStatuses == null || currentStatuses.isEmpty();
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
                                        false));
            }
        }
        if (empty) {
            return;
        }
        Function<String, PersonBadgeStyle> resolver = badgeStyleResolver();
        for (EquipmentMachineStatus s : currentStatuses) {
            cardFlowPane
                    .getChildren()
                    .add(EquipmentStatusCardFactory.createCard(s, opts, ap, resolver, false));
        }
    }

    private DisplayOptions currentDisplayOptions() {
        LocalDate actual = actualDate != null ? actualDate : LocalDate.now();
        LocalDate plan = planDate != null ? planDate : LocalDate.now();
        return new DisplayOptions(
                showAladdinCheckBox == null || showAladdinCheckBox.isSelected(),
                showDispatchCheckBox == null || showDispatchCheckBox.isSelected(),
                actual.format(DATE_FMT),
                plan.format(DATE_FMT));
    }

    private Function<String, PersonBadgeStyle> badgeStyleResolver() {
        return shell != null
                ? shell.personBadgeStyleResolverForGantt()
                : (String __) -> PersonBadgeStyle.defaultStyle();
    }

    private void updateSummaryLabels(LocalDate actual, LocalDate plan) {
        if (actualDateSummaryLabel != null) {
            actualDateSummaryLabel.setText("実績日:" + actual.format(DATE_FMT));
        }
        if (planDateSummaryLabel != null) {
            planDateSummaryLabel.setText("予定日:" + plan.format(DATE_FMT));
        }
        if (machineCountLabel != null) {
            if (cachedSources == null) {
                machineCountLabel.setText("—");
            } else if (currentStatuses.isEmpty()) {
                machineCountLabel.setText("該当なし（非稼働日の可能性）");
            } else {
                machineCountLabel.setText(currentStatuses.size() + "台");
            }
        }
        if (sourceSummaryLabel != null && cachedSources != null) {
            sourceSummaryLabel.setText(
                    "実績="
                            + cachedSources.actualSourceLabel()
                            + "  アラジン="
                            + cachedSources.aladdinSourceLabel()
                            + "  配台="
                            + cachedSources.dispatchSourceLabel());
        }
    }

    private String buildMetaSummary() {
        if (loading.get()) {
            return "データ読込中…";
        }
        return (machineCountLabel != null ? machineCountLabel.getText() : "")
                + "  "
                + (lastUpdatedLabel != null ? lastUpdatedLabel.getText() : "");
    }

    private void updateLastUpdatedLabel() {
        if (lastUpdatedLabel != null) {
            lastUpdatedLabel.setText(
                    "最終更新:"
                            + java.time.LocalTime.now()
                                    .format(DateTimeFormatter.ofPattern("HH:mm:ss")));
        }
    }

    private void updateAutoRefreshTimer() {
        boolean want =
                (tabActive.get() || fullscreenStage.isShowing())
                        && (autoRefreshCheckBox == null || autoRefreshCheckBox.isSelected());
        if (want) {
            if (autoRefreshTimeline == null) {
                autoRefreshTimeline =
                        new Timeline(
                                new KeyFrame(
                                        Duration.seconds(AUTO_REFRESH_SEC),
                                        e -> reloadFromSources()));
                autoRefreshTimeline.setCycleCount(Timeline.INDEFINITE);
            }
            if (autoRefreshTimeline.getStatus() != javafx.animation.Animation.Status.RUNNING) {
                autoRefreshTimeline.playFromStart();
            }
        } else if (autoRefreshTimeline != null) {
            autoRefreshTimeline.stop();
        }
    }
}
