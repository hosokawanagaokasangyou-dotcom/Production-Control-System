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
import javafx.scene.control.Label;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.ToggleGroup;
import javafx.scene.layout.FlowPane;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardBuilder;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.LoadedSources;
import jp.co.pm.ai.desktop.ui.EquipmentStatusCardFactory;
import jp.co.pm.ai.desktop.ui.EquipmentStatusCardFactory.DisplayOptions;
import jp.co.pm.ai.desktop.ui.EquipmentStatusFullscreenStage;

/** メインシェル「ダッシュボード」タブ。 */
public final class EquipmentStatusDashboardTabController {

    private static final int AUTO_REFRESH_SEC = 60;
    private static final int MAX_PLAN_DAY_OFFSET = 14;
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("M/d");

    private MainShellController shell;

    @FXML private ToggleButton fullscreenToggle;
    @FXML private CheckBox autoRefreshCheckBox;
    @FXML private ToggleGroup actualDayToggleGroup;
    @FXML private ToggleButton actualTodayToggle;
    @FXML private ToggleButton actualYesterdayToggle;
    @FXML private ToggleGroup planDayToggleGroup;
    @FXML private ToggleButton planTodayToggle;
    @FXML private ToggleButton planTomorrowToggle;
    @FXML private ToggleButton planDayAfterToggle;
    @FXML private Button planForwardDayButton;
    @FXML private Label lastUpdatedLabel;
    @FXML private CheckBox showAladdinCheckBox;
    @FXML private CheckBox showDispatchCheckBox;
    @FXML private Label planDateSummaryLabel;
    @FXML private Label actualDateSummaryLabel;
    @FXML private Label machineCountLabel;
    @FXML private Label sourceSummaryLabel;
    @FXML private FlowPane cardFlowPane;

    private final EquipmentStatusFullscreenStage fullscreenStage = new EquipmentStatusFullscreenStage();

    private Timeline autoRefreshTimeline;
    private Task<LoadedSources> activeReloadTask;
    private final AtomicBoolean tabActive = new AtomicBoolean(false);
    private final AtomicBoolean suppressUiEvents = new AtomicBoolean(false);

    private LoadedSources cachedSources;
    private List<EquipmentMachineStatus> currentStatuses = List.of();

    private int actualDayOffset;
    private int planDayOffset;

    @FXML
    private void initialize() {
        fullscreenStage.setOnClose(
                () ->
                        Platform.runLater(
                                () -> {
                                    if (fullscreenToggle != null) {
                                        fullscreenToggle.setSelected(false);
                                    }
                                    updateAutoRefreshTimer();
                                }));

        if (actualTodayToggle != null) {
            actualTodayToggle.setUserData(0);
        }
        if (actualYesterdayToggle != null) {
            actualYesterdayToggle.setUserData(-1);
        }
        if (planTodayToggle != null) {
            planTodayToggle.setUserData(0);
        }
        if (planTomorrowToggle != null) {
            planTomorrowToggle.setUserData(1);
        }
        if (planDayAfterToggle != null) {
            planDayAfterToggle.setUserData(2);
        }

        if (actualDayToggleGroup != null) {
            actualDayToggleGroup
                    .selectedToggleProperty()
                    .addListener(
                            (obs, prev, cur) -> {
                                if (suppressUiEvents.get() || cur == null) {
                                    return;
                                }
                                Object ud = cur.getUserData();
                                if (ud instanceof Integer off) {
                                    actualDayOffset = off;
                                    rebuildFromCache();
                                }
                            });
        }
        if (planDayToggleGroup != null) {
            planDayToggleGroup
                    .selectedToggleProperty()
                    .addListener(
                            (obs, prev, cur) -> {
                                if (suppressUiEvents.get() || cur == null) {
                                    return;
                                }
                                Object ud = cur.getUserData();
                                if (ud instanceof Integer off) {
                                    planDayOffset = off;
                                    rebuildFromCache();
                                }
                            });
        }

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
            actualDayOffset = clampActualOffset(s.equipmentStatusDashboardActualDayOffset());
            planDayOffset = clampPlanOffset(s.equipmentStatusDashboardPlanDayOffset());
            selectActualOffsetToggle(actualDayOffset);
            selectPlanOffsetToggle(planDayOffset);
            if (autoRefreshCheckBox != null) {
                autoRefreshCheckBox.setSelected(s.equipmentStatusDashboardAutoRefreshEnabled());
            }
            if (showAladdinCheckBox != null) {
                showAladdinCheckBox.setSelected(s.equipmentStatusDashboardShowAladdinPlans());
            }
            if (showDispatchCheckBox != null) {
                showDispatchCheckBox.setSelected(s.equipmentStatusDashboardShowDispatchPlans());
            }
        } finally {
            suppressUiEvents.set(false);
        }
    }

    public int snapshotActualDayOffset() {
        return actualDayOffset;
    }

    public int snapshotPlanDayOffset() {
        return planDayOffset;
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
    private void onFullscreenToggleAction() {
        if (shell == null || shell.getPrimaryStage() == null) {
            return;
        }
        if (fullscreenToggle != null && fullscreenToggle.isSelected()) {
            fullscreenStage.show(
                    shell.getPrimaryStage(),
                    currentStatuses,
                    currentDisplayOptions(),
                    badgeStyleResolver(),
                    buildMetaSummary());
            updateAutoRefreshTimer();
        } else {
            fullscreenStage.hide();
            updateAutoRefreshTimer();
        }
    }

    @FXML
    private void onPlanForwardDayAction() {
        planDayOffset = Math.min(MAX_PLAN_DAY_OFFSET, planDayOffset + 1);
        selectPlanOffsetToggle(planDayOffset);
        rebuildFromCache();
    }

    private void reloadFromSources() {
        if (shell == null) {
            return;
        }
        if (activeReloadTask != null && activeReloadTask.isRunning()) {
            return;
        }
        activeReloadTask =
                new Task<>() {
                    @Override
                    protected LoadedSources call() throws Exception {
                        return EquipmentStatusDashboardSourceLoader.load(shell.snapshotUiEnv());
                    }
                };
        activeReloadTask.setOnSucceeded(
                e -> {
                    cachedSources = activeReloadTask.getValue();
                    rebuildFromCache();
                    updateLastUpdatedLabel();
                });
        activeReloadTask.setOnFailed(
                e -> {
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

    private void rebuildFromCache() {
        LocalDate anchor = LocalDate.now();
        LocalDate actualDate = anchor.plusDays(actualDayOffset);
        LocalDate planDate = anchor.plusDays(planDayOffset);
        if (cachedSources != null) {
            currentStatuses =
                    EquipmentStatusDashboardBuilder.build(
                            cachedSources.actuals(),
                            cachedSources.aladdin(),
                            cachedSources.dispatch(),
                            actualDate,
                            planDate);
        } else {
            currentStatuses = List.of();
        }
        updateSummaryLabels(actualDate, planDate);
        renderCards();
        if (fullscreenStage.isShowing()) {
            fullscreenStage.rebuildCards(
                    currentStatuses, currentDisplayOptions(), badgeStyleResolver());
        }
    }

    private void renderCards() {
        if (cardFlowPane == null) {
            return;
        }
        cardFlowPane.getChildren().clear();
        DisplayOptions opts = currentDisplayOptions();
        Function<String, PersonBadgeStyle> resolver = badgeStyleResolver();
        for (EquipmentMachineStatus s : currentStatuses) {
            cardFlowPane
                    .getChildren()
                    .add(EquipmentStatusCardFactory.createCard(s, opts, resolver, false));
        }
    }

    private DisplayOptions currentDisplayOptions() {
        LocalDate anchor = LocalDate.now();
        return new DisplayOptions(
                showAladdinCheckBox == null || showAladdinCheckBox.isSelected(),
                showDispatchCheckBox == null || showDispatchCheckBox.isSelected(),
                anchor.plusDays(actualDayOffset).format(DATE_FMT),
                anchor.plusDays(planDayOffset).format(DATE_FMT));
    }

    private Function<String, PersonBadgeStyle> badgeStyleResolver() {
        return shell != null
                ? shell.personBadgeStyleResolverForGantt()
                : (String __) -> PersonBadgeStyle.defaultStyle();
    }

    private void updateSummaryLabels(LocalDate actualDate, LocalDate planDate) {
        if (actualDateSummaryLabel != null) {
            actualDateSummaryLabel.setText("実績:" + actualDate.format(DATE_FMT));
        }
        if (planDateSummaryLabel != null) {
            planDateSummaryLabel.setText("予定:" + planDate.format(DATE_FMT));
        }
        if (machineCountLabel != null) {
            machineCountLabel.setText(currentStatuses.size() + "台");
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
        if (planForwardDayButton != null) {
            planForwardDayButton.setDisable(planDayOffset >= MAX_PLAN_DAY_OFFSET);
        }
    }

    private String buildMetaSummary() {
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
                                        e -> {
                                            actualDayOffset = readActualOffsetFromUi();
                                            planDayOffset = readPlanOffsetFromUi();
                                            reloadFromSources();
                                        }));
                autoRefreshTimeline.setCycleCount(Timeline.INDEFINITE);
            }
            if (autoRefreshTimeline.getStatus() != javafx.animation.Animation.Status.RUNNING) {
                autoRefreshTimeline.playFromStart();
            }
        } else if (autoRefreshTimeline != null) {
            autoRefreshTimeline.stop();
        }
    }

    private int readActualOffsetFromUi() {
        ToggleButton sel =
                actualDayToggleGroup != null
                                && actualDayToggleGroup.getSelectedToggle()
                                        instanceof ToggleButton tb
                        ? tb
                        : null;
        if (sel != null && sel.getUserData() instanceof Integer off) {
            return off;
        }
        return actualDayOffset;
    }

    private int readPlanOffsetFromUi() {
        ToggleButton sel =
                planDayToggleGroup != null
                                && planDayToggleGroup.getSelectedToggle()
                                        instanceof ToggleButton tb
                        ? tb
                        : null;
        if (sel != null && sel.getUserData() instanceof Integer off) {
            return off;
        }
        return planDayOffset;
    }

    private void selectActualOffsetToggle(int offset) {
        actualDayOffset = clampActualOffset(offset);
        ToggleButton target =
                actualDayOffset == -1 ? actualYesterdayToggle : actualTodayToggle;
        if (actualDayToggleGroup != null && target != null) {
            actualDayToggleGroup.selectToggle(target);
        }
    }

    private void selectPlanOffsetToggle(int offset) {
        planDayOffset = clampPlanOffset(offset);
        ToggleButton target = null;
        if (planDayOffset == 0) {
            target = planTodayToggle;
        } else if (planDayOffset == 1) {
            target = planTomorrowToggle;
        } else if (planDayOffset == 2) {
            target = planDayAfterToggle;
        }
        if (planDayToggleGroup != null) {
            if (target != null) {
                planDayToggleGroup.selectToggle(target);
            } else {
                planDayToggleGroup.selectToggle(null);
            }
        }
    }

    private static int clampActualOffset(int offset) {
        return offset <= -1 ? -1 : 0;
    }

    private static int clampPlanOffset(int offset) {
        if (offset < 0) {
            return 0;
        }
        return Math.min(MAX_PLAN_DAY_OFFSET, offset);
    }
}
