package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.util.Duration;
import javafx.beans.binding.Bindings;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.scene.control.Accordion;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.ContextMenu;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.MenuItem;
import javafx.scene.control.MultipleSelectionModel;
import javafx.scene.control.RadioButton;
import javafx.scene.control.ScrollBar;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.control.ToggleGroup;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyCodeCombination;
import javafx.scene.input.KeyCombination;
import javafx.scene.effect.DropShadow;
import javafx.scene.input.KeyEvent;
import javafx.scene.Node;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.AppVersionInfo;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.ui.PersonBadgeNodeFactory;

/** Run/log tab; layout in {@code MainRunTab.fxml}. */
public final class MainRunTabController {

    private static final int MAX_PERSISTED_LOG_LINES = 2000;

    private static final String DEFAULT_FONT_FAMILY_LABEL = "システム既定";

    private static final List<Double> PRESET_FONT_SIZES =
            List.of(9d, 10d, 11d, 12d, 13d, 14d, 15d, 16d, 18d, 20d, 22d, 24d);

    private MainShellController shell;

    @FXML
    private TextField workbookField;

    @FXML
    private TextField scriptDirField;

    @FXML
    private ListView<String> logListView;

    @FXML
    private BorderPane runTabRoot;

    @FXML
    private ComboBox<LogViewFilter> logFilterCombo;

    @FXML
    private TextField logSearchField;

    @FXML
    private ComboBox<String> logFontFamilyCombo;

    @FXML
    private ComboBox<Double> logFontSizeCombo;

    @FXML
    private Label statusLabel;

    @FXML
    private Accordion stage2ProgressAccordion;

    @FXML
    private TitledPane stage2ProgressPane;

    @FXML
    private Label stage2ProgressLabel;

    @FXML
    private TextField stage2ProductionPlanField;

    @FXML
    private TextField stage2MemberScheduleField;

    @FXML
    private StackPane stage1NetworkCacheBadgeHost;

    @FXML
    private StackPane attendanceReadinessBadgeHost;

    @FXML
    private Button stage1RunButton;

    @FXML
    private Button openMachineDeliveryManagementXlsmButton;

    @FXML
    private Label stage1PipelineCheckBlockBadge;

    private static final String STAGE1_RUN_BUTTON_TEXT_DEFAULT = "段階1 実行";

    private static final String STAGE1_RUN_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD =
            "段階1（納期管理ビュー更新中）";

    @FXML
    private RadioButton stage2SkipTodayDispatchRadio;

    @FXML
    private RadioButton todayDispatchRadio;

    @FXML
    private ToggleGroup todayDispatchModeToggleGroup;

    /**
     * 「当日配台する」選択中の {@code PM_AI_STAGE2_SKIP_TODAY_DISPATCH}。
     * ソース取得時刻ポリシー等で更新する。「当日は配台しない」選択中は {@link #snapshotStage2SkipTodayDispatch()} が常に true。
     */
    private boolean stage2SkipTodayDispatchWhenTodayMode;

    /** セッション復元・プログラムからの選択変更中はリスナー副作用を抑止する。 */
    private boolean suppressingTodayDispatchModeListener;

    @FXML
    private ComboBox<String> stage2ResultBookFontCombo;

    @FXML
    private CheckBox skipGeminiApiCheckBox;

    @FXML
    private CheckBox stage1MarkAllExcludeAfterRunCheckBox;

    @FXML
    private Button copyAllLogButton;

    @FXML
    private Button clearLogButton;

    @FXML
    private Label prismPipelineLabel;

    @FXML
    private Label appVersionLabel;

    @FXML
    private Label masterWorkbookOpenHintLabel;

    /** 段階1／段階2 の Python 実行中（メインシェルから同期）。 */
    private boolean stage1RunPipelineBusy;

    /** 納期管理ビュー再読み込み中（メインシェルから同期）。 */
    private boolean deliveryCalendarReloadBlocking;

    private boolean stage1BlockedByPipelineCheck;
    /** 起動直後は未確認のため段階1を抑止（readiness 応答後に解除）。 */
    private boolean stage1BlockedByCalendarNotReady = true;
    private String calendarReadinessBlockTooltip = "";

    private String stage1PipelineCheckBlockTooltip = "";

    private String stage1PipelineCheckBlockBadgeText = "";

    /** ゲスト操作者時は実行・ログタブ本体を抑止する。 */
    private boolean guestSessionFactorySwitchOnly;

    @FXML
    private Label pipelineTimingStage1Label;

    @FXML
    private Label pipelineTimingStage20Label;

    @FXML
    private Label pipelineTimingStage21Label;

    @FXML
    private Label pipelineTimingDispatchTrialLabel;

    @FXML
    private HBox pipelineTimingDispatchTrialRow;

    @FXML
    private Label pipelineTimingSummaryExcelLabel;

    @FXML
    private Label pipelineTimingDeliveryCalendarLabel;

    private Runnable pipelineTimingHistoryListener;

    private final ObservableList<String> logLinesAll = FXCollections.observableArrayList();
    private final FilteredList<String> logLinesVisible =
            new FilteredList<>(logLinesAll, s -> true);

    private Font appliedLogFont = Font.getDefault();

    private final AtomicBoolean suppressLogFontEvents = new AtomicBoolean(false);

    private final AtomicBoolean suppressStage2ResultFontEvents = new AtomicBoolean(false);

    private final AtomicBoolean suppressSkipGeminiApiEvents = new AtomicBoolean(false);

    private final AtomicBoolean suppressStage1MarkAllExcludeAfterRunEvents = new AtomicBoolean(false);

    private final AtomicBoolean suppressRunLogSessionPersistence = new AtomicBoolean(false);

    private double pendingSessionLogScroll = Double.NaN;

    private final Object logAppendLock = new Object();

    private final List<String> logAppendPending = new ArrayList<>();

    private Timeline logAppendFlushTimeline;

    private Timeline logListRefreshDebounceTimeline;

    private static final int LOG_APPEND_BATCH_MAX = 40;

    private static final Duration LOG_APPEND_FLUSH_DELAY = Duration.millis(32);

    private static final Duration LOG_LIST_REFRESH_DEBOUNCE = Duration.millis(120);

    /** この割合以上で末尾にいるときだけ自動 scrollTo（ユーザーが上へスクロール中は追従しない）。 */
    private static final double LOG_AUTO_SCROLL_BOTTOM_THRESHOLD = 0.92;

    /** {@code true} のとき新規ログ行で末尾へ追従する（段階実行開始時は強制オン）。 */
    private boolean logTailFollowEnabled = true;

    private final AtomicBoolean suppressLogTailFollowListener = new AtomicBoolean(false);

    private boolean logTailFollowScrollListenerInstalled;

    @FXML
    private void initialize() {
        logFilterCombo.getItems().setAll(LogViewFilter.values());
        logFilterCombo.setValue(LogViewFilter.ALL);
        logFilterCombo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(LogViewFilter f) {
                        return f != null ? f.getLabel() : "";
                    }

                    @Override
                    public LogViewFilter fromString(String string) {
                        if (string == null || string.isBlank()) {
                            return LogViewFilter.ALL;
                        }
                        for (LogViewFilter v : LogViewFilter.values()) {
                            if (v.getLabel().equals(string)) {
                                return v;
                            }
                        }
                        return LogViewFilter.ALL;
                    }
                });
        logFilterCombo
                .valueProperty()
                .addListener(
                        (o, a, b) -> {
                            refreshLogLinesVisiblePredicate();
                            if (shell != null
                                    && !suppressRunLogSessionPersistence.get()) {
                                shell.scheduleDesktopSessionSave();
                            }
                        });
        if (logSearchField != null) {
            logSearchField
                    .textProperty()
                    .addListener(
                            (o, a, b) -> {
                                refreshLogLinesVisiblePredicate();
                                if (!snapshotLogSearchText().isEmpty()) {
                                    scheduleDebouncedLogListRefresh();
                                }
                            });
        }

        List<String> families = new ArrayList<>();
        families.add(DEFAULT_FONT_FAMILY_LABEL);
        List<String> installed = new ArrayList<>(Font.getFamilies());
        Collections.sort(installed);
        families.addAll(installed);
        logFontFamilyCombo.getItems().setAll(families);
        logFontFamilyCombo.getSelectionModel().selectFirst();

        logFontSizeCombo.getItems().setAll(PRESET_FONT_SIZES);
        logFontSizeCombo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(Double object) {
                        if (object == null) {
                            return "";
                        }
                        if (object == Math.rint(object)) {
                            return String.valueOf(object.intValue());
                        }
                        return object.toString();
                    }

                    @Override
                    public Double fromString(String string) {
                        if (string == null || string.isBlank()) {
                            return null;
                        }
                        return Double.valueOf(string.trim());
                    }
                });
        logFontSizeCombo.setValue(14d);

        Runnable onFontUiChange =
                () -> {
                    if (!suppressLogFontEvents.get()) {
                        applyLogAreaFont();
                        if (shell != null) {
                            shell.scheduleDesktopSessionSave();
                        }
                    }
                };
        logFontFamilyCombo.valueProperty().addListener((o, a, b) -> onFontUiChange.run());
        logFontSizeCombo.valueProperty().addListener((o, a, b) -> onFontUiChange.run());

        setupLogListView();
        refreshLogLinesVisiblePredicate();
        if (copyAllLogButton != null) {
            copyAllLogButton.disableProperty().bind(Bindings.isEmpty(logLinesAll));
        }
        if (clearLogButton != null) {
            clearLogButton.disableProperty().bind(Bindings.isEmpty(logLinesAll));
        }
        applyLogAreaFont();
        installStageRunButtonDepth(stage1RunButton, Color.rgb(14, 116, 144, 0.35));
        wireTodayDispatchModeExclusiveRadios();
        if (prismPipelineLabel != null) {
            prismPipelineLabel.setText(PrismGpuBootstrapStatus.runTabSummary());
        }
        if (stage2ResultBookFontCombo != null) {
            List<String> stage2Families = new ArrayList<>();
            stage2Families.add(DEFAULT_FONT_FAMILY_LABEL);
            stage2Families.addAll(installed);
            stage2ResultBookFontCombo.getItems().setAll(stage2Families);
            stage2ResultBookFontCombo.getSelectionModel().selectFirst();
            stage2ResultBookFontCombo
                    .valueProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (!suppressStage2ResultFontEvents.get()
                                        && shell != null) {
                                    shell.scheduleDesktopSessionSave();
                                }
                            });
        }
        if (skipGeminiApiCheckBox != null) {
            skipGeminiApiCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (!suppressSkipGeminiApiEvents.get() && shell != null) {
                                    shell.scheduleDesktopSessionSave();
                                }
                            });
        }
        if (stage1MarkAllExcludeAfterRunCheckBox != null) {
            stage1MarkAllExcludeAfterRunCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (!suppressStage1MarkAllExcludeAfterRunEvents.get()
                                        && shell != null) {
                                    shell.scheduleDesktopSessionSave();
                                }
                            });
        }
        refreshMachineDeliveryManagementOpenButton();
    }

    /** フラットボタン用のごく弱いドロップシャドウ（パルスなし）。 */
    private static void installStageRunButtonDepth(Button button, Color shadowColor) {
        if (button == null) {
            return;
        }
        DropShadow depth = new DropShadow();
        depth.setColor(shadowColor);
        depth.setRadius(10);
        depth.setSpread(0.12);
        depth.setOffsetY(2);
        button.setEffect(depth);
    }

    private void setupLogListView() {
        logListView.setItems(logLinesVisible);
        applyLogListFixedCellHeight();
        logListView.setFocusTraversable(true);
        logListView.setCellFactory(
                lv ->
                        new ListCell<>() {
                            {
                                RunLogListViewSupport.installOverflowClip(this);
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                getStyleClass()
                                        .removeAll(
                                                "log-cell",
                                                "log-kind-error",
                                                "log-kind-warn",
                                                "log-dark");
                                if (empty || item == null) {
                                    setText(null);
                                    setGraphic(null);
                                    return;
                                }
                                setText(null);
                                setWrapText(false);
                                LogLineKind kind = LogLineKind.classify(item);
                                setGraphic(
                                        RunLogListViewSupport.buildLineGraphic(
                                                item,
                                                appliedLogFont,
                                                resolveLogTextFill(kind),
                                                snapshotLogSearchText()));
                                getStyleClass().add("log-cell");
                                if (shell != null && shell.currentDesktopTheme().isDarkUi()) {
                                    getStyleClass().add("log-dark");
                                }
                                switch (kind) {
                                    case ERROR -> getStyleClass().add("log-kind-error");
                                    case WARN -> getStyleClass().add("log-kind-warn");
                                    case NORMAL -> {
                                        /* default row chrome only */
                                    }
                                }
                            }
                        });
        logListView.widthProperty().addListener((o, a, b) -> scheduleDebouncedLogListRefresh());
        logListView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        installLogClipboardSupport();
        logListView.sceneProperty()
                .addListener(
                        (obs, oldScene, scene) -> {
                            if (scene != null) {
                                Platform.runLater(this::installLogTailFollowScrollListener);
                            }
                        });
        if (logListView.getScene() != null) {
            Platform.runLater(this::installLogTailFollowScrollListener);
        }
    }

    private void installLogTailFollowScrollListener() {
        if (logListView == null || logTailFollowScrollListenerInstalled) {
            return;
        }
        ScrollBar sb = (ScrollBar) logListView.lookup(".scroll-bar:vertical");
        if (sb == null) {
            Platform.runLater(this::installLogTailFollowScrollListener);
            return;
        }
        logTailFollowScrollListenerInstalled = true;
        sb.valueProperty()
                .addListener(
                        (obs, oldVal, newVal) -> {
                            if (suppressLogTailFollowListener.get()) {
                                return;
                            }
                            double proportion = readVerticalScrollProportion(logListView);
                            if (!Double.isFinite(proportion)) {
                                return;
                            }
                            if (proportion >= LOG_AUTO_SCROLL_BOTTOM_THRESHOLD) {
                                logTailFollowEnabled = true;
                            } else if (proportion
                                    <= LOG_AUTO_SCROLL_BOTTOM_THRESHOLD - 0.08) {
                                logTailFollowEnabled = false;
                            }
                        });
    }

    /**
     * 可変行高（{@code setFixedCellSize(-1)}）と折り返しの組み合わせは VirtualFlow が極端なセル数を見積もり、
     * {@code index exceeds maxCellCount} やヒープ枯渇を招くことがある。実測行高に応じた正の固定高で抑える。
     * ログは折り返さず右端でクリップする（{@link RunLogListViewSupport}）。
     */
    private void applyLogListFixedCellHeight() {
        if (logListView == null) {
            return;
        }
        logListView.setFixedCellSize(RunLogListViewSupport.fixedCellSizePx(appliedLogFont));
    }

    private void scheduleDebouncedLogListRefresh() {
        if (logListView == null) {
            return;
        }
        Runnable schedule =
                () -> {
                    if (logListRefreshDebounceTimeline == null) {
                        logListRefreshDebounceTimeline =
                                new Timeline(
                                        new KeyFrame(
                                                LOG_LIST_REFRESH_DEBOUNCE,
                                                e -> {
                                                    if (logListView != null) {
                                                        logListView.refresh();
                                                    }
                                                }));
                        logListRefreshDebounceTimeline.setCycleCount(1);
                    }
                    logListRefreshDebounceTimeline.playFromStart();
                };
        if (Platform.isFxApplicationThread()) {
            schedule.run();
        } else {
            Platform.runLater(schedule);
        }
    }

    private boolean shouldAutoScrollLogToEnd() {
        if (logTailFollowEnabled) {
            return true;
        }
        if (logListView == null) {
            return true;
        }
        double proportion = readVerticalScrollProportion(logListView);
        return !Double.isFinite(proportion)
                || proportion >= LOG_AUTO_SCROLL_BOTTOM_THRESHOLD;
    }

    /** 段階実行・配台試行開始時: ログ末尾追従をオンにし、最新行へスクロールする。 */
    void beginLogTailFollowForRun() {
        if (!Platform.isFxApplicationThread()) {
            Platform.runLater(this::beginLogTailFollowForRun);
            return;
        }
        logTailFollowEnabled = true;
        scrollLogListToEnd(true);
    }

    private void scrollLogListToLastVisibleRowIfNeeded() {
        scrollLogListToEnd(false);
    }

    private void scrollLogListToEnd(boolean force) {
        if (logListView == null || (!force && !shouldAutoScrollLogToEnd())) {
            return;
        }
        int n = logLinesVisible.size();
        if (n <= 0) {
            return;
        }
        Runnable scroll =
                () -> {
                    suppressLogTailFollowListener.set(true);
                    try {
                        logListView.scrollTo(n - 1);
                    } finally {
                        suppressLogTailFollowListener.set(false);
                    }
                };
        if (Platform.isFxApplicationThread()) {
            scroll.run();
            Platform.runLater(scroll);
        } else {
            Platform.runLater(scroll);
        }
    }

    private void scheduleLogAppendFlush(boolean immediate) {
        Runnable schedule =
                () -> {
                    if (immediate) {
                        if (logAppendFlushTimeline != null) {
                            logAppendFlushTimeline.stop();
                        }
                        flushPendingLogAppendsOnFxThread();
                    } else {
                        if (logAppendFlushTimeline == null) {
                            logAppendFlushTimeline =
                                    new Timeline(
                                            new KeyFrame(
                                                    LOG_APPEND_FLUSH_DELAY,
                                                    e -> flushPendingLogAppendsOnFxThread()));
                            logAppendFlushTimeline.setCycleCount(1);
                        }
                        logAppendFlushTimeline.playFromStart();
                    }
                };
        if (Platform.isFxApplicationThread()) {
            schedule.run();
        } else {
            Platform.runLater(schedule);
        }
    }

    /** バッチ追記の残りを UI へ反映する（ポータル同期終了時など）。 */
    void flushPendingLogAppends() {
        if (Platform.isFxApplicationThread()) {
            flushPendingLogAppendsOnFxThread();
        } else {
            Platform.runLater(this::flushPendingLogAppendsOnFxThread);
        }
    }

    private void flushPendingLogAppendsOnFxThread() {
        if (!Platform.isFxApplicationThread()) {
            Platform.runLater(this::flushPendingLogAppendsOnFxThread);
            return;
        }
        List<String> batch;
        synchronized (logAppendLock) {
            if (logAppendPending.isEmpty()) {
                return;
            }
            batch = new ArrayList<>(logAppendPending);
            logAppendPending.clear();
        }
        logLinesAll.addAll(batch);
        scrollLogListToLastVisibleRowIfNeeded();
    }

    private void installLogClipboardSupport() {
        var copyKeys = new KeyCodeCombination(KeyCode.C, KeyCombination.SHORTCUT_DOWN);
        var selectAllKeys = new KeyCodeCombination(KeyCode.A, KeyCombination.SHORTCUT_DOWN);
        logListView.addEventFilter(
                KeyEvent.KEY_PRESSED,
                e -> {
                    if (copyKeys.match(e)) {
                        copySelectedLogLinesToClipboard();
                        e.consume();
                    } else if (selectAllKeys.match(e)) {
                        logListView.getSelectionModel().selectAll();
                        e.consume();
                    }
                });
        MenuItem copySelectedItem =
                new MenuItem(
                        "選択をコピー (Ctrl+C)");
        copySelectedItem.setOnAction(e -> copySelectedLogLinesToClipboard());
        MenuItem copyAllItem =
                new MenuItem(
                        "全ログをコピー（バッファ全行）");
        copyAllItem.setOnAction(e -> copyAllBufferedLogToClipboard());
        MenuItem copyVisibleItem =
                new MenuItem(
                        "表示中のログをコピー");
        copyVisibleItem.setOnAction(e -> copyVisibleLogLinesToClipboard());
        logListView.setContextMenu(
                new ContextMenu(copySelectedItem, copyVisibleItem, copyAllItem));
    }

    /** Full buffer (ignores filter); same as toolbar 全ログをコピー. */
    private void copyAllBufferedLogToClipboard() {
        if (logLinesAll.isEmpty()) {
            return;
        }
        String text = String.join("\n", logLinesAll);
        ClipboardContent cc = new ClipboardContent();
        cc.putString(text);
        Clipboard.getSystemClipboard().setContent(cc);
    }

    private void copyVisibleLogLinesToClipboard() {
        if (logLinesVisible.isEmpty()) {
            return;
        }
        String text = String.join("\n", logLinesVisible);
        ClipboardContent cc = new ClipboardContent();
        cc.putString(text);
        Clipboard.getSystemClipboard().setContent(cc);
    }

    private void copySelectedLogLinesToClipboard() {
        ObservableList<String> visible = logListView.getItems();
        MultipleSelectionModel<String> sm = logListView.getSelectionModel();
        ArrayList<Integer> indices = new ArrayList<>(sm.getSelectedIndices());
        if (indices.isEmpty()) {
            int fi = logListView.getFocusModel().getFocusedIndex();
            if (fi >= 0) {
                indices.add(fi);
            }
        }
        if (indices.isEmpty()) {
            return;
        }
        Collections.sort(indices);
        StringBuilder sb = new StringBuilder();
        for (int i : indices) {
            if (i >= 0 && i < visible.size()) {
                if (sb.length() > 0) {
                    sb.append('\n');
                }
                sb.append(visible.get(i));
            }
        }
        if (sb.length() == 0) {
            return;
        }
        ClipboardContent cc = new ClipboardContent();
        cc.putString(sb.toString());
        Clipboard.getSystemClipboard().setContent(cc);
    }

    /** Reapply row styles when UI theme (dark/light) changes. */
    void refreshLogThemeCells() {
        if (logListView != null) {
            boolean dark = shell != null && shell.currentDesktopTheme().isDarkUi();
            if (dark) {
                if (!logListView.getStyleClass().contains("log-dark-surface")) {
                    logListView.getStyleClass().add("log-dark-surface");
                }
            } else {
                logListView.getStyleClass().remove("log-dark-surface");
            }
            logListView.refresh();
        }
    }

    private void refreshLogLinesVisiblePredicate() {
        LogViewFilter kindFilter =
                logFilterCombo != null && logFilterCombo.getValue() != null
                        ? logFilterCombo.getValue()
                        : LogViewFilter.ALL;
        String search = snapshotLogSearchText();
        if (search.isEmpty()) {
            logLinesVisible.setPredicate(kindFilter::test);
            return;
        }
        String searchLower = search.toLowerCase(Locale.ROOT);
        logLinesVisible.setPredicate(
                line ->
                        line != null
                                && kindFilter.test(line)
                                && line.toLowerCase(Locale.ROOT).contains(searchLower));
    }

    private String snapshotLogSearchText() {
        if (logSearchField == null || logSearchField.getText() == null) {
            return "";
        }
        return logSearchField.getText().trim();
    }

    private Color resolveLogTextFill(LogLineKind kind) {
        boolean dark = shell != null && shell.currentDesktopTheme().isDarkUi();
        return switch (kind) {
            case ERROR -> dark ? Color.web("#fecaca") : Color.web("#991b1b");
            case WARN -> dark ? Color.web("#fde68a") : Color.web("#b45309");
            case NORMAL -> dark ? Color.web("#e2e8f0") : Color.web("#1e293b");
        };
    }

    void bindShell(MainShellController shell) {
        if (this.shell != null && pipelineTimingHistoryListener != null) {
            this.shell.pipelineExecutionTimingHistory().removeChangeListener(pipelineTimingHistoryListener);
        }
        this.shell = shell;
        refreshAppVersionLabel();
        refreshOpenWorkbookHintLabels();
        pipelineTimingHistoryListener = () -> Platform.runLater(this::refreshPipelineExecutionTimingLabels);
        if (shell != null) {
            shell.pipelineExecutionTimingHistory().addChangeListener(pipelineTimingHistoryListener);
            refreshPipelineExecutionTimingLabels();
        }
        refreshLogThemeCells();
        refreshMachineDeliveryManagementOpenButton();
    }

    /** 湖南工場のみ「マシン別納期管理表を開く」を表示する。国分では非表示。 */
    void refreshMachineDeliveryManagementOpenButton() {
        if (openMachineDeliveryManagementXlsmButton == null) {
            return;
        }
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        boolean show = AppPaths.isMachineDeliveryManagementXlsmEnabled(ui);
        openMachineDeliveryManagementXlsmButton.setVisible(show);
        openMachineDeliveryManagementXlsmButton.setManaged(show);
    }

    /** {@link GlobalInitSettingTarget} の工場に合わせてロゴを更新する（ツールバー側）。 */
    void refreshFactorySiteLogo() {
        if (shell != null) {
            shell.refreshShellFactoryOperatorToolbar();
        }
    }

    void refreshFactorySiteComboPresentation() {
        if (shell != null) {
            shell.refreshShellFactorySiteComboPresentation();
        }
    }

    void setFactorySiteComboDisabled(boolean disabled) {
        if (shell != null) {
            shell.setShellFactorySiteComboDisabled(disabled);
        }
    }

    void refreshFactorySiteComboFromStore() {
        if (shell != null) {
            shell.refreshShellFactorySiteComboFromStore();
        }
        refreshMachineDeliveryManagementOpenButton();
    }

    void refreshOperatorUserLabel() {
        if (shell != null) {
            shell.refreshShellOperatorUserLabel();
        }
    }

    /** ゲスト操作者時は実行・ログタブ本体のみ抑止する（工場コンボは最上部ツールバー）。 */
    void setGuestSessionFactorySwitchOnly(boolean guestOnly) {
        guestSessionFactorySwitchOnly = guestOnly;
        if (shell != null) {
            shell.setGuestSessionFactoryToolbar(guestOnly);
        }
        applyGuestSessionFactorySwitchOnlyState();
        applyStage1RunButtonEnabledState();
    }

    private void applyGuestSessionFactorySwitchOnlyState() {
        boolean guestOnly = guestSessionFactorySwitchOnly;
        if (runTabRoot == null) {
            return;
        }
        if (runTabRoot.getCenter() != null) {
            runTabRoot.getCenter().setDisable(guestOnly);
        }
        if (runTabRoot.getBottom() != null) {
            runTabRoot.getBottom().setDisable(guestOnly);
        }
        javafx.scene.Node top = runTabRoot.getTop();
        if (!(top instanceof VBox topVBox) || topVBox.getChildren().isEmpty()) {
            return;
        }
        for (javafx.scene.Node child : topVBox.getChildren()) {
            child.setDisable(guestOnly);
        }
    }

    /**
     * 実行・ログタブの「開く」横ラベルを環境変数（マスタ名・サマリブック）に合わせて更新する。
     */
    void refreshOpenWorkbookHintLabels() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        if (masterWorkbookOpenHintLabel != null) {
            String alt = ui.getOrDefault(AppPaths.KEY_PM_AI_MASTER_WORKBOOK, "").trim();
            if (!alt.isEmpty()) {
                masterWorkbookOpenHintLabel.setText(Path.of(alt).getFileName().toString());
            } else {
                masterWorkbookOpenHintLabel.setText("master.xlsm");
            }
        }
    }

    /** 実行・ログタブの {@code version.txt} 表示を更新する（ポータブル同期後など）。 */
    void refreshAppVersionLabel() {
        if (appVersionLabel == null || shell == null) {
            return;
        }
        Path cwd = Paths.get(System.getProperty("user.dir", "."));
        String v = AppVersionInfo.resolveDisplayedVersion(cwd, shell.snapshotUiEnv());
        appVersionLabel.setText("バージョン " + v);
    }

    /**
     * Restores font controls from session; must run after FXML {@link #initialize}.
     */
    void applyLogFontFromSession(String family, double sizePoints) {
        suppressLogFontEvents.set(true);
        try {
            if (family != null && !family.isBlank()) {
                if (!logFontFamilyCombo.getItems().contains(family)) {
                    int insertAt = 1;
                    logFontFamilyCombo.getItems().add(insertAt, family);
                }
                logFontFamilyCombo.setValue(family);
            } else {
                logFontFamilyCombo.getSelectionModel().selectFirst();
            }
            double effectiveSize =
                    sizePoints > 0 && Double.isFinite(sizePoints)
                            ? sizePoints
                            : Font.getDefault().getSize();
            if (!logFontSizeCombo.getItems().contains(effectiveSize)) {
                List<Double> extended = new ArrayList<>(logFontSizeCombo.getItems());
                extended.add(effectiveSize);
                Collections.sort(extended);
                logFontSizeCombo.getItems().setAll(extended);
            }
            logFontSizeCombo.setValue(effectiveSize);
        } finally {
            suppressLogFontEvents.set(false);
        }
        applyLogAreaFont();
    }

    String snapshotLogFontFamily() {
        String v = logFontFamilyCombo != null ? logFontFamilyCombo.getValue() : null;
        if (v == null || v.equals(DEFAULT_FONT_FAMILY_LABEL)) {
            return "";
        }
        return v;
    }

    double snapshotLogFontSize() {
        Double v = logFontSizeCombo != null ? logFontSizeCombo.getValue() : null;
        if (v == null || !Double.isFinite(v) || v <= 0) {
            return 0d;
        }
        return v;
    }

    private void applyLogAreaFont() {
        if (logFontFamilyCombo == null || logFontSizeCombo == null) {
            return;
        }
        String choice = logFontFamilyCombo.getValue();
        Double szObj = logFontSizeCombo.getValue();
        double size =
                szObj != null && szObj > 0 && Double.isFinite(szObj)
                        ? szObj
                        : Font.getDefault().getSize();
        if (choice == null || choice.equals(DEFAULT_FONT_FAMILY_LABEL)) {
            appliedLogFont = Font.font(size);
        } else {
            appliedLogFont = Font.font(choice, size);
        }
        if (logListView != null) {
            applyLogListFixedCellHeight();
            logListView.refresh();
        }
    }

    @FXML
    private void onStage1RunButtonAction() {
        shell.triggerStage1();
    }

    @FXML
    private void onOpenStage2ProductionPlanAction() {
        openExcelBesideField(stage2ProductionPlanField, "stage2-production-plan");
    }

    @FXML
    private void onOpenStage2MemberScheduleAction() {
        openExcelBesideField(stage2MemberScheduleField, "stage2-member-schedule");
    }

    @FXML
    private void onOpenMasterWorkbookAction() {
        if (shell == null) {
            return;
        }
        Path p =
                AppPaths.resolveMasterWorkbookPathForDesktopOpen(
                        shell.snapshotUiEnv(),
                        shell.effectiveTaskInputWorkbookPathForShell());
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[master-workbook] file not found: "
                            + p
                            + " (set "
                            + AppPaths.KEY_PM_AI_MASTER_WORKBOOK
                            + ", or check "
                            + AppPaths.KEY_PM_AI_REPO_ROOT
                            + ")");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[master-workbook] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[master-workbook] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onOpenMachineDeliveryManagementXlsmAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        if (!AppPaths.isMachineDeliveryManagementXlsmEnabled(ui)) {
            return;
        }
        Optional<Path> resolved = AppPaths.resolveMachineDeliveryManagementXlsm(ui);
        if (resolved.isEmpty()) {
            appendLog(
                    "[machine-delivery-xlsm] path not set (set "
                            + AppPaths.KEY_PM_AI_MACHINE_DELIVERY_MANAGEMENT_XLSM
                            + " in env tab; Konan has a factory default)");
            return;
        }
        Path p = resolved.get();
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[machine-delivery-xlsm] file not found: "
                            + p
                            + " (set "
                            + AppPaths.KEY_PM_AI_MACHINE_DELIVERY_MANAGEMENT_XLSM
                            + ")");
            return;
        }
        try {
            DesktopFileOpener.openFileReadOnly(p);
            appendLog("[machine-delivery-xlsm] opened read-only: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[machine-delivery-xlsm] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onOpenDispatchUsageGuideDocxAction() {
        if (shell == null) {
            return;
        }
        Path p = AppPaths.resolveDispatchUsageGuideDocx(shell.snapshotUiEnv());
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
            appendLog("[dispatch-usage-docx] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[dispatch-usage-docx] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onOpenDispatchRulesHtmlAction() {
        if (shell == null) {
            return;
        }
        Path p = AppPaths.resolveDispatchRulesHtml(shell.snapshotUiEnv());
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[dispatch-rules-html] file not found: "
                            + p
                            + " (expected "
                            + AppPaths.DISPATCH_RULES_HTML_REL
                            + " under "
                            + AppPaths.KEY_PM_AI_REPO_ROOT
                            + ")");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[dispatch-rules-html] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[dispatch-rules-html] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onOpenManualAction() {
        if (shell == null) {
            return;
        }
        Path p = AppPaths.resolveManualIndexHtml(shell.snapshotUiEnv());
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[manual] file not found: "
                            + p
                            + " (publish HTML per manual/README.md, or set "
                            + AppPaths.KEY_PM_AI_REPO_ROOT
                            + " if the repository root is wrong)");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[manual] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[manual] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onCopyAllLogButtonAction() {
        copyAllBufferedLogToClipboard();
    }

    @FXML
    private void onClearLogButtonAction() {
        clearMainRunTabLog();
    }

    /** メイン実行タブのログ一覧を空にする（クリアボタンと同一。ポータルバージョンアップ完了後など）。 */
    void clearMainRunTabLog() {
        Runnable clear =
                () -> {
                    flushPendingLogAppendsOnFxThread();
                    logLinesAll.clear();
                    if (shell != null) {
                        shell.scheduleDesktopSessionSave();
                    }
                };
        if (Platform.isFxApplicationThread()) {
            clear.run();
        } else {
            Platform.runLater(clear);
        }
    }

    /**
     * ポータブルバージョンアップ前: ログフィルタを「すべて」にし、同期行が実行・ログに見えるようにする。
     */
    void prepareRunTabForPortableBundleSync() {
        Runnable prep =
                () -> {
                    if (logFilterCombo != null) {
                        logFilterCombo.setValue(LogViewFilter.ALL);
                    }
                    if (logSearchField != null) {
                        logSearchField.clear();
                    }
                };
        if (Platform.isFxApplicationThread()) {
            prep.run();
        } else {
            Platform.runLater(prep);
        }
    }

    /**
     * バックグラウンドのポータル同期から呼ぶ。通常ログ追記と同じバッチ機構へ委譲する。
     */
    void appendPortableBundleSyncLog(String line) {
        appendLog(line, true);
    }

    /** 同期スレッド終了時に残りを UI へ反映する。 */
    void flushPortableBundleSyncLog() {
        flushPendingLogAppends();
        if (shell != null) {
            shell.scheduleDesktopSessionSave();
        }
    }

    TextField getWorkbookField() {
        return workbookField;
    }

    TextField getScriptDirField() {
        return scriptDirField;
    }

    ListView<String> getLogListView() {
        return logListView;
    }

    Label getStatusLabel() {
        return statusLabel;
    }

    /** 履歴ストアの最新値で実行・ログタブのラベルを更新する。 */
    void refreshPipelineExecutionTimingLabels() {
        if (shell == null) {
            return;
        }
        PipelineExecutionTimingHistoryStore store = shell.pipelineExecutionTimingHistory();
        for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
            applyPipelineExecutionTimingLabel(kind, store.lastDurationMs(kind));
        }
    }

    static String formatPipelineExecutionDuration(long durationMs) {
        if (durationMs < 0L) {
            return "—";
        }
        if (durationMs < 60_000L) {
            return String.format("%.2f秒", durationMs / 1000.0);
        }
        long minutes = durationMs / 60_000L;
        double seconds = (durationMs % 60_000L) / 1000.0;
        return String.format("%d分%.1f秒", minutes, seconds);
    }

    private void applyPipelineExecutionTimingLabel(PipelineExecutionTimingKind kind, long durationMs) {
        Label target =
                switch (kind) {
                    case STAGE1 -> pipelineTimingStage1Label;
                    case STAGE2_0 -> pipelineTimingStage20Label;
                    case STAGE2_1 -> pipelineTimingStage21Label;
                    case DISPATCH_TRIAL -> pipelineTimingDispatchTrialLabel;
                    case SUMMARY_EXCEL -> pipelineTimingSummaryExcelLabel;
                    case DELIVERY_CALENDAR_VIEW -> pipelineTimingDeliveryCalendarLabel;
                };
        if (target != null) {
            target.setText(
                    durationMs >= 0L
                            ? formatPipelineExecutionDuration(durationMs)
                            : "—");
        }
    }

    /** 段階2子プロセスへ渡す {@code PM_AI_STAGE2_SKIP_TODAY_DISPATCH}（排他2択は段階1ボタン横）。 */
    boolean snapshotStage2SkipTodayDispatch() {
        if (!snapshotTodayDispatch()) {
            return true;
        }
        return stage2SkipTodayDispatchWhenTodayMode;
    }

    /** 当日配台する（朝運用・ソース固定）。「当日は配台しない」との排他2択。 */
    boolean snapshotTodayDispatch() {
        return todayDispatchRadio != null && todayDispatchRadio.isSelected();
    }

    /**
     * セッションから当日配台モード（排他2択）を復元する。
     *
     * <p>{@code todayDispatch=true} のとき {@code skipToday} は当日配台モード中の skip 上書きとして保持する。
     * {@code todayDispatch=false} は常に「当日は配台しない」（旧来の skip=false かつ today=false もこちらへ正規化）。
     */
    void applyTodayDispatchModeFromSession(boolean skipToday, boolean todayDispatch) {
        suppressingTodayDispatchModeListener = true;
        try {
            if (todayDispatch && todayDispatchRadio != null) {
                todayDispatchRadio.setSelected(true);
                stage2SkipTodayDispatchWhenTodayMode = skipToday;
            } else if (stage2SkipTodayDispatchRadio != null) {
                stage2SkipTodayDispatchRadio.setSelected(true);
                stage2SkipTodayDispatchWhenTodayMode = false;
            }
        } finally {
            suppressingTodayDispatchModeListener = false;
        }
        if (shell != null) {
            shell.refreshPlanInputNextDayDialogCoupling();
        }
    }

    /**
     * skip_today のみ更新する。当日配台モード中はラジオを切り替えず上書き値だけ変える（取得時刻ポリシー用）。
     * 「当日は配台しない」選択中は skip は常に true のため、呼び出しは無視する。
     */
    void applyStage2SkipTodayDispatchFromSession(boolean skipToday) {
        if (!snapshotTodayDispatch()) {
            return;
        }
        stage2SkipTodayDispatchWhenTodayMode = skipToday;
        if (shell != null) {
            shell.refreshPlanInputNextDayDialogCoupling();
            shell.scheduleDesktopSessionSave();
        }
    }

    private void wireTodayDispatchModeExclusiveRadios() {
        if (stage2SkipTodayDispatchRadio != null) {
            stage2SkipTodayDispatchRadio.setSelected(true);
        }
        if (todayDispatchRadio != null) {
            todayDispatchRadio.setSelected(false);
        }
        stage2SkipTodayDispatchWhenTodayMode = false;
        if (todayDispatchModeToggleGroup != null) {
            todayDispatchModeToggleGroup
                    .selectedToggleProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (suppressingTodayDispatchModeListener) {
                                    return;
                                }
                                if (b == null) {
                                    suppressingTodayDispatchModeListener = true;
                                    try {
                                        if (stage2SkipTodayDispatchRadio != null) {
                                            stage2SkipTodayDispatchRadio.setSelected(true);
                                        }
                                    } finally {
                                        suppressingTodayDispatchModeListener = false;
                                    }
                                    return;
                                }
                                if (b == todayDispatchRadio) {
                                    stage2SkipTodayDispatchWhenTodayMode = false;
                                }
                                if (shell != null) {
                                    shell.refreshPlanInputNextDayDialogCoupling();
                                    shell.scheduleDesktopSessionSave();
                                }
                            });
        }
    }

    /**
     * 段階1／段階2 実行中は段階1ボタンの再実行を無効化する（進捗・中断はメインシェルツールバーのみ）。段階2実行ボタンは
     * {@link PlanInputTabController} 側。
     */
    void setStageRunProgressVisible(boolean stage1Running, boolean stage2Running) {
        stage1RunPipelineBusy = stage1Running || stage2Running;
        applyStage1RunButtonEnabledState();
    }

    void setDeliveryCalendarReloadBlocking(boolean blocking) {
        deliveryCalendarReloadBlocking = blocking;
        applyStage1RunButtonEnabledState();
    }

    void setStage1BlockedByPipelineCheck(boolean blocked, String tooltip, String badgeText) {
        stage1BlockedByPipelineCheck = blocked;
        stage1PipelineCheckBlockTooltip = tooltip != null ? tooltip : "";
        stage1PipelineCheckBlockBadgeText = badgeText != null ? badgeText.strip() : "";
        applyStage1RunButtonEnabledState();
    }

    void setStage1BlockedByPipelineCheck(boolean blocked, String tooltip) {
        setStage1BlockedByPipelineCheck(blocked, tooltip, "");
    }

    void setCalendarReadinessBlocked(boolean blocked, String tooltip) {
        stage1BlockedByCalendarNotReady = blocked;
        calendarReadinessBlockTooltip = tooltip != null ? tooltip.strip() : "";
        applyStage1RunButtonEnabledState();
    }

    private void applyStage1RunButtonEnabledState() {
        if (stage1RunButton == null) {
            return;
        }
        stage1RunButton.setDisable(
                guestSessionFactorySwitchOnly
                        || stage1RunPipelineBusy
                        || deliveryCalendarReloadBlocking
                        || stage1BlockedByPipelineCheck
                        || stage1BlockedByCalendarNotReady);
        if (deliveryCalendarReloadBlocking && !stage1RunPipelineBusy) {
            stage1RunButton.setText(STAGE1_RUN_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD);
            stage1RunButton.setTooltip(null);
        } else if (stage1BlockedByCalendarNotReady && !stage1RunPipelineBusy) {
            stage1RunButton.setText(STAGE1_RUN_BUTTON_TEXT_DEFAULT);
            if (!calendarReadinessBlockTooltip.isBlank()) {
                stage1RunButton.setTooltip(new Tooltip(calendarReadinessBlockTooltip));
            } else {
                stage1RunButton.setTooltip(
                        new Tooltip(
                                "カレンダー正本 JSON が未準備です。"
                                        + " 会社カレンダー・メンバー勤怠・機械カレンダーをセットアップしてください。"));
            }
        } else if (stage1BlockedByPipelineCheck && !stage1RunPipelineBusy) {
            stage1RunButton.setText(STAGE1_RUN_BUTTON_TEXT_DEFAULT);
            if (!stage1PipelineCheckBlockTooltip.isBlank()) {
                stage1RunButton.setTooltip(new Tooltip(stage1PipelineCheckBlockTooltip));
            } else {
                stage1RunButton.setTooltip(
                        new Tooltip("原本転記・計画確認タブで問題を確認してください。"));
            }
        } else {
            stage1RunButton.setText(STAGE1_RUN_BUTTON_TEXT_DEFAULT);
            stage1RunButton.setTooltip(null);
        }
        applyStage1PipelineCheckBlockBadge();
    }

    private void applyStage1PipelineCheckBlockBadge() {
        if (stage1PipelineCheckBlockBadge == null) {
            return;
        }
        boolean show =
                stage1BlockedByPipelineCheck
                        && !stage1RunPipelineBusy
                        && !deliveryCalendarReloadBlocking
                        && !stage1PipelineCheckBlockBadgeText.isBlank();
        stage1PipelineCheckBlockBadge.setText(
                show ? stage1PipelineCheckBlockBadgeText : "");
        stage1PipelineCheckBlockBadge.setManaged(show);
        stage1PipelineCheckBlockBadge.setVisible(show);
        if (show && !stage1PipelineCheckBlockTooltip.isBlank()) {
            stage1PipelineCheckBlockBadge.setTooltip(
                    new Tooltip(stage1PipelineCheckBlockTooltip));
        } else {
            stage1PipelineCheckBlockBadge.setTooltip(null);
        }
    }

    /**
     * ネットワークソースが使えずキャッシュを読んだとき、段階1ボタン左にバッジを表示する。
     */
    void setStage1NetworkCacheBadge(boolean visible, PersonBadgeStyle style, String labelText) {
        Platform.runLater(
                () -> {
                    if (stage1NetworkCacheBadgeHost == null) {
                        return;
                    }
                    stage1NetworkCacheBadgeHost.getChildren().clear();
                    stage1NetworkCacheBadgeHost.setManaged(visible);
                    stage1NetworkCacheBadgeHost.setVisible(visible);
                    if (!visible || style == null) {
                        return;
                    }
                    String t =
                            labelText != null && !labelText.isBlank() ? labelText.strip() : "キャッシュ";
                    StackPane graphic = PersonBadgeNodeFactory.createBadge(t, style, 1.0, 14.0);
                    Tooltip.install(
                            graphic,
                            new Tooltip(
                                    "PM_AI_TASK_INPUT_SOURCE_DIR または "
                                            + "PM_AI_ACTUAL_DETAIL_SOURCE_DIR "
                                            + "を参照できず、リポジトリ配下の最終キャッシュを使用して段階1／段階2に渡します。"));
                    stage1NetworkCacheBadgeHost.getChildren().add(graphic);
                });
    }

    /** 勤怠正本が段階2未準備のとき、段階1ボタン付近に警告バッジを表示する。 */
    void setAttendanceReadinessBadge(
            boolean visible,
            PersonBadgeStyle style,
            String labelText,
            String tooltipText) {
        Platform.runLater(
                () -> {
                    if (attendanceReadinessBadgeHost == null) {
                        return;
                    }
                    attendanceReadinessBadgeHost.getChildren().clear();
                    attendanceReadinessBadgeHost.setManaged(visible);
                    attendanceReadinessBadgeHost.setVisible(visible);
                    if (!visible || style == null) {
                        return;
                    }
                    String t =
                            labelText != null && !labelText.isBlank()
                                    ? labelText.strip()
                                    : "勤怠未準備";
                    StackPane graphic = PersonBadgeNodeFactory.createBadge(t, style, 1.0, 14.0);
                    String tip =
                            tooltipText != null && !tooltipText.isBlank()
                                    ? tooltipText.strip()
                                    : "勤怠正本（attendance-data.json / machine-calendar-data.json）が未準備です。"
                                            + " 会社カレンダー・メンバー勤怠・機械カレンダータブでセットアップしてください。";
                    Tooltip.install(graphic, new Tooltip(tip));
                    attendanceReadinessBadgeHost.getChildren().add(graphic);
                });
    }

    /**
     * Fills read-only paths after stage-2 success (newest files under {@link
     * jp.co.pm.ai.desktop.config.AppPaths#defaultPlanningOutputDir}).
     */
    void setStage2ArtifactPaths(String productionPlanPath, String memberSchedulePath) {
        if (stage2ProductionPlanField != null) {
            stage2ProductionPlanField.setText(productionPlanPath != null ? productionPlanPath : "");
        }
        if (stage2MemberScheduleField != null) {
            stage2MemberScheduleField.setText(memberSchedulePath != null ? memberSchedulePath : "");
        }
        if (shell != null) {
            shell.scheduleDesktopSessionSave();
        }
    }

    void updateStage2Progress(MainRunStage2Progress.State state, String detail) {
        MainRunStage2Progress.apply(stage2ProgressPane, stage2ProgressLabel, state, detail);
        if (stage2ProgressAccordion != null && stage2ProgressPane != null) {
            stage2ProgressAccordion.setExpandedPane(stage2ProgressPane);
            if (state == MainRunStage2Progress.State.COMPLETED) {
                stage2ProgressPane.setExpanded(true);
            }
        }
        if (shell != null) {
            shell.syncStageRunBusyFromStage2Progress(state, detail);
        }
    }

    String snapshotStage2ProductionPlanPath() {
        if (stage2ProductionPlanField == null || stage2ProductionPlanField.getText() == null) {
            return "";
        }
        return stage2ProductionPlanField.getText().trim();
    }

    String snapshotStage2MemberSchedulePath() {
        if (stage2MemberScheduleField == null || stage2MemberScheduleField.getText() == null) {
            return "";
        }
        return stage2MemberScheduleField.getText().trim();
    }

    /** 開発用チェックのうち「全依頼を配台不要」は段階1実行後に毎回オフ。AI API スキップはユーザー選択を保持（セッション永続化）。 */
    void resetDevCheckboxesAfterStage1Run() {
        suppressStage1MarkAllExcludeAfterRunEvents.set(true);
        try {
            if (stage1MarkAllExcludeAfterRunCheckBox != null) {
                stage1MarkAllExcludeAfterRunCheckBox.setSelected(false);
            }
        } finally {
            suppressStage1MarkAllExcludeAfterRunEvents.set(false);
        }
        if (shell != null) {
            shell.scheduleDesktopSessionSave();
        }
    }

    String snapshotStage2ResultBookFont() {
        if (stage2ResultBookFontCombo == null) {
            return "";
        }
        String v = stage2ResultBookFontCombo.getValue();
        if (v == null
                || v.isBlank()
                || v.equals(DEFAULT_FONT_FAMILY_LABEL)) {
            return "";
        }
        return v.trim();
    }

    void applyStage2ResultBookFontFromSession(String family) {
        if (stage2ResultBookFontCombo == null) {
            return;
        }
        suppressStage2ResultFontEvents.set(true);
        try {
            if (family != null && !family.isBlank()) {
                if (!stage2ResultBookFontCombo.getItems().contains(family)) {
                    stage2ResultBookFontCombo.getItems().add(1, family);
                }
                stage2ResultBookFontCombo.setValue(family);
            } else {
                stage2ResultBookFontCombo.getSelectionModel().selectFirst();
            }
        } finally {
            suppressStage2ResultFontEvents.set(false);
        }
    }

    /** 子プロセスへ渡す {@code PM_AI_SKIP_GEMINI_API}（チェックは本タブ「その他」アコーディオン）。 */
    boolean snapshotSkipGeminiApi() {
        return skipGeminiApiCheckBox != null && skipGeminiApiCheckBox.isSelected();
    }

    void applySkipGeminiApiFromSession(boolean skip) {
        if (skipGeminiApiCheckBox == null) {
            return;
        }
        suppressSkipGeminiApiEvents.set(true);
        try {
            skipGeminiApiCheckBox.setSelected(skip);
        } finally {
            suppressSkipGeminiApiEvents.set(false);
        }
    }

    /** 段階1正常終了後に全依頼を配台不要 yes にする（開発用チェック）。 */
    boolean snapshotStage1MarkAllExcludeAfterRun() {
        return stage1MarkAllExcludeAfterRunCheckBox != null
                && stage1MarkAllExcludeAfterRunCheckBox.isSelected();
    }

    void applyStage1MarkAllExcludeAfterRunFromSession(boolean enabled) {
        if (stage1MarkAllExcludeAfterRunCheckBox == null) {
            return;
        }
        suppressStage1MarkAllExcludeAfterRunEvents.set(true);
        try {
            stage1MarkAllExcludeAfterRunCheckBox.setSelected(enabled);
        } finally {
            suppressStage1MarkAllExcludeAfterRunEvents.set(false);
        }
    }

    void appendLog(String line) {
        appendLog(line, true);
    }

    /**
     * Appends one log line. When {@code scrollToEnd} is false, vertical scroll is unchanged (used for boot
     * lines before restoring session scroll).
     */
    void appendLog(String line, boolean scrollToEnd) {
        if (line == null || line.isEmpty()) {
            return;
        }
        List<String> parts = splitLogLineParts(line);
        if (parts.size() > 1) {
            for (String part : parts) {
                appendLogLine(part, scrollToEnd);
            }
            return;
        }
        appendLogLine(line, scrollToEnd);
    }

    private static List<String> splitLogLineParts(String line) {
        if (line == null || line.isEmpty()) {
            return List.of();
        }
        if (line.indexOf('\n') < 0 && line.indexOf('\r') < 0) {
            return List.of(line);
        }
        String normalized = line.replace("\r\n", "\n").replace('\r', '\n');
        String[] raw = normalized.split("\n");
        List<String> parts = new ArrayList<>(raw.length);
        for (String part : raw) {
            if (!part.isEmpty()) {
                parts.add(part);
            }
        }
        return parts;
    }

    private void appendLogLine(String line, boolean scrollToEnd) {
        if (line == null || line.isEmpty()) {
            return;
        }
        if (!scrollToEnd) {
            Runnable addImmediate =
                    () -> {
                        flushPendingLogAppendsOnFxThread();
                        logLinesAll.add(line);
                    };
            if (Platform.isFxApplicationThread()) {
                addImmediate.run();
            } else {
                Platform.runLater(addImmediate);
            }
            return;
        }
        synchronized (logAppendLock) {
            logAppendPending.add(line);
            boolean immediate = logAppendPending.size() >= LOG_APPEND_BATCH_MAX;
            scheduleLogAppendFlush(immediate);
        }
    }

    /**
     * Restores log buffer, filter, and schedules scroll restore after layout (see {@link
     * #flushPendingSessionLogScroll()}).
     */
    void restoreRunLogUiFromSession(
            String filterName, List<String> lines, double scrollProportion) {
        Runnable restore =
                () -> {
                    flushPendingLogAppendsOnFxThread();
                    suppressRunLogSessionPersistence.set(true);
                    try {
                        if (lines != null && !lines.isEmpty()) {
                            List<String> expanded = new ArrayList<>(lines.size());
                            for (String line : lines) {
                                expanded.addAll(splitLogLineParts(line));
                            }
                            logLinesAll.setAll(expanded);
                        } else {
                            logLinesAll.clear();
                        }
                        logFilterCombo.setValue(LogViewFilter.fromStoredName(filterName));
                        refreshLogLinesVisiblePredicate();
                    } finally {
                        suppressRunLogSessionPersistence.set(false);
                    }
                    pendingSessionLogScroll = scrollProportion;
                };
        if (Platform.isFxApplicationThread()) {
            restore.run();
        } else {
            Platform.runLater(restore);
        }
    }

    /** Applies {@link #pendingSessionLogScroll} once the log {@link ListView} is laid out. */
    void flushPendingSessionLogScroll() {
        double p = pendingSessionLogScroll;
        pendingSessionLogScroll = Double.NaN;
        if (!Double.isFinite(p)) {
            return;
        }
        logTailFollowEnabled = p >= LOG_AUTO_SCROLL_BOTTOM_THRESHOLD;
        applyLogScrollProportion(p);
        Platform.runLater(() -> applyLogScrollProportion(p));
    }

    String snapshotLogFilterName() {
        LogViewFilter v = logFilterCombo != null ? logFilterCombo.getValue() : null;
        return v != null ? v.name() : LogViewFilter.ALL.name();
    }

    List<String> snapshotPersistedLogLines() {
        int n = logLinesAll.size();
        if (n <= MAX_PERSISTED_LOG_LINES) {
            return List.copyOf(logLinesAll);
        }
        return List.copyOf(logLinesAll.subList(n - MAX_PERSISTED_LOG_LINES, n));
    }

    /**
     * 実行・ログタブの全文（フィルタ前の蓄積バッファ）を改行連結する。リモートサポート用ログ保存向け。
     */
    String snapshotAllLogText() {
        flushPendingLogAppendsOnFxThread();
        if (logLinesAll.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder(Math.min(logLinesAll.size() * 80, 1 << 20));
        for (int i = 0; i < logLinesAll.size(); i++) {
            if (i > 0) {
                sb.append('\n');
            }
            sb.append(logLinesAll.get(i));
        }
        if (!logLinesAll.isEmpty()) {
            sb.append('\n');
        }
        return sb.toString();
    }

    double snapshotLogScrollProportion() {
        return readVerticalScrollProportion(logListView);
    }

    private static double readVerticalScrollProportion(ListView<?> listView) {
        if (listView == null) {
            return Double.NaN;
        }
        ScrollBar sb = (ScrollBar) listView.lookup(".scroll-bar:vertical");
        if (sb == null) {
            return Double.NaN;
        }
        double min = sb.getMin();
        double max = sb.getMax();
        double v = sb.getValue();
        if (max <= min) {
            return 1d;
        }
        return (v - min) / (max - min);
    }

    private void applyLogScrollProportion(double proportion) {
        if (logListView == null || !Double.isFinite(proportion)) {
            return;
        }
        double p = Math.max(0d, Math.min(1d, proportion));
        suppressLogTailFollowListener.set(true);
        try {
            ScrollBar sb = (ScrollBar) logListView.lookup(".scroll-bar:vertical");
            if (sb == null) {
                return;
            }
            double min = sb.getMin();
            double max = sb.getMax();
            if (max > min) {
                sb.setValue(min + p * (max - min));
            }
        } finally {
            suppressLogTailFollowListener.set(false);
        }
    }

    private void openExcelBesideField(TextField field, String logTag) {
        String raw = field != null && field.getText() != null ? field.getText().trim() : "";
        if (raw.isEmpty()) {
            appendLog("[" + logTag + "] path is empty");
            return;
        }
        Path p = Paths.get(raw);
        if (!Files.isRegularFile(p)) {
            appendLog("[" + logTag + "] file not found: " + p);
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[" + logTag + "] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[" + logTag + "] open failed: " + e.getMessage());
        }
    }
}
