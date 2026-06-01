package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.util.Duration;
import javafx.beans.binding.Bindings;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.ContextMenu;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.MenuItem;
import javafx.scene.control.MultipleSelectionModel;
import javafx.scene.control.OverrunStyle;
import javafx.scene.control.ScrollBar;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyCodeCombination;
import javafx.scene.input.KeyCombination;
import javafx.scene.effect.DropShadow;
import javafx.scene.image.ImageView;
import javafx.scene.input.KeyEvent;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.scene.text.TextFlow;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.AppVersionInfo;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.FactorySiteLogoSupport;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.io.SummaryAiDispatchExportLock;
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
    private TextField stage2ProductionPlanField;

    @FXML
    private TextField stage2MemberScheduleField;

    @FXML
    private StackPane stage1NetworkCacheBadgeHost;

    @FXML
    private Button stage1RunButton;

    private static final String STAGE1_RUN_BUTTON_TEXT_DEFAULT = "段階1 実行";

    private static final String STAGE1_RUN_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD =
            "段階1（納期管理ビュー更新中）";

    @FXML
    private CheckBox stage1ClearCacheAndRunCheckBox;

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

    @FXML
    private Label summaryWorkbookOpenHintLabel;

    @FXML
    private Button openSummaryAiDispatchButton;

    @FXML
    private Button forceUnlockSummaryAiDispatchButton;

    @FXML
    private StackPane factoryLogoHost;

    private Timeline summaryLockPollTimeline;

    /** {@link #refreshSummaryWorkbookOpenLockState()} の前回ポーリング結果（変化時のみ段階2等 UI を同期）。 */
    private Boolean lastPolledSummaryExportLocked;

    /** 段階1／段階2 の Python 実行中（メインシェルから同期）。 */
    private boolean stage1RunPipelineBusy;

    /** 納期管理ビュー再読み込み中（メインシェルから同期）。 */
    private boolean deliveryCalendarReloadBlocking;

    private Tooltip summaryOpenButtonTooltip;

    @FXML
    private ImageView factoryLogoImageView;

    @FXML
    private Label factoryLogoCaptionLabel;

    @FXML
    private Label operatorUserLabel;

    @FXML
    private Button changeOperatorPinButton;

    @FXML
    private Label pipelineTimingStage1Label;

    @FXML
    private Label pipelineTimingStage2Label;

    @FXML
    private Label pipelineTimingStage3Label;

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
                                setGraphic(buildLogLineGraphic(item));
                                setWrapText(false);
                                setTextOverrun(OverrunStyle.CLIP);
                                double w = logListView.getWidth() - 28;
                                if (w > 0) {
                                    setMaxWidth(w);
                                }
                                getStyleClass().add("log-cell");
                                if (shell != null && shell.currentDesktopTheme().isDarkUi()) {
                                    getStyleClass().add("log-dark");
                                }
                                switch (LogLineKind.classify(item)) {
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
    }

    /**
     * 可変行高（{@code setFixedCellSize(-1)}）と折り返しの組み合わせは VirtualFlow が極端なセル数を見積もり、
     * {@code index exceeds maxCellCount} やヒープ枯渇を招くことがある。フォントに応じた正の固定高で抑える。
     * ログは折り返さず右端でクリップするため、セル高は1行分に近い値とする。
     */
    private void applyLogListFixedCellHeight() {
        if (logListView == null) {
            return;
        }
        double lineHeight = appliedLogFont.getSize() * 1.35;
        double cell = Math.clamp(lineHeight * 1.45, 22.0, 52.0);
        logListView.setFixedCellSize(cell);
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
        if (logListView == null) {
            return true;
        }
        double proportion = readVerticalScrollProportion(logListView);
        return !Double.isFinite(proportion)
                || proportion >= LOG_AUTO_SCROLL_BOTTOM_THRESHOLD;
    }

    private void scrollLogListToLastVisibleRowIfNeeded() {
        if (logListView == null || !shouldAutoScrollLogToEnd()) {
            return;
        }
        int n = logLinesVisible.size();
        if (n > 0) {
            logListView.scrollTo(n - 1);
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

    private TextFlow buildLogLineGraphic(String item) {
        TextFlow flow = new TextFlow();
        LogLineKind kind = LogLineKind.classify(item);
        Color baseFill = resolveLogTextFill(kind);
        String search = snapshotLogSearchText();
        if (search.isEmpty()) {
            Text text = new Text(item);
            text.setFont(appliedLogFont);
            text.setFill(baseFill);
            flow.getChildren().add(text);
            return flow;
        }
        String lowerItem = item.toLowerCase(Locale.ROOT);
        String searchLower = search.toLowerCase(Locale.ROOT);
        int from = 0;
        while (from < item.length()) {
            int idx = lowerItem.indexOf(searchLower, from);
            if (idx < 0) {
                Text tail = new Text(item.substring(from));
                tail.setFont(appliedLogFont);
                tail.setFill(baseFill);
                flow.getChildren().add(tail);
                break;
            }
            if (idx > from) {
                Text prefix = new Text(item.substring(from, idx));
                prefix.setFont(appliedLogFont);
                prefix.setFill(baseFill);
                flow.getChildren().add(prefix);
            }
            Text hit = new Text(item.substring(idx, idx + search.length()));
            hit.setFont(appliedLogFont);
            hit.setFill(baseFill);
            hit.getStyleClass().add("pm-log-search-hit");
            flow.getChildren().add(hit);
            from = idx + search.length();
        }
        return flow;
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
        refreshFactorySiteLogo();
        refreshOperatorUserLabel();
        startSummaryExportLockPolling();
        refreshSummaryWorkbookOpenLockState();
        pipelineTimingHistoryListener = () -> Platform.runLater(this::refreshPipelineExecutionTimingLabels);
        if (shell != null) {
            shell.pipelineExecutionTimingHistory().addChangeListener(pipelineTimingHistoryListener);
            refreshPipelineExecutionTimingLabels();
        }
        refreshLogThemeCells();
    }

    private void startSummaryExportLockPolling() {
        stopSummaryExportLockPolling();
        summaryLockPollTimeline =
                new Timeline(
                        new KeyFrame(
                                Duration.seconds(1.5),
                                e -> refreshSummaryWorkbookOpenLockState()));
        summaryLockPollTimeline.setCycleCount(Timeline.INDEFINITE);
        summaryLockPollTimeline.play();
    }

    private void stopSummaryExportLockPolling() {
        if (summaryLockPollTimeline != null) {
            summaryLockPollTimeline.stop();
            summaryLockPollTimeline = null;
        }
    }

    /**
     * 共有ロックファイルの存在（{@link MainShellController#isSummaryAiDispatchExportLocked}）に応じて
     * サマリ「エクセルを開く」を有効／無効にする。他 PC が作成中でもロックが見える。
     */
    void refreshSummaryWorkbookOpenLockState() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path workbook = AppPaths.summaryAiDispatchXlsxPath(ui);
        boolean locked = shell.isSummaryAiDispatchExportLocked();
        if (lastPolledSummaryExportLocked == null || lastPolledSummaryExportLocked != locked) {
            lastPolledSummaryExportLocked = locked;
            shell.refreshSummaryWorkbookLockUi();
        }
        if (openSummaryAiDispatchButton != null) {
            openSummaryAiDispatchButton.setDisable(locked);
        }
        if (forceUnlockSummaryAiDispatchButton != null) {
            forceUnlockSummaryAiDispatchButton.setDisable(!locked);
        }
        if (summaryWorkbookOpenHintLabel != null) {
            String fileName = workbook.getFileName().toString();
            if (locked) {
                String host =
                        SummaryAiDispatchExportLock.readLockInfo(workbook)
                                .map(SummaryAiDispatchExportLock.LockInfo::displayCreator)
                                .orElse("他端末");
                summaryWorkbookOpenHintLabel.setText(fileName + " （作成中: " + host + "）");
                if (openSummaryAiDispatchButton != null) {
                    if (summaryOpenButtonTooltip != null) {
                        Tooltip.uninstall(openSummaryAiDispatchButton, summaryOpenButtonTooltip);
                    }
                    summaryOpenButtonTooltip =
                            new Tooltip(
                                    "サマリ xlsx を作成中です（"
                                            + host
                                            + "）。完了後に開けます。残ロックは「ロック解除」で削除できます。");
                    Tooltip.install(openSummaryAiDispatchButton, summaryOpenButtonTooltip);
                }
            } else {
                summaryWorkbookOpenHintLabel.setText(fileName);
                if (openSummaryAiDispatchButton != null && summaryOpenButtonTooltip != null) {
                    Tooltip.uninstall(openSummaryAiDispatchButton, summaryOpenButtonTooltip);
                    summaryOpenButtonTooltip = null;
                }
            }
        }
    }

    /** {@link GlobalInitSettingTarget} の工場に合わせて実行・ログタブ上部のロゴを更新する。 */
    void refreshFactorySiteLogo() {
        if (factoryLogoHost == null || factoryLogoCaptionLabel == null) {
            return;
        }
        FactorySite site = GlobalInitSettingTarget.load();
        factoryLogoCaptionLabel.setText(site.displayLabelJa());
        factoryLogoHost.getStyleClass().removeIf(c -> c.startsWith("pm-factory-logo-"));
        factoryLogoHost.getStyleClass().add("pm-factory-logo-" + site.name().toLowerCase());
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        if (factoryLogoImageView != null) {
            FactorySiteLogoSupport.applyBrandingOverrideToImageView(factoryLogoImageView, site, ui);
        }
        if (factoryLogoCaptionLabel != null) {
            boolean branding = factoryLogoImageView != null && factoryLogoImageView.isVisible();
            factoryLogoCaptionLabel.setVisible(!branding);
            factoryLogoCaptionLabel.setManaged(!branding);
        }
        Tooltip.install(
                factoryLogoHost,
                new Tooltip(site.displayLabelJa() + "（init_setting 対象工場）"));
    }

    /** 起動時選択した操作者名をヘッダーに表示する。 */
    void refreshOperatorUserLabel() {
        if (operatorUserLabel == null) {
            return;
        }
        String op = FactoryOperatorUserStore.sessionOperatorName();
        operatorUserLabel.setText(op.isBlank() ? "操作者: （未選択）" : "操作者: " + op);
        if (changeOperatorPinButton != null) {
            changeOperatorPinButton.setDisable(op.isBlank() || FactoryOperatorUserStore.isGuestSession());
        }
    }

    @FXML
    private void onChangeOperatorPinAction() {
        if (shell == null) {
            return;
        }
        shell.promptChangeSessionOperatorPin();
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
        if (summaryWorkbookOpenHintLabel != null) {
            summaryWorkbookOpenHintLabel.setText(
                    AppPaths.summaryAiDispatchXlsxPath(ui).getFileName().toString());
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
    private void onOpenSummaryAiDispatchAction() {
        if (shell == null) {
            return;
        }
        Path p = AppPaths.summaryAiDispatchXlsxPath(shell.snapshotUiEnv());
        if (shell.isSummaryAiDispatchExportLocked()) {
            String host =
                    SummaryAiDispatchExportLock.readLockInfo(p)
                            .map(SummaryAiDispatchExportLock.LockInfo::displayCreator)
                            .orElse("他端末");
            appendLog("[summary-ai-dispatch] 作成中のため開けません（" + host + "）");
            return;
        }
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[summary-ai-dispatch] file not found: "
                            + p
                            + " (set "
                            + AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK
                            + " to open another book, or "
                            + AppPaths.KEY_PM_AI_REPO_ROOT
                            + " if the repository root is wrong)");
            return;
        }
        try {
            DesktopFileOpener.openFileReadOnly(p);
            appendLog(
                    "[summary-ai-dispatch] opened (read-only): "
                            + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[summary-ai-dispatch] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onOpenSummaryGenerationHistoryAction() {
        if (shell == null) {
            return;
        }
        shell.selectMainShellTab(MainShellTabId.SUMMARY_AI_DISPATCH_GENERATION);
        appendLog("[summary-generation] サマリ Excel 世代タブを開きました");
    }

    @FXML
    private void onForceUnlockSummaryExportLockAction() {
        if (shell == null) {
            return;
        }
        Path workbook = AppPaths.summaryAiDispatchXlsxPath(shell.snapshotUiEnv());
        if (!shell.isSummaryAiDispatchExportLocked()) {
            refreshSummaryWorkbookOpenLockState();
            return;
        }
        String host =
                SummaryAiDispatchExportLock.readLockInfo(workbook)
                        .map(SummaryAiDispatchExportLock.LockInfo::displayCreator)
                        .orElse("他端末");
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle("サマリ作成ロックの強制解除");
        alert.setHeaderText("ロックファイルを削除します");
        alert.setContentText(
                "出力中の端末（"
                        + host
                        + "）があると、ブック破損や同時書き込みの恐れがあります。\n"
                        + "クラッシュ等で残ったロックの削除にも使えます。\n\n続行しますか？");
        if (alert.showAndWait().orElse(ButtonType.CANCEL) != ButtonType.OK) {
            return;
        }
        boolean removed = SummaryAiDispatchExportLock.forceRemove(workbook);
        appendLog(
                removed
                        ? "[summary-ai-dispatch] ロックを強制削除しました: "
                                + SummaryAiDispatchExportLock.lockFilePath(workbook)
                        : "[summary-ai-dispatch] ロック削除に失敗しました（権限・ネットワークを確認）");
        refreshSummaryWorkbookOpenLockState();
        shell.refreshSummaryWorkbookLockUi();
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
                    case STAGE2, STAGE2_0 -> pipelineTimingStage2Label;
                    case STAGE2_5, STAGE3, STAGE2_1, STAGE3_0, STAGE3_1, STAGE3_2 ->
                            pipelineTimingStage3Label;
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

    private void applyStage1RunButtonEnabledState() {
        if (stage1RunButton == null) {
            return;
        }
        stage1RunButton.setDisable(stage1RunPipelineBusy || deliveryCalendarReloadBlocking);
        if (deliveryCalendarReloadBlocking && !stage1RunPipelineBusy) {
            stage1RunButton.setText(STAGE1_RUN_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD);
        } else {
            stage1RunButton.setText(STAGE1_RUN_BUTTON_TEXT_DEFAULT);
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

    /**
     * When {@code true}, stage-1 deletes {@code ai_remarks_cache.json} (and legacy {@code output/} copy) before
     * spawning Python.
     */
    boolean snapshotStage1ClearCacheAndRun() {
        return stage1ClearCacheAndRunCheckBox != null
                && stage1ClearCacheAndRunCheckBox.isSelected();
    }

    /** 段階1実行後は毎回オフ（ワンショット）。セッションには保存しない。 */
    void resetStage1ClearCacheAndRunCheckbox() {
        if (stage1ClearCacheAndRunCheckBox != null) {
            stage1ClearCacheAndRunCheckBox.setSelected(false);
        }
    }

    /** 開発用チェック（Gemini スキップ・全配台不要）は段階1実行後に毎回オフ。セッションには保存しない。 */
    void resetDevCheckboxesAfterStage1Run() {
        suppressSkipGeminiApiEvents.set(true);
        suppressStage1MarkAllExcludeAfterRunEvents.set(true);
        try {
            if (skipGeminiApiCheckBox != null) {
                skipGeminiApiCheckBox.setSelected(false);
            }
            if (stage1MarkAllExcludeAfterRunCheckBox != null) {
                stage1MarkAllExcludeAfterRunCheckBox.setSelected(false);
            }
        } finally {
            suppressSkipGeminiApiEvents.set(false);
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
                            logLinesAll.setAll(lines);
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
            return 0d;
        }
        return (v - min) / (max - min);
    }

    private void applyLogScrollProportion(double proportion) {
        if (logListView == null || !Double.isFinite(proportion)) {
            return;
        }
        double p = Math.max(0d, Math.min(1d, proportion));
        ScrollBar sb = (ScrollBar) logListView.lookup(".scroll-bar:vertical");
        if (sb == null) {
            return;
        }
        double min = sb.getMin();
        double max = sb.getMax();
        if (max > min) {
            sb.setValue(min + p * (max - min));
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
