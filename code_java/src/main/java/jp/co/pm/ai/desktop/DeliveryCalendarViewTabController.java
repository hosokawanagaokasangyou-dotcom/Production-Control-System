package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Supplier;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.Accordion;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.control.TabPane;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.DeliveryCalendarMainCell;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnDragReorderSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * Displays JSON from {@code pm_ai_delivery_calendar_view.py} (delivery calendar main grid).
 * Source file uses ASCII and Unicode escapes only (safe on CP932 mounts).
 */
public final class DeliveryCalendarViewTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    /**
     * Child stdout lines {@code PM_AI_PROGRESS 0..100}. During Python subprocess, drives {@link #statusLabel} and
     * {@link #deliveryReloadProgressPayloadPrep} only. Tab reload and {@link #deliveryReloadProgressMainCalendar} update
     * in {@link #applyPayloadBody}（アラジン→実績→配台→メイン表の順）。
     */
    private static final String PM_AI_PROGRESS_PREFIX = "PM_AI_PROGRESS ";

    /**
     * Cursor NDJSON filename {@code debug-&lt;id&gt;.log}; align with the active chat debug session id.
     * Writes via {@link AgentDebugLog} so Windows JVM logs mirror to WSL workspace ({@code .cursor/rules/agent-debug-wsl-windows-mirror.mdc}).
     */
    private static final String DEBUG_SESSION_ID_OVERLAY = "ebddd7";

    /**
     * Child stderr is merged into stdout ({@link PythonProcessRunner#runCaptureAsync}); probe scripts may
     * print logging lines before the JSON line. Scan object-shaped lines from bottom to top (same as
     * {@code ActualsStatusTabController.parseActualsPayloadRoot}).
     */
    private static JsonNode parseDeliveryCalendarPayloadRoot(String trimmed)
            throws JsonProcessingException {
        String[] lines = trimmed.split("\\R", -1);
        JsonProcessingException lastLineFailure = null;
        for (int i = lines.length - 1; i >= 0; i--) {
            String ln = lines[i].trim();
            if (ln.isEmpty() || !ln.startsWith("{")) {
                continue;
            }
            try {
                return JSON.readTree(ln);
            } catch (JsonProcessingException e) {
                lastLineFailure = e;
            }
        }
        try {
            return JSON.readTree(trimmed);
        } catch (JsonProcessingException e) {
            if (lastLineFailure != null) {
                throw lastLineFailure;
            }
            throw e;
        }
    }

    @FXML
    private Button refreshButton;

    @FXML
    private Label statusLabel;

    @FXML
    private VBox deliveryReloadProgressContainer;

    @FXML
    private ProgressBar deliveryReloadProgressPayloadPrep;

    @FXML
    private Label deliveryReloadPctPayloadPrep;

    @FXML
    private ProgressBar deliveryReloadProgressMainCalendar;

    @FXML
    private Label deliveryReloadPctMainCalendar;

    @FXML
    private ProgressBar deliveryReloadProgressDispatch;

    @FXML
    private Label deliveryReloadPctDispatch;

    @FXML
    private ProgressBar deliveryReloadProgressActuals;

    @FXML
    private Label deliveryReloadPctActuals;

    @FXML
    private ProgressBar deliveryReloadProgressAladdin;

    @FXML
    private Label deliveryReloadPctAladdin;

    @FXML
    private Label metaLabel;

    @FXML
    private ScrollPane metaScrollPane;

    @FXML
    private TabPane innerTabPane;

    @FXML
    private AladdinProcessingPlanDataTabController aladdinProcessingPlanDataTabController;

    @FXML
    private ProcessingActualsDataTabController processingActualsDataTabController;

    @FXML
    private ResultDispatchTableTabController deliveryCalendarResultDispatchTableTabController;

    @FXML
    private StackPane mainSpreadsheetHost;

    @FXML
    private Slider mainRowHeightSlider;

    @FXML
    private Label mainRowHeightPctLabel;

    @FXML
    private CheckBox mainCellWrapCheck;

    @FXML
    private HBox mainColumnStripHost;

    private MainShellController shell;

    private Stage ownerStage;

    private final SpreadsheetView mainSpreadsheet = new SpreadsheetView();

    private Supplier<RunRequest> requestFactory;

    private final ArrayList<String> mainHeadersRef = new ArrayList<>();

    private final ObservableList<ObservableList<DeliveryCalendarMainCell>> mainRows =
            FXCollections.observableArrayList();

    /**
     * Parallel metadata for {@link #mainRows}: each element is {@code {machineKey, requestId}} for a
     * data row, or {@code null} for section/header rows. Populated during {@link #loadMainCalendar} from
     * the Python JSON before column permutation; row order is never changed by column permutation so
     * indices stay aligned.
     */
    private final ArrayList<String[]> mainRowMeta = new ArrayList<>();

    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayoutMain =
            new AtomicReference<>(List.of());

    private final AtomicInteger headerColumnCountMain = new AtomicInteger(0);

    private final AtomicBoolean suppressMainPersistence = new AtomicBoolean(false);

    private final AtomicReference<TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs>
            mainPresentationPrefs =
                    new AtomicReference<>(TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs.defaults());

    private final AtomicBoolean suppressPresentationUiMain = new AtomicBoolean(false);

    private volatile boolean presentationControlsInstalled;

    private final AtomicBoolean suppressInnerTabSessionPersistence = new AtomicBoolean(false);

    private volatile boolean innerTabPersistenceWired;

    @FXML
    private void initialize() {
        StackPane.setAlignment(mainSpreadsheet, Pos.CENTER_LEFT);
        mainSpreadsheetHost.getChildren().setAll(mainSpreadsheet);
        VBox.setVgrow(mainSpreadsheetHost, Priority.ALWAYS);

        if (metaScrollPane != null && metaLabel != null) {
            metaScrollPane.setFitToWidth(true);
            metaLabel.setWrapText(true);
            metaLabel.prefWidthProperty().bind(metaScrollPane.widthProperty().subtract(18));
        }

        SpreadsheetThemeBridge.install(mainSpreadsheet);
        mainSpreadsheet.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);

        mainColumnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::resetMainColumnWidths,
                                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN,
                                headerColumnCountMain,
                                this::onMainLeadingColumnCommitted,
                                this::onReorderMainColumns,
                                () ->
                                        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                                                ownerStage,
                                                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN,
                                                mainSpreadsheet,
                                                () -> new ArrayList<>(mainHeadersRef))));

        mainSpreadsheet.setGrid(new GridBase(0, 0));

        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                mainSpreadsheetHost, headerColumnCountMain::get);
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.requestFactory = shell::buildDeliveryCalendarRequest;
        this.ownerStage = shell.getPrimaryStage();

        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                mainSpreadsheet,
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN,
                suppressMainPersistence::get,
                () -> new ArrayList<>(mainHeadersRef));

        initDeliveryCalendarPresentationControlsOnce();

        if (processingActualsDataTabController != null) {
            processingActualsDataTabController.bindShell(shell);
        }
        if (aladdinProcessingPlanDataTabController != null) {
            aladdinProcessingPlanDataTabController.bindShell(shell);
        }
        if (deliveryCalendarResultDispatchTableTabController != null) {
            deliveryCalendarResultDispatchTableTabController.bindShell(shell);
            deliveryCalendarResultDispatchTableTabController.setResultDispatchRefreshButtonVisible(false);
        }
        ensureInnerTabPersistenceWired();
    }

    private void ensureInnerTabPersistenceWired() {
        if (innerTabPersistenceWired || innerTabPane == null || shell == null) {
            return;
        }
        innerTabPersistenceWired = true;
        innerTabPane
                .getSelectionModel()
                .selectedIndexProperty()
                .addListener(
                        (obs, a, b) -> {
                            if (suppressInnerTabSessionPersistence.get()) {
                                return;
                            }
                            shell.persistDesktopSessionNow();
                        });
    }

    /** @return ???????????????????? -1 */
    public int snapshotInnerTabSelectedIndex() {
        if (innerTabPane == null) {
            return -1;
        }
        return innerTabPane.getSelectionModel().getSelectedIndex();
    }

    public void applyInnerTabSelectedIndex(int index) {
        if (innerTabPane == null || index < 0) {
            return;
        }
        int n = innerTabPane.getTabs().size();
        if (index >= n) {
            return;
        }
        suppressInnerTabSessionPersistence.set(true);
        try {
            innerTabPane.getSelectionModel().select(index);
        } finally {
            suppressInnerTabSessionPersistence.set(false);
        }
    }

    /**
     * メインシェルで納期管理ビュータブへ切り替えたときに呼ぶ。内側 {@link #innerTabPane} 配下の
     * {@link Accordion} / {@link TitledPane} のうち「操作・ソース」等を閉じる。
     * 各子タブの「データ表」{@link TitledPane} は開いたままにする。
     */
    public void collapseInnerSectionPanesOnShellSelect() {
        Platform.runLater(
                () -> {
                    if (innerTabPane != null) {
                        collapseTitledPanesAndAccordionsUnder(innerTabPane);
                    }
                });
    }

    /** FXML の「データ表」{@link TitledPane} の見出し文言と一致させる。 */
    private static final String INNER_TAB_DATA_TABLE_TITLED_PANE_TEXT = "データ表";

    private static void collapseTitledPanesAndAccordionsUnder(Node node) {
        if (node == null) {
            return;
        }
        if (node instanceof Accordion accordion) {
            accordion.setExpandedPane(null);
        }
        if (node instanceof TitledPane titledPane) {
            if (INNER_TAB_DATA_TABLE_TITLED_PANE_TEXT.equals(titledPane.getText())) {
                titledPane.setExpanded(true);
            } else {
                titledPane.setExpanded(false);
            }
        }
        if (node instanceof Parent parent) {
            for (Node child : parent.getChildrenUnmodifiable()) {
                collapseTitledPanesAndAccordionsUnder(child);
            }
        }
    }

    private void initDeliveryCalendarPresentationControlsOnce() {
        if (presentationControlsInstalled) {
            return;
        }
        if (mainRowHeightSlider == null) {
            return;
        }
        presentationControlsInstalled = true;

        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs lm =
                TableColumnOrderPersistence.loadSpreadsheetTabPresentationPrefs(
                        TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN);
        mainPresentationPrefs.set(lm);

        suppressPresentationUiMain.set(true);
        try {
            double lo = SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN;
            double hi = SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX;
            double pv = lm.rowHeightPercent();
            if (Double.isNaN(pv) || pv <= 0) {
                pv = 100.0;
            }
            pv = Math.min(hi, Math.max(lo, pv));
            mainRowHeightSlider.setMin(lo);
            mainRowHeightSlider.setMax(hi);
            mainRowHeightSlider.setValue(pv);
            mainRowHeightSlider.setMajorTickUnit(250);
            mainRowHeightSlider.setMinorTickCount(4);
            mainRowHeightSlider.setShowTickMarks(true);
            if (mainRowHeightPctLabel != null) {
                mainRowHeightPctLabel.setText(String.format("%.0f%%", pv));
            }
            if (mainCellWrapCheck != null) {
                mainCellWrapCheck.setSelected(lm.cellWrapText());
            }
        } finally {
            suppressPresentationUiMain.set(false);
        }

        SliderCommittedChangeSupport.install(
                mainRowHeightSlider,
                () -> {
                    if (mainRowHeightPctLabel != null && mainRowHeightSlider != null) {
                        mainRowHeightPctLabel.setText(
                                String.format("%.0f%%", mainRowHeightSlider.getValue()));
                    }
                },
                this::commitMainPresentationFromSlider);
        if (mainCellWrapCheck != null) {
            mainCellWrapCheck
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (suppressPresentationUiMain.get()) {
                                    return;
                                }
                                commitMainPresentationFromUi();
                            });
        }
    }

    private void commitMainPresentationFromSlider() {
        if (suppressPresentationUiMain.get()) {
            return;
        }
        commitMainPresentationFromUi();
    }

    private void commitMainPresentationFromUi() {
        if (mainRowHeightSlider == null) {
            return;
        }
        double lo = SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN;
        double hi = SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX;
        double v =
                Math.min(
                        hi,
                        Math.max(lo, mainRowHeightSlider.getValue()));
        boolean wrap = mainCellWrapCheck != null && mainCellWrapCheck.isSelected();
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs next =
                new TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs(v, wrap);
        mainPresentationPrefs.set(next);
        TableColumnOrderPersistence.saveSpreadsheetTabPresentationPrefs(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN, next);
        if (mainRowHeightPctLabel != null) {
            mainRowHeightPctLabel.setText(String.format("%.0f%%", v));
        }
        rebuildMainSpreadsheet();
    }

    private void resetMainColumnWidths() {
        if (mainSpreadsheet == null) {
            return;
        }
        double w = 112;
        for (var c : mainSpreadsheet.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private void onMainLeadingColumnCommitted(int n) {
        headerColumnCountMain.set(n);
        rebuildMainSpreadsheet();
    }

    private void onReorderMainColumns() {
        if (mainHeadersRef.isEmpty()) {
            if (shell != null) {
                shell.appendLog("[delivery-calendar] columns empty (reload first)");
            }
            return;
        }
        if (ownerStage == null) {
            return;
        }
        SpreadsheetColumnReorderDialog.show(ownerStage, new ArrayList<>(mainHeadersRef))
                .ifPresent(
                        perm -> {
                            List<String> oldHeaders = new ArrayList<>(mainHeadersRef);
                            List<String> titleOrder = perm.stream().map(oldHeaders::get).toList();
                            applyPersistedMainColumnOrderAfterLogicalReorder(titleOrder);
                        });
    }

    private void applyPersistedMainColumnOrderAfterLogicalReorder(List<String> titleOrder) {
        if (mainHeadersRef.isEmpty()) {
            return;
        }
        List<String> oldHeaders = new ArrayList<>(mainHeadersRef);
        boolean[] oldVis =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN, oldHeaders.size());
        List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayoutMain.get();
        TableColumnOrderPersistence.applyLogicalColumnOrderDeliveryCalendar(
                mainHeadersRef, mainRows, titleOrder);
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        oldHeaders, oldVis, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN, newVis);
        List<Double> widths =
                TableColumnOrderPersistence.resolveWidthsForHeaders(mainHeadersRef, lay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
        for (int i = 0; i < mainHeadersRef.size(); i++) {
            newLay.add(
                    new TableColumnOrderPersistence.ColumnSpec(mainHeadersRef.get(i), widths.get(i)));
        }
        persistedLayoutMain.set(newLay);
        TableColumnOrderPersistence.saveLayout(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN, newLay);
        rebuildMainSpreadsheet();
    }

    private void rebuildMainSpreadsheet() {
        if (mainHeadersRef.isEmpty()) {
            mainSpreadsheet.setGrid(new GridBase(0, 0));
            return;
        }
        suppressMainPersistence.set(true);
        try {
            final List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            mainHeadersRef, persistedLayoutMain.get(), 112);
            final double widthDefault = 112;
            GridBase grid =
                    SpreadsheetTabularSupport.buildReadOnlyDeliveryCalendarMainGrid(
                            mainHeadersRef, mainRows, headerColumnCountMain.get());
            TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs pres = mainPresentationPrefs.get();
            SpreadsheetTabularSupport.applySpreadsheetGridRowHeightsAndWrap(
                    grid, pres.cellWrapText(), pres.rowHeightPercent());
            mainSpreadsheet.setGrid(grid);
            mainSpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);

            Platform.runLater(
                    () -> {
                        SpreadsheetTabularSupport.applyColumnWidths(mainSpreadsheet, widths, widthDefault);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                mainSpreadsheet, headerColumnCountMain.get());
                        SpreadsheetTabularSupport.applyColumnFilters(mainSpreadsheet);
                        SpreadsheetTabularSupport.pinSpreadsheetFilterRow(mainSpreadsheet);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(mainSpreadsheet);
                        SpreadsheetTabularSupport.refreshSpreadsheetAfterRowPresentationChange(mainSpreadsheet);
                        SpreadsheetColumnDragReorderSupport.refreshAfterGridReady(
                                mainSpreadsheet,
                                suppressMainPersistence::get,
                                () -> new ArrayList<>(mainHeadersRef),
                                headerColumnCountMain.get(),
                                this::applyPersistedMainColumnOrderAfterLogicalReorder);
                        ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                                mainSpreadsheet,
                                () -> new ArrayList<>(mainHeadersRef),
                                () ->
                                        TableColumnOrderPersistence.loadColumnVisibility(
                                                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN,
                                                mainHeadersRef.size()));
                    });
        } finally {
            suppressMainPersistence.set(false);
        }
    }

    @FXML
    private void onClearMainFilters() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(mainSpreadsheet);
    }

    @FXML
    private void onRefreshButtonAction() {
        if (requestFactory == null) {
            statusLabel.setText("初期化待ち");
            return;
        }
        refreshButton.setDisable(true);
        showDeliveryReloadProgress();
        statusLabel.setText("取得中…");
        RunRequest req = requestFactory.get();
        PythonProcessRunner.runCaptureAsyncWithLineTap(
                        req,
                        line ->
                                Platform.runLater(() -> handleDeliveryCalendarProgressLine(line)))
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            refreshButton.setDisable(false);
                                            if (err != null) {
                                                hideDeliveryReloadProgress();
                                                statusLabel.setText("error: " + err.getMessage());
                                                if (shell != null) {
                                                    shell.appendLog("[delivery-calendar] " + err.getMessage());
                                                }
                                                return;
                                            }
                                            if (cap == null) {
                                                hideDeliveryReloadProgress();
                                                statusLabel.setText("no result");
                                                return;
                                            }
                                            applyPayload(cap.stdout());
                                            statusLabel.setText("exit=" + cap.exitCode());
                                        }));
    }

    private void showDeliveryReloadProgress() {
        if (deliveryReloadProgressContainer == null) {
            return;
        }
        resetDeliveryReloadTabSegments();
        deliveryReloadProgressContainer.setManaged(true);
        deliveryReloadProgressContainer.setVisible(true);
    }

    private void hideDeliveryReloadProgress() {
        if (deliveryReloadProgressContainer == null) {
            return;
        }
        resetDeliveryReloadTabSegments();
        deliveryReloadProgressContainer.setVisible(false);
        deliveryReloadProgressContainer.setManaged(false);
    }

    private void resetDeliveryReloadTabSegments() {
        setDeliveryReloadSegmentProgress(
                deliveryReloadProgressPayloadPrep, deliveryReloadPctPayloadPrep, 0.0);
        setDeliveryReloadSegmentProgress(
                deliveryReloadProgressAladdin, deliveryReloadPctAladdin, 0.0);
        setDeliveryReloadSegmentProgress(
                deliveryReloadProgressActuals, deliveryReloadPctActuals, 0.0);
        setDeliveryReloadSegmentProgress(
                deliveryReloadProgressDispatch, deliveryReloadPctDispatch, 0.0);
        setDeliveryReloadSegmentProgress(
                deliveryReloadProgressMainCalendar, deliveryReloadPctMainCalendar, 0.0);
    }

    private static void setDeliveryReloadSegmentProgress(
            ProgressBar bar, Label pctLabel, double p) {
        if (bar == null) {
            return;
        }
        double x = Math.max(0.0, Math.min(1.0, p));
        bar.setProgress(x);
        if (pctLabel != null) {
            pctLabel.setText(String.format("%.0f%%", x * 100.0));
        }
    }

    private void handleDeliveryCalendarProgressLine(String line) {
        if (line == null) {
            return;
        }
        String t = line.strip();
        if (!t.startsWith(PM_AI_PROGRESS_PREFIX)) {
            return;
        }
        try {
            int pct =
                    Integer.parseInt(t.substring(PM_AI_PROGRESS_PREFIX.length()).trim());
            pct = Math.max(0, Math.min(100, pct));
            statusLabel.setText("取得中… ペイロード準備 " + pct + "%");
            double frac = pct / 100.0;
            setDeliveryReloadSegmentProgress(
                    deliveryReloadProgressPayloadPrep, deliveryReloadPctPayloadPrep, frac);
        } catch (NumberFormatException ignored) {
            // ignore malformed progress lines
        }
    }

    private void applyPayload(String stdout) {
        try {
            applyPayloadBody(stdout);
        } finally {
            hideDeliveryReloadProgress();
        }
    }

    private void applyPayloadBody(String stdout) {
        metaLabel.setText("");
        String trimmed = stdout != null ? stdout.trim() : "";
        if (trimmed.isEmpty()) {
            statusLabel.setText("empty stdout");
            return;
        }
        try {
            JsonNode root = parseDeliveryCalendarPayloadRoot(trimmed);
            if (!root.path("ok").asBoolean(false)) {
                statusLabel.setText(root.path("error").asText("failed"));
            }
            JsonNode meta = root.get("meta");
            if (meta != null && meta.isObject()) {
                String na = "\uff08\u672a\u8a2d\u5b9a\uff09";
                String noneResolved = "\uff08\u89e3\u6c7a\u5148\u306a\u3057\uff09";
                String defaultDirSuffix =
                        " \uff08\u74b0\u5883\u5909\u6570\u7a7a\u6b04\u30fb\u65e2\u5b9a\u30d5\u30a9\u30eb\u30c0\uff09";
                StringBuilder sb = new StringBuilder();
                String taskEff = meta.path("pmAiTaskInputSourceDirEffective").asText("");
                if (taskEff.isEmpty()) {
                    taskEff = meta.path("pmAiTaskInputSourceDir").asText("");
                }
                sb.append("PM_AI_TASK_INPUT_SOURCE_DIR: ")
                        .append(taskEff.isEmpty() ? na : taskEff);
                if (meta.path("pmAiTaskInputSourceDirUsesDefaultDir").asBoolean(false)) {
                    sb.append(defaultDirSuffix);
                }
                sb.append("\n");
                String planPath = meta.path("processingPlanPath").asText("");
                sb.append("PM_AI_PROCESSING_PLAN_PATH: ")
                        .append(planPath.isEmpty() ? na : planPath)
                        .append("\n");
                String actDirEff = meta.path("pmAiActualDetailSourceDirEffective").asText("");
                if (actDirEff.isEmpty()) {
                    actDirEff = meta.path("pmAiActualDetailSourceDir").asText("");
                }
                sb.append("PM_AI_ACTUAL_DETAIL_SOURCE_DIR: ")
                        .append(actDirEff.isEmpty() ? na : actDirEff);
                if (meta.path("pmAiActualDetailSourceDirUsesDefaultDir").asBoolean(false)) {
                    sb.append(defaultDirSuffix);
                }
                sb.append("\n");
                String actWbEnv = meta.path("pmAiActualDetailWorkbook").asText("");
                if (!actWbEnv.isEmpty()) {
                    sb.append("PM_AI_ACTUAL_DETAIL_WORKBOOK: ")
                            .append(actWbEnv)
                            .append("\n");
                }
                String actResolved = meta.path("actualDetailWorkbookPath").asText("");
                sb.append(
                                "\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30\uff08"
                                        + "\u89e3\u6c7a\u6e08\u307f\u8aad\u8fbc\u5143\uff09: ")
                        .append(actResolved.isEmpty() ? noneResolved : actResolved)
                        .append("\n");
                sb.append("\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30\u884c\u6570: ")
                        .append(meta.path("actualDetailRowCount").asInt(0))
                        .append("\n");
                String disp = meta.path("dispatchJsonPath").asText("");
                sb.append("\u7d50\u679c_\u914d\u53f0\u8868.json: ")
                        .append(disp.isEmpty() ? noneResolved : disp);
                JsonNode deliveryProbe = meta.get("deliveryCalendarProbe");
                if (deliveryProbe != null && deliveryProbe.isObject()) {
                    String pretty = deliveryProbe.toPrettyString();
                    if (pretty.length() > 16000) {
                        pretty = pretty.substring(0, 16000) + "\n... (truncated)";
                    }
                    sb.append("\n\n--- deliveryCalendarProbe (f73cbb) ---\n");
                    sb.append(pretty);
                }
                metaLabel.setText(sb.toString());
            }

            if (root.path("ok").asBoolean(false)) {
                setDeliveryReloadSegmentProgress(
                        deliveryReloadProgressPayloadPrep, deliveryReloadPctPayloadPrep, 1.0);

                statusLabel.setText("反映中… アラジン加工計画タブ");
                setDeliveryReloadSegmentProgress(
                        deliveryReloadProgressAladdin, deliveryReloadPctAladdin, 0.0);
                if (aladdinProcessingPlanDataTabController != null) {
                    aladdinProcessingPlanDataTabController.reloadAladdinProcessingPlanFromDisk();
                }
                setDeliveryReloadSegmentProgress(
                        deliveryReloadProgressAladdin, deliveryReloadPctAladdin, 1.0);

                statusLabel.setText("反映中… 加工実績タブ");
                setDeliveryReloadSegmentProgress(
                        deliveryReloadProgressActuals, deliveryReloadPctActuals, 0.0);
                if (processingActualsDataTabController != null) {
                    processingActualsDataTabController.reloadProcessingActualsFromDisk();
                }
                setDeliveryReloadSegmentProgress(
                        deliveryReloadProgressActuals, deliveryReloadPctActuals, 1.0);

                statusLabel.setText("反映中… 配台結果タブ");
                setDeliveryReloadSegmentProgress(
                        deliveryReloadProgressDispatch, deliveryReloadPctDispatch, 0.0);
                if (deliveryCalendarResultDispatchTableTabController != null) {
                    deliveryCalendarResultDispatchTableTabController.reloadResultDispatchTableFromDisk();
                }
                setDeliveryReloadSegmentProgress(
                        deliveryReloadProgressDispatch, deliveryReloadPctDispatch, 1.0);
            }

            statusLabel.setText("反映中… メイン表");
            setDeliveryReloadSegmentProgress(
                    deliveryReloadProgressMainCalendar, deliveryReloadPctMainCalendar, 0.0);
            JsonNode mainCal = root.get("mainCalendar");
            if (mainCal != null && mainCal.isObject()) {
                loadMainCalendar(mainCal);
            } else {
                mainHeadersRef.clear();
                mainRows.clear();
                rebuildMainSpreadsheet();
            }
            setDeliveryReloadSegmentProgress(
                    deliveryReloadProgressMainCalendar, deliveryReloadPctMainCalendar, 1.0);

            // #region agent log
            if (shell != null && root.path("ok").asBoolean(false)) {
                Map<String, Object> dbg = new LinkedHashMap<>();
                dbg.put("ok", true);
                dbg.put(
                        "actualDetailRowCount",
                        meta != null && meta.isObject()
                                ? meta.path("actualDetailRowCount").asInt(0)
                                : -1);
                dbg.put("pythonProbeNdjson", ".cursor/debug-f73cbb.log");
                AgentDebugLog.appendStructured(
                        shell.snapshotUiEnv(),
                        "f73cbb",
                        "H_JAVA_TRIGGER",
                        "DeliveryCalendarViewTabController.applyPayload",
                        "delivery_calendar_ok",
                        dbg);
            }
            // #endregion

        } catch (Exception e) {
            statusLabel.setText("parse: " + e.getMessage());
            if (shell != null) {
                shell.appendLog("[delivery-calendar] parse " + e.getMessage());
            }
        }
    }

    private void loadMainCalendar(JsonNode mainCal) {
        mainHeadersRef.clear();
        JsonNode cols = mainCal.get("columns");
        if (cols != null && cols.isArray()) {
            for (JsonNode c : cols) {
                mainHeadersRef.add(c.asText(""));
            }
        }
        mainRows.clear();
        mainRowMeta.clear();
        JsonNode rows = mainCal.get("rows");
        if (rows != null && rows.isArray()) {
            for (JsonNode row : rows) {
                String kind = row.path("kind").asText("data");
                if ("section".equals(kind)) {
                    mainRowMeta.add(null);
                } else {
                    mainRowMeta.add(new String[]{
                        row.path("machineKey").asText(""),
                        row.path("requestId").asText("")
                    });
                }
                ObservableList<DeliveryCalendarMainCell> line = FXCollections.observableArrayList();
                JsonNode cells = row.get("cells");
                if (cells != null && cells.isArray()) {
                    for (JsonNode cell : cells) {
                        line.add(parseDeliveryCalendarMainCell(cell));
                    }
                }
                mainRows.add(line);
            }
        }

        List<TableColumnOrderPersistence.ColumnSpec> lay =
                TableColumnOrderPersistence.loadLayout(
                        TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN);
        persistedLayoutMain.set(lay);
        List<String> beforeHeaders = new ArrayList<>(mainHeadersRef);
        boolean[] visBefore =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN, beforeHeaders.size());
        List<String> titleOrder =
                lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
        TableColumnOrderPersistence.applyLogicalColumnOrderDeliveryCalendar(
                mainHeadersRef, mainRows, titleOrder);
        boolean[] visAfter =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        beforeHeaders, visBefore, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN, visAfter);

        overlayChildTabValues();
        rebuildMainSpreadsheet();
    }

    // -------------------------------------------------------------------------
    // Child-tab data overlay (replaces Python-computed p/a/d with Java-side data)
    // -------------------------------------------------------------------------

    /** Pattern for calendar column-header dates like {@code 2026?4?21?(?)}. */
    private static final java.util.regex.Pattern CAL_DATE_HDR =
            java.util.regex.Pattern.compile("(\\d{4})\u5e74(\\d{1,2})\u6708(\\d{1,2})\u65e5\\([\u6708\u706b\u6c34\u6728\u91d1\u571f\u65e5]\\)");

    /** Pattern for Aladdin date column headers: {@code yyyy/MM/dd}. */
    private static final java.util.regex.Pattern ALADDIN_DATE_COL =
            java.util.regex.Pattern.compile("\\d{4}/\\d{2}/\\d{2}");

    private static final String COL_MK_NAME = "\u6a5f\u68b0\u540d"; // ???
    private static final String COL_TID     = "\u4f9d\u983cNO";     // ??NO
    private static final String COL_KAKOU_DATE = "\u52a0\u5de5\u65e5"; // ???
    private static final String COL_ACT_QTY = "\u5b9f\u52a0\u5de5\u6570"; // ????
    private static final String COL_MFG_COND = "\u88fd\u9020\u6761\u4ef6(\u5185\u8a33)"; // ????(??)
    private static final String MFG_COND_LENGTH = "\u9577\u3055"; // ??
    private static final String COL_DIS_DATE = "\u914d\u53f0\u65e5"; // ???
    private static final String COL_DIS_QTY  = "\u5f53\u65e5\u914d\u53f0\u6570\u91cf"; // ??????

    /**
     * Replaces triple-cell p/a/d values in {@link #mainRows} using shaped JSON cache files written by
     * child tabs on each load (falling back to in-memory child-tab data when the files are absent).
     * This is called after {@link #loadMainCalendar} so column permutation is already applied to
     * {@link #mainHeadersRef}; row order is unchanged so {@link #mainRowMeta} indices stay aligned.
     */
    private void overlayChildTabValues() {
        if (mainRowMeta.isEmpty() || mainHeadersRef.isEmpty()) {
            return;
        }

        // Map column position \u2192 normalised date string "yyyy/MM/dd" for every date column
        Map<Integer, String> calDateByIdx = new LinkedHashMap<>();
        for (int i = 0; i < mainHeadersRef.size(); i++) {
            String ds = parseDateHeader(mainHeadersRef.get(i));
            if (ds != null) {
                calDateByIdx.put(i, ds);
            }
        }
        if (calDateByIdx.isEmpty()) {
            return;
        }

        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();

        // --- Aladdin plan: prefer shaped JSON cache, fall back to in-memory ---
        List<String> planHeaders;
        List<List<String>> planRows;
        Path aladdinJsonPath = AppPaths.resolveShapedAladdinPlanJsonPath(ui);
        if (Files.isRegularFile(aladdinJsonPath)) {
            try {
                JsonTableIo.ArrayTable t = JsonTableIo.loadArrayTable(aladdinJsonPath);
                planHeaders = t.columns();
                planRows = t.rows();
            } catch (Exception ex) {
                if (shell != null) {
                    shell.appendLog(
                            "[delivery-calendar] aladdin shaped JSON load failed: " + ex.getMessage());
                }
                planHeaders = aladdinProcessingPlanDataTabController.getShapedHeaders();
                planRows = aladdinProcessingPlanDataTabController.getShapedRows();
            }
        } else {
            planHeaders = aladdinProcessingPlanDataTabController.getShapedHeaders();
            planRows = aladdinProcessingPlanDataTabController.getShapedRows();
        }

        // --- Processing actuals: prefer shaped JSON cache, fall back to in-memory ---
        List<String> actHeaders;
        List<List<String>> actRows;
        Path actualsJsonPath = AppPaths.resolveShapedProcessingActualsJsonPath(ui);
        if (Files.isRegularFile(actualsJsonPath)) {
            try {
                JsonTableIo.ArrayTable t = JsonTableIo.loadArrayTable(actualsJsonPath);
                actHeaders = t.columns();
                actRows = t.rows();
            } catch (Exception ex) {
                if (shell != null) {
                    shell.appendLog(
                            "[delivery-calendar] actuals shaped JSON load failed: " + ex.getMessage());
                }
                actHeaders = processingActualsDataTabController.getUnfilteredShapedHeaders();
                actRows = processingActualsDataTabController.getUnfilteredShapedRows();
            }
        } else {
            actHeaders = processingActualsDataTabController.getUnfilteredShapedHeaders();
            actRows = processingActualsDataTabController.getUnfilteredShapedRows();
        }

        // --- Dispatch: load from \u7d50\u679c_\u914d\u53f0\u8868.json when present (restart-safe without opening tab) ---
        Path dispatchJsonPath = AppPaths.resolveResultDispatchTableJsonPath(ui);
        List<String> disHeaders = deliveryCalendarResultDispatchTableTabController.getShapedHeaders();
        List<List<String>> disRows = deliveryCalendarResultDispatchTableTabController.getShapedRows();
        String dispatchSource = disRows.isEmpty() ? "none" : "memory";
        if (Files.isRegularFile(dispatchJsonPath)) {
            try {
                JsonTableIo.SheetTable st = JsonTableIo.loadFlatTable(dispatchJsonPath);
                if (!st.columns().isEmpty() && !st.rows().isEmpty()) {
                    disHeaders = new ArrayList<>(st.columns());
                    disRows = sheetTableToRowLists(st);
                    dispatchSource = "file";
                }
            } catch (Exception ex) {
                if (shell != null) {
                    shell.appendLog(
                            "[delivery-calendar] dispatch flat JSON load failed: " + ex.getMessage());
                }
            }
        }

        Map<String, Map<String, Map<String, Double>>> planLookup =
                buildAladdinPlanLookup(planHeaders, planRows);
        Map<String, Map<String, Map<String, Double>>> actualLookup =
                buildActualLookup(actHeaders, actRows);
        Map<String, Map<String, Map<String, Double>>> dispatchLookup =
                buildDispatchLookup(disHeaders, disRows);

        // region agent log
        Map<String, Object> overlayData = new LinkedHashMap<>();
        overlayData.put("planMachines", planLookup.size());
        overlayData.put("actualMachines", actualLookup.size());
        overlayData.put("dispatchMachines", dispatchLookup.size());
        overlayData.put("calDateCols", calDateByIdx.size());
        overlayData.put("mainRows", mainRows.size());
        overlayData.put("aladdinJsonExists", Files.isRegularFile(aladdinJsonPath));
        overlayData.put("actualsJsonExists", Files.isRegularFile(actualsJsonPath));
        overlayData.put("dispatchJsonExists", Files.isRegularFile(dispatchJsonPath));
        overlayData.put("dispatchSource", dispatchSource);
        overlayData.put("dispatchJsonPath", dispatchJsonPath.toString());
        AgentDebugLog.appendStructured(
                ui,
                DEBUG_SESSION_ID_OVERLAY,
                "OVERLAY",
                "DeliveryCalendarViewTabController.java:overlayChildTabValues",
                "lookup_sizes",
                overlayData);
        // endregion

        // Replace TripleQty cells with values from child tab data
        for (int i = 0; i < mainRows.size(); i++) {
            String[] meta = i < mainRowMeta.size() ? mainRowMeta.get(i) : null;
            if (meta == null) {
                continue; // section row
            }
            String mk = meta[0];
            String tid = meta[1];
            if (mk.isEmpty() && tid.isEmpty()) {
                continue;
            }
            ObservableList<DeliveryCalendarMainCell> row = mainRows.get(i);
            for (Map.Entry<Integer, String> e : calDateByIdx.entrySet()) {
                int j = e.getKey();
                if (j >= row.size()) {
                    continue;
                }
                String dateStr = e.getValue();
                double p = lookupQty(planLookup, mk, tid, dateStr);
                double a = lookupQty(actualLookup, mk, tid, dateStr);
                double d = lookupQty(dispatchLookup, mk, tid, dateStr);
                String sp = Math.abs(p) > 1e-12 ? formatQtyShort(p) : "";
                String sa = Math.abs(a) > 1e-12 ? formatQtyShort(a) : "";
                String sd = Math.abs(d) > 1e-12 ? formatQtyShort(d) : "";
                row.set(j, new DeliveryCalendarMainCell.TripleQty(sp, sa, sd));
            }
        }
    }

    private static double lookupQty(
            Map<String, Map<String, Map<String, Double>>> lookup, String mk, String tid, String dateStr) {
        Map<String, Map<String, Double>> byTid = lookup.get(mk);
        if (byTid == null) {
            return 0.0;
        }
        Map<String, Double> byDate = byTid.get(tid);
        if (byDate == null) {
            return 0.0;
        }
        Double q = byDate.get(dateStr);
        return q != null ? q : 0.0;
    }

    /**
     * Parses a calendar column-header string like {@code 2026?4?21?(?)} to {@code "2026/04/21"};
     * returns {@code null} if the string does not match the pattern.
     */
    private static String parseDateHeader(String header) {
        if (header == null) {
            return null;
        }
        var m = CAL_DATE_HDR.matcher(header);
        if (!m.matches()) {
            return null;
        }
        int y = Integer.parseInt(m.group(1));
        int mo = Integer.parseInt(m.group(2));
        int d = Integer.parseInt(m.group(3));
        return String.format("%04d/%02d/%02d", y, mo, d);
    }

    /**
     * NFKC-normalise a machine name to a match key (mirrors Python {@code _normalize_equipment_match_key}).
     */
    private static String normalizeEquipmentMatchKey(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String t = java.text.Normalizer.normalize(val, java.text.Normalizer.Form.NFKC);
        t = t.replace('\u00a0', ' ').replace('\u3000', ' ');
        t = t.replaceAll("[\u200b\u200c\u200d\ufeff]", "");
        return t.replaceAll("\\s+", " ").strip();
    }

    /** Format a double the same way Python {@code _format_qty_short} does. */
    private static String formatQtyShort(double v) {
        long rounded = Math.round(v);
        if (Math.abs(v - rounded) < 1e-9) {
            return Long.toString(rounded);
        }
        String s = String.format("%.4f", v).replaceAll("0+$", "").replaceAll("\\.$", "");
        return s.isEmpty() ? "0" : s;
    }

    private static double parseCellDouble(String s) {
        if (s == null || s.isBlank()) {
            return 0.0;
        }
        try {
            return Double.parseDouble(s.strip());
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private static int colIdx(List<String> headers, String title) {
        for (int i = 0; i < headers.size(); i++) {
            if (title.equals(headers.get(i))) {
                return i;
            }
        }
        return -1;
    }

    private static String cellAt(List<String> row, int idx) {
        return (idx >= 0 && idx < row.size() && row.get(idx) != null) ? row.get(idx) : "";
    }

    /** Converts {@link JsonTableIo.SheetTable} row maps to aligned row lists for lookup builders. */
    private static List<List<String>> sheetTableToRowLists(JsonTableIo.SheetTable st) {
        List<String> cols = st.columns();
        List<List<String>> out = new ArrayList<>(st.rows().size());
        for (Map<String, String> m : st.rows()) {
            List<String> line = new ArrayList<>(cols.size());
            for (String c : cols) {
                line.add(m != null ? m.getOrDefault(c, "") : "");
            }
            out.add(line);
        }
        return out;
    }

    /**
     * Normalises a date string to {@code yyyy/MM/dd}.  Accepts {@code yyyy/MM/dd} (unchanged),
     * {@code yyyy-MM-dd} or {@code yyyy-MM-ddThh:mm:ss} (ISO, replaces {@code -} and drops time).
     */
    private static String normaliseDateStr(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String s = raw.strip();
        // ISO datetime: 2026-04-21 or 2026-04-21T00:00:00
        if (s.length() >= 10 && s.charAt(4) == '-' && s.charAt(7) == '-') {
            return s.substring(0, 10).replace('-', '/');
        }
        // Already yyyy/MM/dd
        if (s.length() == 10 && s.charAt(4) == '/' && s.charAt(7) == '/') {
            return s;
        }
        // Try flexible y/m/d variants like yyyy/M/d
        try {
            String[] parts = s.split("[/\\-]");
            if (parts.length == 3) {
                int y = Integer.parseInt(parts[0].strip());
                int mo = Integer.parseInt(parts[1].strip());
                int d = Integer.parseInt(parts[2].strip());
                return String.format("%04d/%02d/%02d", y, mo, d);
            }
        } catch (NumberFormatException ignored) { /* fall through */ }
        return null;
    }

    /**
     * Builds plan qty lookup from Aladdin shaped data.
     * Key: {@code normalizedMk -> tid -> "yyyy/MM/dd" -> sumQty}.
     * Date columns are identified by the {@code yyyy/MM/dd} header pattern.
     */
    private static Map<String, Map<String, Map<String, Double>>> buildAladdinPlanLookup(
            List<String> headers, List<List<String>> rows) {
        int mkIdx  = colIdx(headers, COL_MK_NAME);
        int tidIdx = colIdx(headers, COL_TID);
        if (mkIdx < 0 || tidIdx < 0) {
            return Map.of();
        }
        Map<Integer, String> dateCols = new LinkedHashMap<>();
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (h != null && ALADDIN_DATE_COL.matcher(h).matches()) {
                dateCols.put(i, h);
            }
        }
        if (dateCols.isEmpty()) {
            return Map.of();
        }
        Map<String, Map<String, Map<String, Double>>> result = new LinkedHashMap<>();
        for (List<String> row : rows) {
            String mk  = normalizeEquipmentMatchKey(cellAt(row, mkIdx));
            String tid = cellAt(row, tidIdx).strip();
            if (mk.isEmpty() || tid.isEmpty()) {
                continue;
            }
            for (Map.Entry<Integer, String> e : dateCols.entrySet()) {
                double qty = parseCellDouble(cellAt(row, e.getKey()));
                if (Math.abs(qty) > 1e-12) {
                    result.computeIfAbsent(mk, k -> new LinkedHashMap<>())
                          .computeIfAbsent(tid, k -> new LinkedHashMap<>())
                          .merge(e.getValue(), qty, Double::sum);
                }
            }
        }
        return result;
    }

    /**
     * Builds actual qty lookup from <em>unfiltered</em> shaped actuals data,
     * applying the same {@code ????(??) == "??"} filter Python uses.
     * Aggregation: MAX per {@code (mk, tid, date)}.
     * Key: {@code normalizedMk -> tid -> "yyyy/MM/dd" -> maxQty}.
     */
    private static Map<String, Map<String, Map<String, Double>>> buildActualLookup(
            List<String> headers, List<List<String>> rows) {
        int mkIdx   = colIdx(headers, COL_MK_NAME);
        int tidIdx  = colIdx(headers, COL_TID);
        int dateIdx = colIdx(headers, COL_KAKOU_DATE);
        int qtyIdx  = colIdx(headers, COL_ACT_QTY);
        int condIdx = colIdx(headers, COL_MFG_COND);
        if (mkIdx < 0 || tidIdx < 0 || dateIdx < 0 || qtyIdx < 0) {
            return Map.of();
        }
        Map<String, Map<String, Map<String, Double>>> result = new LinkedHashMap<>();
        for (List<String> row : rows) {
            // Filter: ????(??) == "??" (when column is present)
            if (condIdx >= 0) {
                String cond = cellAt(row, condIdx).strip();
                if (!MFG_COND_LENGTH.equals(cond)) {
                    continue;
                }
            }
            String mk  = normalizeEquipmentMatchKey(cellAt(row, mkIdx));
            String tid = cellAt(row, tidIdx).strip();
            String ds  = normaliseDateStr(cellAt(row, dateIdx));
            if (mk.isEmpty() || tid.isEmpty() || ds == null) {
                continue;
            }
            double qty = parseCellDouble(cellAt(row, qtyIdx));
            if (qty <= 1e-12) {
                continue;
            }
            // MAX aggregation (mirrors Python)
            result.computeIfAbsent(mk, k -> new LinkedHashMap<>())
                  .computeIfAbsent(tid, k -> new LinkedHashMap<>())
                  .merge(ds, qty, Math::max);
        }
        return result;
    }

    /**
     * Builds dispatch qty lookup from shaped dispatch table data.
     * Key: {@code normalizedMk -> tid -> "yyyy/MM/dd" -> sumQty}.
     * The {@code ???} column comes from JSON and is in ISO {@code yyyy-MM-dd} format.
     */
    private static Map<String, Map<String, Map<String, Double>>> buildDispatchLookup(
            List<String> headers, List<List<String>> rows) {
        int mkIdx   = colIdx(headers, COL_MK_NAME);
        int tidIdx  = colIdx(headers, COL_TID);
        int dateIdx = colIdx(headers, COL_DIS_DATE);
        int qtyIdx  = colIdx(headers, COL_DIS_QTY);
        if (mkIdx < 0 || tidIdx < 0 || dateIdx < 0 || qtyIdx < 0) {
            return Map.of();
        }
        Map<String, Map<String, Map<String, Double>>> result = new LinkedHashMap<>();
        for (List<String> row : rows) {
            String mk  = normalizeEquipmentMatchKey(cellAt(row, mkIdx));
            String tid = cellAt(row, tidIdx).strip();
            String ds  = normaliseDateStr(cellAt(row, dateIdx));
            if (mk.isEmpty() || tid.isEmpty() || ds == null) {
                continue;
            }
            double qty = parseCellDouble(cellAt(row, qtyIdx));
            if (Math.abs(qty) > 1e-12) {
                result.computeIfAbsent(mk, k -> new LinkedHashMap<>())
                      .computeIfAbsent(tid, k -> new LinkedHashMap<>())
                      .merge(ds, qty, Double::sum);
            }
        }
        return result;
    }

    private static DeliveryCalendarMainCell parseDeliveryCalendarMainCell(JsonNode cell) {
        if (cell == null || cell.isNull() || cell.isMissingNode()) {
            return new DeliveryCalendarMainCell.PlainText("");
        }
        if (cell.isObject() && cell.has("triple")) {
            JsonNode t = cell.get("triple");
            return new DeliveryCalendarMainCell.TripleQty(
                    t.path("p").asText(""),
                    t.path("a").asText(""),
                    t.path("d").asText(""));
        }
        return new DeliveryCalendarMainCell.PlainText(cell.asText(""));
    }
}
