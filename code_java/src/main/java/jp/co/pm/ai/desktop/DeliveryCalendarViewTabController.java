package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
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
import com.fasterxml.jackson.databind.node.ObjectNode;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.control.TabPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.DeliveryCalendarMainCell;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnDragReorderSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * Displays JSON from {@code pm_ai_delivery_calendar_view.py} (delivery calendar + plan comparison).
 * Source file uses ASCII and Unicode escapes only (safe on CP932 mounts).
 */
public final class DeliveryCalendarViewTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    // #region agent log
    private static final String AGENT_DEBUG_LOG_PATH =
            "/mnt/c/????AI??????_JAVA/.cursor/debug-ebddd7.log";

    private static void agentDebugNdjson(
            String hypothesisId, String location, String message, ObjectNode data) {
        try {
            ObjectNode root = JSON.createObjectNode();
            root.put("sessionId", "ebddd7");
            root.put("hypothesisId", hypothesisId);
            root.put("location", location);
            root.put("message", message);
            root.put("timestamp", System.currentTimeMillis());
            root.set("data", data);
            Files.writeString(
                    Path.of(AGENT_DEBUG_LOG_PATH),
                    JSON.writeValueAsString(root) + "\n",
                    StandardOpenOption.CREATE,
                    StandardOpenOption.APPEND);
        } catch (Exception ignored) {
            // debug instrumentation only
        }
    }

    // #endregion

    private static final Map<String, String> COMPARE_HEADER_JP = compareHeaderJp();

    private static Map<String, String> compareHeaderJp() {
        Map<String, String> m = new LinkedHashMap<>();
        m.put("machine_key", "\u6a5f\u68b0\u30ad\u30fc");
        m.put("machine_display", "\u6a5f\u68b0\u540d\u8868\u793a");
        m.put("request_id", "\u4f9d\u983cNO");
        m.put("calendar_date", "\u66c6\u65e5");
        m.put("qty_dispatch_json", "\u6570\u91cf_\u7d50\u679c\u914d\u53f0\u8868");
        m.put(
                "qty_task_input_aladdin",
                "\u6570\u91cf_\u30bf\u30b9\u30af\u5165\u529b\u30a2\u30e9\u30b8\u30f3");
        m.put("delta", "\u5dee\u5206");
        return Map.copyOf(m);
    }

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
    private Label metaLabel;

    @FXML
    private TabPane innerTabPane;

    @FXML
    private StackPane mainSpreadsheetHost;

    @FXML
    private StackPane compareSpreadsheetHost;

    @FXML
    private Slider mainRowHeightSlider;

    @FXML
    private Label mainRowHeightPctLabel;

    @FXML
    private CheckBox mainCellWrapCheck;

    @FXML
    private HBox mainColumnStripHost;

    @FXML
    private Slider compareRowHeightSlider;

    @FXML
    private Label compareRowHeightPctLabel;

    @FXML
    private CheckBox compareCellWrapCheck;

    @FXML
    private HBox compareColumnStripHost;

    private MainShellController shell;

    private Stage ownerStage;

    private final SpreadsheetView mainSpreadsheet = new SpreadsheetView();

    private final SpreadsheetView compareSpreadsheet = new SpreadsheetView();

    private Supplier<RunRequest> requestFactory;

    private final ArrayList<String> mainHeadersRef = new ArrayList<>();

    private final ObservableList<ObservableList<DeliveryCalendarMainCell>> mainRows =
            FXCollections.observableArrayList();

    private final ArrayList<String> compareHeadersRef = new ArrayList<>();

    private final ObservableList<ObservableList<String>> compareRows = FXCollections.observableArrayList();

    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayoutMain =
            new AtomicReference<>(List.of());

    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayoutCompare =
            new AtomicReference<>(List.of());

    private final AtomicInteger headerColumnCountMain = new AtomicInteger(0);

    private final AtomicInteger headerColumnCountCompare = new AtomicInteger(0);

    private final AtomicBoolean suppressMainPersistence = new AtomicBoolean(false);

    private final AtomicBoolean suppressComparePersistence = new AtomicBoolean(false);

    private final AtomicReference<TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs>
            mainPresentationPrefs =
                    new AtomicReference<>(TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs.defaults());

    private final AtomicReference<TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs>
            comparePresentationPrefs =
                    new AtomicReference<>(TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs.defaults());

    private final AtomicBoolean suppressPresentationUiMain = new AtomicBoolean(false);

    private final AtomicBoolean suppressPresentationUiCompare = new AtomicBoolean(false);

    private volatile boolean presentationControlsInstalled;

    @FXML
    private void initialize() {
        StackPane.setAlignment(mainSpreadsheet, Pos.TOP_LEFT);
        mainSpreadsheetHost.getChildren().setAll(mainSpreadsheet);
        VBox.setVgrow(mainSpreadsheetHost, Priority.ALWAYS);
        mainSpreadsheet.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        mainSpreadsheet.prefWidthProperty().bind(mainSpreadsheetHost.widthProperty());
        mainSpreadsheet.prefHeightProperty().bind(mainSpreadsheetHost.heightProperty());

        StackPane.setAlignment(compareSpreadsheet, Pos.TOP_LEFT);
        compareSpreadsheetHost.getChildren().setAll(compareSpreadsheet);
        VBox.setVgrow(compareSpreadsheetHost, Priority.ALWAYS);
        compareSpreadsheet.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        compareSpreadsheet.prefWidthProperty().bind(compareSpreadsheetHost.widthProperty());
        compareSpreadsheet.prefHeightProperty().bind(compareSpreadsheetHost.heightProperty());

        SpreadsheetTabularSupport.installDeliveryCalendarSpreadsheetChrome(mainSpreadsheet);
        SpreadsheetTabularSupport.installDeliveryCalendarSpreadsheetChrome(compareSpreadsheet);
        mainSpreadsheet.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        compareSpreadsheet.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        SpreadsheetTabularSupport.installFullRowDataSelection(mainSpreadsheet);
        SpreadsheetTabularSupport.installFullRowDataSelection(compareSpreadsheet);

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
        compareColumnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::resetCompareColumnWidths,
                                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE,
                                headerColumnCountCompare,
                                this::onCompareLeadingColumnCommitted,
                                this::onReorderCompareColumns,
                                () ->
                                        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                                                ownerStage,
                                                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE,
                                                compareSpreadsheet,
                                                () -> new ArrayList<>(compareHeadersRef))));

        mainSpreadsheet.setGrid(new GridBase(0, 0));
        compareSpreadsheet.setGrid(new GridBase(0, 0));
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
        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                compareSpreadsheet,
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE,
                suppressComparePersistence::get,
                () -> new ArrayList<>(compareHeadersRef));

        initDeliveryCalendarPresentationControlsOnce();
    }

    private void initDeliveryCalendarPresentationControlsOnce() {
        if (presentationControlsInstalled) {
            return;
        }
        if (mainRowHeightSlider == null || compareRowHeightSlider == null) {
            return;
        }
        presentationControlsInstalled = true;

        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs lm =
                TableColumnOrderPersistence.loadSpreadsheetTabPresentationPrefs(
                        TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN);
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs lc =
                TableColumnOrderPersistence.loadSpreadsheetTabPresentationPrefs(
                        TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE);
        mainPresentationPrefs.set(lm);
        comparePresentationPrefs.set(lc);

        suppressPresentationUiMain.set(true);
        suppressPresentationUiCompare.set(true);
        try {
            mainRowHeightSlider.setMin(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN);
            mainRowHeightSlider.setMax(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX);
            mainRowHeightSlider.setValue(lm.rowHeightPercent());
            mainRowHeightSlider.setMajorTickUnit(250);
            mainRowHeightSlider.setMinorTickCount(4);
            mainRowHeightSlider.setShowTickMarks(true);
            if (mainRowHeightPctLabel != null) {
                mainRowHeightPctLabel.setText(String.format("%.0f%%", lm.rowHeightPercent()));
            }
            if (mainCellWrapCheck != null) {
                mainCellWrapCheck.setSelected(lm.cellWrapText());
            }

            compareRowHeightSlider.setMin(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN);
            compareRowHeightSlider.setMax(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX);
            compareRowHeightSlider.setValue(lc.rowHeightPercent());
            compareRowHeightSlider.setMajorTickUnit(250);
            compareRowHeightSlider.setMinorTickCount(4);
            compareRowHeightSlider.setShowTickMarks(true);
            if (compareRowHeightPctLabel != null) {
                compareRowHeightPctLabel.setText(String.format("%.0f%%", lc.rowHeightPercent()));
            }
            if (compareCellWrapCheck != null) {
                compareCellWrapCheck.setSelected(lc.cellWrapText());
            }
        } finally {
            suppressPresentationUiMain.set(false);
            suppressPresentationUiCompare.set(false);
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

        SliderCommittedChangeSupport.install(
                compareRowHeightSlider,
                () -> {
                    if (compareRowHeightPctLabel != null && compareRowHeightSlider != null) {
                        compareRowHeightPctLabel.setText(
                                String.format("%.0f%%", compareRowHeightSlider.getValue()));
                    }
                },
                this::commitComparePresentationFromSlider);
        if (compareCellWrapCheck != null) {
            compareCellWrapCheck
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (suppressPresentationUiCompare.get()) {
                                    return;
                                }
                                commitComparePresentationFromUi();
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
        double v = mainRowHeightSlider.getValue();
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

    private void commitComparePresentationFromSlider() {
        if (suppressPresentationUiCompare.get()) {
            return;
        }
        commitComparePresentationFromUi();
    }

    private void commitComparePresentationFromUi() {
        if (compareRowHeightSlider == null) {
            return;
        }
        double v = compareRowHeightSlider.getValue();
        boolean wrap = compareCellWrapCheck != null && compareCellWrapCheck.isSelected();
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs next =
                new TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs(v, wrap);
        comparePresentationPrefs.set(next);
        TableColumnOrderPersistence.saveSpreadsheetTabPresentationPrefs(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE, next);
        if (compareRowHeightPctLabel != null) {
            compareRowHeightPctLabel.setText(String.format("%.0f%%", v));
        }
        rebuildCompareSpreadsheet();
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

    private void resetCompareColumnWidths() {
        if (compareSpreadsheet == null) {
            return;
        }
        double w = 112;
        for (var c : compareSpreadsheet.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private void onMainLeadingColumnCommitted(int n) {
        headerColumnCountMain.set(n);
        rebuildMainSpreadsheet();
    }

    private void onCompareLeadingColumnCommitted(int n) {
        headerColumnCountCompare.set(n);
        rebuildCompareSpreadsheet();
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

    private void onReorderCompareColumns() {
        if (compareHeadersRef.isEmpty()) {
            if (shell != null) {
                shell.appendLog("[delivery-calendar] columns empty (reload first)");
            }
            return;
        }
        if (ownerStage == null) {
            return;
        }
        SpreadsheetColumnReorderDialog.show(ownerStage, new ArrayList<>(compareHeadersRef))
                .ifPresent(
                        perm -> {
                            List<String> oldHeaders = new ArrayList<>(compareHeadersRef);
                            List<String> titleOrder = perm.stream().map(oldHeaders::get).toList();
                            applyPersistedCompareColumnOrderAfterLogicalReorder(titleOrder);
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

    private void applyPersistedCompareColumnOrderAfterLogicalReorder(List<String> titleOrder) {
        if (compareHeadersRef.isEmpty()) {
            return;
        }
        List<String> oldHeaders = new ArrayList<>(compareHeadersRef);
        boolean[] oldVis =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE, oldHeaders.size());
        List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayoutCompare.get();
        TableColumnOrderPersistence.applyLogicalColumnOrder(compareHeadersRef, compareRows, titleOrder);
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        oldHeaders, oldVis, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE, newVis);
        List<Double> widths =
                TableColumnOrderPersistence.resolveWidthsForHeaders(compareHeadersRef, lay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
        for (int i = 0; i < compareHeadersRef.size(); i++) {
            newLay.add(
                    new TableColumnOrderPersistence.ColumnSpec(compareHeadersRef.get(i), widths.get(i)));
        }
        persistedLayoutCompare.set(newLay);
        TableColumnOrderPersistence.saveLayout(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE, newLay);
        rebuildCompareSpreadsheet();
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
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(mainSpreadsheet);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                mainSpreadsheet, headerColumnCountMain.get());
                        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(mainSpreadsheet);
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

    private void rebuildCompareSpreadsheet() {
        if (compareHeadersRef.isEmpty()) {
            compareSpreadsheet.setGrid(new GridBase(0, 0));
            return;
        }
        suppressComparePersistence.set(true);
        try {
            final List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            compareHeadersRef, persistedLayoutCompare.get(), 112);
            final double widthDefault = 112;
            GridBase grid =
                    SpreadsheetTabularSupport.buildReadOnlyPlainGrid(
                            compareHeadersRef, compareRows, headerColumnCountCompare.get());
            TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs pres =
                    comparePresentationPrefs.get();
            SpreadsheetTabularSupport.applySpreadsheetGridRowHeightsAndWrap(
                    grid, pres.cellWrapText(), pres.rowHeightPercent());
            compareSpreadsheet.setGrid(grid);
            compareSpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);

            Platform.runLater(
                    () -> {
                        SpreadsheetTabularSupport.applyColumnWidths(compareSpreadsheet, widths, widthDefault);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(compareSpreadsheet);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                compareSpreadsheet, headerColumnCountCompare.get());
                        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(compareSpreadsheet);
                        SpreadsheetTabularSupport.refreshSpreadsheetAfterRowPresentationChange(compareSpreadsheet);
                        SpreadsheetColumnDragReorderSupport.refreshAfterGridReady(
                                compareSpreadsheet,
                                suppressComparePersistence::get,
                                () -> new ArrayList<>(compareHeadersRef),
                                headerColumnCountCompare.get(),
                                this::applyPersistedCompareColumnOrderAfterLogicalReorder);
                        ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                                compareSpreadsheet,
                                () -> new ArrayList<>(compareHeadersRef),
                                () ->
                                        TableColumnOrderPersistence.loadColumnVisibility(
                                                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE,
                                                compareHeadersRef.size()));
                    });
        } finally {
            suppressComparePersistence.set(false);
        }
    }

    @FXML
    private void onClearMainFilters() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(mainSpreadsheet);
    }

    @FXML
    private void onClearCompareFilters() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(compareSpreadsheet);
    }

    @FXML
    private void onRefreshButtonAction() {
        if (requestFactory == null) {
            statusLabel.setText("\u521d\u671f\u5316\u5f85\u3061");
            return;
        }
        refreshButton.setDisable(true);
        statusLabel.setText("\u53d6\u5f97\u4e2d...");
        RunRequest req = requestFactory.get();
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            refreshButton.setDisable(false);
                                            if (err != null) {
                                                statusLabel.setText("error: " + err.getMessage());
                                                if (shell != null) {
                                                    shell.appendLog("[delivery-calendar] " + err.getMessage());
                                                }
                                                return;
                                            }
                                            if (cap == null) {
                                                statusLabel.setText("no result");
                                                return;
                                            }
                                            statusLabel.setText("exit=" + cap.exitCode());
                                            applyPayload(cap.stdout());
                                        }));
    }

    private void applyPayload(String stdout) {
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
                StringBuilder sb = new StringBuilder();
                if (meta.hasNonNull("processingPlanPath")) {
                    sb.append("PM_AI_PROCESSING_PLAN_PATH: ")
                            .append(meta.get("processingPlanPath").asText())
                            .append("\n");
                }
                if (meta.hasNonNull("dispatchJsonPath")) {
                    sb.append("\u7d50\u679c_\u914d\u53f0\u8868.json: ")
                            .append(meta.get("dispatchJsonPath").asText());
                }
                metaLabel.setText(sb.toString());
            }

            JsonNode mainCal = root.get("mainCalendar");
            if (mainCal != null && mainCal.isObject()) {
                loadMainCalendar(mainCal);
            } else {
                mainHeadersRef.clear();
                mainRows.clear();
                rebuildMainSpreadsheet();
            }

            JsonNode cmp = root.get("planCompareTable");
            if (cmp != null && cmp.isObject()) {
                loadCompareTable(cmp);
            } else {
                compareHeadersRef.clear();
                compareRows.clear();
                rebuildCompareSpreadsheet();
            }
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
        JsonNode rows = mainCal.get("rows");
        if (rows != null && rows.isArray()) {
            for (JsonNode row : rows) {
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

        // #region agent log
        int tripleTotal = 0;
        int tripleNonEmpty = 0;
        for (ObservableList<DeliveryCalendarMainCell> line : mainRows) {
            for (DeliveryCalendarMainCell c : line) {
                if (c instanceof DeliveryCalendarMainCell.TripleQty t) {
                    tripleTotal++;
                    if (!t.plan().isBlank()
                            || !t.actual().isBlank()
                            || !t.dispatch().isBlank()) {
                        tripleNonEmpty++;
                    }
                }
            }
        }
        ObjectNode dbg = JSON.createObjectNode();
        dbg.put("tripleTotal", tripleTotal);
        dbg.put("tripleNonEmpty", tripleNonEmpty);
        dbg.put("mainHeaderCount", mainHeadersRef.size());
        agentDebugNdjson(
                "H5",
                "DeliveryCalendarViewTabController.loadMainCalendar",
                "parsed cells before column reorder",
                dbg);
        // #endregion

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
        // #region agent log
        ObjectNode dbgOrder = JSON.createObjectNode();
        dbgOrder.put("savedOrderSize", titleOrder.size());
        StringBuilder sbOrder = new StringBuilder();
        int lim = Math.min(12, titleOrder.size());
        for (int i = 0; i < lim; i++) {
            if (i > 0) {
                sbOrder.append(" | ");
            }
            sbOrder.append(titleOrder.get(i));
        }
        dbgOrder.put("savedOrderSample", sbOrder.toString());
        agentDebugNdjson(
                "H6",
                "DeliveryCalendarViewTabController.loadMainCalendar",
                "persisted column title order sample",
                dbgOrder);
        // #endregion
        TableColumnOrderPersistence.applyLogicalColumnOrderDeliveryCalendar(
                mainHeadersRef, mainRows, titleOrder);
        boolean[] visAfter =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        beforeHeaders, visBefore, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN, visAfter);

        rebuildMainSpreadsheet();
    }

    private void loadCompareTable(JsonNode cmp) {
        compareHeadersRef.clear();
        JsonNode cols = cmp.get("columns");
        if (cols != null && cols.isArray()) {
            for (JsonNode c : cols) {
                String key = c.asText("");
                compareHeadersRef.add(COMPARE_HEADER_JP.getOrDefault(key, key));
            }
        }
        compareRows.clear();
        JsonNode rows = cmp.get("rows");
        if (rows != null && rows.isArray()) {
            for (JsonNode row : rows) {
                ObservableList<String> line = FXCollections.observableArrayList();
                if (row.isArray()) {
                    for (JsonNode cell : row) {
                        line.add(cell.asText(""));
                    }
                }
                compareRows.add(line);
            }
        }

        List<TableColumnOrderPersistence.ColumnSpec> lay =
                TableColumnOrderPersistence.loadLayout(
                        TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE);
        persistedLayoutCompare.set(lay);
        List<String> beforeHeaders = new ArrayList<>(compareHeadersRef);
        boolean[] visBefore =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE,
                        beforeHeaders.size());
        List<String> titleOrder =
                lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
        TableColumnOrderPersistence.applyLogicalColumnOrder(
                compareHeadersRef, compareRows, titleOrder);
        boolean[] visAfter =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        beforeHeaders, visBefore, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_COMPARE, visAfter);

        rebuildCompareSpreadsheet();
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
