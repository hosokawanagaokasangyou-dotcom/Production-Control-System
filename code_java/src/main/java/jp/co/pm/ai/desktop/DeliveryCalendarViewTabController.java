package jp.co.pm.ai.desktop;

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
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.control.TabPane;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TextArea;
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

    @FXML
    private TextArea mainSelectedCellInfo;

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

    private final AtomicBoolean suppressInnerTabSessionPersistence = new AtomicBoolean(false);

    private volatile boolean innerTabPersistenceWired;

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

        if (metaScrollPane != null && metaLabel != null) {
            metaScrollPane.setFitToWidth(true);
            metaLabel.setWrapText(true);
            metaLabel.prefWidthProperty().bind(metaScrollPane.widthProperty().subtract(18));
        }

        SpreadsheetTabularSupport.installDeliveryCalendarSpreadsheetChrome(mainSpreadsheet);
        SpreadsheetTabularSupport.installDeliveryCalendarSpreadsheetChrome(compareSpreadsheet);
        mainSpreadsheet.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        compareSpreadsheet.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        /*
         * \u5168\u884c\u9078\u629e\u62e1\u5f35\u306f\u30af\u30ea\u30c3\u30af\u5217\u3068 TablePosition \u306e\u5217\u304c\u305a\u308c\u308b
         * \uff08\u62e1\u5f35\u5f8c\u306e\u30d5\u30a9\u30fc\u30ab\u30b9\u304c\u5148\u982d\u5217\u306b\u79fb\u308b\u305f\u3081\uff09\u3002
         * \u4e0b\u90e8\u306e\u30bb\u30eb\u5185\u5bb9\u8868\u793a\u3068\u5408\u308f\u305b\u308b\u305f\u3081\u3001\u30e1\u30a4\u30f3\u8868\u3067\u306e\u307f\u62e1\u5f35\u3092\u7121\u52b9\u5316\u3002
         */
        SpreadsheetTabularSupport.installFullRowDataSelection(mainSpreadsheet, () -> true);
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

        installMainSelectedCellInfoListener();
    }

    /**
     * \u9078\u629e\u30bb\u30eb\u306e\u30e2\u30c7\u30eb\u5024\uff08{@link DeliveryCalendarMainCell.TripleQty} \u307e\u305f\u306f
     * {@link DeliveryCalendarMainCell.PlainText}\uff09\u3092 {@link #mainSelectedCellInfo} \u306b\u8868\u793a\u3059\u308b\u3002
     */
    private void installMainSelectedCellInfoListener() {
        if (mainSelectedCellInfo == null) {
            return;
        }
        mainSelectedCellInfo.setText("\u30bb\u30eb\u3092\u30af\u30ea\u30c3\u30af\u3059\u308b\u3068\u3001"
                + "\u305d\u306e\u30bb\u30eb\u306e\u30e2\u30c7\u30eb\u4e0a\u306e\u5024\u304c\u3053\u3053\u306b\u51fa\u307e\u3059\u3002");
        mainSpreadsheet
                .getSelectionModel()
                .getSelectedCells()
                .addListener(
                        (javafx.collections.ListChangeListener<? super TablePosition>) ch ->
                                Platform.runLater(this::refreshMainSelectedCellInfo));
    }

    private void refreshMainSelectedCellInfo() {
        if (mainSelectedCellInfo == null) {
            return;
        }
        var sm = mainSpreadsheet.getSelectionModel();
        TablePosition<?, ?> pos = sm.getFocusedCell();
        if (pos == null || pos.getRow() < 0) {
            var sel = sm.getSelectedCells();
            if (sel == null || sel.isEmpty()) {
                mainSelectedCellInfo.setText("(\u672a\u9078\u629e)");
                return;
            }
            pos = sel.getFirst();
        }
        if (pos == null || pos.getRow() < 0) {
            mainSelectedCellInfo.setText("(\u672a\u9078\u629e)");
            return;
        }
        /* TablePosition#getRow \u306f\u30d3\u30e5\u30fc\u884c\u3002\u30b0\u30ea\u30c3\u30c9\u884c\u306f SpreadsheetView#getModelRow \u3092\u4f7f\u3046\u3002 */
        int viewRow = pos.getRow();
        int gridRow = mainSpreadsheet.getModelRow(viewRow);
        int col = pos.getColumn();
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        StringBuilder sb = new StringBuilder();
        String header = (col >= 0 && col < mainHeadersRef.size()) ? mainHeadersRef.get(col) : "";
        sb.append("viewRow=").append(viewRow)
                .append("  gridRow=").append(gridRow)
                .append("  col=").append(col)
                .append("  header=").append(header)
                .append('\n');
        if (gridRow < 0) {
            sb.append("(\u884c\u306e\u89e3\u6c7a\u5931\u8d25 viewRow=").append(viewRow).append(")");
            mainSelectedCellInfo.setText(sb.toString());
            return;
        }
        if (gridRow < firstData) {
            sb.append("(\u30d5\u30a3\u30eb\u30bf\u884c)");
            mainSelectedCellInfo.setText(sb.toString());
            return;
        }
        int dataRow = gridRow - firstData;
        if (dataRow < 0 || dataRow >= mainRows.size()) {
            sb.append("(\u30c7\u30fc\u30bf\u884c\u30aa\u30fc\u30d0\u30fc)");
            mainSelectedCellInfo.setText(sb.toString());
            return;
        }
        ObservableList<DeliveryCalendarMainCell> line = mainRows.get(dataRow);
        if (col < 0 || col >= line.size()) {
            sb.append("(\u5217\u30aa\u30fc\u30d0\u30fc)");
            mainSelectedCellInfo.setText(sb.toString());
            return;
        }
        DeliveryCalendarMainCell mc = line.get(col);
        sb.append("dataRow=").append(dataRow).append('\n');
        if (mc instanceof DeliveryCalendarMainCell.TripleQty t) {
            sb.append("type: TripleQty\n")
                    .append("  plan(\u30bf\u30b9\u30af\u5165\u529b)     : ").append(quoteForCellInfo(t.plan())).append('\n')
                    .append("  actual(\u5b9f\u7e3e\u660e\u7d30)       : ").append(quoteForCellInfo(t.actual())).append('\n')
                    .append("  dispatch(\u7d50\u679c_\u914d\u53f0\u8868) : ").append(quoteForCellInfo(t.dispatch())).append('\n');
        } else if (mc instanceof DeliveryCalendarMainCell.PlainText pt) {
            sb.append("type: PlainText\n").append("  text: ").append(quoteForCellInfo(pt.text())).append('\n');
        } else {
            sb.append("type: (null)\n");
        }
        mainSelectedCellInfo.setText(sb.toString());
    }

    private static String quoteForCellInfo(String s) {
        if (s == null) return "(null)";
        if (s.isEmpty()) return "(\u7a7a\u6587\u5b57)  len=0";
        return "\"" + s + "\"  len=" + s.length();
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

        if (processingActualsDataTabController != null) {
            processingActualsDataTabController.bindShell(shell);
        }
        if (aladdinProcessingPlanDataTabController != null) {
            aladdinProcessingPlanDataTabController.bindShell(shell);
        }
        if (deliveryCalendarResultDispatchTableTabController != null) {
            deliveryCalendarResultDispatchTableTabController.bindShell(shell);
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
                    grid,
                    pres.cellWrapText(),
                    pres.rowHeightPercent(),
                    SpreadsheetTabularSupport.DELIVERY_CALENDAR_ROW_HEIGHT_BASE_PX,
                    SpreadsheetTabularSupport.DELIVERY_CALENDAR_ROW_HEIGHT_MIN_PX);
            mainSpreadsheet.setGrid(grid);
            mainSpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);

            Platform.runLater(
                    () -> {
                        SpreadsheetTabularSupport.applyColumnWidths(mainSpreadsheet, widths, widthDefault);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(mainSpreadsheet);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                mainSpreadsheet, headerColumnCountMain.get());
                        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(mainSpreadsheet);
                        SpreadsheetTabularSupport.refreshSpreadsheetAfterRowPresentationChange(
                                mainSpreadsheet, true);
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
