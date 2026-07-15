package jp.co.pm.ai.desktop;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;

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
import javafx.scene.control.ProgressBar;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchDeadlineJudgment;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchInteractiveConsolidator;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchStage3Support;
import jp.co.pm.ai.desktop.io.AladdinProcessingPlanSourceReloader;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.io.DispatchAladdinEntryWorkbookExporter;
import jp.co.pm.ai.desktop.reconciliation.RequestFormOriginalIndexLookup;
import jp.co.pm.ai.desktop.ui.DispatchAladdinEntryGenerationDialog;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnDragReorderSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetRequestFormOriginalHeaderStyle;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * Loads {@link AppPaths#RESULT_DISPATCH_TABLE_JSON_BASENAME} into ControlsFX {@link SpreadsheetView}. Layout
 * {@code ResultDispatchTableTab.fxml}.
 */
public final class ResultDispatchTableTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    public record AladdinEntryExportOutcome(
            DispatchAladdinEntryWorkbookExporter.ExportResult result,
            List<String> warnings,
            Exception error) {
        public boolean succeeded() {
            return error == null && result != null;
        }
    }

    private static final String HINT_TEXT =
            "PM_AI_RESULT_DISPATCH_TABLE_DIR またはデフォルトの code/output/"
                    + " 配下の JSON を表示します。再読みで"
                    + "最新化します。"
                    + " ControlsFX SpreadsheetView （段階1成形結果と同じ"
                    + "列フィルタ）。";

    @FXML
    private Button refreshButton;

    @FXML private Button aladdinEntryReloadExportButton;

    @FXML
    private Button aladdinEntryLocalExportButton;

    @FXML
    private Button aladdinEntryOpenLocalLatestButton;

    @FXML
    private Button aladdinEntryOpenLocalGenerationsButton;

    @FXML
    private Button aladdinEntryOpenLatestButton;

    @FXML
    private Button aladdinEntryOpenGenerationsButton;

    @FXML
    private TitledPane operationsSourceTitledPane;

    @FXML
    private ProgressBar aladdinEntryExportProgress;

    @FXML
    private Label aladdinEntryExportProgressLabel;

    @FXML
    private Label dataStageBadgeLabel;

    @FXML
    private Label statusLabel;

    @FXML
    private Label pathLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private Slider resultDispatchRowHeightSlider;

    @FXML
    private Label resultDispatchRowHeightPctLabel;

    @FXML
    private CheckBox resultDispatchCellWrapCheck;

    @FXML
    private HBox columnStripHost;

    @FXML
    private StackPane spreadsheetHost;

    private MainShellController shell;

    private Stage ownerStage;

    private final SpreadsheetView spreadsheetView = new SpreadsheetView();

    private final List<String> headersRef = new ArrayList<>();

    private ObservableList<ObservableList<String>> rows;

    private final AtomicBoolean suppressColumnPersistence = new AtomicBoolean(false);

    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());

    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    private final AtomicReference<TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs>
            spreadsheetTabPrefs =
                    new AtomicReference<>(
                            TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs.defaults());

    private final AtomicBoolean suppressPresentationUiEvents = new AtomicBoolean(false);

    private volatile boolean resultDispatchPresentationHooksInstalled;

    private boolean embeddedInDeliveryCalendar;

    private ButtonAttentionGlow aladdinEntryOpenLatestGlow;

    @FXML
    private void initialize() {
        hintLabel.setText(HINT_TEXT);
        rows = FXCollections.observableArrayList();

        StackPane.setAlignment(spreadsheetView, Pos.CENTER_LEFT);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
        SpreadsheetThemeBridge.install(spreadsheetView);
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(spreadsheetView);

        columnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::applyDynamicColumnWidths,
                                TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                                headerColumnCount,
                                this::onLeadingColumnCountCommitted,
                                this::onReorderColumns,
                                () ->
                                        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                                                ownerStage,
                                                TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                                                spreadsheetView,
                                                () -> new ArrayList<>(headersRef))));

        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                spreadsheetHost, headerColumnCount::get);
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        ownerStage = shell.getPrimaryStage();

        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                spreadsheetView,
                TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                suppressColumnPersistence::get,
                () -> new ArrayList<>(headersRef));

        initResultDispatchSpreadsheetPresentationControls();

        Platform.runLater(() -> reloadFromDisk(false));
    }

    /**
     * 納期管理ビューに埋め込んだときは親の「再読込」で JSON を更新するため、本タブの再読みボタンを隠す。
     * メインシェル単独タブでは {@code true} のまま。
     */
    void setResultDispatchRefreshButtonVisible(boolean visible) {
        if (refreshButton != null) {
            refreshButton.setVisible(visible);
            refreshButton.setManaged(visible);
        }
    }

    /** 納期管理ビューに埋め込んだときのみ {@code true}（納期判定列を付与）。 */
    void setEmbeddedInDeliveryCalendar(boolean embedded) {
        embeddedInDeliveryCalendar = embedded;
    }

    /** 親（納期管理ビュー）の再読込成功後に呼ぶ。 */
    public void reloadResultDispatchTableFromDisk() {
        reloadFromDisk(false);
    }

    /**
     * 納期管理ビュー内「配台結果」子タブ表示時: 「操作・ソース」を開き、必要なら「最新を開く」を光らせる。
     *
     * @param promptExcelExportAttention 段階2 完了直後など、生成済みExcelを開くグローを付ける
     */
    public void onEmbeddedDispatchResultTabShown(boolean promptExcelExportAttention) {
        if (!embeddedInDeliveryCalendar) {
            return;
        }
        expandOperationsSourcePane();
        if (promptExcelExportAttention) {
            startAladdinOpenAttentionGlow();
        }
    }

    private void expandOperationsSourcePane() {
        if (operationsSourceTitledPane != null) {
            operationsSourceTitledPane.setExpanded(true);
        }
    }

    private void startAladdinOpenAttentionGlow() {
        expandOperationsSourcePane();
        if (aladdinEntryOpenLatestGlow == null && aladdinEntryOpenLatestButton != null) {
            aladdinEntryOpenLatestGlow = new ButtonAttentionGlow(aladdinEntryOpenLatestButton);
        }
        if (aladdinEntryOpenLatestGlow != null) {
            aladdinEntryOpenLatestGlow.startIfIdle();
        }
    }

    private void dismissAladdinOpenAttentionGlow() {
        ButtonAttentionGlow.stopAll(aladdinEntryOpenLatestGlow);
    }

    private void onLeadingColumnCountCommitted(int n) {
        headerColumnCount.set(n);
        rebuildSpreadsheet();
    }

    private void initResultDispatchSpreadsheetPresentationControls() {
        if (resultDispatchPresentationHooksInstalled) {
            return;
        }
        if (resultDispatchRowHeightSlider == null) {
            return;
        }
        resultDispatchPresentationHooksInstalled = true;
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs loaded =
                TableColumnOrderPersistence.loadSpreadsheetTabPresentationPrefs(
                        TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE);
        spreadsheetTabPrefs.set(loaded);
        suppressPresentationUiEvents.set(true);
        try {
            resultDispatchRowHeightSlider.setMin(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN);
            resultDispatchRowHeightSlider.setMax(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX);
            resultDispatchRowHeightSlider.setValue(loaded.rowHeightPercent());
            resultDispatchRowHeightSlider.setMajorTickUnit(250);
            resultDispatchRowHeightSlider.setMinorTickCount(4);
            resultDispatchRowHeightSlider.setShowTickMarks(true);
            if (resultDispatchRowHeightPctLabel != null) {
                resultDispatchRowHeightPctLabel.setText(
                        String.format("%.0f%%", loaded.rowHeightPercent()));
            }
            if (resultDispatchCellWrapCheck != null) {
                resultDispatchCellWrapCheck.setSelected(loaded.cellWrapText());
            }
        } finally {
            suppressPresentationUiEvents.set(false);
        }
        SliderCommittedChangeSupport.install(
                resultDispatchRowHeightSlider,
                () -> {
                    if (resultDispatchRowHeightPctLabel != null
                            && resultDispatchRowHeightSlider != null) {
                        resultDispatchRowHeightPctLabel.setText(
                                String.format("%.0f%%", resultDispatchRowHeightSlider.getValue()));
                    }
                },
                this::commitResultDispatchSpreadsheetPresentationFromSlider);
        if (resultDispatchCellWrapCheck != null) {
            resultDispatchCellWrapCheck
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (suppressPresentationUiEvents.get()) {
                                    return;
                                }
                                commitResultDispatchSpreadsheetPresentationFromUi();
                            });
        }
    }

    private void commitResultDispatchSpreadsheetPresentationFromSlider() {
        if (suppressPresentationUiEvents.get()) {
            return;
        }
        commitResultDispatchSpreadsheetPresentationFromUi();
    }

    private void commitResultDispatchSpreadsheetPresentationFromUi() {
        if (resultDispatchRowHeightSlider == null) {
            return;
        }
        double v = resultDispatchRowHeightSlider.getValue();
        boolean wrap =
                resultDispatchCellWrapCheck != null && resultDispatchCellWrapCheck.isSelected();
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs next =
                new TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs(v, wrap);
        spreadsheetTabPrefs.set(next);
        TableColumnOrderPersistence.saveSpreadsheetTabPresentationPrefs(
                TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE, next);
        if (resultDispatchRowHeightPctLabel != null) {
            resultDispatchRowHeightPctLabel.setText(String.format("%.0f%%", v));
        }
        rebuildSpreadsheet();
    }

    private void onReorderColumns() {
        if (headersRef.isEmpty()) {
            shell.appendLog(
                    "[result-dispatch-json] 列がありません（先に再読み）");
            return;
        }
        boolean[] visForDialog =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                        headersRef.size());
        SpreadsheetColumnReorderDialog.show(
                        ownerStage, new ArrayList<>(headersRef), visForDialog)
                .ifPresent(
                        perm -> {
                            List<String> oldHeaders = new ArrayList<>(headersRef);
                            List<String> titleOrder = perm.stream().map(oldHeaders::get).toList();
                            applyPersistedColumnOrderAfterLogicalReorder(titleOrder);
                        });
    }

    private void applyPersistedColumnOrderAfterLogicalReorder(List<String> titleOrder) {
        if (headersRef.isEmpty()) {
            return;
        }
        List<String> oldHeaders = new ArrayList<>(headersRef);
        boolean[] oldVis =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE, oldHeaders.size());
        List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
        TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        oldHeaders, oldVis, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE, newVis);
        List<Double> widths =
                TableColumnOrderPersistence.resolveWidthsForHeaders(headersRef, lay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
        for (int i = 0; i < headersRef.size(); i++) {
            newLay.add(
                    new TableColumnOrderPersistence.ColumnSpec(headersRef.get(i), widths.get(i)));
        }
        persistedLayout.set(newLay);
        TableColumnOrderPersistence.saveLayout(
                TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE, newLay);
        rebuildSpreadsheet();
    }

    private void applyDynamicColumnWidths() {
        double w = 112;
        for (var c : spreadsheetView.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private void rebuildSpreadsheet() {
        if (headersRef.isEmpty()) {
            spreadsheetView.setGrid(new GridBase(0, 0));
            return;
        }
        suppressColumnPersistence.set(true);
        try {
            final List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            headersRef, persistedLayout.get(), 112);
            final double widthDefault = 112;

            int deadlineJudgmentColIdx = -1;
            if (embeddedInDeliveryCalendar) {
                deadlineJudgmentColIdx =
                        headersRef.indexOf(ResultDispatchDeadlineJudgment.COL_TITLE);
            }
            GridBase grid =
                    SpreadsheetTabularSupport.buildReadOnlyPlainGrid(
                            headersRef, rows, headerColumnCount.get(), deadlineJudgmentColIdx);
            TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs pres =
                    spreadsheetTabPrefs.get();
            SpreadsheetTabularSupport.applySpreadsheetGridRowHeightsAndWrap(
                    grid, pres.cellWrapText(), pres.rowHeightPercent());
            spreadsheetView.setGrid(grid);

            Platform.runLater(
                    () -> {
                        SpreadsheetTabularSupport.applyColumnWidths(
                                spreadsheetView, widths, widthDefault);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                spreadsheetView, headerColumnCount.get());
                        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(spreadsheetView);
                        SpreadsheetTabularSupport.pinSpreadsheetFilterRow(spreadsheetView);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(spreadsheetView);
                        SpreadsheetTabularSupport.refreshSpreadsheetAfterRowPresentationChange(
                                spreadsheetView, true);
                        SpreadsheetColumnDragReorderSupport.refreshAfterGridReady(
                                spreadsheetView,
                                suppressColumnPersistence::get,
                                () -> new ArrayList<>(headersRef),
                                headerColumnCount.get(),
                                this::applyPersistedColumnOrderAfterLogicalReorder);
                        ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                                spreadsheetView,
                                () -> new ArrayList<>(headersRef),
                                () ->
                                        TableColumnOrderPersistence.loadColumnVisibility(
                                                TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                                                headersRef.size()),
                                () ->
                                        SpreadsheetTabularSupport.reapplySpreadsheetColumnChrome(
                                                spreadsheetView, headerColumnCount.get()));
                        SpreadsheetRequestFormOriginalHeaderStyle.applyWhenReady(
                                spreadsheetView, new ArrayList<>(headersRef));
                    });
        } finally {
            suppressColumnPersistence.set(false);
        }
    }

    @FXML
    private void onRefreshButtonAction() {
        reloadFromDisk(true);
    }

    private void reloadFromDisk(boolean userCompletionDialog) {
        if (shell == null) {
            return;
        }
        if (refreshButton != null) {
            refreshButton.setDisable(true);
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path path = AppPaths.resolveResultDispatchTableJsonPath(ui);
        pathLabel.setText(path.toString());
        if (!Files.isRegularFile(path)) {
            statusLabel.setText("ファイルなし");
            applyEmpty();
            if (refreshButton != null) {
                refreshButton.setDisable(false);
            }
            if (userCompletionDialog) {
                shell.showWarningDialog("再読み", "ファイルが見つかりません。\n" + path);
            }
            return;
        }
        try {
            String raw = Files.readString(path, StandardCharsets.UTF_8);
            JsonNode root = JSON.readTree(raw);
            JsonNode columnsNode = root.get("columns");
            JsonNode rowsNode = root.get("rows");
            if (columnsNode == null || !columnsNode.isArray() || rowsNode == null || !rowsNode.isArray()) {
                statusLabel.setText("JSON 構造が不正");
                applyEmpty();
                if (refreshButton != null) {
                    refreshButton.setDisable(false);
                }
                if (userCompletionDialog) {
                    shell.showWarningDialog(
                            "再読み",
                            "JSON の構造が不正です（columns / rows が必要です）。\n" + path);
                }
                return;
            }
            List<String> headerOrder = new ArrayList<>();
            for (JsonNode c : columnsNode) {
                if (c != null && c.isTextual()) {
                    headerOrder.add(c.asText(""));
                }
            }
            List<Map<String, String>> rowMaps = new ArrayList<>();
            for (JsonNode r : rowsNode) {
                if (r == null || !r.isObject()) {
                    continue;
                }
                LinkedHashMap<String, String> row = new LinkedHashMap<>();
                for (String h : headerOrder) {
                    row.put(h, formatCell(r.get(h)));
                }
                rowMaps.add(row);
            }
            ResultDispatchInteractiveConsolidator.consolidatePlanAndTimelineRowsInPlace(
                    headerOrder, rowMaps);
            boolean stage3 = ResultDispatchStage3Support.hasStage3ActualColumn(headerOrder);
            if (stage3) {
                ResultDispatchStage3Support.applyStage3DisplayQuantities(headerOrder, rowMaps);
                ResultDispatchStage3Support.removeRedundantActualColumnFromMaps(headerOrder, rowMaps);
            }
            ResultDispatchStage3Support.applyPlanningStageBadgeFromDispatchJson(
                    dataStageBadgeLabel, path);
            statusLabel.setText(rowMaps.size() + " 行");

            headersRef.clear();
            headersRef.addAll(headerOrder);
            rows.clear();
            for (Map<String, String> map : rowMaps) {
                ObservableList<String> line = FXCollections.observableArrayList();
                for (String h : headersRef) {
                    line.add(map.getOrDefault(h, ""));
                }
                rows.add(line);
            }

            List<TableColumnOrderPersistence.ColumnSpec> lay =
                    TableColumnOrderPersistence.loadLayout(
                            TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE);
            persistedLayout.set(lay);
            List<String> beforeHeaders = new ArrayList<>(headersRef);
            boolean[] visBefore =
                    TableColumnOrderPersistence.loadColumnVisibility(
                            TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                            beforeHeaders.size());
            List<String> titleOrder =
                    lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
            TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
            boolean[] visAfter =
                    TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                            beforeHeaders, visBefore, titleOrder);
            TableColumnOrderPersistence.saveColumnVisibility(
                    TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE, visAfter);

            injectDeadlineJudgmentColumnIfNeeded();

            rebuildSpreadsheet();
            if (userCompletionDialog) {
                shell.showInformationDialog(
                        "再読み完了",
                        "結果_配台表を読み込みました。\n" + path + "\n行数: " + rowMaps.size());
            }
        } catch (Exception ex) {
            statusLabel.setText("error");
            shell.appendLog("[result-dispatch-json] " + ex.getMessage());
            applyEmpty();
            if (userCompletionDialog) {
                shell.showErrorDialog(
                        "再読みエラー",
                        ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
        } finally {
            if (refreshButton != null) {
                refreshButton.setDisable(false);
            }
        }
    }

    private void applyEmpty() {
        headersRef.clear();
        rows.clear();
        persistedLayout.set(List.of());
        spreadsheetView.setGrid(new GridBase(0, 0));
    }

    private void injectDeadlineJudgmentColumnIfNeeded() {
        if (!embeddedInDeliveryCalendar || headersRef.isEmpty()) {
            return;
        }
        final String col = ResultDispatchDeadlineJudgment.COL_TITLE;
        int colIdx = headersRef.indexOf(col);
        if (colIdx < 0) {
            int afterAnswer = headersRef.indexOf("回答納期");
            colIdx = afterAnswer >= 0 ? afterAnswer + 1 : headersRef.size();
            headersRef.add(colIdx, col);
            for (ObservableList<String> row : rows) {
                row.add(colIdx, "");
            }
        }
        for (int r = 0; r < rows.size(); r++) {
            ObservableList<String> line = rows.get(r);
            LinkedHashMap<String, String> map = new LinkedHashMap<>();
            for (int c = 0; c < headersRef.size(); c++) {
                String h = headersRef.get(c);
                if (ResultDispatchDeadlineJudgment.COL_TITLE.equals(h)) {
                    continue;
                }
                String v = c < line.size() && line.get(c) != null ? line.get(c) : "";
                map.put(h, v);
            }
            line.set(colIdx, ResultDispatchDeadlineJudgment.judgmentOkNg(map));
        }
    }

    private static String formatCell(JsonNode n) {
        if (n == null || n.isNull()) {
            return "";
        }
        if (n.isBoolean()) {
            return n.asBoolean() ? "true" : "false";
        }
        if (n.isInt() || n.isLong()) {
            return Long.toString(n.longValue());
        }
        if (n.isDouble() || n.isFloat() || n.isBigDecimal()) {
            double d = n.asDouble();
            if (Double.isFinite(d) && d == Math.rint(d) && Math.abs(d) < 1e15) {
                return Long.toString((long) d);
            }
            return n.asText("");
        }
        if (n.isTextual()) {
            String t = n.asText("");
            if (t.length() >= 19 && t.charAt(10) == 'T' && t.charAt(4) == '-') {
                return t.substring(0, 10);
            }
            return t;
        }
        return n.asText("");
    }

    @FXML
    private void onReloadAladdinPlanAndExportEntryWorkbookAction() {
        runAladdinEntryExport(true, DispatchAladdinEntryWorkbookExporter.Destination.SHARED);
    }

    @FXML
    private void onExportAladdinEntryWorkbookLocalAction() {
        runAladdinEntryExport(false, DispatchAladdinEntryWorkbookExporter.Destination.LOCAL);
    }

    private void runAladdinEntryExport(
            boolean reloadAladdinPlanFromSource,
            DispatchAladdinEntryWorkbookExporter.Destination destination) {
        runAladdinEntryExport(reloadAladdinPlanFromSource, destination, true, null);
    }

    public void exportSharedAladdinEntryWorkbookAfterStage2(
            Consumer<AladdinEntryExportOutcome> completion) {
        runAladdinEntryExport(
                false,
                DispatchAladdinEntryWorkbookExporter.Destination.SHARED,
                false,
                completion);
    }

    private void runAladdinEntryExport(
            boolean reloadAladdinPlanFromSource,
            DispatchAladdinEntryWorkbookExporter.Destination destination,
            boolean showCompletionDialog,
            Consumer<AladdinEntryExportOutcome> completion) {
        if (shell == null) {
            if (completion != null) {
                completion.accept(
                        new AladdinEntryExportOutcome(
                                null, List.of(), new IllegalStateException("画面の初期化が完了していません。")));
            }
            return;
        }
        DispatchAladdinEntryWorkbookExporter.Destination dest =
                destination != null
                        ? destination
                        : DispatchAladdinEntryWorkbookExporter.Destination.SHARED;
        dismissAladdinOpenAttentionGlow();
        Map<String, String> ui = shell.snapshotUiEnv();
        setAladdinEntryExportButtonsDisabled(true);
        showAladdinEntryExportProgress(
                ProgressBar.INDETERMINATE_PROGRESS,
                reloadAladdinPlanFromSource
                        ? "アラジン加工計画読込中…"
                        : "依頼書原本 目次読込中…");
        Thread worker =
                new Thread(
                        () -> {
                            List<String> warnings = new ArrayList<>();
                            try {
                                if (reloadAladdinPlanFromSource) {
                                    AladdinProcessingPlanSourceReloader.ReloadResult reload =
                                            AladdinProcessingPlanSourceReloader
                                                    .reloadNewestFromDiskAndSaveShapedJson(ui);
                                    shell.appendLog(
                                            "[aladdin-entry-export] 加工計画再読込: "
                                                    + reload.sourceFile()
                                                    + " ("
                                                    + reload.rowCount()
                                                    + " 行 × "
                                                    + reload.columnCount()
                                                    + " 列)");
                                    Platform.runLater(shell::refreshAladdinProcessingPlanTabFromDisk);
                                }
                                Map<String, DispatchAladdinEntrySheetBuilder.IndexInfo> index =
                                        RequestFormOriginalIndexLookup.loadByIraiNoKey(
                                                ui,
                                                warnings,
                                                (processed, total) ->
                                                        Platform.runLater(
                                                                () ->
                                                                        showAladdinEntryExportProgress(
                                                                                total > 0
                                                                                        ? (double) processed / total
                                                                                        : ProgressBar.INDETERMINATE_PROGRESS,
                                                                                "依頼書原本 目次読込中… ("
                                                                                        + processed
                                                                                        + "/"
                                                                                        + total
                                                                                        + ")")));
                                Platform.runLater(
                                        () ->
                                                showAladdinEntryExportProgress(
                                                        ProgressBar.INDETERMINATE_PROGRESS,
                                                        dest == DispatchAladdinEntryWorkbookExporter
                                                                        .Destination.LOCAL
                                                                ? "Excel 作成中（ローカル）…"
                                                                : "Excel 作成中…"));
                                DispatchAladdinEntryWorkbookExporter.ExportResult result =
                                        DispatchAladdinEntryWorkbookExporter.writeFromCachedSources(
                                                ui, index, dest);
                                Platform.runLater(
                                        () ->
                                                finishAladdinEntryExport(
                                                        result,
                                                        warnings,
                                                        null,
                                                        reloadAladdinPlanFromSource,
                                                        dest,
                                                        showCompletionDialog,
                                                        completion));
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () ->
                                                finishAladdinEntryExport(
                                                        null,
                                                        warnings,
                                                        ex,
                                                        reloadAladdinPlanFromSource,
                                                        dest,
                                                        showCompletionDialog,
                                                        completion));
                            }
                        },
                        reloadAladdinPlanFromSource
                                ? "aladdin-entry-reload-export"
                                : dest == DispatchAladdinEntryWorkbookExporter.Destination.LOCAL
                                        ? "aladdin-entry-local-export"
                                        : "aladdin-entry-export");
        worker.setDaemon(true);
        worker.start();
    }

    private void setAladdinEntryExportButtonsDisabled(boolean disabled) {
        if (aladdinEntryReloadExportButton != null) {
            aladdinEntryReloadExportButton.setDisable(disabled);
        }
        if (aladdinEntryLocalExportButton != null) {
            aladdinEntryLocalExportButton.setDisable(disabled);
        }
    }

    private void showAladdinEntryExportProgress(double progress, String text) {
        if (aladdinEntryExportProgress != null) {
            aladdinEntryExportProgress.setVisible(true);
            aladdinEntryExportProgress.setManaged(true);
            aladdinEntryExportProgress.setProgress(progress);
        }
        if (aladdinEntryExportProgressLabel != null) {
            aladdinEntryExportProgressLabel.setVisible(true);
            aladdinEntryExportProgressLabel.setManaged(true);
            aladdinEntryExportProgressLabel.setText(text);
        }
    }

    private void hideAladdinEntryExportProgress() {
        if (aladdinEntryExportProgress != null) {
            aladdinEntryExportProgress.setVisible(false);
            aladdinEntryExportProgress.setManaged(false);
        }
        if (aladdinEntryExportProgressLabel != null) {
            aladdinEntryExportProgressLabel.setVisible(false);
            aladdinEntryExportProgressLabel.setManaged(false);
        }
    }

    private void finishAladdinEntryExport(
            DispatchAladdinEntryWorkbookExporter.ExportResult result,
            List<String> warnings,
            Exception error,
            boolean reloadedAladdinPlan,
            DispatchAladdinEntryWorkbookExporter.Destination destination,
            boolean showCompletionDialog,
            Consumer<AladdinEntryExportOutcome> completion) {
        setAladdinEntryExportButtonsDisabled(false);
        hideAladdinEntryExportProgress();
        boolean local =
                destination == DispatchAladdinEntryWorkbookExporter.Destination.LOCAL;
        String dialogTitle =
                local
                        ? "アラジン入力用Excel出力（ローカル）"
                        : reloadedAladdinPlan
                                ? "アラジン加工計画読込→Excel出力"
                                : "アラジン入力用Excel出力";
        for (String w : warnings) {
            shell.appendLog("[aladdin-entry-export] warn: " + w);
        }
        if (error != null) {
            shell.appendLog("[aladdin-entry-export] error: " + error.getMessage());
            if (showCompletionDialog) {
                shell.showErrorDialog(
                        dialogTitle,
                        error.getMessage() != null ? error.getMessage() : error.toString());
            }
            if (completion != null) {
                completion.accept(new AladdinEntryExportOutcome(null, List.copyOf(warnings), error));
            }
            return;
        }
        shell.appendLog("[aladdin-entry-export] " + result.latestPath());
        StringBuilder sb =
                new StringBuilder(
                                local
                                        ? "アラジン入力用配台計画をローカルへ出力しました。\n\n最新: "
                                        : reloadedAladdinPlan
                                                ? "アラジン加工計画を再読込し、入力用配台計画 Excel を出力しました。\n\n最新: "
                                                : "アラジン入力用配台計画を出力しました。\n\n最新: ")
                        .append(result.latestPath())
                        .append("\n世代: ")
                        .append(result.generationPath());
        if (!warnings.isEmpty()) {
            sb.append("\n\n警告:\n").append(String.join("\n", warnings));
        }
        if (showCompletionDialog) {
            shell.showInformationDialog(dialogTitle, sb.toString());
        }
        if (!local) {
            startAladdinOpenAttentionGlow();
        }
        if (completion != null) {
            completion.accept(new AladdinEntryExportOutcome(result, List.copyOf(warnings), null));
        }
    }

    @FXML
    private void onOpenLatestAladdinEntryWorkbookAction() {
        dismissAladdinOpenAttentionGlow();
        if (shell == null) {
            return;
        }
        Path latest = AppPaths.aladdinEntryDispatchPlanXlsxPath(shell.snapshotUiEnv());
        if (!Files.isRegularFile(latest)) {
            shell.showWarningDialog(
                    "最新を開く",
                    "ファイルがまだありません。段階2を実行するか「アラジン加工計画読込→Excel出力」を実行してください。\n"
                            + latest);
            return;
        }
        try {
            DesktopFileOpener.openFileReadOnly(latest);
        } catch (Exception ex) {
            shell.showErrorDialog(
                    "最新を開く", "ファイルを開けませんでした。\n" + latest + "\n" + ex.getMessage());
        }
    }

    @FXML
    private void onOpenLocalLatestAladdinEntryWorkbookAction() {
        if (shell == null) {
            return;
        }
        Path latest = AppPaths.aladdinEntryDispatchPlanLocalXlsxPath(shell.snapshotUiEnv());
        if (!Files.isRegularFile(latest)) {
            shell.showWarningDialog(
                    "ローカル最新を開く",
                    "ファイルがまだありません。先に「ローカルへ出力」を実行してください。\n" + latest);
            return;
        }
        try {
            DesktopFileOpener.openFileReadOnly(latest);
        } catch (Exception ex) {
            shell.showErrorDialog(
                    "ローカル最新を開く",
                    "ファイルを開けませんでした。\n" + latest + "\n" + ex.getMessage());
        }
    }

    @FXML
    private void onOpenAladdinEntryGenerationsAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        DispatchAladdinEntryGenerationDialog.show(
                ownerStage, ui, DispatchAladdinEntryWorkbookExporter.currentOperatorDirName(ui));
    }

    @FXML
    private void onOpenLocalAladdinEntryGenerationsAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        DispatchAladdinEntryGenerationDialog.show(
                ownerStage,
                ui,
                DispatchAladdinEntryWorkbookExporter.currentOperatorDirName(ui),
                DispatchAladdinEntryWorkbookExporter.Destination.LOCAL);
    }

    void clearColumnFiltersAndSort() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(spreadsheetView);
    }

    @FXML
    private void onClearColumnFiltersAction() {
        clearColumnFiltersAndSort();
    }

    /** Snapshot of current shaped headers (after column-order permutation). Thread-safe defensive copy. */
    List<String> getShapedHeaders() {
        return new ArrayList<>(headersRef);
    }

    /** Snapshot of current shaped rows (after column-order permutation). Thread-safe defensive copy. */
    List<List<String>> getShapedRows() {
        List<List<String>> out = new ArrayList<>(rows.size());
        for (var r : rows) {
            out.add(new ArrayList<>(r));
        }
        return out;
    }
}
