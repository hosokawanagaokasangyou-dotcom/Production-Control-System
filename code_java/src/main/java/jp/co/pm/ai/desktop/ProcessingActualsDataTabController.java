package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Path;
import java.text.Collator;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.TreeSet;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.TaskInputSourceRawGridIo;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnDragReorderSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * Raw spreadsheet for the machining actual-detail workbook, resolved via {@link NetworkSourceDirResolver}
 * ({@link AppPaths#KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK} / {@link AppPaths#KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR}).
 * Display applies {@link TaskInputSourceRawGridIo#applyProcessingActualsDisplaySteps}. Optional sheet:
 * {@link AppPaths#KEY_PM_AI_ACTUAL_DETAIL_SHEET}. Rows can be filtered by combo selection for column
 * {@link #HEADER_PRODUCT_CONDITION_BREAKDOWN}. FXML: {@code ProcessingActualsDataTab.fxml}.
 */

public final class ProcessingActualsDataTabController {

    /** Header label after shaping (Excel row 5); must match the workbook column title. */
    private static final String HEADER_PRODUCT_CONDITION_BREAKDOWN =
            "\u88fd\u54c1\u6761\u4ef6(\u5185\u8a33)";

    /** Combo first row: no row filter (show full shaped table). */
    private static final String PRODUCT_CONDITION_FILTER_ALL = "\uff08\u5168\u884c\uff09";

    /** Combo entry matching rows whose cell in the column is blank. */
    private static final String PRODUCT_CONDITION_EMPTY_DISPLAY = "\uff08\u7a7a\u767d\uff09";

    private static final String HINT_TEXT =
            "\u5148\u982d4\u884c\u3092\u9664\u53bb\u3057\u3001\u539f\u7a3f\u306e5\u884c\u76ee\u3092\u5217\u898b\u51fa\u3057\u306b\u3057\u307e\u3059\u3002"
                    + " \u2462 \u30b3\u30f3\u30dc\u3067\u300c"
                    + HEADER_PRODUCT_CONDITION_BREAKDOWN
                    + "\u300d\u306e\u5024\u3092\u9078\u629e\u3057\u3001\u8868\u793a\u884c\u3092\u7d5e\u308a\u8fbc\u307f\u307e\u3059\u3002"
                    + " \u30c7\u30fc\u30bf\u306f PM_AI_ACTUAL_DETAIL_WORKBOOK \u307e\u305f\u306f"
                    + " PM_AI_ACTUAL_DETAIL_SOURCE_DIR \u304b\u3089\u89e3\u6c7a\u3055\u308c\u308b Excel\uff08\u307e\u305f\u306f CSV\uff09\u3092\u8aad\u307f\u8fbc\u307f\u307e\u3059\u3002"
                    + " PM_AI_ACTUAL_DETAIL_SHEET \u3067\u30b7\u30fc\u30c8\u540d\u3092\u6307\u5b9a\u3067\u304d\u307e\u3059\u3002"
                    + " \u30cd\u30c3\u30c8\u30ef\u30fc\u30af\u672a\u5230\u9054\u6642\u306f\u30ed\u30fc\u30ab\u30eb\u30ad\u30e3\u30c3\u30b7\u30e5\u3092\u8a66\u884c\u3057\u307e\u3059\u3002"
                    + " \u9078\u629e\u5024\u306f\u81ea\u52d5\u4fdd\u5b58\u3055\u308c\u3001\u6b21\u56de\u8d77\u52d5\u6642\u306b\u5fa9\u5143\u3055\u308c\u307e\u3059\u3002";


    @FXML
    private Button refreshButton;

    @FXML
    private Label statusLabel;

    @FXML
    private Label dirLabel;

    @FXML
    private Label pathLabel;

    @FXML
    private ComboBox<String> sheetCombo;

    @FXML
    private ComboBox<String> productConditionBreakdownFilterCombo;

    @FXML
    private Label hintLabel;

    @FXML
    private Slider rowHeightSlider;

    @FXML
    private Label rowHeightPctLabel;

    @FXML
    private CheckBox cellWrapCheck;

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

    private final AtomicReference<TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs> spreadsheetTabPrefs =
            new AtomicReference<>(TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs.defaults());

    private final AtomicBoolean suppressPresentationUiEvents = new AtomicBoolean(false);

    private final AtomicBoolean suppressSheetUi = new AtomicBoolean(false);

    private volatile Path loadedPath;

    private volatile boolean presentationHooksInstalled;

    private final AtomicBoolean suppressFilterUi = new AtomicBoolean(false);

    /** Full shaped grid before {@link #applyProductBreakdownFilter}; used when the filter selection changes. */
    private final List<String> unfilteredShapedHeaders = new ArrayList<>();

    private final List<List<String>> unfilteredShapedRows = new ArrayList<>();

    @FXML
    private void initialize() {
        hintLabel.setText(HINT_TEXT);
        rows = FXCollections.observableArrayList();

        StackPane.setAlignment(spreadsheetView, Pos.CENTER_LEFT);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        SpreadsheetThemeBridge.install(spreadsheetView);

        columnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::applyDynamicColumnWidths,
                                TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW,
                                headerColumnCount,
                                this::onLeadingColumnCountCommitted,
                                this::onReorderColumns,
                                () ->
                                        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                                                ownerStage,
                                                TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW,
                                                spreadsheetView,
                                                () -> new ArrayList<>(headersRef))));

        sheetCombo
                .getSelectionModel()
                .selectedIndexProperty()
                .addListener(
                        (obs, a, b) -> {
                            if (suppressSheetUi.get() || loadedPath == null) {
                                return;
                            }
                            if (!isExcelPath(loadedPath)) {
                                return;
                            }
                            int idx = sheetCombo.getSelectionModel().getSelectedIndex();
                            if (idx < 0) {
                                return;
                            }
                            Platform.runLater(() -> applyLoadedFile(loadedPath, idx, false));
                        });

        if (productConditionBreakdownFilterCombo != null) {
            productConditionBreakdownFilterCombo
                    .getSelectionModel()
                    .selectedItemProperty()
                    .addListener(
                            (o, oldV, newV) -> {
                                if (suppressFilterUi.get()) {
                                    return;
                                }
                                TableColumnOrderPersistence.saveProcessingActualsProductConditionBreakdownFilter(
                                        persistProductConditionSelection(newV));
                                Platform.runLater(this::refilterFromSnapshotIfPossible);
                            });
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        ownerStage = shell.getPrimaryStage();

        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                spreadsheetView,
                TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW,
                suppressColumnPersistence::get,
                () -> new ArrayList<>(headersRef));

        initSpreadsheetPresentationControls();

        Platform.runLater(this::reloadFromSourceDir);
    }

    private void onLeadingColumnCountCommitted(int n) {
        headerColumnCount.set(n);
        rebuildSpreadsheet();
    }

    private void initSpreadsheetPresentationControls() {
        if (presentationHooksInstalled) {
            return;
        }
        if (rowHeightSlider == null) {
            return;
        }
        presentationHooksInstalled = true;
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs loaded =
                TableColumnOrderPersistence.loadSpreadsheetTabPresentationPrefs(
                        TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW);
        spreadsheetTabPrefs.set(loaded);
        suppressPresentationUiEvents.set(true);
        try {
            rowHeightSlider.setMin(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN);
            rowHeightSlider.setMax(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX);
            rowHeightSlider.setValue(loaded.rowHeightPercent());
            rowHeightSlider.setMajorTickUnit(250);
            rowHeightSlider.setMinorTickCount(4);
            rowHeightSlider.setShowTickMarks(true);
            if (rowHeightPctLabel != null) {
                rowHeightPctLabel.setText(String.format("%.0f%%", loaded.rowHeightPercent()));
            }
            if (cellWrapCheck != null) {
                cellWrapCheck.setSelected(loaded.cellWrapText());
            }
        } finally {
            suppressPresentationUiEvents.set(false);
        }
        SliderCommittedChangeSupport.install(
                rowHeightSlider,
                () -> {
                    if (rowHeightPctLabel != null && rowHeightSlider != null) {
                        rowHeightPctLabel.setText(String.format("%.0f%%", rowHeightSlider.getValue()));
                    }
                },
                this::commitPresentationFromSlider);
        if (cellWrapCheck != null) {
            cellWrapCheck
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (suppressPresentationUiEvents.get()) {
                                    return;
                                }
                                commitPresentationFromUi();
                            });
        }
    }

    private void commitPresentationFromSlider() {
        if (suppressPresentationUiEvents.get()) {
            return;
        }
        commitPresentationFromUi();
    }

    private void commitPresentationFromUi() {
        if (rowHeightSlider == null) {
            return;
        }
        double v = rowHeightSlider.getValue();
        boolean wrap = cellWrapCheck != null && cellWrapCheck.isSelected();
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs next =
                new TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs(v, wrap);
        spreadsheetTabPrefs.set(next);
        TableColumnOrderPersistence.saveSpreadsheetTabPresentationPrefs(
                TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW, next);
        if (rowHeightPctLabel != null) {
            rowHeightPctLabel.setText(String.format("%.0f%%", v));
        }
        rebuildSpreadsheet();
    }

    private void onReorderColumns() {
        if (headersRef.isEmpty()) {
            if (shell != null) {
                shell.appendLog("[processing-actuals-detail] no columns; reload source.");
            }
            return;
        }
        SpreadsheetColumnReorderDialog.show(ownerStage, new ArrayList<>(headersRef))
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
                        TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW, oldHeaders.size());
        List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
        TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(oldHeaders, oldVis, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW, newVis);
        List<Double> widths =
                TableColumnOrderPersistence.resolveWidthsForHeaders(headersRef, lay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
        for (int i = 0; i < headersRef.size(); i++) {
            newLay.add(new TableColumnOrderPersistence.ColumnSpec(headersRef.get(i), widths.get(i)));
        }
        persistedLayout.set(newLay);
        TableColumnOrderPersistence.saveLayout(
                TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW, newLay);
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

            GridBase grid = SpreadsheetTabularSupport.buildReadOnlyPlainGrid(headersRef, rows);
            TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs pres = spreadsheetTabPrefs.get();
            SpreadsheetTabularSupport.applySpreadsheetGridRowHeightsAndWrap(
                    grid, pres.cellWrapText(), pres.rowHeightPercent());
            spreadsheetView.setGrid(grid);

            Platform.runLater(
                    () -> {
                        SpreadsheetTabularSupport.applyColumnWidths(spreadsheetView, widths, widthDefault);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(spreadsheetView);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                spreadsheetView, headerColumnCount.get());
                        SpreadsheetTabularSupport.applyColumnFilters(spreadsheetView);
                        SpreadsheetTabularSupport.refreshSpreadsheetAfterRowPresentationChange(spreadsheetView);
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
                                                TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW,
                                                headersRef.size()));
                    });
        } finally {
            suppressColumnPersistence.set(false);
        }
    }

    @FXML
    private void onRefreshButtonAction() {
        reloadFromSourceDir();
    }

    private void reloadFromSourceDir() {
        if (shell == null) {
            return;
        }
        refreshButton.setDisable(true);
        try {
            Map<String, String> ui = shell.snapshotUiEnv();
            Path dir = AppPaths.resolveActualDetailSourceDir(ui);
            dirLabel.setText(dir != null ? dir.toString() : "(\u672a\u8a2d\u5b9a)");
            NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
            Optional<Path> resolved = r.actualDetailPath();
            if (resolved.isEmpty()) {
                statusLabel.setText("\u30d5\u30a1\u30a4\u30eb\u306a\u3057\u307e\u305f\u306f\u53c2\u7167\u4e0d\u53ef");
                pathLabel.setText("");
                sheetCombo.setDisable(true);
                sheetCombo.getItems().clear();
                loadedPath = null;
                applyEmpty();
                return;
            }
            Path file = resolved.get().toAbsolutePath().normalize();
            loadedPath = file;
            pathLabel.setText(file.toString());

            String low = file.getFileName().toString().toLowerCase(Locale.ROOT);
            if (low.endsWith(".pq") || low.endsWith(".parquet")) {
                statusLabel.setText("Parquet \u672a\u5bfe\u5fdc");
                sheetCombo.setDisable(true);
                sheetCombo.getItems().clear();
                applyEmpty();
                return;
            }

            if (isExcelPath(file)) {
                suppressSheetUi.set(true);
                try {
                    List<String> names = TaskInputSourceRawGridIo.listExcelSheetNames(file);
                    sheetCombo.getItems().setAll(names);
                    sheetCombo.setDisable(names.isEmpty());
                    if (!names.isEmpty()) {
                        sheetCombo.getSelectionModel().select(0);
                        selectPreferredSheet(ui, names);
                    }
                } catch (IOException ex) {
                    statusLabel.setText("\u30b7\u30fc\u30c8\u4e00\u89a7\u30a8\u30e9\u30fc");
                    if (shell != null) {
                        shell.appendLog(
                                "[processing-actuals-detail] "
                                        + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
                    }
                    sheetCombo.setDisable(true);
                    sheetCombo.getItems().clear();
                    applyEmpty();
                    return;
                } finally {
                    suppressSheetUi.set(false);
                }
                applyLoadedFile(file, sheetCombo.getSelectionModel().getSelectedIndex(), true);
            } else {
                sheetCombo.setDisable(true);
                sheetCombo.getItems().clear();
                applyLoadedFile(file, 0, true);
            }
        } finally {
            refreshButton.setDisable(false);
        }
    }

    private void selectPreferredSheet(Map<String, String> ui, List<String> names) {
        if (sheetCombo == null || names == null || names.isEmpty() || ui == null) {
            return;
        }
        String want = ui.get(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SHEET);
        if (want != null) {
            want = want.strip();
        }
        if (want == null || want.isEmpty()) {
            return;
        }
        int ix = names.indexOf(want);
        if (ix >= 0) {
            sheetCombo.getSelectionModel().select(ix);
        }
    }

    private void applyLoadedFile(Path file, int excelSheetIndex, boolean showErrorsInStatus) {
        try {
            PlanInputTabularIo.TabularSheet shaped =
                    TaskInputSourceRawGridIo.applyProcessingActualsDisplaySteps(
                            TaskInputSourceRawGridIo.readRaw(file, excelSheetIndex));
            rememberShapedSnapshot(shaped);
            populateProductConditionFilterChoices(shaped);
            PlanInputTabularIo.TabularSheet tab = applyProductBreakdownFilter(shaped);
            populateFromFilteredSheet(tab);
        } catch (Exception ex) {
            if (showErrorsInStatus) {
                statusLabel.setText("\u8aad\u8fbc\u30a8\u30e9\u30fc");
            }
            if (shell != null) {
                shell.appendLog(
                        "[processing-actuals-detail] "
                                + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            }
            applyEmpty();
        }
    }

    /**
     * Fills the product-condition combo from distinct values in the shaped sheet and restores the last
     * persisted selection when still present.
     */
    private void populateProductConditionFilterChoices(PlanInputTabularIo.TabularSheet shaped) {
        if (productConditionBreakdownFilterCombo == null || shaped == null) {
            return;
        }
        ObservableList<String> items = productConditionBreakdownFilterCombo.getItems();
        suppressFilterUi.set(true);
        try {
            items.clear();
            items.add(PRODUCT_CONDITION_FILTER_ALL);
            int col = indexOfProductBreakdownColumn(shaped.headers());
            if (col >= 0) {
                Collator collator = Collator.getInstance(Locale.JAPANESE);
                TreeSet<String> sorted = new TreeSet<>(collator);
                boolean anyEmpty = false;
                for (List<String> row : shaped.rows()) {
                    String cell =
                            col < row.size() && row.get(col) != null ? row.get(col).strip() : "";
                    if (cell.isEmpty()) {
                        anyEmpty = true;
                    } else {
                        sorted.add(cell);
                    }
                }
                items.addAll(sorted);
                if (anyEmpty) {
                    items.add(PRODUCT_CONDITION_EMPTY_DISPLAY);
                }
            }
            String saved =
                    TableColumnOrderPersistence.loadProcessingActualsProductConditionBreakdownFilter();
            if (saved != null && !saved.isEmpty() && items.contains(saved)) {
                productConditionBreakdownFilterCombo.getSelectionModel().select(saved);
            } else {
                productConditionBreakdownFilterCombo.getSelectionModel().selectFirst();
            }
        } finally {
            suppressFilterUi.set(false);
        }
    }

    /** Stored string matches combo items (empty string = {@link #PRODUCT_CONDITION_FILTER_ALL}). */
    private static String persistProductConditionSelection(String selectedItem) {
        if (selectedItem == null || PRODUCT_CONDITION_FILTER_ALL.equals(selectedItem)) {
            return "";
        }
        return selectedItem;
    }

    private void rememberShapedSnapshot(PlanInputTabularIo.TabularSheet shaped) {
        unfilteredShapedHeaders.clear();
        unfilteredShapedRows.clear();
        unfilteredShapedHeaders.addAll(shaped.headers());
        for (List<String> r : shaped.rows()) {
            unfilteredShapedRows.add(new ArrayList<>(r));
        }
    }

    private void clearShapedSnapshot() {
        unfilteredShapedHeaders.clear();
        unfilteredShapedRows.clear();
    }

    /**
     * Keeps rows whose {@link #HEADER_PRODUCT_CONDITION_BREAKDOWN} cell equals the combo selection.
     * {@link #PRODUCT_CONDITION_FILTER_ALL} leaves all rows. Unknown column: no filtering (with log).
     */
    private PlanInputTabularIo.TabularSheet applyProductBreakdownFilter(
            PlanInputTabularIo.TabularSheet shaped) {
        String sel =
                productConditionBreakdownFilterCombo != null
                        ? productConditionBreakdownFilterCombo.getSelectionModel().getSelectedItem()
                        : null;
        if (sel == null || PRODUCT_CONDITION_FILTER_ALL.equals(sel)) {
            return shaped;
        }
        String want = PRODUCT_CONDITION_EMPTY_DISPLAY.equals(sel) ? "" : sel;
        List<String> headers = shaped.headers();
        int col = indexOfProductBreakdownColumn(headers);
        if (col < 0) {
            if (shell != null) {
                shell.appendLog(
                        "[processing-actuals-detail] missing header column: "
                                + HEADER_PRODUCT_CONDITION_BREAKDOWN);
            }
            return shaped;
        }
        List<List<String>> out = new ArrayList<>();
        for (List<String> row : shaped.rows()) {
            String cell = col < row.size() && row.get(col) != null ? row.get(col).strip() : "";
            if (want.equals(cell)) {
                out.add(new ArrayList<>(row));
            }
        }
        return new PlanInputTabularIo.TabularSheet(new ArrayList<>(headers), out);
    }

    private static int indexOfProductBreakdownColumn(List<String> headers) {
        if (headers == null) {
            return -1;
        }
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (HEADER_PRODUCT_CONDITION_BREAKDOWN.equals(h != null ? h.strip() : "")) {
                return i;
            }
        }
        return -1;
    }

    private void refilterFromSnapshotIfPossible() {
        if (unfilteredShapedHeaders.isEmpty()) {
            return;
        }
        List<List<String>> copyRows = new ArrayList<>();
        for (List<String> r : unfilteredShapedRows) {
            copyRows.add(new ArrayList<>(r));
        }
        PlanInputTabularIo.TabularSheet shaped =
                new PlanInputTabularIo.TabularSheet(
                        new ArrayList<>(unfilteredShapedHeaders), copyRows);
        PlanInputTabularIo.TabularSheet tab = applyProductBreakdownFilter(shaped);
        populateFromFilteredSheet(tab);
    }

    private void populateFromFilteredSheet(PlanInputTabularIo.TabularSheet tab) {
        List<TableColumnOrderPersistence.ColumnSpec> lay =
                TableColumnOrderPersistence.loadLayout(
                        TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW);
        persistedLayout.set(lay);
        List<String> beforeHeaders = new ArrayList<>(tab.headers());
        boolean[] visBefore =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW,
                        beforeHeaders.size());
        List<String> titleOrder =
                lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();

        headersRef.clear();
        headersRef.addAll(tab.headers());
        rows.clear();
        for (List<String> r : tab.rows()) {
            rows.add(FXCollections.observableArrayList(r));
        }

        TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
        boolean[] visAfter =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        beforeHeaders, visBefore, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW, visAfter);

        statusLabel.setText(
                rows.size()
                        + " \u884c \u00d7 "
                        + headersRef.size()
                        + " \u5217"
                        + filterActiveSuffix());
        rebuildSpreadsheet();
    }

    private String filterActiveSuffix() {
        String sel =
                productConditionBreakdownFilterCombo != null
                        ? productConditionBreakdownFilterCombo.getSelectionModel().getSelectedItem()
                        : null;
        if (sel == null || PRODUCT_CONDITION_FILTER_ALL.equals(sel)) {
            return "";
        }
        return " [\u7d5e\u308a]";
    }

    private void applyEmpty() {
        suppressFilterUi.set(true);
        try {
            if (productConditionBreakdownFilterCombo != null) {
                productConditionBreakdownFilterCombo.getItems().clear();
                productConditionBreakdownFilterCombo.getItems().add(PRODUCT_CONDITION_FILTER_ALL);
                productConditionBreakdownFilterCombo.getSelectionModel().selectFirst();
            }
        } finally {
            suppressFilterUi.set(false);
        }
        clearShapedSnapshot();
        headersRef.clear();
        rows.clear();
        persistedLayout.set(List.of());
        spreadsheetView.setGrid(new GridBase(0, 0));
    }

    void clearColumnFiltersAndSort() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(spreadsheetView);
    }

    @FXML
    private void onClearColumnFiltersAction() {
        clearColumnFiltersAndSort();
    }

    private static boolean isExcelPath(Path p) {
        if (p == null || p.getFileName() == null) {
            return false;
        }
        String low = p.getFileName().toString().toLowerCase(Locale.ROOT);
        return low.endsWith(".xlsx")
                || low.endsWith(".xlsm")
                || low.endsWith(".xltx")
                || low.endsWith(".xltm");
    }
}
