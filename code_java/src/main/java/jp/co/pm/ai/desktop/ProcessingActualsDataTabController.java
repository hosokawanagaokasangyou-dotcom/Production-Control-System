package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Path;
import java.text.Collator;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.TreeSet;
import java.util.concurrent.CancellationException;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
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
import jp.co.pm.ai.desktop.io.JsonTableIo;
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
 * Shaped JSON on disk uses {@link #projectShapedForJsonExport}（表示列・先頭見出し列・必須見出し）and optional {@code columns_all}.
 * Display applies {@link TaskInputSourceRawGridIo#applyProcessingActualsDisplaySteps}（検査NO行より前を削除、
 * マーカー欠落時は先頭4行） then
 * {@link TaskInputSourceRawGridIo#applyProcessingActualsDateTimeColumns} ({@link #shapeLoaded} と同一)。
 * その後コントローラ側では（チャット履歴で確定した順）
 * 製造条件(内訳)コンボによる行フィルタを先に適用し、続けて
 * {@link TaskInputSourceRawGridIo#applyProcessingActualsDedupeByQuadKey} を適用してからグリッドへ載せる。
 * メモリ／JSON のスナップショットはコンボ・重複除去の前（日時列まで済んだ shaped）を保持する。
 * Optional sheet:
 * {@link AppPaths#KEY_PM_AI_ACTUAL_DETAIL_SHEET}. Rows can be filtered by combo selection for column
 * {@link #HEADER_MANUFACTURING_CONDITION_BREAKDOWN}. FXML: {@code ProcessingActualsDataTab.fxml}.
 */

public final class ProcessingActualsDataTabController {

    /** Header label after shaping (Excel row 5); must match the workbook column title. */
    private static final String HEADER_MANUFACTURING_CONDITION_BREAKDOWN =
            "\u88fd\u9020\u6761\u4ef6(\u5185\u8a33)";

    /** Combo first row: no row filter (show full shaped table). */
    private static final String MANUFACTURING_CONDITION_FILTER_ALL = "\uff08\u5168\u884c\uff09";

    /** Combo entry matching rows whose cell in the column is blank. */
    private static final String MANUFACTURING_CONDITION_EMPTY_DISPLAY = "\uff08\u7a7a\u767d\uff09";

    /**
     * キャッシュ JSON に必ず含める見出し（納期オーバーレイ・配台重複除去などが参照）。投影後も欠けないようにする。
     */
    private static final Set<String> PROCESSING_ACTUALS_JSON_REQUIRED_HEADERS =
            Set.of(
                    "\u5de5\u7a0b\u540d",
                    "\u6a5f\u68b0\u540d",
                    "\u4f9d\u983cNO",
                    "\u4f9d\u983c\uff2e\uff2f",
                    "\u52a0\u5de5\u65e5",
                    "\u5b9f\u52a0\u5de5\u6570",
                    "\u88fd\u9020\u6761\u4ef6(\u5185\u8a33)",
                    "\u52a0\u5de5\u958b\u59cb\u65e5\u6642",
                    "\u52a0\u5de5\u7d42\u4e86\u65e5\u6642");

    private static final String HINT_TEXT =
            "シート先頭から、いずれかのセルに「検査NO」または「検査ＮＯ」が含まれる行（列見出し行）の直前までを除去し、その行を列見出しとします。"
                    + " 当該文字列が見つからない場合は従来どおり先頭4行を除去します。"
                    + " \u2462 \u30b3\u30f3\u30dc\u3067\u300c"
                    + HEADER_MANUFACTURING_CONDITION_BREAKDOWN
                    + "\u300d\u306e\u5024\u3092\u9078\u629e\u3057\u3001\u8868\u793a\u884c\u3092\u7d5e\u308a\u8fbc\u307f\u307e\u3059\u3002"
                    + " \u30c7\u30fc\u30bf\u306f PM_AI_ACTUAL_DETAIL_WORKBOOK \u307e\u305f\u306f"
                    + " PM_AI_ACTUAL_DETAIL_SOURCE_DIR \u304b\u3089\u89e3\u6c7a\u3055\u308c\u308b Excel\uff08\u307e\u305f\u306f CSV\uff09\u3092\u8aad\u307f\u8fbc\u307f\u307e\u3059\u3002"
                    + " PM_AI_ACTUAL_DETAIL_SHEET \u3067\u30b7\u30fc\u30c8\u540d\u3092\u6307\u5b9a\u3067\u304d\u307e\u3059\u3002"
                    + " \u30cd\u30c3\u30c8\u30ef\u30fc\u30af\u672a\u5230\u9054\u6642\u306f\u30ed\u30fc\u30ab\u30eb\u30ad\u30e3\u30c3\u30b7\u30e5\u3092\u8a66\u884c\u3057\u307e\u3059\u3002"
                    + " \u9078\u629e\u5024\u306f\u81ea\u52d5\u4fdd\u5b58\u3055\u308c\u3001\u6b21\u56de\u8d77\u52d5\u6642\u306b\u5fa9\u5143\u3055\u308c\u307e\u3059\u3002"
                    + " \u52a0\u5de5\u65e5\u30fb\u958b\u59cb/\u7d42\u4e86\u306e\u6642\u5206\u304b\u3089"
                    + "\u52a0\u5de5\u958b\u59cb\u65e5\u6642\u30fb\u52a0\u5de5\u7d42\u4e86\u65e5\u6642\u5217\u3092\u4ed8\u52a0\u3057\u307e\u3059\u3002"
                    + " \u7d9e\u308a\u8fbc\u307f\u5f8c\u3001\u540c\u4e00\u5de5\u7a0b\u540d\u30fb\u6a5f\u68b0\u540d\u30fb\u4f9d\u983cNO\uff08\u307e\u305f\u306f\u4f9d\u983c\uff2e\uff2f\uff09\u30fb\u52a0\u5de5\u65e5\u306e\u884c\u306f"
                    + "\u5148\u982d\u884c\u306e\u307f\u6b8b\u3057\u307e\u3059\u3002"
                    + " PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES で元ファイルの読込上限（バイト、既定 20MiB、0 で無制限）を変更できます。";

    /** {@link ColumnVisibilityDialog} / JSON と同一の必須見出しに対応するマスク（表示は強制）。 */
    private boolean[] mandatoryVisibilityMaskForHeaders(List<String> headers) {
        if (headers == null) {
            return null;
        }
        boolean[] m = new boolean[headers.size()];
        for (int i = 0; i < headers.size(); i++) {
            String t = headers.get(i) != null ? headers.get(i).strip() : "";
            m[i] = PROCESSING_ACTUALS_JSON_REQUIRED_HEADERS.contains(t);
        }
        return m;
    }

    @FXML
    private Label statusLabel;

    @FXML
    private ProgressBar loadProgressBar;

    @FXML
    private Label dirLabel;

    @FXML
    private Label pathLabel;

    @FXML
    private ComboBox<String> sheetCombo;

    @FXML
    private ComboBox<String> manufacturingConditionBreakdownFilterCombo;

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

    /** 重複した古い {@link Task} の結果を UI に適用しないための世代番号。 */
    private final AtomicLong reloadGeneration = new AtomicLong();

    private volatile Task<?> activeReloadTask;

    private volatile boolean presentationHooksInstalled;

    private final AtomicBoolean suppressFilterUi = new AtomicBoolean(false);

    /**
     * Grid after datetime shaping and before manufacturing filter / dedupe; used when the combo selection
     * changes.
     */
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
                                                () -> new ArrayList<>(headersRef),
                                                mandatoryVisibilityMaskForHeaders(
                                                        new ArrayList<>(headersRef)))));

        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                spreadsheetHost, headerColumnCount::get);

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
                            reloadSheetFromDiskAsync(false);
                        });

        if (manufacturingConditionBreakdownFilterCombo != null) {
            manufacturingConditionBreakdownFilterCombo
                    .getSelectionModel()
                    .selectedItemProperty()
                    .addListener(
                            (o, oldV, newV) -> {
                                if (suppressFilterUi.get()) {
                                    return;
                                }
                                TableColumnOrderPersistence.saveProcessingActualsManufacturingConditionBreakdownFilter(
                                        persistManufacturingConditionSelection(newV));
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

        /*
         * 起動時は xlsx を開かない（ヒープ・OOM 回避）。読込は納期管理ビュー上部の「再読込」成功時、
         * または {@link #reloadProcessingActualsFromDisk} が親から呼ばれたときのみ。
         */
        refreshSourcePathLabelsOnly();
    }

    /**
     * 納期管理ビューで「再読込」が成功したあと、加工実績ブックを読み込むために親から呼ばれる。
     */
    public void reloadProcessingActualsFromDisk() {
        reloadFromSourceDir();
    }

    /** POI でブックを開かず、解決済みパスと案内文言だけ更新する（起動時・セッション復元直後）。 */
    private void refreshSourcePathLabelsOnly() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path dir = AppPaths.resolveActualDetailSourceDir(ui);
        dirLabel.setText(dir != null ? dir.toString() : "(\u672a\u8a2d\u5b9a)");
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
        Optional<Path> resolved = r.actualDetailPath();
        if (resolved.isEmpty()) {
            statusLabel.setText(
                    "\u30d5\u30a1\u30a4\u30eb\u306a\u3057\u307e\u305f\u306f\u53c2\u7167\u4e0d\u53ef");
            pathLabel.setText("");
            loadedPath = null;
            applyEmpty();
            return;
        }
        Path file = resolved.get().toAbsolutePath().normalize();
        pathLabel.setText(file.toString());
        loadedPath = null;
        statusLabel.setText("未読込 — 納期管理ビュー上部の「再読込」で Excel を読み込みます");
        sheetCombo.getItems().clear();
        sheetCombo.setDisable(true);
        applyEmpty();
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
        oldVis =
                ColumnVisibilitySupport.mergeMandatoryIntoVisibility(
                        oldVis, mandatoryVisibilityMaskForHeaders(oldHeaders));
        List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
        TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(oldHeaders, oldVis, titleOrder);
        newVis =
                ColumnVisibilitySupport.mergeMandatoryIntoVisibility(
                        newVis, mandatoryVisibilityMaskForHeaders(headersRef));
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
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                spreadsheetView, headerColumnCount.get());
                        SpreadsheetTabularSupport.applyColumnFilters(spreadsheetView);
                        SpreadsheetTabularSupport.pinSpreadsheetFilterRow(spreadsheetView);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(spreadsheetView);
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
                                        ColumnVisibilitySupport.mergeMandatoryIntoVisibility(
                                                TableColumnOrderPersistence.loadColumnVisibility(
                                                        TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW,
                                                        headersRef.size()),
                                                mandatoryVisibilityMaskForHeaders(
                                                        new ArrayList<>(headersRef))));
                    });
        } finally {
            suppressColumnPersistence.set(false);
        }
    }

    /** Excel / CSV の読込・成形は POI が重いためワーカースレッドで実行する。 */
    private record ActualsReloadPayload(
            boolean excel,
            List<String> sheetNames,
            int selectedSheetIndex,
            PlanInputTabularIo.TabularSheet shaped) {}

    /** Display steps + datetime columns after {@link TaskInputSourceRawGridIo#readRaw(Path, int)}. */
    private static PlanInputTabularIo.TabularSheet applyActualsShapingAfterRaw(
            PlanInputTabularIo.TabularSheet raw) {
        PlanInputTabularIo.TabularSheet stepped =
                TaskInputSourceRawGridIo.applyProcessingActualsDisplaySteps(raw);
        return TaskInputSourceRawGridIo.applyProcessingActualsDateTimeColumns(stepped);
    }

    private static PlanInputTabularIo.TabularSheet shapeLoaded(Path file, int excelSheetIndex)
            throws IOException {
        PlanInputTabularIo.TabularSheet raw =
                TaskInputSourceRawGridIo.readRaw(file, excelSheetIndex);
        return applyActualsShapingAfterRaw(raw);
    }

    private static final long PROGRESS_MAX = 10_000L;

    private static int preferredSheetIndex(List<String> names, Map<String, String> ui) {
        if (names == null || names.isEmpty()) {
            return 0;
        }
        String want = ui != null ? ui.get(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SHEET) : null;
        if (want != null) {
            want = want.strip();
        }
        if (want == null || want.isEmpty()) {
            return 0;
        }
        int ix = names.indexOf(want);
        return ix >= 0 ? ix : 0;
    }

    private void bindLoadProgressToTask(Task<?> task) {
        if (loadProgressBar == null || task == null) {
            return;
        }
        loadProgressBar.setManaged(true);
        loadProgressBar.setVisible(true);
        loadProgressBar.progressProperty().bind(task.progressProperty());
    }

    private void unbindLoadProgress() {
        if (loadProgressBar == null) {
            return;
        }
        loadProgressBar.progressProperty().unbind();
        loadProgressBar.setProgress(0);
        loadProgressBar.setVisible(false);
        loadProgressBar.setManaged(false);
    }

    private void reloadFromSourceDir() {
        if (shell == null) {
            return;
        }
        if (activeReloadTask != null) {
            activeReloadTask.cancel(true);
        }
        final long gen = reloadGeneration.incrementAndGet();
        final Map<String, String> uiSnap = new HashMap<>(shell.snapshotUiEnv());

        Path dir = AppPaths.resolveActualDetailSourceDir(uiSnap);
        dirLabel.setText(dir != null ? dir.toString() : "(\u672a\u8a2d\u5b9a)");
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(uiSnap);
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
        statusLabel.setText("\u8aad\u8fbc\u4e2d\u2026");

        String low = file.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".pq") || low.endsWith(".parquet")) {
            statusLabel.setText("Parquet \u672a\u5bfe\u5fdc");
            sheetCombo.setDisable(true);
            sheetCombo.getItems().clear();
            applyEmpty();
            return;
        }

        try {
            AppPaths.ensureActualDetailRawFileWithinLimit(file, uiSnap);
        } catch (IOException ex) {
            statusLabel.setText(
                    ex.getMessage() != null && !ex.getMessage().isBlank()
                            ? ex.getMessage()
                            : "加工実績の元データが上限を超えているため読込を中止しました");
            if (shell != null) {
                shell.appendLog("[processing-actuals-detail] " + ex.getMessage());
            }
            sheetCombo.setDisable(true);
            sheetCombo.getItems().clear();
            loadedPath = null;
            applyEmpty();
            return;
        }

        Task<ActualsReloadPayload> task =
                new Task<>() {
                    @Override
                    protected ActualsReloadPayload call() throws Exception {
                        updateProgress(0, PROGRESS_MAX);
                        if (isCancelled()) {
                            return null;
                        }
                        if (isExcelPath(file)) {
                            updateProgress(100, PROGRESS_MAX);
                            List<String> names = TaskInputSourceRawGridIo.listExcelSheetNames(file);
                            if (names.isEmpty()) {
                                throw new IOException("Excel \u306b\u30b7\u30fc\u30c8\u304c\u3042\u308a\u307e\u305b\u3093");
                            }
                            int idx = preferredSheetIndex(names, uiSnap);
                            if (idx >= names.size()) {
                                idx = 0;
                            }
                            updateProgress(200, PROGRESS_MAX);
                            PlanInputTabularIo.TabularSheet raw =
                                    TaskInputSourceRawGridIo.readRaw(
                                            file,
                                            idx,
                                            p -> {
                                                if (isCancelled()) {
                                                    throw new CancellationException();
                                                }
                                                long w =
                                                        200L
                                                                + (long)
                                                                        Math.min(
                                                                                7400L,
                                                                                Math.round(p * 7400d));
                                                updateProgress(w, PROGRESS_MAX);
                                            });
                            if (isCancelled()) {
                                return null;
                            }
                            updateProgress(7700, PROGRESS_MAX);
                            PlanInputTabularIo.TabularSheet stepped =
                                    TaskInputSourceRawGridIo.applyProcessingActualsDisplaySteps(raw);
                            updateProgress(8600, PROGRESS_MAX);
                            PlanInputTabularIo.TabularSheet shaped =
                                    TaskInputSourceRawGridIo.applyProcessingActualsDateTimeColumns(stepped);
                            updateProgress(PROGRESS_MAX, PROGRESS_MAX);
                            return new ActualsReloadPayload(true, names, idx, shaped);
                        }
                        PlanInputTabularIo.TabularSheet raw =
                                TaskInputSourceRawGridIo.readRaw(
                                        file,
                                        0,
                                        p -> {
                                            if (isCancelled()) {
                                                throw new CancellationException();
                                            }
                                            long w =
                                                    (long)
                                                            Math.min(
                                                                    8200L,
                                                                    Math.round(p * 8200d));
                                            updateProgress(w, PROGRESS_MAX);
                                        });
                        if (isCancelled()) {
                            return null;
                        }
                        updateProgress(8300, PROGRESS_MAX);
                        PlanInputTabularIo.TabularSheet stepped =
                                TaskInputSourceRawGridIo.applyProcessingActualsDisplaySteps(raw);
                        updateProgress(9100, PROGRESS_MAX);
                        PlanInputTabularIo.TabularSheet shaped =
                                TaskInputSourceRawGridIo.applyProcessingActualsDateTimeColumns(stepped);
                        updateProgress(PROGRESS_MAX, PROGRESS_MAX);
                        return new ActualsReloadPayload(false, List.of(), 0, shaped);
                    }
                };

        activeReloadTask = task;
        bindLoadProgressToTask(task);
        task.setOnSucceeded(
                ev -> {
                    activeReloadTask = null;
                    try {
                        if (gen != reloadGeneration.get()) {
                            return;
                        }
                        ActualsReloadPayload p = task.getValue();
                        if (p == null) {
                            return;
                        }
                        applyReloadPayloadOnFx(p, true);
                    } finally {
                        unbindLoadProgress();
                    }
                });
        task.setOnFailed(
                ev -> {
                    activeReloadTask = null;
                    try {
                        if (gen != reloadGeneration.get()) {
                            return;
                        }
                        Throwable ex = task.getException();
                        statusLabel.setText("\u8aad\u8fbc\u30a8\u30e9\u30fc");
                        if (shell != null) {
                            shell.appendLog(
                                    "[processing-actuals-detail] "
                                            + (ex != null && ex.getMessage() != null
                                                    ? ex.getMessage()
                                                    : String.valueOf(ex)));
                        }
                        applyEmpty();
                    } finally {
                        unbindLoadProgress();
                    }
                });
        task.setOnCancelled(
                ev -> {
                    activeReloadTask = null;
                    unbindLoadProgress();
                });
        new Thread(task, "processing-actuals-reload").start();
    }

    private void applyReloadPayloadOnFx(ActualsReloadPayload p, boolean showErrorsInStatus) {
        if (p.excel()) {
            suppressSheetUi.set(true);
            try {
                sheetCombo.getItems().setAll(p.sheetNames());
                sheetCombo.setDisable(p.sheetNames().isEmpty());
                if (!p.sheetNames().isEmpty()) {
                    sheetCombo.getSelectionModel().select(p.selectedSheetIndex());
                }
            } finally {
                suppressSheetUi.set(false);
            }
        } else {
            sheetCombo.setDisable(true);
            sheetCombo.getItems().clear();
        }
        applyShapedToUi(p.shaped(), showErrorsInStatus);
    }

    /** シートタブ変更時: 選択インデックスでブックを再読込（ワーカースレッド）。 */
    private void reloadSheetFromDiskAsync(boolean showErrorsInStatus) {
        Path path = loadedPath;
        if (shell == null || path == null || !isExcelPath(path)) {
            return;
        }
        int idx = sheetCombo.getSelectionModel().getSelectedIndex();
        if (idx < 0) {
            return;
        }
        Map<String, String> uiSnap = new HashMap<>(shell.snapshotUiEnv());
        try {
            AppPaths.ensureActualDetailRawFileWithinLimit(path, uiSnap);
        } catch (IOException ex) {
            if (showErrorsInStatus) {
                statusLabel.setText(
                        ex.getMessage() != null && !ex.getMessage().isBlank()
                                ? ex.getMessage()
                                : "加工実績の元データが上限を超えているため読込を中止しました");
            }
            if (shell != null) {
                shell.appendLog("[processing-actuals-detail] " + ex.getMessage());
            }
            applyEmpty();
            return;
        }
        if (activeReloadTask != null) {
            activeReloadTask.cancel(true);
        }
        final long gen = reloadGeneration.incrementAndGet();
        Task<PlanInputTabularIo.TabularSheet> task =
                new Task<>() {
                    @Override
                    protected PlanInputTabularIo.TabularSheet call() throws Exception {
                        updateProgress(0, PROGRESS_MAX);
                        PlanInputTabularIo.TabularSheet raw =
                                TaskInputSourceRawGridIo.readRaw(
                                        path,
                                        idx,
                                        p -> {
                                            if (isCancelled()) {
                                                throw new CancellationException();
                                            }
                                            long w =
                                                    150L
                                                            + (long)
                                                                    Math.min(
                                                                            8050L,
                                                                            Math.round(p * 8050d));
                                            updateProgress(w, PROGRESS_MAX);
                                        });
                        if (isCancelled()) {
                            return null;
                        }
                        updateProgress(8300, PROGRESS_MAX);
                        PlanInputTabularIo.TabularSheet stepped =
                                TaskInputSourceRawGridIo.applyProcessingActualsDisplaySteps(raw);
                        updateProgress(9100, PROGRESS_MAX);
                        PlanInputTabularIo.TabularSheet shaped =
                                TaskInputSourceRawGridIo.applyProcessingActualsDateTimeColumns(stepped);
                        updateProgress(PROGRESS_MAX, PROGRESS_MAX);
                        return shaped;
                    }
                };
        activeReloadTask = task;
        bindLoadProgressToTask(task);
        task.setOnSucceeded(
                ev -> {
                    activeReloadTask = null;
                    try {
                        if (gen != reloadGeneration.get()) {
                            return;
                        }
                        PlanInputTabularIo.TabularSheet shaped = task.getValue();
                        applyShapedToUi(shaped, showErrorsInStatus);
                    } finally {
                        unbindLoadProgress();
                    }
                });
        task.setOnFailed(
                ev -> {
                    activeReloadTask = null;
                    try {
                        if (gen != reloadGeneration.get()) {
                            return;
                        }
                        Throwable ex = task.getException();
                        if (showErrorsInStatus) {
                            statusLabel.setText("\u8aad\u8fbc\u30a8\u30e9\u30fc");
                        }
                        if (shell != null) {
                            shell.appendLog(
                                    "[processing-actuals-detail] "
                                            + (ex != null && ex.getMessage() != null
                                                    ? ex.getMessage()
                                                    : String.valueOf(ex)));
                        }
                        applyEmpty();
                    } finally {
                        unbindLoadProgress();
                    }
                });
        task.setOnCancelled(
                ev -> {
                    activeReloadTask = null;
                    unbindLoadProgress();
                });
        new Thread(task, "processing-actuals-sheet").start();
    }

    /**
     * キャッシュ用 JSON 向けに shaped を列投影する。含める列: 列表示 ON、または見出し列（先頭 {@link #headerColumnCount}
     * 列）、またはオーバーレイ等が参照する必須見出し。順序は現在の {@link #headersRef} に合わせる。
     */
    private PlanInputTabularIo.TabularSheet projectShapedForJsonExport(PlanInputTabularIo.TabularSheet shaped) {
        if (shaped == null || headersRef.isEmpty()) {
            return shaped;
        }
        List<String> shHeaders = shaped.headers();
        Map<String, Integer> titleToFirstIndex = new HashMap<>();
        for (int i = 0; i < shHeaders.size(); i++) {
            String key = shHeaders.get(i) != null ? shHeaders.get(i).strip() : "";
            titleToFirstIndex.putIfAbsent(key, i);
        }
        boolean[] vis =
                ColumnVisibilitySupport.mergeMandatoryIntoVisibility(
                        TableColumnOrderPersistence.loadColumnVisibility(
                                TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW,
                                headersRef.size()),
                        mandatoryVisibilityMaskForHeaders(headersRef));
        int lead = Math.max(0, headerColumnCount.get());
        List<String> outTitles = new ArrayList<>();
        for (int i = 0; i < headersRef.size(); i++) {
            String title = headersRef.get(i) != null ? headersRef.get(i).strip() : "";
            boolean visible = i < vis.length && vis[i];
            boolean headingCol = i < lead;
            boolean required = PROCESSING_ACTUALS_JSON_REQUIRED_HEADERS.contains(title);
            if (visible || headingCol || required) {
                outTitles.add(headersRef.get(i));
            }
        }
        if (outTitles.isEmpty()) {
            return shaped;
        }
        List<List<String>> outRows = new ArrayList<>(shaped.rows().size());
        for (List<String> row : shaped.rows()) {
            List<String> line = new ArrayList<>(outTitles.size());
            for (String colTitle : outTitles) {
                String k = colTitle != null ? colTitle.strip() : "";
                Integer ix = titleToFirstIndex.get(k);
                if (ix == null) {
                    line.add("");
                } else {
                    line.add(ix < row.size() && row.get(ix) != null ? row.get(ix) : "");
                }
            }
            outRows.add(line);
        }
        return new PlanInputTabularIo.TabularSheet(outTitles, outRows);
    }

    private void applyShapedToUi(
            PlanInputTabularIo.TabularSheet shaped, boolean showErrorsInStatus) {
        try {
            // 1) shapeLoaded 済み（検査NO行より前を除去、またはフォールバックで先頭4行除去 → ヘッダ行・日時列）。コンボ／四キー除去は未適用。
            rememberShapedSnapshot(shaped);
            populateManufacturingConditionFilterChoices(shaped);
            // 2) 製造条件(内訳)コンボ → 3) 工程名・機械名・依頼NO・加工日の四キー重複は先頭行のみ（フィルタ後集合に対して）
            PlanInputTabularIo.TabularSheet tab = applyManufacturingFilterThenQuadDedupe(shaped);
            populateFromFilteredSheet(tab);
            if (shell != null && !headersRef.isEmpty()) {
                try {
                    java.nio.file.Path savePath =
                            AppPaths.resolveShapedProcessingActualsJsonPath(shell.snapshotUiEnv());
                    PlanInputTabularIo.TabularSheet projected = projectShapedForJsonExport(shaped);
                    JsonTableIo.saveArrayTable(
                            savePath,
                            projected.headers(),
                            projected.rows(),
                            shaped.headers());
                } catch (Exception saveEx) {
                    shell.appendLog(
                            "[processing-actuals-detail] shaped JSON save failed: "
                                    + saveEx.getMessage());
                }
            }
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
     * Applies manufacturing-condition combo filtering first, then quad-key dedupe on the remaining rows.
     * Order（チャット履歴「4.と5.の順番を入れ替え」）: フィルタを先、重複除去を後。
     */
    private PlanInputTabularIo.TabularSheet applyManufacturingFilterThenQuadDedupe(
            PlanInputTabularIo.TabularSheet shaped) {
        PlanInputTabularIo.TabularSheet filtered = applyManufacturingConditionFilter(shaped);
        return TaskInputSourceRawGridIo.applyProcessingActualsDedupeByQuadKey(filtered);
    }

    /**
     * Fills the product-condition combo from distinct values in the shaped sheet and restores the last
     * persisted selection when still present.
     */
    private void populateManufacturingConditionFilterChoices(PlanInputTabularIo.TabularSheet shaped) {
        if (manufacturingConditionBreakdownFilterCombo == null || shaped == null) {
            return;
        }
        ObservableList<String> items = manufacturingConditionBreakdownFilterCombo.getItems();
        suppressFilterUi.set(true);
        try {
            items.clear();
            items.add(MANUFACTURING_CONDITION_FILTER_ALL);
            int col = indexOfManufacturingConditionColumn(shaped.headers());
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
                    items.add(MANUFACTURING_CONDITION_EMPTY_DISPLAY);
                }
            }
            String saved =
                    TableColumnOrderPersistence.loadProcessingActualsManufacturingConditionBreakdownFilter();
            if (saved != null && !saved.isEmpty() && items.contains(saved)) {
                manufacturingConditionBreakdownFilterCombo.getSelectionModel().select(saved);
            } else {
                manufacturingConditionBreakdownFilterCombo.getSelectionModel().selectFirst();
            }
        } finally {
            suppressFilterUi.set(false);
        }
    }

    /** Stored string matches combo items (empty string = {@link #MANUFACTURING_CONDITION_FILTER_ALL}). */
    private static String persistManufacturingConditionSelection(String selectedItem) {
        if (selectedItem == null || MANUFACTURING_CONDITION_FILTER_ALL.equals(selectedItem)) {
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
     * Keeps rows whose {@link #HEADER_MANUFACTURING_CONDITION_BREAKDOWN} cell equals the combo selection.
     * {@link #MANUFACTURING_CONDITION_FILTER_ALL} leaves all rows. Unknown column: no filtering (with log).
     */
    private PlanInputTabularIo.TabularSheet applyManufacturingConditionFilter(
            PlanInputTabularIo.TabularSheet shaped) {
        String sel =
                manufacturingConditionBreakdownFilterCombo != null
                        ? manufacturingConditionBreakdownFilterCombo.getSelectionModel().getSelectedItem()
                        : null;
        if (sel == null || MANUFACTURING_CONDITION_FILTER_ALL.equals(sel)) {
            return shaped;
        }
        String want = MANUFACTURING_CONDITION_EMPTY_DISPLAY.equals(sel) ? "" : sel;
        List<String> headers = shaped.headers();
        int col = indexOfManufacturingConditionColumn(headers);
        if (col < 0) {
            if (shell != null) {
                shell.appendLog(
                        "[processing-actuals-detail] missing header column: "
                                + HEADER_MANUFACTURING_CONDITION_BREAKDOWN);
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

    private static int indexOfManufacturingConditionColumn(List<String> headers) {
        if (headers == null) {
            return -1;
        }
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (HEADER_MANUFACTURING_CONDITION_BREAKDOWN.equals(h != null ? h.strip() : "")) {
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
        PlanInputTabularIo.TabularSheet tab = applyManufacturingFilterThenQuadDedupe(shaped);
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
        visBefore =
                ColumnVisibilitySupport.mergeMandatoryIntoVisibility(
                        visBefore, mandatoryVisibilityMaskForHeaders(beforeHeaders));
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
        visAfter =
                ColumnVisibilitySupport.mergeMandatoryIntoVisibility(
                        visAfter, mandatoryVisibilityMaskForHeaders(headersRef));
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
                manufacturingConditionBreakdownFilterCombo != null
                        ? manufacturingConditionBreakdownFilterCombo.getSelectionModel().getSelectedItem()
                        : null;
        if (sel == null || MANUFACTURING_CONDITION_FILTER_ALL.equals(sel)) {
            return "";
        }
        return " [\u7d5e\u308a]";
    }

    private void applyEmpty() {
        suppressFilterUi.set(true);
        try {
            if (manufacturingConditionBreakdownFilterCombo != null) {
                manufacturingConditionBreakdownFilterCombo.getItems().clear();
                manufacturingConditionBreakdownFilterCombo.getItems().add(MANUFACTURING_CONDITION_FILTER_ALL);
                manufacturingConditionBreakdownFilterCombo.getSelectionModel().selectFirst();
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

    /**
     * Snapshot of shaped headers <em>before</em> manufacturing-condition filter and dedupe.
     * This is the same data Python's {@code _aggregate_daily_actual_qty_aladdin_max} uses as its
     * source, so calendar overlay applies the same "長さ" filter independently.
     */
    List<String> getUnfilteredShapedHeaders() {
        return new ArrayList<>(unfilteredShapedHeaders);
    }

    /** Snapshot of shaped rows before filter/dedupe (see {@link #getUnfilteredShapedHeaders()}). */
    List<List<String>> getUnfilteredShapedRows() {
        return new ArrayList<>(unfilteredShapedRows);
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
