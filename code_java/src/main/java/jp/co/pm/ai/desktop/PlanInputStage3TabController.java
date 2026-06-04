package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.effect.DropShadow;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.GridChange;
import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.io.ExcelCellReadSupport;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.PlanInputDeprecatedOverrideColumnSupport;
import jp.co.pm.ai.desktop.ui.PlanInputEditedCellMarks;
import jp.co.pm.ai.desktop.ui.PlanInputDateColumnSupport;
import jp.co.pm.ai.desktop.ui.PlanInputProcessSequenceRowOrder;
import jp.co.pm.ai.desktop.ui.PlanInputStage3DispatchableViolationSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetMultiColumnFilterCoordinator;
import jp.co.pm.ai.desktop.ui.SpreadsheetPlanInputCellEditSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetPlanInputRowDragSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/**
 * 「配台計画_タスク入力3.0」タブ。段階3.0 前処理（入力3表生成）で書き出した枝番タスクを
 * {@link PlanInputTabController} と同じ ControlsFX {@link SpreadsheetView}（配色・列操作・編集）で表示する。
 */
public class PlanInputStage3TabController {

    private static final String COL_DISPATCH_TRIAL_ORDER = "配台試行順番";

    /** 入力3表シート名（Python 側 PLAN_INPUT_STAGE3_SHEET_NAME の既定と一致）。 */
    public static final String STAGE3_SHEET_NAME = "配台計画_タスク入力3.0";

    private static final String HINT_TEXT =
            "配台計画手動修正タブで「入力3表を生成」した枝番タスク（"
                    + STAGE3_SHEET_NAME
                    + " シート）を表示・編集します。"
                    + " 保存は入力1表シートを維持したまま入力3表シートのみ上書きします。";

    @FXML private Button stage30RunButton;
    @FXML private Button stage31RunButton;
    @FXML private Button stage32RunButton;
    @FXML private Button reloadButton;
    @FXML private Button saveButton;
    @FXML private Button addRowButton;
    @FXML private Button removeRowsButton;
    @FXML private Label pathLabel;
    @FXML private Label statusLabel;
    @FXML private Label hintLabel;
    @FXML private Label stage3ValidationWarningLabel;
    @FXML private TextField rowSearchField;
    @FXML private TextField colWidthField;
    @FXML private HBox tableOperationBar;
    @FXML private HBox columnStripHost;
    @FXML private StackPane spreadsheetHost;

    private final SpreadsheetView spreadsheetView = new SpreadsheetView();
    private final List<String> headersRef = new ArrayList<>();
    private ObservableList<ObservableList<String>> rows;
    private final AtomicBoolean suppressColumnOrderPersistence = new AtomicBoolean(false);
    private final AtomicBoolean suppressDirtyFromGridEvents = new AtomicBoolean(false);
    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());
    private final AtomicInteger headerColumnCount = new AtomicInteger(0);
    private final AtomicReference<Stage2RollUnitLengthTables> cachedRollUnitHighlightTables =
            new AtomicReference<>();

    /** sidecar 名前空間（入力3表シート専用。入力1表のマークと混ざらないよう分ける）。 */
    private static final String EDIT_MARKS_NAMESPACE = "stage3";
    private final Map<String, String> editBaselineByMarkKey = new LinkedHashMap<>();
    private final Set<String> editMarksPersistedAtLoad = new LinkedHashSet<>();
    private final Set<String> editedCellMarks = new LinkedHashSet<>();

    private GridBase currentGrid;
    private EventHandler<GridChange> gridChangeHandler;
    private boolean cellEditHooksInstalled;
    private boolean stageRunPipelineBusy;

    /** 段階3.0/3.1/3.2 ボタン押下〜完了まで（3.1 ウィザード表示中を含む）。 */
    private boolean stage3RunButtonsLocked;
    private boolean tableDirtySinceSave;

    private MainShellController shell;
    private Stage ownerStage;

    @FXML
    private void initialize() {
        if (hintLabel != null) {
            hintLabel.setText(HINT_TEXT);
        }
        if (colWidthField != null) {
            colWidthField.setText("112");
        }
        installStageRunButtonDepth(stage30RunButton, Color.rgb(120, 81, 169, 0.35));
        installStageRunButtonDepth(stage31RunButton, Color.rgb(120, 81, 169, 0.35));
        installStageRunButtonDepth(stage32RunButton, Color.rgb(120, 81, 169, 0.35));

        StackPane.setAlignment(spreadsheetView, Pos.TOP_LEFT);
        spreadsheetView.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        rows = FXCollections.observableArrayList();
        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        SpreadsheetTabularSupport.installFullRowDataSelection(
                spreadsheetView,
                SpreadsheetPlanInputRowDragSupport::skipFullRowExpansionDuringPlanInputRowDrag);
        SpreadsheetThemeBridge.install(spreadsheetView);
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(spreadsheetView);
        SpreadsheetTabularSupport.installSpreadsheetClickSelectionAlign(spreadsheetView);
        SpreadsheetPlanInputRowDragSupport.install(
                spreadsheetView,
                SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex(),
                rows,
                this::finishRowReorderAfterDnD);
        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                spreadsheetHost,
                headerColumnCount::get,
                SpreadsheetPlanInputRowDragSupport::skipFullRowExpansionDuringPlanInputRowDrag);
        if (rowSearchField != null) {
            rowSearchField
                    .textProperty()
                    .addListener(
                            (obs, prev, cur) ->
                                    SpreadsheetMultiColumnFilterCoordinator.setRowTextSearchQuery(
                                            spreadsheetView, cur));
        }
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        ownerStage = shell.getPrimaryStage();

        columnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::applyDynamicColumnWidths,
                                TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3,
                                headerColumnCount,
                                this::onLeadingColumnCountCommitted,
                                this::onReorderColumns,
                                () ->
                                        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                                                ownerStage,
                                                TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3,
                                                spreadsheetView,
                                                () -> new ArrayList<>(headersRef))));

        if (!cellEditHooksInstalled) {
            SpreadsheetPlanInputCellEditSupport.install(
                    spreadsheetView,
                    ownerStage,
                    SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex(),
                    headersRef,
                    rows,
                    () -> {
                        markTableDirtySinceSave();
                        rebuildSpreadsheet();
                    });
            cellEditHooksInstalled = true;
        }

        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                spreadsheetView,
                TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3,
                suppressColumnOrderPersistence::get,
                () -> new ArrayList<>(headersRef));

        Platform.runLater(this::reloadFromDisk);
    }

    void setStageRunProgressVisible(boolean stage1Running, boolean pipelineBusy) {
        stageRunPipelineBusy = stage1Running || pipelineBusy;
        applyStageRunButtonEnabledState();
        applyPipelineBusyControlsState();
    }

    void setStage3RunButtonsLocked(boolean locked) {
        stage3RunButtonsLocked = locked;
        applyStageRunButtonEnabledState();
    }

    boolean isStage3RunButtonsLocked() {
        return stage3RunButtonsLocked;
    }

    boolean isTableDirtySinceSave() {
        return tableDirtySinceSave;
    }

    void invalidateRollUnitHighlightCacheAndRefresh() {
        cachedRollUnitHighlightTables.set(null);
        if (!headersRef.isEmpty()) {
            rebuildSpreadsheet();
        }
    }

    @FXML
    private void onStage30RunButtonAction() {
        if (shell != null && !stage3RunButtonsLocked && !stageRunPipelineBusy) {
            shell.triggerStage30();
        }
    }

    @FXML
    private void onStage31RunButtonAction() {
        if (shell != null && !stage3RunButtonsLocked && !stageRunPipelineBusy) {
            shell.triggerStage31();
        }
    }

    @FXML
    private void onStage32RunButtonAction() {
        if (shell != null && !stage3RunButtonsLocked && !stageRunPipelineBusy) {
            shell.triggerStage32();
        }
    }

    @FXML
    private void onReloadButtonAction() {
        reloadFromDisk();
    }

    @FXML
    private void onSaveButtonAction() {
        Path workbook = resolveWorkbookPath();
        if (workbook == null) {
            if (shell != null) {
                shell.showWarningDialog("保存", "入力3表の元ブックが見つかりません。");
            }
            return;
        }
        if (headersRef.isEmpty()) {
            if (shell != null) {
                shell.showWarningDialog("保存", "保存する表データがありません。");
            }
            return;
        }
        try {
            List<List<String>> dataRows = copyRowsForSave();
            PlanInputTabularIo.writeExcelSheetPreservingOthers(
                    workbook,
                    STAGE3_SHEET_NAME,
                    new PlanInputTabularIo.TabularSheet(headersRef, dataRows));
            PlanInputEditedCellMarks.save(workbook, editedCellMarks, EDIT_MARKS_NAMESPACE);
            clearTableDirtySinceSave();
            setStatus("保存済み " + dataRows.size() + " 行");
            if (shell != null) {
                shell.appendLog("[plan-input-stage3] saved " + workbook);
                shell.showInformationDialog(
                        "保存完了",
                        "入力3表を保存しました。\n" + workbook + "\n行数: " + dataRows.size());
            }
        } catch (Exception ex) {
            setStatus("保存エラー: " + ex.getMessage());
            if (shell != null) {
                shell.appendLog("[plan-input-stage3] save error: " + ex.getMessage());
                shell.showErrorDialog(
                        "保存エラー",
                        ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
        }
    }

    @FXML
    private void onAddRowButtonAction() {
        if (headersRef.isEmpty()) {
            setStatus("先に入力3表を読み込んでください。");
            return;
        }
        ObservableList<String> r = FXCollections.observableArrayList();
        for (int i = 0; i < headersRef.size(); i++) {
            r.add("");
        }
        rows.add(r);
        markTableDirtySinceSave();
        rebuildSpreadsheet();
    }

    @FXML
    private void onRemoveRowsButtonAction() {
        var cells = spreadsheetView.getSelectionModel().getSelectedCells();
        if (cells.isEmpty()) {
            return;
        }
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        List<Integer> sorted =
                cells.stream()
                        .map(TablePosition::getRow)
                        .filter(gr -> gr >= firstData)
                        .map(gr -> gr - firstData)
                        .distinct()
                        .sorted(Comparator.reverseOrder())
                        .collect(Collectors.toList());
        for (int r : sorted) {
            if (r >= 0 && r < rows.size()) {
                rows.remove(r);
            }
        }
        markTableDirtySinceSave();
        rebuildSpreadsheet();
    }

    @FXML
    private void onRowUpAction() {
        int i = selectedDataRowIndex();
        if (i <= 0) {
            return;
        }
        int colIdx = focusedColumnIndex();
        markTableDirtySinceSave();
        swapDataRowsInMemory(i - 1, i);
        scheduleRowReorderPresentation(i - 1, colIdx);
    }

    @FXML
    private void onRowDownAction() {
        int i = selectedDataRowIndex();
        if (i < 0 || i >= rows.size() - 1) {
            return;
        }
        int colIdx = focusedColumnIndex();
        markTableDirtySinceSave();
        swapDataRowsInMemory(i, i + 1);
        scheduleRowReorderPresentation(i + 1, colIdx);
    }

    @FXML
    private void onClearColumnFiltersAction() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(spreadsheetView);
        if (rowSearchField != null) {
            rowSearchField.clear();
        }
    }

    /** 入力3表シートをディスクから読み込み、表に反映する。 */
    public void reloadFromDisk() {
        Path workbook = resolveWorkbookPath();
        if (pathLabel != null) {
            pathLabel.setText(workbook != null ? workbook.toString() : "（未設定）");
        }
        if (workbook == null || !Files.isRegularFile(workbook)) {
            setStatus("入力3表の元ブックが見つかりません。配台計画手動修正タブで「入力3表を生成」してください。");
            clearTableData();
            return;
        }
        try {
            PlanInputTabularIo.TabularSheet sheet =
                    PlanInputTabularIo.read(workbook, STAGE3_SHEET_NAME);
            applyLoadedSheet(sheet);
            clearTableDirtySinceSave();
            setStatus(sheet.rows().size() + " 行");
        } catch (Exception ex) {
            clearTableData();
            setStatus(
                    "入力3表シート「"
                            + STAGE3_SHEET_NAME
                            + "」を読み込めません。"
                            + "段階3.0 前処理（入力3表生成）が未実行の可能性があります。詳細: "
                            + ex.getMessage());
        }
    }

    private void applyLoadedSheet(PlanInputTabularIo.TabularSheet sheet) {
        headersRef.clear();
        headersRef.addAll(sheet.headers());
        rows.clear();
        for (List<String> line : sheet.rows()) {
            ObservableList<String> r = FXCollections.observableArrayList(line);
            while (r.size() < headersRef.size()) {
                r.add("");
            }
            while (r.size() > headersRef.size()) {
                r.remove(r.size() - 1);
            }
            rows.add(r);
        }
        PlanInputDeprecatedOverrideColumnSupport.migrateAndDropDeprecatedOverrideColumns(
                headersRef, rows);
        normalizePlanInputDateOnlyColumns();
        List<TableColumnOrderPersistence.ColumnSpec> lay =
                TableColumnOrderPersistence.loadLayout(
                        TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3);
        persistedLayout.set(lay);
        List<String> fileHeaders = new ArrayList<>(headersRef);
        List<String> titleOrder =
                lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
        TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
        boolean[] visForHeaders =
                TableColumnOrderPersistence.resolveVisibilityAfterSheetLoad(
                        TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3,
                        fileHeaders,
                        titleOrder,
                        headersRef);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3, visForHeaders);
        renumberDispatchTrialOrderColumn();
        loadPlanInputEditMarks(resolveWorkbookPath());
        rebuildSpreadsheet(false);
    }

    /** 読込時: 編集差分の基準値を記録し、sidecar JSON のマークを取り込む。 */
    private void loadPlanInputEditMarks(Path workbook) {
        editBaselineByMarkKey.clear();
        editBaselineByMarkKey.putAll(PlanInputEditedCellMarks.captureBaseline(headersRef, rows));
        Set<String> loaded =
                PlanInputEditedCellMarks.filterToPresentRows(
                        headersRef,
                        rows,
                        PlanInputEditedCellMarks.load(workbook, EDIT_MARKS_NAMESPACE));
        editMarksPersistedAtLoad.clear();
        editMarksPersistedAtLoad.addAll(loaded);
        editedCellMarks.clear();
        editedCellMarks.addAll(loaded);
    }

    /** 現在の表から編集マークを再計算し、変化があれば sidecar JSON へ保存する。 */
    private void refreshAndPersistPlanInputEditMarks() {
        PlanInputEditedCellMarks.recompute(
                headersRef, rows, editBaselineByMarkKey, editMarksPersistedAtLoad, editedCellMarks);
        Path workbook = resolveWorkbookPath();
        if (workbook != null) {
            PlanInputEditedCellMarks.save(workbook, editedCellMarks, EDIT_MARKS_NAMESPACE);
        }
    }

    private void clearTableData() {
        headersRef.clear();
        if (rows != null) {
            rows.clear();
        }
        if (rowSearchField != null) {
            rowSearchField.clear();
        }
        onClearColumnFiltersAction();
        rebuildSpreadsheet(false);
    }

    private Path resolveWorkbookPath() {
        if (shell == null) {
            return null;
        }
        String p = shell.stage3PlanInputWorkbookPath();
        if (p == null || p.isBlank()) {
            return null;
        }
        try {
            return Path.of(p.trim());
        } catch (Exception ex) {
            return null;
        }
    }

    private List<List<String>> copyRowsForSave() {
        List<List<String>> dataRows = new ArrayList<>();
        for (ObservableList<String> r : rows) {
            List<String> copy = new ArrayList<>(r);
            while (copy.size() < headersRef.size()) {
                copy.add("");
            }
            while (copy.size() > headersRef.size()) {
                copy.remove(copy.size() - 1);
            }
            dataRows.add(copy);
        }
        return dataRows;
    }

    private void markTableDirtySinceSave() {
        tableDirtySinceSave = true;
        applyStageRunButtonEnabledState();
    }

    private void clearTableDirtySinceSave() {
        tableDirtySinceSave = false;
        applyStageRunButtonEnabledState();
    }

    private void applyStageRunButtonEnabledState() {
        boolean disable = stageRunPipelineBusy || stage3RunButtonsLocked || tableDirtySinceSave;
        Tooltip runningTip =
                new Tooltip("段階3の処理が実行中です。完了までお待ちください。");
        Tooltip dirtyTip =
                new Tooltip(
                        "入力3表に未保存の変更があります。「保存」または「入力3表を再読込」で確定してから実行してください。");
        for (Button b : List.of(stage30RunButton, stage31RunButton, stage32RunButton)) {
            if (b == null) {
                continue;
            }
            b.setDisable(disable);
            if (stage3RunButtonsLocked || stageRunPipelineBusy) {
                b.setTooltip(runningTip);
            } else if (tableDirtySinceSave) {
                b.setTooltip(dirtyTip);
            } else {
                b.setTooltip(null);
            }
        }
    }

    /** 段階1／2／3 パイプライン実行中は表操作・再読込を無効化する。 */
    private void applyPipelineBusyControlsState() {
        boolean disable = stageRunPipelineBusy;
        if (reloadButton != null) {
            reloadButton.setDisable(disable);
        }
        if (saveButton != null) {
            saveButton.setDisable(disable);
        }
        if (tableOperationBar != null) {
            tableOperationBar.setDisable(disable);
        }
        if (colWidthField != null) {
            colWidthField.setDisable(disable);
        }
        if (rowSearchField != null) {
            rowSearchField.setDisable(disable);
        }
        if (columnStripHost != null) {
            columnStripHost.setDisable(disable);
        }
    }

    private void applyDynamicColumnWidths() {
        double w = 112;
        try {
            w = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
        } catch (NumberFormatException ignored) {
        }
        for (var c : spreadsheetView.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private Stage2RollUnitLengthTables rollUnitHighlightTablesCached() {
        return cachedRollUnitHighlightTables.updateAndGet(
                cur -> {
                    if (cur != null) {
                        return cur;
                    }
                    if (shell == null) {
                        return Stage2RollUnitLengthTables.empty();
                    }
                    try {
                        return Stage2RollUnitLengthTables.load(shell.snapshotUiEnv());
                    } catch (Exception e) {
                        return Stage2RollUnitLengthTables.empty();
                    }
                });
    }

    private void rebuildSpreadsheet() {
        rebuildSpreadsheet(true, SpreadsheetTabularSupport.GridAttachMode.STANDARD);
    }

    private void rebuildSpreadsheet(boolean preserveColumnFilters) {
        SpreadsheetTabularSupport.GridAttachMode attach =
                preserveColumnFilters
                        ? SpreadsheetTabularSupport.GridAttachMode.STANDARD
                        : SpreadsheetTabularSupport.GridAttachMode.FULL_RESET;
        rebuildSpreadsheet(preserveColumnFilters, attach);
    }

    private void rebuildSpreadsheet(
            boolean preserveColumnFilters, SpreadsheetTabularSupport.GridAttachMode attachMode) {
        boolean rowReorderRefresh =
                attachMode == SpreadsheetTabularSupport.GridAttachMode.IN_PLACE;
        if (headersRef.isEmpty()) {
            detachGridHandler();
            GridBase empty = new GridBase(0, 0);
            SpreadsheetTabularSupport.attachGridToSpreadsheetView(
                    spreadsheetView, empty, SpreadsheetTabularSupport.GridAttachMode.STANDARD);
            currentGrid = empty;
            updateStage3DispatchableViolationWarning();
            return;
        }
        final Map<Integer, Set<String>> columnFilterSnapshot =
                preserveColumnFilters
                        ? SpreadsheetMultiColumnFilterCoordinator.copyColumnAllowedByIndex(
                                spreadsheetView)
                        : Map.of();
        suppressColumnOrderPersistence.set(true);
        suppressDirtyFromGridEvents.set(true);
        try {
            detachGridHandler();
            double colW = 112;
            try {
                colW = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
            } catch (NumberFormatException ignored) {
            }
            final List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            headersRef, persistedLayout.get(), colW);
            final double widthDefault = colW;

            GridBase grid =
                    SpreadsheetTabularSupport.buildPlanInputGrid(
                            headersRef,
                            rows,
                            false,
                            headerColumnCount.get(),
                            rollUnitHighlightTablesCached());
            int firstDataRow = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
            refreshAndPersistPlanInputEditMarks();
            PlanInputEditedCellMarks.applyHighlights(
                    grid, headersRef, rows, firstDataRow, editedCellMarks);
            PlanInputStage3DispatchableViolationSupport.applyViolationHighlights(
                    grid, headersRef, rows, firstDataRow);
            updateStage3DispatchableViolationWarning();
            var rowSync =
                    SpreadsheetTabularSupport.newRowsSyncHandler(rows, headersRef, firstDataRow);
            gridChangeHandler =
                    ev -> {
                        rowSync.handle(ev);
                        refreshAndPersistPlanInputEditMarks();
                        PlanInputEditedCellMarks.applyHighlights(
                                currentGrid, headersRef, rows, firstDataRow, editedCellMarks);
                        PlanInputStage3DispatchableViolationSupport.applyViolationHighlights(
                                currentGrid, headersRef, rows, firstDataRow);
                        updateStage3DispatchableViolationWarning();
                        if (!suppressDirtyFromGridEvents.get()) {
                            markTableDirtySinceSave();
                        }
                    };
            grid.addEventHandler(GridChange.GRID_CHANGE_EVENT, gridChangeHandler);
            currentGrid = grid;
            SpreadsheetTabularSupport.attachGridToSpreadsheetView(
                    spreadsheetView, grid, attachMode);

            Platform.runLater(
                    () -> {
                        try {
                            if (rowReorderRefresh) {
                                SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                        spreadsheetView, headerColumnCount.get());
                                SpreadsheetTabularSupport.applyColumnFiltersWithDialog(
                                        spreadsheetView);
                                SpreadsheetMultiColumnFilterCoordinator.restoreColumnAllowedSnapshot(
                                        spreadsheetView, columnFilterSnapshot);
                                SpreadsheetTabularSupport.pinSpreadsheetFilterRow(spreadsheetView);
                                String q =
                                        rowSearchField != null && rowSearchField.getText() != null
                                                ? rowSearchField.getText().trim()
                                                : "";
                                if (!q.isEmpty()) {
                                    SpreadsheetMultiColumnFilterCoordinator.setRowTextSearchQuery(
                                            spreadsheetView, q);
                                }
                                return;
                            }
                            SpreadsheetTabularSupport.applyColumnWidths(
                                    spreadsheetView, widths, widthDefault);
                            SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                    spreadsheetView, headerColumnCount.get());
                            SpreadsheetTabularSupport.applyColumnFiltersWithDialog(spreadsheetView);
                            SpreadsheetMultiColumnFilterCoordinator.restoreColumnAllowedSnapshot(
                                    spreadsheetView, columnFilterSnapshot);
                            SpreadsheetTabularSupport.pinSpreadsheetFilterRow(spreadsheetView);
                            SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(
                                    spreadsheetView);
                            ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                                    spreadsheetView,
                                    () -> new ArrayList<>(headersRef),
                                    () ->
                                            TableColumnOrderPersistence.loadColumnVisibility(
                                                    TableColumnOrderPersistence
                                                            .TableId.PLAN_INPUT_STAGE3,
                                                    headersRef.size()));
                            String q =
                                    rowSearchField != null && rowSearchField.getText() != null
                                            ? rowSearchField.getText().trim()
                                            : "";
                            if (!q.isEmpty()) {
                                SpreadsheetMultiColumnFilterCoordinator.setRowTextSearchQuery(
                                        spreadsheetView, q);
                            }
                        } finally {
                            suppressDirtyFromGridEvents.set(false);
                        }
                    });
        } catch (Throwable t) {
            suppressDirtyFromGridEvents.set(false);
            throw t;
        } finally {
            suppressColumnOrderPersistence.set(false);
        }
    }

    private void detachGridHandler() {
        if (currentGrid != null && gridChangeHandler != null) {
            currentGrid.removeEventHandler(GridChange.GRID_CHANGE_EVENT, gridChangeHandler);
        }
        gridChangeHandler = null;
        currentGrid = null;
    }

    private void renumberDispatchTrialOrderColumn() {
        PlanInputProcessSequenceRowOrder.stabilizeAndRenumberDispatchTrialOrder(headersRef, rows);
    }

    private int selectedDataRowIndex() {
        var sm = spreadsheetView.getSelectionModel();
        TablePosition<?, ?> pos = sm.getFocusedCell();
        if (pos == null || pos.getRow() < 0) {
            var cells = sm.getSelectedCells();
            if (cells.isEmpty()) {
                return -1;
            }
            pos = cells.getFirst();
        }
        int gridRow = spreadsheetView.getModelRow(pos.getRow());
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int idx = gridRow - firstData;
        if (idx >= 0 && idx < rows.size()) {
            return idx;
        }
        return -1;
    }

    private int focusedColumnIndex() {
        var sm = spreadsheetView.getSelectionModel();
        TablePosition<?, ?> pos = sm.getFocusedCell();
        if (pos != null && pos.getColumn() >= 0) {
            return pos.getColumn();
        }
        var cells = sm.getSelectedCells();
        if (cells != null && !cells.isEmpty()) {
            int c = cells.getFirst().getColumn();
            if (c >= 0) {
                return c;
            }
        }
        return 0;
    }

    private void finishRowReorderAfterDnD() {
        renumberDispatchTrialOrderColumn();
        markTableDirtySinceSave();
        applyRowReorderPresentation(-1, -1);
    }

    private void scheduleRowReorderPresentation(int focusDataRow, int focusColumn) {
        Platform.runLater(() -> applyRowReorderPresentation(focusDataRow, focusColumn));
    }

    private void applyRowReorderPresentation(int focusDataRow, int focusColumn) {
        rebuildSpreadsheet(true, SpreadsheetTabularSupport.GridAttachMode.IN_PLACE);
        if (focusDataRow >= 0) {
            focusCellAfterReorder(focusDataRow, focusColumn);
        }
    }

    private void swapDataRowsInMemory(int a, int b) {
        if (a < 0 || b < 0 || a >= rows.size() || b >= rows.size() || a == b) {
            return;
        }
        ObservableList<String> moved = rows.get(a);
        rows.set(a, rows.get(b));
        rows.set(b, moved);
        renumberDispatchTrialOrderColumn();
    }

    private void focusCellAfterReorder(int dataRowIndex, int columnIndex) {
        if (dataRowIndex < 0 || dataRowIndex >= rows.size()) {
            return;
        }
        var cols = spreadsheetView.getColumns();
        if (cols.isEmpty()) {
            return;
        }
        int c = Math.max(0, Math.min(columnIndex, cols.size() - 1));
        SpreadsheetColumn scol = cols.get(c);
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int modelGridRow = firstData + dataRowIndex;
        Platform.runLater(
                () -> {
                    int viewRow = spreadsheetView.getViewRow(modelGridRow);
                    if (viewRow < 0) {
                        return;
                    }
                    var sm = spreadsheetView.getSelectionModel();
                    SpreadsheetTabularSupport.safeClearSpreadsheetSelection(spreadsheetView);
                    try {
                        sm.clearAndSelect(viewRow, scol);
                        sm.focus(viewRow, scol);
                    } catch (RuntimeException ex) {
                        SpreadsheetTabularSupport.safeClearSpreadsheetSelection(spreadsheetView);
                    }
                });
    }

    private void onLeadingColumnCountCommitted(int n) {
        headerColumnCount.set(n);
        markTableDirtySinceSave();
        rebuildSpreadsheet();
    }

    private void onReorderColumns() {
        if (headersRef.isEmpty() || shell == null) {
            return;
        }
        boolean[] visForDialog =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3, headersRef.size());
        SpreadsheetColumnReorderDialog.show(
                        ownerStage, new ArrayList<>(headersRef), visForDialog)
                .ifPresent(
                        perm -> {
                            List<String> oldHeaders = new ArrayList<>(headersRef);
                            boolean[] oldVis =
                                    TableColumnOrderPersistence.loadColumnVisibility(
                                            TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3,
                                            oldHeaders.size());
                            List<String> titleOrder = perm.stream().map(oldHeaders::get).toList();
                            List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
                            TableColumnOrderPersistence.applyLogicalColumnOrder(
                                    headersRef, rows, titleOrder);
                            boolean[] newVis =
                                    TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                                            oldHeaders, oldVis, titleOrder);
                            TableColumnOrderPersistence.saveColumnVisibility(
                                    TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3, newVis);
                            double colW = 112;
                            try {
                                colW =
                                        Math.max(
                                                40,
                                                Double.parseDouble(colWidthField.getText().trim()));
                            } catch (NumberFormatException ignored) {
                            }
                            List<Double> widths =
                                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                                            headersRef, lay, colW);
                            List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
                            for (int i = 0; i < headersRef.size(); i++) {
                                newLay.add(
                                        new TableColumnOrderPersistence.ColumnSpec(
                                                headersRef.get(i), widths.get(i)));
                            }
                            persistedLayout.set(newLay);
                            TableColumnOrderPersistence.saveLayout(
                                    TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3, newLay);
                            markTableDirtySinceSave();
                            rebuildSpreadsheet(false);
                        });
    }

    private void normalizePlanInputDateOnlyColumns() {
        if (headersRef.isEmpty() || rows == null || rows.isEmpty()) {
            return;
        }
        List<Integer> colIdx = new ArrayList<>();
        for (int c = 0; c < headersRef.size(); c++) {
            if (PlanInputDateColumnSupport.isEditableDateColumn(headersRef.get(c))) {
                colIdx.add(c);
            }
        }
        if (colIdx.isEmpty()) {
            return;
        }
        for (ObservableList<String> row : rows) {
            for (int c : colIdx) {
                if (c < row.size()) {
                    String v = row.get(c);
                    if (v != null && !v.isEmpty()) {
                        row.set(c, ExcelCellReadSupport.stripMidnightDateTimeSuffix(v));
                    }
                }
            }
        }
    }

    private void updateStage3DispatchableViolationWarning() {
        if (stage3ValidationWarningLabel == null) {
            return;
        }
        int violations = PlanInputStage3DispatchableViolationSupport.countViolations(headersRef, rows);
        boolean show = violations > 0;
        stage3ValidationWarningLabel.setVisible(show);
        stage3ValidationWarningLabel.setManaged(show);
        if (show) {
            stage3ValidationWarningLabel.setText(
                    PlanInputStage3DispatchableViolationSupport.warningMessage(violations));
        } else {
            stage3ValidationWarningLabel.setText("");
        }
    }

    private void setStatus(String msg) {
        Runnable apply = () -> statusLabel.setText(msg);
        if (Platform.isFxApplicationThread()) {
            apply.run();
        } else {
            Platform.runLater(apply);
        }
    }

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
}
