package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
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
import jp.co.pm.ai.desktop.ui.PlanInputDateColumnSupport;
import jp.co.pm.ai.desktop.ui.PlanInputRawInputDateShift;
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
    @FXML private Button shiftRawInputDateMinusOneButton;
    @FXML private Button clearRawInputDateOverrideButton;
    @FXML private Label pathLabel;
    @FXML private Label statusLabel;
    @FXML private Label hintLabel;
    @FXML private TextField rowSearchField;
    @FXML private TextField colWidthField;
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

    private GridBase currentGrid;
    private EventHandler<GridChange> gridChangeHandler;
    private boolean cellEditHooksInstalled;
    private boolean stageRunPipelineBusy;
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

        StackPane.setAlignment(spreadsheetView, Pos.CENTER_LEFT);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        rows = FXCollections.observableArrayList();
        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        SpreadsheetTabularSupport.installFullRowDataSelection(
                spreadsheetView,
                SpreadsheetPlanInputRowDragSupport::skipFullRowExpansionDuringPlanInputRowDrag);
        SpreadsheetThemeBridge.install(spreadsheetView);
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(spreadsheetView);
        SpreadsheetPlanInputRowDragSupport.install(
                spreadsheetView,
                SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex(),
                rows,
                () -> {
                    renumberDispatchTrialOrderColumn();
                    markTableDirtySinceSave();
                    rebuildSpreadsheet();
                });
        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                spreadsheetHost, headerColumnCount::get);
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
        if (shell != null) {
            shell.triggerStage30();
        }
    }

    @FXML
    private void onStage31RunButtonAction() {
        if (shell != null) {
            shell.triggerStage31();
        }
    }

    @FXML
    private void onStage32RunButtonAction() {
        if (shell != null) {
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
            clearTableDirtySinceSave();
            setStatus("保存しました: " + workbook + " （" + dataRows.size() + " 行）");
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
    private void onShiftRawInputDateMinusOneAction() {
        if (shell == null || headersRef.isEmpty()) {
            return;
        }
        int updated = PlanInputRawInputDateShift.applyMinusOneDayToAllOverrides(headersRef, rows);
        if (updated == PlanInputRawInputDateShift.MISSING_OVERRIDE_COLUMN) {
            shell.showErrorDialog(
                    "原反投入日の前倒し",
                    "列「"
                            + PlanInputRawInputDateShift.COL_RAW_INPUT_DATE_OVERRIDE
                            + "」がありません。表を読み込んでから実行してください。");
            return;
        }
        if (updated == 0) {
            shell.showInformationDialog(
                    "原反投入日の前倒し",
                    "原反投入日（または上書き）を解釈できる行がありませんでした。");
            return;
        }
        markTableDirtySinceSave();
        rebuildSpreadsheet();
    }

    @FXML
    private void onClearRawInputDateOverrideAction() {
        if (shell == null || headersRef.isEmpty()) {
            return;
        }
        int cleared = PlanInputRawInputDateShift.clearAllOverrides(headersRef, rows);
        if (cleared == PlanInputRawInputDateShift.MISSING_OVERRIDE_COLUMN) {
            shell.showErrorDialog(
                    "原反投入日上書きのクリア",
                    "列「"
                            + PlanInputRawInputDateShift.COL_RAW_INPUT_DATE_OVERRIDE
                            + "」がありません。表を読み込んでから実行してください。");
            return;
        }
        if (cleared == 0) {
            shell.showInformationDialog(
                    "原反投入日上書きのクリア",
                    "「"
                            + PlanInputRawInputDateShift.COL_RAW_INPUT_DATE_OVERRIDE
                            + "」に値がある行がありませんでした。");
            return;
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
        swapDataRows(i - 1, i);
        focusCellAfterReorder(i - 1, colIdx);
    }

    @FXML
    private void onRowDownAction() {
        int i = selectedDataRowIndex();
        if (i < 0 || i >= rows.size() - 1) {
            return;
        }
        int colIdx = focusedColumnIndex();
        markTableDirtySinceSave();
        swapDataRows(i, i + 1);
        focusCellAfterReorder(i + 1, colIdx);
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
            setStatus("入力3表: " + sheet.rows().size() + " 行（" + workbook + "）");
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
        normalizePlanInputDateOnlyColumns();
        List<TableColumnOrderPersistence.ColumnSpec> lay =
                TableColumnOrderPersistence.loadLayout(
                        TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3);
        persistedLayout.set(lay);
        List<String> beforeHeaders = new ArrayList<>(headersRef);
        boolean[] visBefore =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3, beforeHeaders.size());
        List<String> titleOrder =
                lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
        TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
        boolean[] visAfter =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        beforeHeaders, visBefore, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.PLAN_INPUT_STAGE3, visAfter);
        rebuildSpreadsheet(false);
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
        boolean disable = stageRunPipelineBusy || tableDirtySinceSave;
        Tooltip dirtyTip =
                new Tooltip(
                        "入力3表に未保存の変更があります。「保存」または「入力3表を再読込」で確定してから実行してください。");
        for (Button b : List.of(stage30RunButton, stage31RunButton, stage32RunButton)) {
            if (b == null) {
                continue;
            }
            b.setDisable(disable);
            b.setTooltip(tableDirtySinceSave && !stageRunPipelineBusy ? dirtyTip : null);
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
        rebuildSpreadsheet(true);
    }

    private void rebuildSpreadsheet(boolean preserveColumnFilters) {
        if (headersRef.isEmpty()) {
            detachGridHandler();
            GridBase empty = new GridBase(0, 0);
            spreadsheetView.getSelectionModel().clearSelection();
            spreadsheetView.setGrid(empty);
            currentGrid = empty;
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
            var rowSync =
                    SpreadsheetTabularSupport.newRowsSyncHandler(rows, headersRef, firstDataRow);
            gridChangeHandler =
                    ev -> {
                        rowSync.handle(ev);
                        if (!suppressDirtyFromGridEvents.get()) {
                            markTableDirtySinceSave();
                        }
                    };
            grid.addEventHandler(GridChange.GRID_CHANGE_EVENT, gridChangeHandler);
            currentGrid = grid;
            spreadsheetView.getSelectionModel().clearSelection();
            spreadsheetView.setGrid(grid);

            Platform.runLater(
                    () -> {
                        try {
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
        int col = headersRef.indexOf(COL_DISPATCH_TRIAL_ORDER);
        if (col < 0) {
            return;
        }
        for (int i = 0; i < rows.size(); i++) {
            ObservableList<String> r = rows.get(i);
            while (r.size() <= col) {
                r.add("");
            }
            r.set(col, Integer.toString(i + 1));
        }
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

    private void swapDataRows(int a, int b) {
        if (a < 0 || b < 0 || a >= rows.size() || b >= rows.size() || a == b) {
            return;
        }
        ObservableList<String> moved = rows.get(a);
        rows.set(a, rows.get(b));
        rows.set(b, moved);
        renumberDispatchTrialOrderColumn();
        rebuildSpreadsheet();
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
                    sm.clearSelection();
                    sm.clearAndSelect(viewRow, scol);
                    sm.focus(viewRow, scol);
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
