package jp.co.pm.ai.desktop;

import java.io.File;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.effect.DropShadow;
import javafx.scene.layout.HBox;
import javafx.scene.paint.Color;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.GridChange;
import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
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
 * 配台計画_タスク入力タブ。レイアウトは {@code PlanInputTab.fxml}。
 *
 * <p>段階2の「当日は配台しない」オプション（{@code PM_AI_STAGE2_SKIP_TODAY_DISPATCH}）のチェックボックスは本タブに配置する。
 *
 * <p>ControlsFX {@link SpreadsheetView} で先頭固定列をネイティブに扱う。
 */
public final class PlanInputTabController {

    /** planning_core の {@code RESULT_TASK_COL_DISPATCH_TRIAL_ORDER} 相当（段階1タスク入力の並び順）。 */
    private static final String COL_DISPATCH_TRIAL_ORDER = "配台試行順番";

    public static final String ENV_PM_AI_PLAN_INPUT_PATH = AppPaths.KEY_PM_AI_PLAN_INPUT_PATH;
    public static final String ENV_TASK_PLAN_SHEET = "TASK_PLAN_SHEET";

    public static final String DEFAULT_PLAN_INPUT_SHEET_NAME =
            "配台計画_タスク入力";

    private static final String HINT_TEXT =
            "PM_AI_PLAN_INPUT_PATH に読み込む表ファイルのパスを指定。"
                    + "（段階2 load_planning_tasks_df: CSV / Parquet / xlsx 対応）。"
                    + "Excel のときはシート名も指定（TASK_PLAN_SHEET / この欄）。"
                    + " .xlsx 保存はデータのみ（マクロは含みません）。";

    private Stage ownerStage;

    private MainShellController shell;

    @FXML
    private TextField pathField;

    @FXML
    private TextField sheetField;

    @FXML
    private Button browseButton;

    @FXML
    private Button loadButton;

    @FXML
    private Button saveButton;

    @FXML
    private Button addRowButton;

    @FXML
    private Button removeRowsButton;

    @FXML
    private Button stage2RunButton;

    @FXML
    private CheckBox stage2SkipTodayDispatchCheckBox;

    @FXML
    private HBox columnStripHost;

    @FXML
    private TextField colWidthField;

    @FXML
    private Label hintLabel;

    @FXML
    private TextField rowSearchField;

    @FXML
    private StackPane spreadsheetHost;

    private final SpreadsheetView spreadsheetView = new SpreadsheetView();

    private final List<String> headersRef = new ArrayList<>();
    private ObservableList<ObservableList<String>> rows;
    private final AtomicBoolean suppressColumnOrderPersistence = new AtomicBoolean(false);
    private final AtomicBoolean suppressPlanInputDirtyFromGridEvents = new AtomicBoolean(false);
    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());
    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    private final AtomicReference<Stage2RollUnitLengthTables> cachedRollUnitHighlightTables =
            new AtomicReference<>();

    private GridBase currentGrid;
    private EventHandler<GridChange> gridChangeHandler;

    private boolean planInputCellEditHooksInstalled;

    /** 段階1／段階2 の Python 実行中（メインシェルから同期）。 */
    private boolean stage2RunPipelineBusy;

    /** 配台計画手動修正タブの表が未保存のとき、段階2を抑止する。 */
    private boolean stage2BlockedByDispatchUnsavedEdit;

    /**
     * 配台計画_タスク入力タブの表を手動変更したが「保存」または「再読み」でディスクと同期していないとき、段階2を抑止する。
     */
    private boolean stage2BlockedByUnsavedPlanInputTableEdit;

    @FXML
    private void initialize() {
        pathField.setPromptText(
                "PM_AI_PLAN_INPUT_PATH （.csv / .xlsx / .xlsm）");
        sheetField.setText(DEFAULT_PLAN_INPUT_SHEET_NAME);
        sheetField.setPromptText("Excel sheet name (TASK_PLAN_SHEET / TASK_PLAN_SHEET)");
        colWidthField.setText("112");
        hintLabel.setText(HINT_TEXT);

        installStageRunButtonDepth(stage2RunButton, Color.rgb(194, 65, 12, 0.35));
        if (stage2SkipTodayDispatchCheckBox != null) {
            stage2SkipTodayDispatchCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (shell != null) {
                                    shell.scheduleDesktopSessionSave();
                                }
                            });
        }

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
                    markPlanInputTableDirtySinceSave();
                    rebuildSpreadsheet();
                });

        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                spreadsheetHost, headerColumnCount::get);

        rowSearchField
                .textProperty()
                .addListener(
                        (obs, prev, cur) ->
                                SpreadsheetMultiColumnFilterCoordinator.setRowTextSearchQuery(
                                        spreadsheetView, cur));
    }

    /** 実行・ログタブの段階ボタンと同系のごく弱いドロップシャドウ。 */
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

    /** Renumbers dispatch-trial-order column to 1..n after row reorder (DnD, etc.). */
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

    void bindShell(MainShellController shell) {
        this.shell = shell;
        ownerStage = shell.getPrimaryStage();

        columnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::applyDynamicColumnWidths,
                                TableColumnOrderPersistence.TableId.PLAN_INPUT,
                                headerColumnCount,
                                this::onLeadingColumnCountCommitted,
                                this::onReorderColumns,
                                () ->
                                        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                                                ownerStage,
                                                TableColumnOrderPersistence.TableId.PLAN_INPUT,
                                                spreadsheetView,
                                                () -> new ArrayList<>(headersRef))));

        shell.acceptReloadAfterStage1PlanInput(
                () -> {
                    Map<String, String> env = shell.snapshotUiEnv();
                    if (env != null) {
                        pathField.setText(AppPaths.defaultStage1PlanTasksPath(env).toString());
                    }
                    sheetField.setText(AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
                    loadFromCurrentPath(false);
                });

        if (!planInputCellEditHooksInstalled) {
            SpreadsheetPlanInputCellEditSupport.install(
                    spreadsheetView,
                    ownerStage,
                    SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex(),
                    headersRef,
                    rows,
                    () -> {
                        markPlanInputTableDirtySinceSave();
                        rebuildSpreadsheet();
                    });
            planInputCellEditHooksInstalled = true;
        }

        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                spreadsheetView,
                TableColumnOrderPersistence.TableId.PLAN_INPUT,
                suppressColumnOrderPersistence::get,
                () -> new ArrayList<>(headersRef));

        Platform.runLater(
                () -> {
                    syncFromEnv();
                    if (!pathField.getText().isBlank()) {
                        loadFromCurrentPath(false);
                    }
                });
        shell.syncPlanInputStage2ButtonFromDispatchDirty();
    }

    /**
     * 段階1／段階2 実行中は再実行を無効化する（{@link MainShellController#applyRunTabGating} から）。
     */
    void setStageRunProgressVisible(boolean stage1Running, boolean stage2Running) {
        stage2RunPipelineBusy = stage1Running || stage2Running;
        applyStage2RunButtonEnabledState();
    }

    /**
     * 配台計画手動修正の表に未保存の変更があるとき {@code blocked} を true にする（保存または「再読み」で false）。
     */
    void setStage2BlockedByUnsavedDispatchEdit(boolean blocked) {
        stage2BlockedByDispatchUnsavedEdit = blocked;
        applyStage2RunButtonEnabledState();
    }

    /** タスク入力表が「保存」または「再読み」後と同期しているか（段階2実行可否）。 */
    boolean isPlanInputTableDirtySinceSave() {
        return stage2BlockedByUnsavedPlanInputTableEdit;
    }

    private void markPlanInputTableDirtySinceSave() {
        stage2BlockedByUnsavedPlanInputTableEdit = true;
        applyStage2RunButtonEnabledState();
    }

    private void clearPlanInputTableDirtySinceSave() {
        stage2BlockedByUnsavedPlanInputTableEdit = false;
        applyStage2RunButtonEnabledState();
    }

    private void applyStage2RunButtonEnabledState() {
        if (stage2RunButton == null) {
            return;
        }
        boolean disable =
                stage2RunPipelineBusy
                        || stage2BlockedByDispatchUnsavedEdit
                        || stage2BlockedByUnsavedPlanInputTableEdit;
        if (stage2RunButton != null) {
            stage2RunButton.setDisable(disable);
        }
        if (stage2RunPipelineBusy) {
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(null);
            }
        } else if (stage2BlockedByDispatchUnsavedEdit) {
            Tooltip blockedTip =
                    new Tooltip(
                            "配台計画手動修正タブに未保存の変更があります。「保存 (JSON+xlsx)」または「再読み」で確定してから実行してください。");
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(blockedTip);
            }
        } else if (stage2BlockedByUnsavedPlanInputTableEdit) {
            Tooltip blockedTip =
                    new Tooltip(
                            "配台計画_タスク入力タブの表に未保存の変更があります。「保存」または「再読み」で確定してから実行してください。");
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(blockedTip);
            }
        } else {
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(null);
            }
        }
    }

    /** 段階2子プロセスへ渡す {@code PM_AI_STAGE2_SKIP_TODAY_DISPATCH}（チェックは本タブ）。 */
    boolean snapshotStage2SkipTodayDispatch() {
        return stage2SkipTodayDispatchCheckBox != null && stage2SkipTodayDispatchCheckBox.isSelected();
    }

    void applyStage2SkipTodayDispatchFromSession(boolean skipToday) {
        if (stage2SkipTodayDispatchCheckBox != null) {
            stage2SkipTodayDispatchCheckBox.setSelected(skipToday);
        }
    }

    @FXML
    private void onStage2RunButtonAction() {
        if (shell != null) {
            shell.triggerStage2();
        }
    }

    @FXML
    private void onRowUpAction() {
        int i = selectedPlanInputDataIndex();
        if (i <= 0) {
            return;
        }
        int colIdx = planInputFocusedColumnIndex();
        markPlanInputTableDirtySinceSave();
        swapPlanInputDataRows(i - 1, i);
        focusPlanInputCellAfterReorder(i - 1, colIdx);
    }

    @FXML
    private void onRowDownAction() {
        int i = selectedPlanInputDataIndex();
        if (i < 0 || i >= rows.size() - 1) {
            return;
        }
        int colIdx = planInputFocusedColumnIndex();
        markPlanInputTableDirtySinceSave();
        swapPlanInputDataRows(i, i + 1);
        focusPlanInputCellAfterReorder(i + 1, colIdx);
    }

    /** Selected data row index in {@link #rows}, or -1. Uses model row when filters/sort change view order. */
    private int selectedPlanInputDataIndex() {
        var sm = spreadsheetView.getSelectionModel();
        TablePosition pos = sm.getFocusedCell();
        if (pos == null || pos.getRow() < 0) {
            var cells = sm.getSelectedCells();
            if (cells.isEmpty()) {
                return -1;
            }
            pos = cells.getFirst();
        }
        int viewRow = pos.getRow();
        int gridRow = spreadsheetView.getModelRow(viewRow);
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int idx = gridRow - firstData;
        if (idx >= 0 && idx < rows.size()) {
            return idx;
        }
        return -1;
    }

    private int planInputFocusedColumnIndex() {
        var sm = spreadsheetView.getSelectionModel();
        TablePosition pos = sm.getFocusedCell();
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

    private void swapPlanInputDataRows(int a, int b) {
        if (a < 0 || b < 0 || a >= rows.size() || b >= rows.size() || a == b) {
            return;
        }
        ObservableList<String> moved = rows.get(a);
        rows.set(a, rows.get(b));
        rows.set(b, moved);
        renumberDispatchTrialOrderColumn();
        rebuildSpreadsheet();
    }

    /**
     * After reorder, keep selection on the same logical data row and column (handles filtered/sorted view rows).
     */
    private void focusPlanInputCellAfterReorder(int dataRowIndex, int columnIndex) {
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
        markPlanInputTableDirtySinceSave();
        rebuildSpreadsheet();
    }

    private void onReorderColumns() {
        if (headersRef.isEmpty()) {
            shell.appendLog(
                    "[plan-input] "
                            + "ヘッダーが無いため列を"
                            + "並べ替えられません");
            return;
        }
        boolean[] visForDialog =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.PLAN_INPUT, headersRef.size());
        SpreadsheetColumnReorderDialog.show(
                        ownerStage, new ArrayList<>(headersRef), visForDialog)
                .ifPresent(
                        perm -> {
                            List<String> oldHeaders = new ArrayList<>(headersRef);
                            boolean[] oldVis =
                                    TableColumnOrderPersistence.loadColumnVisibility(
                                            TableColumnOrderPersistence.TableId.PLAN_INPUT,
                                            oldHeaders.size());
                            List<String> titleOrder = perm.stream().map(oldHeaders::get).toList();
                            List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
                            TableColumnOrderPersistence.applyLogicalColumnOrder(
                                    headersRef, rows, titleOrder);
                            boolean[] newVis =
                                    TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                                            oldHeaders, oldVis, titleOrder);
                            TableColumnOrderPersistence.saveColumnVisibility(
                                    TableColumnOrderPersistence.TableId.PLAN_INPUT, newVis);
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
                                    TableColumnOrderPersistence.TableId.PLAN_INPUT, newLay);
                            markPlanInputTableDirtySinceSave();
                            rebuildSpreadsheet(false);
                        });
    }

    @FXML
    private void onBrowseButtonAction() {
        FileChooser ch = new FileChooser();
        ch.setTitle(
                "配台計画_タスク入力 — "
                        + "ファイルを開く");
        ch.getExtensionFilters()
                .addAll(
                        new FileChooser.ExtensionFilter("Tabular", "*.csv", "*.xlsx", "*.xlsm"),
                        new FileChooser.ExtensionFilter("All", "*.*"));
        File f = ch.showOpenDialog(ownerStage);
        if (f != null) {
            pathField.setText(f.getAbsolutePath());
        }
    }

    @FXML
    private void onLoadButtonAction() {
        syncFromEnv();
        loadFromCurrentPath(true);
    }

    @FXML
    private void onSaveButtonAction() {
        if (pathField.getText().isBlank()) {
            shell.appendLog("[plan-input] save: path is empty");
            shell.showWarningDialog("保存", "保存先のパスが空です。");
            return;
        }
        Path path = Path.of(pathField.getText().trim());
        try {
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
            PlanInputTabularIo.write(
                    path,
                    sheetField.getText().trim().isEmpty()
                            ? DEFAULT_PLAN_INPUT_SHEET_NAME
                            : sheetField.getText().trim(),
                    new PlanInputTabularIo.TabularSheet(headersRef, dataRows));
            shell.appendLog("[plan-input] saved " + path);
            clearPlanInputTableDirtySinceSave();
            shell.showInformationDialog(
                    "保存完了",
                    "配台計画_タスク入力を保存しました。\n"
                            + path
                            + "\n行数: "
                            + dataRows.size());
        } catch (Exception ex) {
            shell.appendLog("[plan-input] save error: " + ex.getMessage());
            shell.showErrorDialog(
                    "保存エラー",
                    ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onAddRowButtonAction() {
        if (headersRef.isEmpty()) {
            shell.appendLog("[plan-input] load a file first (headers required)");
            return;
        }
        ObservableList<String> r = FXCollections.observableArrayList();
        for (int i = 0; i < headersRef.size(); i++) {
            r.add("");
        }
        rows.add(r);
        markPlanInputTableDirtySinceSave();
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
        shell.appendLog("[plan-input] removed " + sorted.size() + " row(s)");
        markPlanInputTableDirtySinceSave();
        rebuildSpreadsheet();
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
                        return Stage2RollUnitLengthTables.load(
                                AppPaths.resolveRepoRoot(shell.snapshotUiEnv()));
                    } catch (Exception e) {
                        return Stage2RollUnitLengthTables.empty();
                    }
                });
    }

    private void rebuildSpreadsheet() {
        rebuildSpreadsheet(true);
    }

    /**
     * @param preserveColumnFilters {@code true} のとき、再構築前の列フィルタ（許容値集合）を復元する。ファイル読込・
     *     論理列並べ替え後は {@code false}（列インデックスが変わるため）。
     */
    private void rebuildSpreadsheet(boolean preserveColumnFilters) {
        if (headersRef.isEmpty()) {
            detachGridHandler();
            GridBase empty = new GridBase(0, 0);
            // DnD 直後など setGrid 内の選択検証が旧インデックスで IndexOutOfBounds になるのを防ぐ
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
        suppressPlanInputDirtyFromGridEvents.set(true);
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
                        if (!suppressPlanInputDirtyFromGridEvents.get()) {
                            markPlanInputTableDirtySinceSave();
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
                                                    TableColumnOrderPersistence.TableId.PLAN_INPUT,
                                                    headersRef.size()));
                            String q =
                                    rowSearchField.getText() != null
                                            ? rowSearchField.getText().trim()
                                            : "";
                            if (!q.isEmpty()) {
                                SpreadsheetMultiColumnFilterCoordinator.setRowTextSearchQuery(
                                        spreadsheetView, q);
                            }
                        } finally {
                            suppressPlanInputDirtyFromGridEvents.set(false);
                        }
                    });
        } catch (Throwable t) {
            suppressPlanInputDirtyFromGridEvents.set(false);
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

    private void applyLoaded() {
        rebuildSpreadsheet(false);
    }

    private void syncFromEnv() {
        Map<String, String> env = shell.snapshotUiEnv();
        if (env != null) {
            String p = trim(env.get(ENV_PM_AI_PLAN_INPUT_PATH));
            if (!p.isEmpty() && pathField.getText().isBlank()) {
                pathField.setText(p);
            }
            String sh = trim(env.get(ENV_TASK_PLAN_SHEET));
            if (!sh.isEmpty() && sheetField.getText().isBlank()) {
                sheetField.setText(sh);
            }
        }
    }

    private void loadFromCurrentPath(boolean showCompletionDialog) {
        Path path = Path.of(pathField.getText().trim());
        if (!java.nio.file.Files.isRegularFile(path)) {
            shell.appendLog("[plan-input] file not found: " + path);
            if (showCompletionDialog) {
                shell.showWarningDialog("読込", "ファイルが見つかりません。\n" + path);
            }
            return;
        }
        String sheetName = sheetField.getText().trim();
        if (sheetName.isEmpty()) {
            sheetName = DEFAULT_PLAN_INPUT_SHEET_NAME;
        }
        try {
            PlanInputTabularIo.TabularSheet sheet = PlanInputTabularIo.read(path, sheetName);
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
            List<TableColumnOrderPersistence.ColumnSpec> lay =
                    TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.PLAN_INPUT);
            persistedLayout.set(lay);
            List<String> beforeHeaders = new ArrayList<>(headersRef);
            boolean[] visBefore =
                    TableColumnOrderPersistence.loadColumnVisibility(
                            TableColumnOrderPersistence.TableId.PLAN_INPUT, beforeHeaders.size());
            List<String> titleOrder =
                    lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
            TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
            boolean[] visAfter =
                    TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                            beforeHeaders, visBefore, titleOrder);
            TableColumnOrderPersistence.saveColumnVisibility(
                    TableColumnOrderPersistence.TableId.PLAN_INPUT, visAfter);
            applyLoaded();
            clearPlanInputTableDirtySinceSave();
            shell.appendLog(
                    "[plan-input] loaded rows="
                            + rows.size()
                            + " cols="
                            + headersRef.size()
                            + " path="
                            + path);
            if (showCompletionDialog) {
                shell.showInformationDialog(
                        "読込完了",
                        "配台計画_タスク入力を読み込みました。\n"
                                + path
                                + "\n行数: "
                                + rows.size()
                                + " / 列数: "
                                + headersRef.size());
            }
        } catch (Exception ex) {
            shell.appendLog("[plan-input] load error: " + ex.getMessage());
            if (showCompletionDialog) {
                shell.showErrorDialog(
                        "読込エラー",
                        ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
        }
    }

    private static String trim(String s) {
        return s != null ? s.trim() : "";
    }

    String snapshotPlanInputPath() {
        return pathField.getText() != null ? pathField.getText().trim() : "";
    }

    String snapshotPlanInputSheet() {
        return sheetField.getText() != null ? sheetField.getText().trim() : "";
    }

    void restoreDesktopSessionPaths(String path, String sheet) {
        if (path != null && !path.isBlank()) {
            pathField.setText(path.trim());
        }
        if (sheet != null && !sheet.isBlank()) {
            sheetField.setText(sheet.trim());
        }
    }

    void clearColumnFiltersAndSort() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(spreadsheetView);
    }

    @FXML
    private void onClearColumnFiltersAction() {
        clearColumnFiltersAndSort();
        rowSearchField.clear();
    }
}
