package jp.co.pm.ai.desktop;

import java.io.File;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
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
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
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
import jp.co.pm.ai.desktop.ui.SpreadsheetPlanInputCellEditSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetPlanInputRowDragSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * 配台計画_タスク入力タブ。レイアウトは {@code PlanInputTab.fxml}。
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
    private HBox columnStripHost;

    @FXML
    private TextField colWidthField;

    @FXML
    private Label hintLabel;

    @FXML
    private StackPane spreadsheetHost;

    private final SpreadsheetView spreadsheetView = new SpreadsheetView();

    private final List<String> headersRef = new ArrayList<>();
    private ObservableList<ObservableList<String>> rows;
    private final AtomicBoolean suppressColumnOrderPersistence = new AtomicBoolean(false);
    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());
    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    private GridBase currentGrid;
    private EventHandler<GridChange> gridChangeHandler;

    private boolean planInputCellEditHooksInstalled;

    @FXML
    private void initialize() {
        pathField.setPromptText(
                "PM_AI_PLAN_INPUT_PATH （.csv / .xlsx / .xlsm）");
        sheetField.setText(DEFAULT_PLAN_INPUT_SHEET_NAME);
        sheetField.setPromptText("Excel sheet name (TASK_PLAN_SHEET / TASK_PLAN_SHEET)");
        colWidthField.setText("112");
        hintLabel.setText(HINT_TEXT);

        StackPane.setAlignment(spreadsheetView, Pos.CENTER_LEFT);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        rows = FXCollections.observableArrayList();
        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        SpreadsheetTabularSupport.installFullRowDataSelection(
                spreadsheetView,
                SpreadsheetPlanInputRowDragSupport::skipFullRowExpansionDuringPlanInputRowDrag);
        SpreadsheetThemeBridge.install(spreadsheetView);
        SpreadsheetPlanInputRowDragSupport.install(
                spreadsheetView,
                SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex(),
                rows,
                () -> {
                    renumberDispatchTrialOrderColumn();
                    rebuildSpreadsheet();
                });
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
                    loadFromCurrentPath();
                });

        if (!planInputCellEditHooksInstalled) {
            SpreadsheetPlanInputCellEditSupport.install(
                    spreadsheetView,
                    ownerStage,
                    SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex(),
                    headersRef,
                    rows,
                    this::rebuildSpreadsheet);
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
                        loadFromCurrentPath();
                    }
                });
    }

    @FXML
    private void onRowUpAction() {
        int i = selectedPlanInputDataIndex();
        if (i <= 0) {
            return;
        }
        int colIdx = planInputFocusedColumnIndex();
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
                        SpreadsheetColumnReorderDialog.show(ownerStage, new ArrayList<>(headersRef))
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
                            rebuildSpreadsheet();
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
        loadFromCurrentPath();
    }

    @FXML
    private void onSaveButtonAction() {
        Path path = Path.of(pathField.getText().trim());
        if (pathField.getText().isBlank()) {
            shell.appendLog("[plan-input] save: path is empty");
            return;
        }
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
        } catch (Exception ex) {
            shell.appendLog("[plan-input] save error: " + ex.getMessage());
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

    private void rebuildSpreadsheet() {
        if (headersRef.isEmpty()) {
            detachGridHandler();
            GridBase empty = new GridBase(0, 0);
            spreadsheetView.setGrid(empty);
            currentGrid = empty;
            return;
        }
        suppressColumnOrderPersistence.set(true);
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

            GridBase grid = SpreadsheetTabularSupport.buildPlanInputGrid(headersRef, rows, false);
            gridChangeHandler =
                    SpreadsheetTabularSupport.newRowsSyncHandler(
                            rows, headersRef, SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex());
            grid.addEventHandler(GridChange.GRID_CHANGE_EVENT, gridChangeHandler);
            currentGrid = grid;
            spreadsheetView.setGrid(grid);

            Platform.runLater(
                    () -> {
                        SpreadsheetTabularSupport.applyColumnWidths(spreadsheetView, widths, widthDefault);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                spreadsheetView, headerColumnCount.get());
                        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(spreadsheetView);
                        SpreadsheetTabularSupport.pinSpreadsheetFilterRow(spreadsheetView);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(spreadsheetView);
                        ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                                spreadsheetView,
                                () -> new ArrayList<>(headersRef),
                                () ->
                                        TableColumnOrderPersistence.loadColumnVisibility(
                                                TableColumnOrderPersistence.TableId.PLAN_INPUT,
                                                headersRef.size()));
                    });
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
        rebuildSpreadsheet();
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

    private void loadFromCurrentPath() {
        Path path = Path.of(pathField.getText().trim());
        if (!java.nio.file.Files.isRegularFile(path)) {
            shell.appendLog("[plan-input] file not found: " + path);
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
            shell.appendLog(
                    "[plan-input] loaded rows="
                            + rows.size()
                            + " cols="
                            + headersRef.size()
                            + " path="
                            + path);
        } catch (Exception ex) {
            shell.appendLog("[plan-input] load error: " + ex.getMessage());
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
    }
}
