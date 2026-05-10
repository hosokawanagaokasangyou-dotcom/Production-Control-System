package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
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
 * 環境変数 {@link AppPaths#KEY_PM_AI_TASK_INPUT_SOURCE_DIR} で指定するソースフォルダから最新の加工計画ファイルを読み込み、
 * {@link SpreadsheetView} に表示する。レイアウトは {@code AladdinProcessingPlanDataTab.fxml}。
 */

public final class AladdinProcessingPlanDataTabController {

    private static final String HINT_TEXT =
            "PM_AI_TASK_INPUT_SOURCE_DIR で解決したフォルダ内の最新ファイル（csv / xlsx）から加工計画DATAを読み込みます。"
                    + " 複数ある場合は更新日時が最も新しいものを使います。"
                    + " Excel は「Excel シート」でシートを選べます。Parquet は未対応です。";


    @FXML
    private Label statusLabel;

    @FXML
    private Label dirLabel;

    @FXML
    private Label pathLabel;

    @FXML
    private ComboBox<String> sheetCombo;

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

    @FXML
    private void initialize() {
        hintLabel.setText(HINT_TEXT);
        rows = FXCollections.observableArrayList();

        StackPane.setAlignment(spreadsheetView, Pos.CENTER_LEFT);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
        SpreadsheetThemeBridge.install(spreadsheetView);

        columnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::applyDynamicColumnWidths,
                                TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW,
                                headerColumnCount,
                                this::onLeadingColumnCountCommitted,
                                this::onReorderColumns,
                                () ->
                                        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                                                ownerStage,
                                                TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW,
                                                spreadsheetView,
                                                () -> new ArrayList<>(headersRef))));

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
                            Platform.runLater(() -> applyLoadedFile(loadedPath, idx, false));
                        });
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        ownerStage = shell.getPrimaryStage();

        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                spreadsheetView,
                TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW,
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
                        TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW);
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
                TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW, next);
        if (rowHeightPctLabel != null) {
            rowHeightPctLabel.setText(String.format("%.0f%%", v));
        }
        rebuildSpreadsheet();
    }

    private void onReorderColumns() {
        if (headersRef.isEmpty()) {
            if (shell != null) {
                shell.appendLog("[aladdin-plan-data] 列見出しが空のため並べ替えできません");
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
                        TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW, oldHeaders.size());
        List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
        TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(oldHeaders, oldVis, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW, newVis);
        List<Double> widths =
                TableColumnOrderPersistence.resolveWidthsForHeaders(headersRef, lay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
        for (int i = 0; i < headersRef.size(); i++) {
            newLay.add(new TableColumnOrderPersistence.ColumnSpec(headersRef.get(i), widths.get(i)));
        }
        persistedLayout.set(newLay);
        TableColumnOrderPersistence.saveLayout(
                TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW, newLay);
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
                        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(spreadsheetView);
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
                                        TableColumnOrderPersistence.loadColumnVisibility(
                                                TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW,
                                                headersRef.size()));
                    });
        } finally {
            suppressColumnPersistence.set(false);
        }
    }

    /** 納期管理ビュー上部の「再読込」成功後に親から呼ばれる。 */
    public void reloadAladdinProcessingPlanFromDisk() {
        reloadFromSourceDir();
    }

    private void reloadFromSourceDir() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path dir = AppPaths.resolveTaskInputSourceDir(ui);
        dirLabel.setText(dir != null ? dir.toString() : "(未設定)");
        if (dir == null || !Files.isDirectory(dir)) {
            statusLabel.setText("ソースフォルダがありません");
            pathLabel.setText("");
            sheetCombo.setDisable(true);
            sheetCombo.getItems().clear();
            loadedPath = null;
            applyEmpty();
            return;
        }
        Optional<Path> newest = NetworkSourceDirResolver.newestTaskInputFileInDirectory(dir);
        if (newest.isEmpty()) {
            statusLabel.setText("読込対象ファイルがありません");
            pathLabel.setText("");
            sheetCombo.setDisable(true);
            sheetCombo.getItems().clear();
            loadedPath = null;
            applyEmpty();
            return;
        }
        Path file = newest.get().toAbsolutePath().normalize();
        loadedPath = file;
        pathLabel.setText(file.toString());

        String low = file.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".pq") || low.endsWith(".parquet")) {
            statusLabel.setText("Parquet は未対応です");
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
                }
            } catch (IOException ex) {
                statusLabel.setText("Excel シート一覧の取得に失敗しました");
                if (shell != null) {
                    shell.appendLog(
                            "[aladdin-plan-data] "
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
    }

    private void applyLoadedFile(Path file, int excelSheetIndex, boolean showErrorsInStatus) {
        try {
            PlanInputTabularIo.TabularSheet tab =
                    TaskInputSourceRawGridIo.applyAladdinProcessingPlanDisplaySteps(
                            TaskInputSourceRawGridIo.readRaw(file, excelSheetIndex));
            // Persist shaped plan (pre-column-order) for DeliveryCalendarView overlay JSON cache
            if (shell != null) {
                try {
                    java.nio.file.Path savePath =
                            AppPaths.resolveShapedAladdinPlanJsonPath(shell.snapshotUiEnv());
                    JsonTableIo.saveArrayTable(savePath, tab.headers(), tab.rows());
                } catch (Exception saveEx) {
                    shell.appendLog(
                            "[aladdin-plan-data] shaped JSON save failed: "
                                    + saveEx.getMessage());
                }
            }
            List<TableColumnOrderPersistence.ColumnSpec> lay =
                    TableColumnOrderPersistence.loadLayout(
                            TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW);
            persistedLayout.set(lay);
            List<String> beforeHeaders = new ArrayList<>(tab.headers());
            boolean[] visBefore =
                    TableColumnOrderPersistence.loadColumnVisibility(
                            TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW,
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
                    TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW, visAfter);

            statusLabel.setText(rows.size() + " 行 × " + headersRef.size() + " 列");
            rebuildSpreadsheet();
        } catch (Exception ex) {
            if (showErrorsInStatus) {
                statusLabel.setText("読込エラー");
            }
            if (shell != null) {
                shell.appendLog(
                        "[aladdin-plan-data] "
                                + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            }
            applyEmpty();
        }
    }

    private void applyEmpty() {
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
