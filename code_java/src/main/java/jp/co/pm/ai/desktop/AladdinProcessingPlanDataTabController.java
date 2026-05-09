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
import javafx.scene.text.Text;
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
 * {@link AppPaths#KEY_PM_AI_TASK_INPUT_SOURCE_DIR} 直下�?�最新タスク入力ファイルを読み、�?��?�シート（また�?� CSV 全体）を
 * 生表として {@link SpreadsheetView} に表示する。レイアウト�?� {@code AladdinProcessingPlanDataTab.fxml}�?
 */
public final class AladdinProcessingPlanDataTabController {

    private static final String HINT_TEXT =
            "PM_AI_TASK_INPUT_SOURCE_DIR で�?定したフォルダ�?で、加工計画DATA相当�?�拡張子�?csv / xlsx 等）�?�"
                    + "更新時刻が最新の1ファイルを表示します�?Excel は列見�?�しを「�??1…」とした�?えでシート上�?�全行を�?ータ行とします�?"
                    + " ネットワーク未到達時はフォルダが開けず空表示になります�?";

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

    @FXML
    private Text metaText;

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

        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
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
                shell.appendLog("[aladdin-plan-data] 列がありません?���?�に再読み?�?");
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
                                                TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW,
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
            Path dir = AppPaths.resolveTaskInputSourceDir(ui);
            dirLabel.setText(dir != null ? dir.toString() : "(未設�?)");
            if (dir == null || !Files.isDirectory(dir)) {
                statusLabel.setText("フォルダなしまた�?�未到�?");
                metaText.setText("");
                pathLabel.setText("");
                sheetCombo.setDisable(true);
                sheetCombo.getItems().clear();
                loadedPath = null;
                applyEmpty();
                return;
            }
            Optional<Path> newest = NetworkSourceDirResolver.newestTaskInputFileInDirectory(dir);
            if (newest.isEmpty()) {
                statusLabel.setText("該当ファイルな�?");
                metaText.setText("");
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
                statusLabel.setText("Parquet は未対応で�?");
                metaText.setText(file.toString());
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
                    statusLabel.setText("シート一覧エラー");
                    metaText.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
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

    private void applyLoadedFile(Path file, int excelSheetIndex, boolean showErrorsInStatus) {
        try {
            PlanInputTabularIo.TabularSheet tab = TaskInputSourceRawGridIo.readRaw(file, excelSheetIndex);
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

            statusLabel.setText(rows.size() + " �? �? " + headersRef.size() + " �?");
            metaText.setText(
                    "file="
                            + file
                            + (isExcelPath(file)
                                    ? ", sheetIndex=" + excelSheetIndex
                                    : ", csv"));

            rebuildSpreadsheet();
        } catch (Exception ex) {
            if (showErrorsInStatus) {
                statusLabel.setText("読込エラー");
                metaText.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
            if (shell != null) {
                shell.appendLog("[aladdin-plan-data] " + ex.getMessage());
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
