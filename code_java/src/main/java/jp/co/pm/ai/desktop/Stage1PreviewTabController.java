package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnDragReorderSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/** Stage1 task-input preview table; layout {@code Stage1PreviewTab.fxml}. Uses ControlsFX {@link SpreadsheetView}. */
public final class Stage1PreviewTabController {

    public static final String DEFAULT_STAGE1_PREVIEW_SHEET = AppPaths.STAGE1_TASK_INPUT_PREVIEW_SHEET;

    private Stage ownerStage;

    private MainShellController shell;

    @FXML
    private Label hintLabel;

    @FXML
    private TextField pathField;

    @FXML
    private TextField sheetField;

    @FXML
    private Button fromEnvButton;

    @FXML
    private Button browseButton;

    @FXML
    private Button loadButton;

    @FXML
    private HBox columnStripHost;

    @FXML
    private TextField colWidthField;

    @FXML
    private Slider stage1RowHeightSlider;

    @FXML
    private Label stage1RowHeightPctLabel;

    @FXML
    private CheckBox stage1CellWrapCheck;

    @FXML
    private StackPane spreadsheetHost;

    private final SpreadsheetView spreadsheetView = new SpreadsheetView();

    private final List<String> headersRef = new ArrayList<>();
    private ObservableList<ObservableList<String>> rows;
    private final AtomicBoolean suppressColumnOrderPersistence = new AtomicBoolean(false);
    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());
    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    private final AtomicReference<TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs>
            spreadsheetTabPrefs =
                    new AtomicReference<>(
                            TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs.defaults());

    private final AtomicBoolean suppressPresentationUiEvents = new AtomicBoolean(false);

    private volatile boolean stage1PresentationHooksInstalled;

    @FXML
    private void initialize() {
        pathField.setPromptText(
                "PM_AI_OUTPUT_DIR またはリポ根直下/output/"
                        + AppPaths.STAGE1_TASK_INPUT_PREVIEW_FILENAME
                        + " (問合せ取込後・タスク一覧化前)");
        sheetField.setText(DEFAULT_STAGE1_PREVIEW_SHEET);
        sheetField.setPromptText("Excel sheet name");
        colWidthField.setText("112");

        hintLabel.setText(buildHintText());

        StackPane.setAlignment(spreadsheetView, Pos.CENTER_LEFT);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        rows = FXCollections.observableArrayList();
        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        SpreadsheetThemeBridge.install(spreadsheetView);
    }

    private static String buildHintText() {
        return "問合せ xlsx を読み込み、ヘッダー行と"
                + "列名を整えた直後（依頼NO"
                + "がある行のみ）。"
                + " マスタ・配台試行順付与前の"
                + " stage1_task_input_table.xlsx"
                + " シート「"
                + DEFAULT_STAGE1_PREVIEW_SHEET
                + "」を表示します。";
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        ownerStage = shell.getPrimaryStage();

        columnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::applyDynamicColumnWidths,
                                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW,
                                headerColumnCount,
                                this::onLeadingColumnCountCommitted,
                                this::onReorderColumns,
                                () ->
                                        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                                                ownerStage,
                                                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW,
                                                spreadsheetView,
                                                () -> new ArrayList<>(headersRef))));

        shell.acceptReloadAfterStage1Preview(
                () -> {
                    fillPathFromEnv();
                    sheetField.setText(DEFAULT_STAGE1_PREVIEW_SHEET);
                    loadButton.fire();
                });

        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                spreadsheetView,
                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW,
                suppressColumnOrderPersistence::get,
                () -> new ArrayList<>(headersRef));

        initStage1SpreadsheetPresentationControls();

        javafx.application.Platform.runLater(
                () -> {
                    if (pathField.getText().isBlank()) {
                        fillPathFromEnv();
                    }
                    if (!pathField.getText().isBlank()) {
                        loadButton.fire();
                    }
                });
    }

    private void onLeadingColumnCountCommitted(int n) {
        headerColumnCount.set(n);
        rebuildSpreadsheet();
    }

    private void initStage1SpreadsheetPresentationControls() {
        if (stage1PresentationHooksInstalled) {
            return;
        }
        if (stage1RowHeightSlider == null) {
            return;
        }
        stage1PresentationHooksInstalled = true;
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs loaded =
                TableColumnOrderPersistence.loadSpreadsheetTabPresentationPrefs(
                        TableColumnOrderPersistence.TableId.STAGE1_PREVIEW);
        spreadsheetTabPrefs.set(loaded);
        suppressPresentationUiEvents.set(true);
        try {
            stage1RowHeightSlider.setMin(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN);
            stage1RowHeightSlider.setMax(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX);
            stage1RowHeightSlider.setValue(loaded.rowHeightPercent());
            stage1RowHeightSlider.setMajorTickUnit(250);
            stage1RowHeightSlider.setMinorTickCount(4);
            stage1RowHeightSlider.setShowTickMarks(true);
            if (stage1RowHeightPctLabel != null) {
                stage1RowHeightPctLabel.setText(
                        String.format("%.0f%%", loaded.rowHeightPercent()));
            }
            if (stage1CellWrapCheck != null) {
                stage1CellWrapCheck.setSelected(loaded.cellWrapText());
            }
        } finally {
            suppressPresentationUiEvents.set(false);
        }
        SliderCommittedChangeSupport.install(
                stage1RowHeightSlider,
                () -> {
                    if (stage1RowHeightPctLabel != null && stage1RowHeightSlider != null) {
                        stage1RowHeightPctLabel.setText(
                                String.format("%.0f%%", stage1RowHeightSlider.getValue()));
                    }
                },
                this::commitStage1SpreadsheetPresentationFromSlider);
        if (stage1CellWrapCheck != null) {
            stage1CellWrapCheck
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (suppressPresentationUiEvents.get()) {
                                    return;
                                }
                                commitStage1SpreadsheetPresentationFromUi();
                            });
        }
    }

    private void commitStage1SpreadsheetPresentationFromSlider() {
        if (suppressPresentationUiEvents.get()) {
            return;
        }
        commitStage1SpreadsheetPresentationFromUi();
    }

    private void commitStage1SpreadsheetPresentationFromUi() {
        if (stage1RowHeightSlider == null) {
            return;
        }
        double v = stage1RowHeightSlider.getValue();
        boolean wrap = stage1CellWrapCheck != null && stage1CellWrapCheck.isSelected();
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs next =
                new TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs(v, wrap);
        spreadsheetTabPrefs.set(next);
        TableColumnOrderPersistence.saveSpreadsheetTabPresentationPrefs(
                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW, next);
        if (stage1RowHeightPctLabel != null) {
            stage1RowHeightPctLabel.setText(String.format("%.0f%%", v));
        }
        rebuildSpreadsheet();
    }

    private void onReorderColumns() {
        if (headersRef.isEmpty()) {
            shell.appendLog("[stage1-preview] 列がありません（先に読み込み）");
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

    /**
     * ダイアログまたはヘッダードラッグで確定した見出し順へ論理列を揃え、列幅レイアウトを保存して再構築する。
     */
    private void applyPersistedColumnOrderAfterLogicalReorder(List<String> titleOrder) {
        if (headersRef.isEmpty()) {
            return;
        }
        List<String> oldHeaders = new ArrayList<>(headersRef);
        boolean[] oldVis =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId.STAGE1_PREVIEW, oldHeaders.size());
        List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
        TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        oldHeaders, oldVis, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW, newVis);
        double colW = readColWidthFieldOrDefault();
        List<Double> widths =
                TableColumnOrderPersistence.resolveWidthsForHeaders(headersRef, lay, colW);
        List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
        for (int i = 0; i < headersRef.size(); i++) {
            newLay.add(
                    new TableColumnOrderPersistence.ColumnSpec(headersRef.get(i), widths.get(i)));
        }
        persistedLayout.set(newLay);
        TableColumnOrderPersistence.saveLayout(
                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW, newLay);
        rebuildSpreadsheet();
    }

    private double readColWidthFieldOrDefault() {
        double colW = 112;
        try {
            colW = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
        } catch (NumberFormatException ignored) {
        }
        return colW;
    }

    @FXML
    private void onFromEnvButtonAction() {
        fillPathFromEnv();
    }

    @FXML
    private void onBrowseButtonAction() {
        FileChooser ch = new FileChooser();
        ch.setTitle(AppPaths.STAGE1_TASK_INPUT_PREVIEW_FILENAME);
        ch.getExtensionFilters()
                .addAll(
                        new FileChooser.ExtensionFilter("Excel", "*.xlsx", "*.xlsm"),
                        new FileChooser.ExtensionFilter("All", "*.*"));
        String cur = pathField.getText();
        if (cur != null && !cur.isBlank()) {
            try {
                Path p = Path.of(cur.trim());
                if (java.nio.file.Files.isRegularFile(p) && p.getParent() != null) {
                    ch.setInitialDirectory(p.getParent().toFile());
                }
            } catch (Exception ignored) {
            }
        }
        var f = ch.showOpenDialog(ownerStage);
        if (f != null) {
            pathField.setText(f.getAbsolutePath());
        }
    }

    @FXML
    private void onLoadButtonAction() {
        if (pathField.getText().isBlank()) {
            fillPathFromEnv();
        }
        Path path = Path.of(pathField.getText().trim());
        if (!java.nio.file.Files.isRegularFile(path)) {
            shell.appendLog("[stage1-preview] file not found: " + path);
            return;
        }
        String sheet = sheetField.getText().trim();
        if (sheet.isEmpty()) {
            sheet = DEFAULT_STAGE1_PREVIEW_SHEET;
        }
        try {
            PlanInputTabularIo.TabularSheet sh = PlanInputTabularIo.read(path, sheet);
            headersRef.clear();
            headersRef.addAll(sh.headers());
            rows.clear();
            for (List<String> line : sh.rows()) {
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
                    TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.STAGE1_PREVIEW);
            persistedLayout.set(lay);
            List<String> beforeHeaders = new ArrayList<>(headersRef);
            boolean[] visBefore =
                    TableColumnOrderPersistence.loadColumnVisibility(
                            TableColumnOrderPersistence.TableId.STAGE1_PREVIEW, beforeHeaders.size());
            List<String> titleOrder =
                    lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
            TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
            boolean[] visAfter =
                    TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                            beforeHeaders, visBefore, titleOrder);
            TableColumnOrderPersistence.saveColumnVisibility(
                    TableColumnOrderPersistence.TableId.STAGE1_PREVIEW, visAfter);
            applyLoaded();
            shell.appendLog(
                    "[stage1-preview] loaded rows="
                            + rows.size()
                            + " cols="
                            + headersRef.size()
                            + " path="
                            + path);
        } catch (Exception ex) {
            shell.appendLog("[stage1-preview] load error: " + ex.getMessage());
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

    private void rebuildSpreadsheet() {
        if (headersRef.isEmpty()) {
            spreadsheetView.setGrid(new GridBase(0, 0));
            return;
        }
        suppressColumnOrderPersistence.set(true);
        try {
            double colW = readColWidthFieldOrDefault();
            final List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            headersRef, persistedLayout.get(), colW);
            final double widthDefault = colW;

            GridBase grid = SpreadsheetTabularSupport.buildStage1PreviewGrid(headersRef, rows);
            TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs pres =
                    spreadsheetTabPrefs.get();
            SpreadsheetTabularSupport.applySpreadsheetGridRowHeightsAndWrap(
                    grid, pres.cellWrapText(), pres.rowHeightPercent());
            spreadsheetView.setGrid(grid);

            javafx.application.Platform.runLater(
                    () -> {
                        SpreadsheetTabularSupport.applyColumnWidths(spreadsheetView, widths, widthDefault);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(spreadsheetView);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                spreadsheetView, headerColumnCount.get());
                        SpreadsheetTabularSupport.applyColumnFilters(spreadsheetView);
                        SpreadsheetTabularSupport.refreshSpreadsheetAfterRowPresentationChange(
                                spreadsheetView);
                        SpreadsheetColumnDragReorderSupport.refreshAfterGridReady(
                                spreadsheetView,
                                suppressColumnOrderPersistence::get,
                                () -> new ArrayList<>(headersRef),
                                headerColumnCount.get(),
                                this::applyPersistedColumnOrderAfterLogicalReorder);
                        ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                                spreadsheetView,
                                () -> new ArrayList<>(headersRef),
                                () ->
                                        TableColumnOrderPersistence.loadColumnVisibility(
                                                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW,
                                                headersRef.size()));
                    });
        } finally {
            suppressColumnOrderPersistence.set(false);
        }
    }

    private void applyLoaded() {
        rebuildSpreadsheet();
    }

    private void fillPathFromEnv() {
        Map<String, String> env = shell.snapshotUiEnv();
        if (env == null) {
            return;
        }
        pathField.setText(AppPaths.defaultStage1TaskInputPreviewPath(env).toString());
    }

    String snapshotStage1PreviewPath() {
        return pathField.getText() != null ? pathField.getText().trim() : "";
    }

    String snapshotStage1PreviewSheet() {
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
