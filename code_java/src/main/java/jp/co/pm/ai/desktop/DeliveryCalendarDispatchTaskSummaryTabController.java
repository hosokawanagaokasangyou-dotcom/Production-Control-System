package jp.co.pm.ai.desktop;

import java.nio.charset.StandardCharsets;
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

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.ListChangeListener;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TableView;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchDeadlineJudgment;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchInteractiveConsolidator;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchStage3Support;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchTaskSummaryConsolidator;
import jp.co.pm.ai.desktop.dispatch.TaskIdLeadingAlphaPrefix;
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
 * 納期管理ビュー「配台結果（タスク集約）」: 日別配台結果を 1 タスク 1 行にまとめ、依頼NO 先頭英字でタブ分割する。
 */
public final class DeliveryCalendarDispatchTaskSummaryTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String COL_TID = "依頼NO";
    private static final String COL_START = "加工開始日時";
    private static final String COL_END = "加工終了日時";

    private static final List<String> SORT_COLUMN_LABELS =
            List.of(COL_TID, COL_START, COL_END);

    @FXML
    private Button refreshButton;

    @FXML
    private Label dataStageBadgeLabel;

    @FXML
    private Label statusLabel;

    @FXML
    private Label pathLabel;

    @FXML
    private ComboBox<String> sortColumnCombo;

    @FXML
    private ToggleButton sortDescendingToggle;

    @FXML
    private Slider rowHeightSlider;

    @FXML
    private Label rowHeightPctLabel;

    @FXML
    private CheckBox cellWrapCheck;

    @FXML
    private HBox columnStripHost;

    @FXML
    private TabPane prefixTabPane;

    @FXML
    private StackPane spreadsheetHost;

    @FXML
    private Label dailyScheduleTitleLabel;

    @FXML
    private TableView<DailyScheduleRow> dailyScheduleTable;

    private MainShellController shell;

    private Stage ownerStage;

    private final SpreadsheetView spreadsheetView = new SpreadsheetView();

    private final List<String> headersRef = new ArrayList<>();

    private ObservableList<ObservableList<String>> rows;

    /** 接頭辞タブ選択前の全集約行（列順は {@link #headersRef}）。 */
    private final List<Map<String, String>> consolidatedRows = new ArrayList<>();

    /** 表示中の集約行（{@link #rows} と同順）。 */
    private final List<Map<String, String>> displayedConsolidatedRows = new ArrayList<>();

    /** タスクキー → 日別配台行（reload 時に構築）。 */
    private final Map<String, List<Map<String, String>>> dailyRowsByGroupKey = new LinkedHashMap<>();

    private final ObservableList<DailyScheduleRow> dailyScheduleItems =
            FXCollections.observableArrayList();

    private final AtomicBoolean suppressColumnPersistence = new AtomicBoolean(false);

    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());

    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    private final AtomicReference<TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs>
            spreadsheetTabPrefs =
                    new AtomicReference<>(
                            TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs.defaults());

    private final AtomicBoolean suppressPresentationUiEvents = new AtomicBoolean(false);

    private final AtomicBoolean suppressPrefixTabEvents = new AtomicBoolean(false);

    private volatile boolean presentationHooksInstalled;

    @FXML
    private void initialize() {
        rows = FXCollections.observableArrayList();

        StackPane.setAlignment(spreadsheetView, Pos.CENTER_LEFT);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
        SpreadsheetThemeBridge.install(spreadsheetView);
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(spreadsheetView);
        installDailyScheduleSelectionListener();
        initDailyScheduleTable();

        sortColumnCombo.setItems(FXCollections.observableArrayList(SORT_COLUMN_LABELS));
        sortColumnCombo.getSelectionModel().select(COL_TID);
        sortColumnCombo
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((obs, a, b) -> applyPrefixFilterAndSort());
        sortDescendingToggle
                .selectedProperty()
                .addListener((obs, a, b) -> applyPrefixFilterAndSort());

        if (prefixTabPane != null) {
            prefixTabPane
                    .getSelectionModel()
                    .selectedIndexProperty()
                    .addListener(
                            (obs, a, b) -> {
                                if (suppressPrefixTabEvents.get()) {
                                    return;
                                }
                                applyPrefixFilterAndSort();
                            });
        }

        columnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::applyDynamicColumnWidths,
                                TableColumnOrderPersistence.TableId
                                        .DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
                                headerColumnCount,
                                this::onLeadingColumnCountCommitted,
                                this::onReorderColumns,
                                () ->
                                        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                                                ownerStage,
                                                TableColumnOrderPersistence.TableId
                                                        .DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
                                                spreadsheetView,
                                                () -> new ArrayList<>(headersRef))));
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.ownerStage = shell.getPrimaryStage();
        initPresentationControlsOnce();
        reloadFromDisk();
    }

    void setRefreshButtonVisible(boolean visible) {
        if (refreshButton != null) {
            refreshButton.setVisible(visible);
            refreshButton.setManaged(visible);
        }
    }

    void reloadFromDisk() {
        reloadFromDisk(false);
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
        if (pathLabel != null) {
            pathLabel.setText(path.toString());
        }
        if (!Files.isRegularFile(path)) {
            if (statusLabel != null) {
                statusLabel.setText("ファイルなし");
            }
            applyEmpty();
            finishReload(userCompletionDialog, path, false, "ファイルが見つかりません。\n" + path);
            return;
        }
        try {
            String raw = Files.readString(path, StandardCharsets.UTF_8);
            JsonNode root = JSON.readTree(raw);
            JsonNode columnsNode = root.get("columns");
            JsonNode rowsNode = root.get("rows");
            if (columnsNode == null || !columnsNode.isArray() || rowsNode == null || !rowsNode.isArray()) {
                if (statusLabel != null) {
                    statusLabel.setText("JSON 構造が不正");
                }
                applyEmpty();
                finishReload(
                        userCompletionDialog,
                        path,
                        false,
                        "JSON の構造が不正です（columns / rows が必要です）。\n" + path);
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
            List<Map<String, String>> summaryRows =
                    ResultDispatchTaskSummaryConsolidator.consolidate(headerOrder, rowMaps);
            dailyRowsByGroupKey.clear();
            dailyRowsByGroupKey.putAll(
                    ResultDispatchTaskSummaryConsolidator.indexDailyRowsByTaskGroup(rowMaps));
            ResultDispatchStage3Support.applyPlanningStageBadgeFromDispatchJson(
                    dataStageBadgeLabel, path);

            headersRef.clear();
            headersRef.addAll(headerOrder);
            consolidatedRows.clear();
            consolidatedRows.addAll(summaryRows);

            applyPersistedColumnLayout();
            injectDeadlineJudgmentColumnIfNeeded();
            rebuildPrefixTabs();
            applyPrefixFilterAndSort();

            if (statusLabel != null) {
                statusLabel.setText(summaryRows.size() + " タスク / " + rowMaps.size() + " 行（日別）");
            }
            finishReload(
                    userCompletionDialog,
                    path,
                    true,
                    "配台結果（タスク集約）を読み込みました。\n" + path + "\nタスク数: " + summaryRows.size());
        } catch (Exception ex) {
            if (statusLabel != null) {
                statusLabel.setText("error");
            }
            shell.appendLog("[dispatch-task-summary] " + ex.getMessage());
            applyEmpty();
            finishReload(
                    userCompletionDialog,
                    path,
                    false,
                    ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    private void finishReload(
            boolean userCompletionDialog, Path path, boolean ok, String message) {
        if (refreshButton != null) {
            refreshButton.setDisable(false);
        }
        if (userCompletionDialog && shell != null) {
            if (ok) {
                shell.showInformationDialog("再読み完了", message);
            } else {
                shell.showWarningDialog("再読み", message);
            }
        }
    }

    private void applyEmpty() {
        headersRef.clear();
        consolidatedRows.clear();
        displayedConsolidatedRows.clear();
        dailyRowsByGroupKey.clear();
        dailyScheduleItems.clear();
        rows.clear();
        persistedLayout.set(List.of());
        spreadsheetView.setGrid(new GridBase(0, 0));
        updateDailyScheduleTitle(null);
        rebuildPrefixTabs();
    }

    private void rebuildPrefixTabs() {
        if (prefixTabPane == null) {
            return;
        }
        suppressPrefixTabEvents.set(true);
        try {
            prefixTabPane.getTabs().clear();
            if (consolidatedRows.isEmpty()) {
                return;
            }
            Set<String> prefixes = new LinkedHashSet<>();
            for (Map<String, String> row : consolidatedRows) {
                prefixes.add(TaskIdLeadingAlphaPrefix.extract(row.get(COL_TID)));
            }
            List<String> ordered = new ArrayList<>(prefixes);
            ordered.sort(
                    (a, b) -> {
                        if (TaskIdLeadingAlphaPrefix.OTHER.equals(a)) {
                            return 1;
                        }
                        if (TaskIdLeadingAlphaPrefix.OTHER.equals(b)) {
                            return -1;
                        }
                        return a.compareTo(b);
                    });
            for (String prefix : ordered) {
                long count =
                        consolidatedRows.stream()
                                .filter(
                                        r ->
                                                prefix.equals(
                                                        TaskIdLeadingAlphaPrefix.extract(
                                                                r.get(COL_TID))))
                                .count();
                Tab tab = new Tab(prefix + " (" + count + ")");
                tab.setClosable(false);
                prefixTabPane.getTabs().add(tab);
            }
            if (!prefixTabPane.getTabs().isEmpty()) {
                prefixTabPane.getSelectionModel().select(0);
            }
        } finally {
            suppressPrefixTabEvents.set(false);
        }
    }

    private void applyPrefixFilterAndSort() {
        if (headersRef.isEmpty()) {
            rows.clear();
            rebuildSpreadsheet();
            return;
        }
        String selectedPrefix = selectedPrefixLabel();
        String sortCol =
                sortColumnCombo != null && sortColumnCombo.getValue() != null
                        ? sortColumnCombo.getValue()
                        : COL_TID;
        boolean desc =
                sortDescendingToggle != null && sortDescendingToggle.isSelected();
        List<Map<String, String>> filtered = new ArrayList<>();
        for (Map<String, String> row : consolidatedRows) {
            if (selectedPrefix == null
                    || selectedPrefix.equals(TaskIdLeadingAlphaPrefix.extract(row.get(COL_TID)))) {
                filtered.add(row);
            }
        }
        filtered.sort(buildRowComparator(sortCol, desc));

        displayedConsolidatedRows.clear();
        displayedConsolidatedRows.addAll(filtered);
        rows.clear();
        for (Map<String, String> map : filtered) {
            ObservableList<String> line = FXCollections.observableArrayList();
            for (String h : headersRef) {
                line.add(map.getOrDefault(h, ""));
            }
            rows.add(line);
        }
        injectDeadlineJudgmentIntoObservableRows();
        rebuildSpreadsheet();
        refreshDailySchedulePane();
    }

    private void initDailyScheduleTable() {
        if (dailyScheduleTable == null) {
            return;
        }
        dailyScheduleTable.setItems(dailyScheduleItems);
        dailyScheduleTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        dailyScheduleTable.getColumns().setAll(buildDailyScheduleColumns());
    }

    private List<TableColumn<DailyScheduleRow, ?>> buildDailyScheduleColumns() {
        List<TableColumn<DailyScheduleRow, ?>> cols = new ArrayList<>();
        cols.add(textColumn(ResultDispatchSchema.COL_DISPATCH_DATE, "dispatchDate", 72));
        cols.add(textColumn(COL_START, "startDateTime", 96));
        cols.add(textColumn(COL_END, "endDateTime", 96));
        cols.add(textColumn(ResultDispatchSchema.COL_DISPATCH_QTY, "dispatchQty", 64));
        cols.add(textColumn("メンバー名", "memberName", 80));
        return cols;
    }

    private static TableColumn<DailyScheduleRow, String> textColumn(
            String title, String property, double minWidth) {
        TableColumn<DailyScheduleRow, String> col = new TableColumn<>(title);
        col.setCellValueFactory(new PropertyValueFactory<>(property));
        col.setMinWidth(minWidth);
        col.setSortable(false);
        return col;
    }

    private void installDailyScheduleSelectionListener() {
        spreadsheetView
                .getSelectionModel()
                .getSelectedCells()
                .addListener(
                        (ListChangeListener<TablePosition>) change -> {
                            while (change.next()) {
                                if (change.wasAdded() || change.wasRemoved()) {
                                    Platform.runLater(this::refreshDailySchedulePane);
                                }
                            }
                        });
    }

    private void refreshDailySchedulePane() {
        if (dailyScheduleTable == null) {
            return;
        }
        int idx = selectedDataRowIndex();
        if (idx < 0 || idx >= displayedConsolidatedRows.size()) {
            dailyScheduleItems.clear();
            updateDailyScheduleTitle(null);
            return;
        }
        Map<String, String> taskRow = displayedConsolidatedRows.get(idx);
        String groupKey = ResultDispatchTaskSummaryConsolidator.taskGroupKey(taskRow);
        List<Map<String, String>> daily =
                ResultDispatchTaskSummaryConsolidator.sortedDailyScheduleRows(
                        dailyRowsByGroupKey.getOrDefault(groupKey, List.of()));
        dailyScheduleItems.setAll(DailyScheduleRow.fromMaps(daily));
        updateDailyScheduleTitle(taskRow);
    }

    private void updateDailyScheduleTitle(Map<String, String> taskRow) {
        if (dailyScheduleTitleLabel == null) {
            return;
        }
        if (taskRow == null) {
            dailyScheduleTitleLabel.setText("配台予定（日別）— 行を選択");
            return;
        }
        String tid = nz(taskRow.get(COL_TID));
        String proc = nz(taskRow.get(ResultDispatchSchema.COL_PROCESS));
        String mach = nz(taskRow.get(ResultDispatchSchema.COL_MACHINE));
        dailyScheduleTitleLabel.setText(
                "配台予定（日別）— " + tid + " / " + proc + " / " + mach);
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

    /** 右ペイン用: 日別配台行の表示モデル。 */
    public static final class DailyScheduleRow {
        private final String dispatchDate;
        private final String startDateTime;
        private final String endDateTime;
        private final String dispatchQty;
        private final String memberName;

        DailyScheduleRow(
                String dispatchDate,
                String startDateTime,
                String endDateTime,
                String dispatchQty,
                String memberName) {
            this.dispatchDate = dispatchDate;
            this.startDateTime = startDateTime;
            this.endDateTime = endDateTime;
            this.dispatchQty = dispatchQty;
            this.memberName = memberName;
        }

        public String getDispatchDate() {
            return dispatchDate;
        }

        public String getStartDateTime() {
            return startDateTime;
        }

        public String getEndDateTime() {
            return endDateTime;
        }

        public String getDispatchQty() {
            return dispatchQty;
        }

        public String getMemberName() {
            return memberName;
        }

        static List<DailyScheduleRow> fromMaps(List<Map<String, String>> daily) {
            List<DailyScheduleRow> out = new ArrayList<>(daily.size());
            for (Map<String, String> row : daily) {
                out.add(
                        new DailyScheduleRow(
                                nz(row.get(ResultDispatchSchema.COL_DISPATCH_DATE)),
                                nz(row.get(COL_START)),
                                nz(row.get(COL_END)),
                                nz(row.get(ResultDispatchSchema.COL_DISPATCH_QTY)),
                                nz(row.get("メンバー名"))));
            }
            return out;
        }
    }

    private String selectedPrefixLabel() {
        if (prefixTabPane == null || prefixTabPane.getTabs().isEmpty()) {
            return null;
        }
        Tab tab = prefixTabPane.getSelectionModel().getSelectedItem();
        if (tab == null || tab.getText() == null || tab.getText().isBlank()) {
            return null;
        }
        String text = tab.getText().strip();
        int paren = text.indexOf(" (");
        return paren > 0 ? text.substring(0, paren) : text;
    }

    private static Comparator<Map<String, String>> buildRowComparator(String sortCol, boolean desc) {
        Comparator<Map<String, String>> cmp;
        if (COL_START.equals(sortCol) || COL_END.equals(sortCol)) {
            cmp =
                    Comparator.comparing(
                            (Map<String, String> r) ->
                                    ResultDispatchDeadlineJudgment.parseDispatchDateTime(
                                            nz(r.get(sortCol))),
                            Comparator.nullsLast(Comparator.naturalOrder()));
        } else {
            cmp =
                    Comparator.comparing(
                            (Map<String, String> r) -> nz(r.get(sortCol)),
                            String.CASE_INSENSITIVE_ORDER);
        }
        if (desc) {
            cmp = cmp.reversed();
        }
        return cmp.thenComparing(r -> nz(r.get(COL_TID)), String.CASE_INSENSITIVE_ORDER);
    }

    private void injectDeadlineJudgmentColumnIfNeeded() {
        if (headersRef.isEmpty()) {
            return;
        }
        final String col = ResultDispatchDeadlineJudgment.COL_TITLE;
        int colIdx = headersRef.indexOf(col);
        if (colIdx < 0) {
            int afterAnswer = headersRef.indexOf("回答納期");
            colIdx = afterAnswer >= 0 ? afterAnswer + 1 : headersRef.size();
            headersRef.add(colIdx, col);
        }
    }

    private void injectDeadlineJudgmentIntoObservableRows() {
        final String col = ResultDispatchDeadlineJudgment.COL_TITLE;
        int colIdx = headersRef.indexOf(col);
        if (colIdx < 0) {
            return;
        }
        for (ObservableList<String> line : rows) {
            while (line.size() < headersRef.size()) {
                line.add("");
            }
            LinkedHashMap<String, String> map = new LinkedHashMap<>();
            for (int c = 0; c < headersRef.size(); c++) {
                String h = headersRef.get(c);
                if (col.equals(h)) {
                    continue;
                }
                map.put(h, c < line.size() && line.get(c) != null ? line.get(c) : "");
            }
            line.set(colIdx, ResultDispatchDeadlineJudgment.judgmentOkNg(map));
        }
    }

    private void applyPersistedColumnLayout() {
        List<TableColumnOrderPersistence.ColumnSpec> lay =
                TableColumnOrderPersistence.loadLayout(
                        TableColumnOrderPersistence.TableId
                                .DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY);
        persistedLayout.set(lay);
        List<String> beforeHeaders = new ArrayList<>(headersRef);
        boolean[] visBefore =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId
                                .DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
                        beforeHeaders.size());
        List<String> titleOrder =
                lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
        if (!titleOrder.isEmpty()) {
            reorderConsolidatedMaps(titleOrder);
            TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
            boolean[] visAfter =
                    TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                            beforeHeaders, visBefore, titleOrder);
            TableColumnOrderPersistence.saveColumnVisibility(
                    TableColumnOrderPersistence.TableId
                            .DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
                    visAfter);
        }
    }

    private void reorderConsolidatedMaps(List<String> titleOrder) {
        List<String> oldHeaders = new ArrayList<>(headersRef);
        headersRef.clear();
        for (String t : titleOrder) {
            if (oldHeaders.contains(t)) {
                headersRef.add(t);
            }
        }
        for (String h : oldHeaders) {
            if (!headersRef.contains(h)) {
                headersRef.add(h);
            }
        }
    }

    private void initPresentationControlsOnce() {
        if (presentationHooksInstalled || rowHeightSlider == null) {
            return;
        }
        presentationHooksInstalled = true;
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs prefs =
                TableColumnOrderPersistence.loadSpreadsheetTabPresentationPrefs(
                        TableColumnOrderPersistence.TableId
                                .DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY);
        spreadsheetTabPrefs.set(prefs);
        suppressPresentationUiEvents.set(true);
        try {
            rowHeightSlider.setValue(prefs.rowHeightPercent());
            if (rowHeightPctLabel != null) {
                rowHeightPctLabel.setText(String.format("%.0f%%", prefs.rowHeightPercent()));
            }
            if (cellWrapCheck != null) {
                cellWrapCheck.setSelected(prefs.cellWrapText());
            }
        } finally {
            suppressPresentationUiEvents.set(false);
        }
        SliderCommittedChangeSupport.install(
                rowHeightSlider,
                () -> {
                    if (rowHeightPctLabel != null && rowHeightSlider != null) {
                        rowHeightPctLabel.setText(
                                String.format("%.0f%%", rowHeightSlider.getValue()));
                    }
                },
                this::commitSpreadsheetPresentationFromSlider);
        if (cellWrapCheck != null) {
            cellWrapCheck
                    .selectedProperty()
                    .addListener((obs, a, b) -> commitSpreadsheetPresentationFromUi());
        }
        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                spreadsheetView,
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
                suppressColumnPersistence::get,
                () -> new ArrayList<>(headersRef));
    }

    private void commitSpreadsheetPresentationFromSlider() {
        if (suppressPresentationUiEvents.get()) {
            return;
        }
        commitSpreadsheetPresentationFromUi();
    }

    private void commitSpreadsheetPresentationFromUi() {
        if (rowHeightSlider == null) {
            return;
        }
        double v = rowHeightSlider.getValue();
        boolean wrap = cellWrapCheck != null && cellWrapCheck.isSelected();
        TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs next =
                new TableColumnOrderPersistence.SpreadsheetTabPresentationPrefs(v, wrap);
        spreadsheetTabPrefs.set(next);
        TableColumnOrderPersistence.saveSpreadsheetTabPresentationPrefs(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY, next);
        if (rowHeightPctLabel != null) {
            rowHeightPctLabel.setText(String.format("%.0f%%", v));
        }
        rebuildSpreadsheet();
    }

    private void onLeadingColumnCountCommitted(int count) {
        headerColumnCount.set(count);
        rebuildSpreadsheet();
    }

    private void onReorderColumns() {
        if (headersRef.isEmpty() || shell == null) {
            return;
        }
        boolean[] visForDialog =
                TableColumnOrderPersistence.loadColumnVisibility(
                        TableColumnOrderPersistence.TableId
                                .DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
                        headersRef.size());
        SpreadsheetColumnReorderDialog.show(ownerStage, new ArrayList<>(headersRef), visForDialog)
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
                        TableColumnOrderPersistence.TableId
                                .DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
                        oldHeaders.size());
        List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
        reorderConsolidatedMaps(titleOrder);
        TableColumnOrderPersistence.applyLogicalColumnOrder(headersRef, rows, titleOrder);
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        oldHeaders, oldVis, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
                newVis);
        List<Double> widths =
                TableColumnOrderPersistence.resolveWidthsForHeaders(headersRef, lay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
        for (int i = 0; i < headersRef.size(); i++) {
            newLay.add(
                    new TableColumnOrderPersistence.ColumnSpec(headersRef.get(i), widths.get(i)));
        }
        persistedLayout.set(newLay);
        TableColumnOrderPersistence.saveLayout(
                TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
                newLay);
        applyPrefixFilterAndSort();
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
            int deadlineJudgmentColIdx = headersRef.indexOf(ResultDispatchDeadlineJudgment.COL_TITLE);
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
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(
                                spreadsheetView);
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
                                                TableColumnOrderPersistence.TableId
                                                        .DELIVERY_CALENDAR_DISPATCH_TASK_SUMMARY,
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

    void clearColumnFiltersAndSort() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(spreadsheetView);
    }

    @FXML
    private void onClearColumnFiltersAction() {
        clearColumnFiltersAndSort();
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

    private static String nz(String v) {
        return v != null ? v : "";
    }
}
