package jp.co.pm.ai.desktop;

import java.io.File;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
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
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.GridChange;
import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.DispatchPlanInputInteractiveCoverageCheck;
import jp.co.pm.ai.desktop.dispatch.DispatchPlanInputInteractiveCoverageCheck.TaskKey;
import jp.co.pm.ai.desktop.io.ExcelCellReadSupport;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.PlanInputDeprecatedOverrideColumnSupport;
import jp.co.pm.ai.desktop.ui.PlanInputEditedCellMarks;
import jp.co.pm.ai.desktop.ui.PlanInputProcessSequenceRowOrder;
import jp.co.pm.ai.desktop.ui.PlanInputDateColumnSupport;
import jp.co.pm.ai.desktop.ui.PlanInputRawInputDateShift;
import jp.co.pm.ai.desktop.ui.PlanInputUnprocessedDispatchRemainingMismatchSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetMultiColumnFilterCoordinator;
import jp.co.pm.ai.desktop.ui.SpreadsheetPlanInputCellEditSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetPlanInputRowDragSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetricsResult;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;
import jp.co.pm.ai.desktop.ui.Stage2InProgressNextDayDispatchDialog;

/**
 * 配台計画_タスク入力タブ。レイアウトは {@code PlanInputTab.fxml}。
 *
 * <p>段階2の「当日は配台しない」オプション（{@code PM_AI_STAGE2_SKIP_TODAY_DISPATCH}）および加工途中の翌日配台量設定は本タブに配置する。
 *
 * <p>ControlsFX {@link SpreadsheetView} で先頭固定列をネイティブに扱う。
 */
public final class PlanInputTabController {

    /** planning_core の {@code RESULT_TASK_COL_DISPATCH_TRIAL_ORDER} 相当（段階1タスク入力の並び順）。 */
    private static final String COL_DISPATCH_TRIAL_ORDER = "配台試行順番";

    public static final String ENV_PM_AI_PLAN_INPUT_PATH = AppPaths.KEY_PM_AI_PLAN_INPUT_PATH;
    public static final String ENV_TASK_PLAN_SHEET = "TASK_PLAN_SHEET";

    public static final String DEFAULT_PLAN_INPUT_SHEET_NAME = AppPaths.STAGE1_PLAN_OUTPUT_SHEET;

    private static final String HINT_TEXT =
            "PM_AI_PLAN_INPUT_PATH に読み込む表ファイルのパスを指定。"
                    + "（段階2 load_planning_tasks_df: CSV / Parquet / xlsx 対応）。"
                    + "Excel のときはシート名も指定（TASK_PLAN_SHEET / この欄）。"
                    + " .xlsx 保存はデータのみ（マクロは含みません）。";

    private static final String DEBUG_SESSION_ID = "5a9d50";

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
    private Button shiftRawInputDateMinusOneButton;

    @FXML
    private Button stage2RunButton;

    @FXML
    private Button stage21RunButton;

    private static final String STAGE2_RUN_BUTTON_TEXT_DEFAULT = "段階2 実行";

    private static final String STAGE2_RUN_BUTTON_TEXT_SUMMARY_LOCKED =
            "段階2（サマリエクセル更新中）";

    private static final String STAGE2_RUN_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD =
            "段階2（納期管理ビュー更新中）";

    @FXML
    private CheckBox stage2SkipTodayDispatchCheckBox;

    @FXML
    private CheckBox stage2InProgressNextDayPromptCheckBox;

    @FXML
    private HBox columnStripHost;

    @FXML
    private TextField colWidthField;

    @FXML
    private Label hintLabel;

    @FXML
    private Label planInputValidationWarningLabel;

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

    /** 読込直後の全セル値（markKey→値）。編集差分判定の基準。 */
    private final Map<String, String> editBaselineByMarkKey = new LinkedHashMap<>();
    /** 読込時に sidecar JSON にあったマーク（基準値に戻しても保持する集合）。 */
    private final Set<String> editMarksPersistedAtLoad = new LinkedHashSet<>();
    /** 元の値から書き換えたセルのマーク（行キー\u0001列見出し）。 */
    private final Set<String> editedCellMarks = new LinkedHashSet<>();

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

    /** 納期管理ビュー再読み込み中（メインシェルから同期）。 */
    private boolean deliveryCalendarReloadBlocking;

    @FXML
    private void initialize() {
        pathField.setPromptText(
                "PM_AI_PLAN_INPUT_PATH （.csv / .xlsx / .xlsm）");
        sheetField.setText(DEFAULT_PLAN_INPUT_SHEET_NAME);
        sheetField.setPromptText("Excel sheet name (TASK_PLAN_SHEET / TASK_PLAN_SHEET)");
        colWidthField.setText("112");
        hintLabel.setText(HINT_TEXT);

        installStageRunButtonDepth(stage2RunButton, Color.rgb(194, 65, 12, 0.35));
        installStageRunButtonDepth(stage21RunButton, Color.rgb(194, 65, 12, 0.35));
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
        if (stage2InProgressNextDayPromptCheckBox != null) {
            stage2InProgressNextDayPromptCheckBox.setSelected(true);
            stage2InProgressNextDayPromptCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (shell != null) {
                                    shell.scheduleDesktopSessionSave();
                                }
                            });
        }
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
                this::finishPlanInputRowReorderAfterDnD);

        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                spreadsheetHost,
                headerColumnCount::get,
                SpreadsheetPlanInputRowDragSupport::skipFullRowExpansionDuringPlanInputRowDrag);

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

    /** 行並べ替え後: §A-1（加工内容順）を維持しつつ配台試行順番を 1..n に振り直す。 */
    private void renumberDispatchTrialOrderColumn() {
        // #region agent log
        if (shell != null && rows != null && !rows.isEmpty()) {
            int colTid = headersRef.indexOf(PlanInputProcessSequenceRowOrder.COL_TASK_ID);
            int colProc = headersRef.indexOf(PlanInputProcessSequenceRowOrder.COL_PROCESS);
            int colDto = headersRef.indexOf(PlanInputProcessSequenceRowOrder.COL_DISPATCH_TRIAL_ORDER);
            List<String> sample = new ArrayList<>();
            for (int i = 0; i < Math.min(8, rows.size()); i++) {
                ObservableList<String> r = rows.get(i);
                sample.add(
                        i
                                + ":"
                                + (colTid >= 0 && colTid < r.size() ? r.get(colTid) : "")
                                + "/"
                                + (colProc >= 0 && colProc < r.size() ? r.get(colProc) : "")
                                + "/dto="
                                + (colDto >= 0 && colDto < r.size() ? r.get(colDto) : ""));
            }
            AgentDebugLog.appendStructured(
                    shell.snapshotUiEnv(),
                    "a2361b",
                    "H6",
                    "PlanInputTabController:renumberDispatchTrialOrderColumn:before",
                    "before stabilize",
                    Map.of("rowCount", rows.size(), "orderHead", sample, "runId", "reorder-fix"));
        }
        // #endregion
        PlanInputProcessSequenceRowOrder.stabilizeAndRenumberDispatchTrialOrder(headersRef, rows);
        // #region agent log
        if (shell != null && rows != null && !rows.isEmpty()) {
            int colTid = headersRef.indexOf(PlanInputProcessSequenceRowOrder.COL_TASK_ID);
            int colProc = headersRef.indexOf(PlanInputProcessSequenceRowOrder.COL_PROCESS);
            int colDto = headersRef.indexOf(PlanInputProcessSequenceRowOrder.COL_DISPATCH_TRIAL_ORDER);
            List<String> sample = new ArrayList<>();
            for (int i = 0; i < Math.min(8, rows.size()); i++) {
                ObservableList<String> r = rows.get(i);
                sample.add(
                        i
                                + ":"
                                + (colTid >= 0 && colTid < r.size() ? r.get(colTid) : "")
                                + "/"
                                + (colProc >= 0 && colProc < r.size() ? r.get(colProc) : "")
                                + "/dto="
                                + (colDto >= 0 && colDto < r.size() ? r.get(colDto) : ""));
            }
            AgentDebugLog.appendStructured(
                    shell.snapshotUiEnv(),
                    "a2361b",
                    "H6",
                    "PlanInputTabController:renumberDispatchTrialOrderColumn:after",
                    "after stabilize",
                    Map.of("rowCount", rows.size(), "orderHead", sample, "runId", "reorder-fix"));
        }
        // #endregion
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

    void setDeliveryCalendarReloadBlocking(boolean blocking) {
        deliveryCalendarReloadBlocking = blocking;
        applyStage2RunButtonEnabledState();
    }

    /**
     * 配台計画手動修正の表に未保存の変更があるとき {@code blocked} を true にする（保存または「再読み」で false）。
     */
    void setStage2BlockedByUnsavedDispatchEdit(boolean blocked) {
        stage2BlockedByDispatchUnsavedEdit = blocked;
        applyStage2RunButtonEnabledState();
    }

    /** ロックファイルの有無に合わせて段階2ボタン表示を更新する（{@link MainShellController#isSummaryAiDispatchExportLocked}）。 */
    void refreshSummaryExportLockPresentation() {
        applyStage2RunButtonEnabledState();
    }

    private boolean isSummaryExportLockedByLockFile() {
        return shell != null && shell.isSummaryAiDispatchExportLocked();
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
        boolean disable =
                stage2RunPipelineBusy
                        || deliveryCalendarReloadBlocking
                        || isSummaryExportLockedByLockFile()
                        || stage2BlockedByDispatchUnsavedEdit
                        || stage2BlockedByUnsavedPlanInputTableEdit;
        if (stage2RunButton != null) {
            stage2RunButton.setDisable(disable);
        }
        if (stage21RunButton != null) {
            stage21RunButton.setDisable(disable);
        }
        if (stage2RunPipelineBusy) {
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(null);
            }
            if (stage21RunButton != null) {
                stage21RunButton.setTooltip(null);
            }
        } else if (deliveryCalendarReloadBlocking) {
            Tooltip blockedTip =
                    new Tooltip("納期管理ビューを再読み込み中です。完了後に実行してください。");
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(blockedTip);
            }
            if (stage21RunButton != null) {
                stage21RunButton.setTooltip(blockedTip);
            }
        } else if (isSummaryExportLockedByLockFile()) {
            Tooltip blockedTip =
                    new Tooltip(
                            "サマリ xlsx を作成中です。完了後に実行するか、実行・ログタブの「ロック解除」を使用してください。");
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(blockedTip);
            }
            if (stage21RunButton != null) {
                stage21RunButton.setTooltip(blockedTip);
            }
        } else if (stage2BlockedByDispatchUnsavedEdit) {
            Tooltip blockedTip =
                    new Tooltip(
                            "配台計画手動修正タブに未保存の変更があります。「保存 (JSON+xlsx)」または「再読み」で確定してから実行してください。");
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(blockedTip);
            }
            if (stage21RunButton != null) {
                stage21RunButton.setTooltip(blockedTip);
            }
        } else if (stage2BlockedByUnsavedPlanInputTableEdit) {
            Tooltip blockedTip =
                    new Tooltip(
                            "配台計画_タスク入力タブの表に未保存の変更があります。「保存」または「再読み」で確定してから実行してください。");
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(blockedTip);
            }
            if (stage21RunButton != null) {
                stage21RunButton.setTooltip(blockedTip);
            }
        } else {
            if (stage2RunButton != null) {
                stage2RunButton.setTooltip(null);
            }
            if (stage21RunButton != null) {
                stage21RunButton.setTooltip(
                        new Tooltip(
                                "残業/休出シミュ付きフル再配台（成功時はメイン output へ正本反映）"));
            }
        }
        if (stage2RunButton != null) {
            if (deliveryCalendarReloadBlocking) {
                stage2RunButton.setText(STAGE2_RUN_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD);
            } else if (isSummaryExportLockedByLockFile()) {
                stage2RunButton.setText(STAGE2_RUN_BUTTON_TEXT_SUMMARY_LOCKED);
            } else {
                stage2RunButton.setText(STAGE2_RUN_BUTTON_TEXT_DEFAULT);
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

    boolean snapshotStage2InProgressNextDayPrompt() {
        return stage2InProgressNextDayPromptCheckBox == null
                || stage2InProgressNextDayPromptCheckBox.isSelected();
    }

    void applyStage2InProgressNextDayPromptFromSession(boolean prompt) {
        if (stage2InProgressNextDayPromptCheckBox != null) {
            stage2InProgressNextDayPromptCheckBox.setSelected(prompt);
        }
    }

    /**
     * 段階2直前ダイアログ用: 実加工数が正の行（加工途中相当）。配台計画除外・完了キーワード行は除く。
     */
    List<Stage2InProgressNextDayDispatchDialog.Row> collectInProgressRowsForNextDayDialog() {
        Map<String, String> rowMap = new java.util.LinkedHashMap<>();
        int colTask = headersRef.indexOf("依頼NO");
        int colProcess = headersRef.indexOf("工程名");
        int colMachine = headersRef.indexOf("機械名");
        Stage2RollUnitLengthTables tables = cachedRollUnitHighlightTables.get();
        if (tables == null && shell != null) {
            try {
                tables = Stage2RollUnitLengthTables.load(shell.snapshotUiEnv());
                cachedRollUnitHighlightTables.set(tables);
            } catch (Exception ignored) {
                tables = Stage2RollUnitLengthTables.empty();
            }
        }
        if (tables == null) {
            tables = Stage2RollUnitLengthTables.empty();
        }
        List<Stage2InProgressNextDayDispatchDialog.Row> out = new ArrayList<>();
        for (ObservableList<String> cells : rows) {
            rowMap.clear();
            for (int c = 0; c < headersRef.size(); c++) {
                String h = headersRef.get(c);
                String v = c < cells.size() && cells.get(c) != null ? cells.get(c) : "";
                rowMap.put(h, v);
            }
            if (isPlanRowExcludedFromStage2Queue(rowMap)) {
                continue;
            }
            double actual =
                    Stage2RollUnitLengthTables.parseFloatSafe(rowMap.getOrDefault("実加工数", ""), 0.0);
            if (actual <= 1e-12) {
                continue;
            }
            String taskId = colTask >= 0 ? cellAt(cells, colTask) : rowMap.getOrDefault("依頼NO", "");
            String process = colProcess >= 0 ? cellAt(cells, colProcess) : rowMap.getOrDefault("工程名", "");
            String machine = colMachine >= 0 ? cellAt(cells, colMachine) : rowMap.getOrDefault("機械名", "");
            if (taskId.isBlank()) {
                continue;
            }
            double remaining =
                    Stage2PlanRowDispatchQtyMetrics.compute(rowMap, tables)
                            .map(Stage2PlanRowDispatchQtyMetricsResult::remainingM)
                            .orElse(0.0);
            Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo =
                    Stage2PlanRowDispatchQtyMetrics.dispatchSimulatorUnitMFromPlanRow(
                            rowMap, tables);
            out.add(
                    new Stage2InProgressNextDayDispatchDialog.Row(
                            taskId, process, machine, actual, remaining, unitInfo));
        }
        return out;
    }

    /**
     * 段階2後の整合確認用: 配台不要オフ（かつ配台計画除外・完了でない）行の (依頼NO, 工程名, 機械名)。
     */
    /**
     * 配台対象行の依頼NO → 原反投入日。
     * 段階2／3 後の午前配台率警告分析用。
     */
    Map<String, LocalDate> collectEffectiveRawInputDateByTaskId() {
        Map<String, String> rowMap = new LinkedHashMap<>();
        int colTask = headersRef.indexOf("依頼NO");
        LinkedHashMap<String, LocalDate> out = new LinkedHashMap<>();
        for (ObservableList<String> cells : rows) {
            rowMap.clear();
            for (int c = 0; c < headersRef.size(); c++) {
                String h = headersRef.get(c);
                String v = c < cells.size() && cells.get(c) != null ? cells.get(c) : "";
                rowMap.put(h, v);
            }
            if (!DispatchPlanInputInteractiveCoverageCheck.isEligiblePlanInputRow(rowMap)) {
                continue;
            }
            String taskId = colTask >= 0 ? cellAt(cells, colTask) : rowMap.getOrDefault("依頼NO", "");
            if (taskId.isBlank()) {
                continue;
            }
            LocalDate effective = effectiveRawInputDate(rowMap);
            if (effective != null) {
                out.putIfAbsent(taskId.strip(), effective);
            }
        }
        return Map.copyOf(out);
    }

    private static LocalDate effectiveRawInputDate(Map<String, String> rowMap) {
        return PlanInputDateColumnSupport.parseCellValue(
                        rowMap.get(PlanInputRawInputDateShift.COL_RAW_INPUT_DATE))
                .orElse(null);
    }

    List<TaskKey> collectEligibleTaskKeysForDispatchCoverage() {
        Map<String, String> rowMap = new LinkedHashMap<>();
        int colTask = headersRef.indexOf("依頼NO");
        int colProcess = headersRef.indexOf("工程名");
        int colMachine = headersRef.indexOf("機械名");
        LinkedHashMap<String, TaskKey> deduped = new LinkedHashMap<>();
        for (ObservableList<String> cells : rows) {
            rowMap.clear();
            for (int c = 0; c < headersRef.size(); c++) {
                String h = headersRef.get(c);
                String v = c < cells.size() && cells.get(c) != null ? cells.get(c) : "";
                rowMap.put(h, v);
            }
            if (!DispatchPlanInputInteractiveCoverageCheck.isEligiblePlanInputRow(rowMap)) {
                continue;
            }
            String taskId = colTask >= 0 ? cellAt(cells, colTask) : rowMap.getOrDefault("依頼NO", "");
            String process =
                    colProcess >= 0 ? cellAt(cells, colProcess) : rowMap.getOrDefault("工程名", "");
            String machine =
                    colMachine >= 0 ? cellAt(cells, colMachine) : rowMap.getOrDefault("機械名", "");
            TaskKey key = new TaskKey(taskId, process, machine);
            if (!key.isComplete()) {
                continue;
            }
            deduped.putIfAbsent(key.identityToken(), key);
        }
        return List.copyOf(deduped.values());
    }

    private static String cellAt(ObservableList<String> cells, int col) {
        if (col < 0 || col >= cells.size()) {
            return "";
        }
        String v = cells.get(col);
        return v != null ? v.strip() : "";
    }

    /**
     * 依頼NO・工程名・機械名が一致する配台計画行（先頭一致）。手動修正タブの配台ロール単位解決用。
     */
    public Optional<Map<String, String>> findPlanRowMapByKeys(
            String taskId, String process, String machine) {
        String tid = taskId != null ? taskId.strip() : "";
        if (tid.isEmpty() || headersRef.isEmpty()) {
            return Optional.empty();
        }
        String proc = process != null ? process.strip() : "";
        String mach = machine != null ? machine.strip() : "";
        int colTask = headersRef.indexOf("依頼NO");
        int colProcess = headersRef.indexOf("工程名");
        int colMachine = headersRef.indexOf("機械名");
        for (ObservableList<String> cells : rows) {
            String rowTid = colTask >= 0 ? cellAt(cells, colTask) : "";
            if (!tid.equals(rowTid)) {
                continue;
            }
            String rowProc = colProcess >= 0 ? cellAt(cells, colProcess) : "";
            String rowMach = colMachine >= 0 ? cellAt(cells, colMachine) : "";
            if (!proc.equals(rowProc) || !mach.equals(rowMach)) {
                continue;
            }
            LinkedHashMap<String, String> row = new LinkedHashMap<>();
            for (int c = 0; c < headersRef.size(); c++) {
                row.put(headersRef.get(c), cellAt(cells, c));
            }
            return Optional.of(row);
        }
        return Optional.empty();
    }

    /** 試走ラボ用: 依頼NO・工程名・機械名ラベル一覧。 */
    public List<String> listPlanInputTaskLabels() {
        if (headersRef.isEmpty() || rows == null || rows.isEmpty()) {
            return List.of();
        }
        int colTask = headersRef.indexOf("依頼NO");
        int colProcess = headersRef.indexOf("工程名");
        int colMachine = headersRef.indexOf("機械名");
        List<String> out = new ArrayList<>();
        for (ObservableList<String> cells : rows) {
            String tid = colTask >= 0 ? cellAt(cells, colTask) : "";
            if (tid.isEmpty()) {
                continue;
            }
            String proc = colProcess >= 0 ? cellAt(cells, colProcess) : "";
            String mach = colMachine >= 0 ? cellAt(cells, colMachine) : "";
            out.add(tid + " / " + proc + " / " + mach);
        }
        return out;
    }

    /** 試走ラボ用: ラベルから行 Map を返す。 */
    public Optional<Map<String, String>> findPlanRowMapByLabel(String label) {
        if (label == null || label.isBlank() || headersRef.isEmpty()) {
            return Optional.empty();
        }
        String[] parts = label.split(" / ", 3);
        String tid = parts.length > 0 ? parts[0].strip() : "";
        String proc = parts.length > 1 ? parts[1].strip() : "";
        String mach = parts.length > 2 ? parts[2].strip() : "";
        Optional<Map<String, String>> found = findPlanRowMapByKeys(tid, proc, mach);
        if (found.isPresent()) {
            Map<String, String> row = new LinkedHashMap<>(found.get());
            row.putIfAbsent("task_id", tid);
            return Optional.of(row);
        }
        return Optional.empty();
    }

    private static boolean isPlanRowExcludedFromStage2Queue(Map<String, String> row) {
        return DispatchPlanInputInteractiveCoverageCheck.isExcludedFromDispatchCoverage(row);
    }

    @FXML
    private void onStage2RunButtonAction() {
        if (shell != null) {
            shell.triggerStage2();
        }
    }

    @FXML
    private void onStage21RunButtonAction() {
        if (shell != null) {
            shell.launchStage21OvertimeSimulationWizard();
        }
    }

    @FXML
    private void onShiftRawInputDateMinusOneAction() {
        if (shell == null || headersRef.isEmpty()) {
            return;
        }
        int updated = PlanInputRawInputDateShift.applyMinusOneDayToAllRows(headersRef, rows);
        if (updated == PlanInputRawInputDateShift.MISSING_RAW_INPUT_DATE_COLUMN) {
            shell.showErrorDialog(
                    "原反投入日の前倒し",
                    "列「"
                            + PlanInputRawInputDateShift.COL_RAW_INPUT_DATE
                            + "」がありません。表を読み込んでから実行してください。");
            return;
        }
        if (updated == 0) {
            shell.showInformationDialog(
                    "原反投入日の前倒し",
                    "原反投入日を解釈できる行がありませんでした。");
            return;
        }
        markPlanInputTableDirtySinceSave();
        rebuildSpreadsheet();
        shell.appendLog(
                "[plan-input] 原反投入日を1日前倒し: "
                        + updated
                        + " 行の「"
                        + PlanInputRawInputDateShift.COL_RAW_INPUT_DATE
                        + "」を更新しました。");
    }

    @FXML
    private void onRowUpAction() {
        int i = selectedPlanInputDataIndex();
        if (i <= 0) {
            return;
        }
        int colIdx = planInputFocusedColumnIndex();
        markPlanInputTableDirtySinceSave();
        swapPlanInputDataRowsInMemory(i - 1, i);
        schedulePlanInputRowReorderPresentation(i - 1, colIdx);
    }

    @FXML
    private void onRowDownAction() {
        int i = selectedPlanInputDataIndex();
        if (i < 0 || i >= rows.size() - 1) {
            return;
        }
        int colIdx = planInputFocusedColumnIndex();
        markPlanInputTableDirtySinceSave();
        swapPlanInputDataRowsInMemory(i, i + 1);
        schedulePlanInputRowReorderPresentation(i + 1, colIdx);
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

    /** DnD 完了後（{@link SpreadsheetPlanInputRowDragSupport} が model を移動済み）。 */
    private void finishPlanInputRowReorderAfterDnD() {
        renumberDispatchTrialOrderColumn();
        markPlanInputTableDirtySinceSave();
        applyPlanInputRowReorderPresentation(-1, -1);
    }

    /** ↑↓ ボタン: 行入替後のグリッド再構築は次 FX パルスで行う（選択検証の IndexOutOfBounds 回避）。 */
    private void schedulePlanInputRowReorderPresentation(int focusDataRow, int focusColumn) {
        Platform.runLater(() -> applyPlanInputRowReorderPresentation(focusDataRow, focusColumn));
    }

    private void applyPlanInputRowReorderPresentation(int focusDataRow, int focusColumn) {
        rebuildSpreadsheet(true, SpreadsheetTabularSupport.GridAttachMode.IN_PLACE);
        if (focusDataRow >= 0) {
            focusPlanInputCellAfterReorder(focusDataRow, focusColumn);
        }
    }

    private void swapPlanInputDataRowsInMemory(int a, int b) {
        if (a < 0 || b < 0 || a >= rows.size() || b >= rows.size() || a == b) {
            return;
        }
        ObservableList<String> moved = rows.get(a);
        rows.set(a, rows.get(b));
        rows.set(b, moved);
        renumberDispatchTrialOrderColumn();
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

    /** 段階2前後の外部更新後に表をディスクから再読込（完了ダイアログなし）。 */
    void reloadQuietlyFromDisk() {
        syncFromEnv();
        loadFromCurrentPath(false);
    }

    /** 段階1「キャッシュをクリアして実行」等で、表表示を空にする（ディスク削除は {@link Stage1AiCacheClearer}）。 */
    void clearTableForStage1CacheClear() {
        headersRef.clear();
        if (rows != null) {
            rows.clear();
        }
        if (rowSearchField != null) {
            rowSearchField.clear();
        }
        clearColumnFiltersAndSort();
        clearPlanInputTableDirtySinceSave();
        rebuildSpreadsheet();
        if (shell != null) {
            shell.appendLog("[plan-input] キャッシュクリアに伴い表を空にしました。");
        }
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
            PlanInputEditedCellMarks.save(path, editedCellMarks);
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
                        return Stage2RollUnitLengthTables.load(shell.snapshotUiEnv());
                    } catch (Exception e) {
                        return Stage2RollUnitLengthTables.empty();
                    }
                });
    }

    /**
     * 材料・製品種類情報（{@code code/} テーブル）の保存・再読込後に呼ぶ。ロール長ハイライト用キャッシュを破棄し、表を再構築する。
     */
    void invalidateRollUnitHighlightCacheAndRefresh() {
        cachedRollUnitHighlightTables.set(null);
        if (!headersRef.isEmpty()) {
            rebuildSpreadsheet();
        }
    }

    private void rebuildSpreadsheet() {
        rebuildSpreadsheet(true, SpreadsheetTabularSupport.GridAttachMode.STANDARD);
    }

    /**
     * @param preserveColumnFilters {@code true} のとき、再構築前の列フィルタ（許容値集合）を復元する。ファイル読込・
     *     論理列並べ替え後は {@code false}（列インデックスが変わるため）。
     */
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
            updatePlanInputUnprocessedDispatchRemainingWarning();
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
            refreshAndPersistPlanInputEditMarks();
            PlanInputEditedCellMarks.applyHighlights(
                    grid, headersRef, rows, firstDataRow, editedCellMarks);
            PlanInputUnprocessedDispatchRemainingMismatchSupport.applyViolationHighlights(
                    grid, headersRef, rows, firstDataRow);
            updatePlanInputUnprocessedDispatchRemainingWarning();
            var rowSync =
                    SpreadsheetTabularSupport.newRowsSyncHandler(rows, headersRef, firstDataRow);
            gridChangeHandler =
                    ev -> {
                        rowSync.handle(ev);
                        refreshAndPersistPlanInputEditMarks();
                        PlanInputEditedCellMarks.applyHighlights(
                                currentGrid, headersRef, rows, firstDataRow, editedCellMarks);
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.applyViolationHighlights(
                                currentGrid, headersRef, rows, firstDataRow);
                        updatePlanInputUnprocessedDispatchRemainingWarning();
                        if (!suppressPlanInputDirtyFromGridEvents.get()) {
                            markPlanInputTableDirtySinceSave();
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
                                        rowSearchField.getText() != null
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

    private void updatePlanInputUnprocessedDispatchRemainingWarning() {
        if (planInputValidationWarningLabel == null) {
            return;
        }
        List<String> mismatchTaskIds =
                PlanInputUnprocessedDispatchRemainingMismatchSupport.collectMismatchTaskIds(
                        headersRef, rows);
        boolean show = !mismatchTaskIds.isEmpty();
        planInputValidationWarningLabel.setVisible(show);
        planInputValidationWarningLabel.setManaged(show);
        if (show) {
            planInputValidationWarningLabel.setText(
                    PlanInputUnprocessedDispatchRemainingMismatchSupport.warningMessage(
                            mismatchTaskIds));
        } else {
            planInputValidationWarningLabel.setText("");
        }
        // #region agent log
        if (shell != null) {
            AgentDebugLog.appendStructured(
                    shell.snapshotUiEnv(),
                    DEBUG_SESSION_ID,
                    "A",
                    "PlanInputTabController.updatePlanInputUnprocessedDispatchRemainingWarning",
                    "plan-input unprocessed vs dispatch-remaining mismatch scan",
                    Map.of(
                            "mismatchCount",
                            mismatchTaskIds.size(),
                            "mismatchTaskIds",
                            mismatchTaskIds,
                            "headersHasUnprocessed",
                            headersRef.contains(
                                    PlanInputUnprocessedDispatchRemainingMismatchSupport
                                            .COL_UNPROCESSED),
                            "headersHasDispatchRemaining",
                            headersRef.contains(
                                    PlanInputUnprocessedDispatchRemainingMismatchSupport
                                            .COL_DISPATCH_REMAINING),
                            "warningVisible",
                            show));
        }
        // #endregion
    }

    private void applyLoaded() {
        rebuildSpreadsheet(false);
    }

    /** 読込時: 編集差分の基準値を記録し、sidecar JSON のマークを取り込む。 */
    private void loadPlanInputEditMarks(Path path) {
        editBaselineByMarkKey.clear();
        editBaselineByMarkKey.putAll(PlanInputEditedCellMarks.captureBaseline(headersRef, rows));
        Set<String> loaded =
                PlanInputEditedCellMarks.filterToPresentRows(
                        headersRef, rows, PlanInputEditedCellMarks.load(path));
        editMarksPersistedAtLoad.clear();
        editMarksPersistedAtLoad.addAll(loaded);
        editedCellMarks.clear();
        editedCellMarks.addAll(loaded);
    }

    /** 現在の表から編集マークを再計算し、変化があれば sidecar JSON へ保存する。 */
    private void refreshAndPersistPlanInputEditMarks() {
        PlanInputEditedCellMarks.recompute(
                headersRef, rows, editBaselineByMarkKey, editMarksPersistedAtLoad, editedCellMarks);
        Path path = currentPlanInputPathOrNull();
        if (path != null) {
            PlanInputEditedCellMarks.save(path, editedCellMarks);
        }
    }

    private Path currentPlanInputPathOrNull() {
        String p = pathField.getText() != null ? pathField.getText().trim() : "";
        if (p.isEmpty()) {
            return null;
        }
        try {
            return Path.of(p);
        } catch (RuntimeException ex) {
            return null;
        }
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
            PlanInputDeprecatedOverrideColumnSupport.migrateAndDropDeprecatedOverrideColumns(
                    headersRef, rows);
            normalizePlanInputDateOnlyColumns();
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
            loadPlanInputEditMarks(path);
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

    /** 日付列のセルから深夜時刻サフィックスを除く。 */
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

    String snapshotPlanInputPath() {
        return pathField.getText() != null ? pathField.getText().trim() : "";
    }

    public String snapshotPlanInputSheet() {
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
