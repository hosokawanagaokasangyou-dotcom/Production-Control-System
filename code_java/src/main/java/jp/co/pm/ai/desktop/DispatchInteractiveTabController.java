package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Base64;
import java.util.BitSet;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.HashSet;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.beans.property.ObjectProperty;
import javafx.beans.property.ReadOnlyObjectWrapper;
import javafx.beans.property.SimpleObjectProperty;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.geometry.Rectangle2D;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.text.Font;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.TabPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableView;
import javafx.scene.control.TablePosition;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.MouseEvent;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.stage.Window;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.GridChange;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.DispatchTrialLogUiStore;
import jp.co.pm.ai.desktop.config.DispatchTrialLogUiStore.DispatchTrialLogUiSnapshot;
import jp.co.pm.ai.desktop.dispatch.DispatchTrialConsistency;
import jp.co.pm.ai.desktop.dispatch.DispatchTrialShortages;
import jp.co.pm.ai.desktop.dispatch.DispatchTrialShortages.DispatchQtyShortfallRow;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchDocument;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchJsonIo;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPivot;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPythonExport;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchTrialPython;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnDragReorderSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetRowReorderDragGhost;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * Interactive pivot editor for result dispatch JSON (ControlsFX SpreadsheetView: task-by-day columns +
 * process+machine-by-day).
 */
public final class DispatchInteractiveTabController {

    private record ReloadBundle(ResultDispatchDocument doc) {}

    private record DispatchSaveOutcome(Path jsonPath, String xlsxStdoutLine) {}

    private static final String DND_PREFIX = "pm-dispatch-dnd|wide|";
    private static final String DND_V2_MARKER = "v2|";
    /** Drag payload for reordering wide-grid profile rows (leading columns only). */
    private static final String DND_ROW_PREFIX = "pm-dispatch-dnd|wide|row|v2|";

    /** One spreadsheet column per calendar day (wide grid date axis). */
    private static final int DAY_SLOT_COLUMNS = 1;

    /**
     * Width for date columns where every data cell is calendar-blocked (gray); keeps holiday bands from wasting
     * horizontal space ({@code SpreadsheetView} uses pixels; value is converted from typographic points).
     */
    private static final double BLOCKED_DATE_COLUMN_PREF_PT = 5.0;

    /**
     * Minimum width for normal date columns so {@code YYYY-MM-DD} plus ControlsFX filter header stays readable
     * (not truncated to "2026-...").
     */
    private static final double MIN_DATE_COLUMN_WIDTH_PX = 112.0;

    /** Fully-blocked (holiday) date columns: stay narrow but wide enough for a short header glyph. */
    private static final double MIN_BLOCKED_DATE_COLUMN_WIDTH_PX = 40.0;

    /** 日付列で数量が正のとき。 */
    private static final String DATE_CELL_STYLE_POSITIVE_QTY =
            "-fx-background-color: #e8f5e9; -fx-text-fill: black;";

    /** 配台試行でタイムライン実績が目標に届かないセル。 */
    private static final String DATE_CELL_STYLE_SHORTFALL =
            "-fx-background-color: #b71c1c; -fx-text-fill: white; -fx-font-weight: bold;";

    private static final List<String> WIDE_STATIC_HEADERS =
            List.of(
                    ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER,
                    ResultDispatchSchema.COL_PROCESS,
                    ResultDispatchSchema.COL_MACHINE,
                    "加工内容",
                    "依頼NO",
                    "換算数量",
                    "計画合計");

    /** 「工程+機械×日」ビューの先頭固定列（日付ブロックの直前まで）。 */
    private static final List<String> BY_DAY_STATIC_HEADERS =
            List.of(
                    ResultDispatchSchema.COL_PROCESS,
                    ResultDispatchSchema.COL_MACHINE,
                    "加工内容");

    private record WideGridBundle(
            GridBase grid,
            List<Map<String, String>> profiles,
            List<WideRow> rowItems,
            boolean[] blockedCols,
            int staticCols,
            int dayCount) {}

    private record ByDayGridBundle(GridBase grid, boolean[] blockedCols, int staticCols, int dayCount) {}

    private record FullGridRebuild(List<LocalDate> axis, WideGridBundle wide, ByDayGridBundle byDay) {}

    /**
     * In-memory {@link #doc} differs from last successful「保存」to disk. Dispatch trial reads JSON from disk, so the
     * button stays disabled until save (row order, cell edits, DnD moves, etc.).
     */
    private boolean dispatchDocDirtySinceSave;

    /** True while load/rebuild progress UI disables the toolbar ({@link #setReloadInteractionDisabled}). */
    private boolean reloadInteractionDisabled;

    /** True while 段階1／段階2 pipeline is running ({@link #setStageRunProgressVisible}). */
    private boolean stagePipelineBusy;

    /** Avoid treating programmatic grid updates as user edits ({@link #onWideGridChange}). */
    private final AtomicBoolean suppressDispatchGridDirty = new AtomicBoolean(false);

    @FXML
    private ProgressIndicator busyIndicator;

    @FXML
    private Button loadButton;

    @FXML
    private Button saveButton;

    @FXML
    private Button stage2RunButton;

    @FXML
    private Button dispatchTrialButton;

    @FXML
    private Button wideRowUpButton;

    @FXML
    private Button wideRowDownButton;

    @FXML
    private Label statusLabel;

    @FXML
    private ProgressBar reloadProgressBar;

    @FXML
    private Label jsonPathLabel;

    @FXML
    private VBox dispatchShortfallPanel;

    @FXML
    private TableView<DispatchQtyShortfallRow> dispatchShortfallTable;

    @FXML
    private TabPane innerTabPane;

    private final AtomicBoolean suppressInnerTabSessionPersistence = new AtomicBoolean(false);

    private volatile boolean innerTabPersistenceWired;

    @FXML
    private StackPane wideSpreadsheetHost;

    @FXML
    private StackPane byDaySpreadsheetHost;

    private final SpreadsheetView wideSpreadsheet = new SpreadsheetView();
    private final SpreadsheetView byDaySpreadsheet = new SpreadsheetView();

    private MainShellController shell;

    private ResultDispatchDocument doc = ResultDispatchDocument.empty();

    private List<LocalDate> dateAxis = new ArrayList<>();

    /**
     * 日付列のユーザー希望順（両グリッド共通）。{@code null} のときは {@link #computeDateAxisList()} の自然順を使う。
     */
    private List<LocalDate> preferredDateAxisOrder;

    /** 列ドラッグ並べ替え起因の {@link #rebuildGrids()} 中はヘッダ変更コールバックを無視する。 */
    private final AtomicBoolean suppressColumnReorderPersistence = new AtomicBoolean(false);

    private final List<Map<String, String>> wideProfiles = new ArrayList<>();

    /** Parallel to {@link #wideProfiles} rows in the wide grid. */
    private final List<WideRow> wideRowItems = new ArrayList<>();

    private List<DispatchQtyShortfallRow> lastDispatchShortfallRows = List.of();

    /** {@link DispatchTrialShortages.FullBundle#shortageHints()}（op_shortage / as_shortage）。試行後ダイアログ用。 */
    private List<DispatchTrialShortages.ShortageHint> lastDispatchShortageHints = List.of();

    private final Set<String> dispatchWideShortfallKeys = new HashSet<>();

    private final Set<String> dispatchByDayShortfallKeys = new HashSet<>();

    @FXML
    private void initialize() {
        StackPane.setAlignment(wideSpreadsheet, Pos.TOP_LEFT);
        wideSpreadsheetHost.getChildren().setAll(wideSpreadsheet);
        VBox.setVgrow(wideSpreadsheetHost, javafx.scene.layout.Priority.ALWAYS);
        wideSpreadsheet.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        wideSpreadsheet.prefWidthProperty().bind(wideSpreadsheetHost.widthProperty());
        wideSpreadsheet.prefHeightProperty().bind(wideSpreadsheetHost.heightProperty());

        StackPane.setAlignment(byDaySpreadsheet, Pos.TOP_LEFT);
        byDaySpreadsheetHost.getChildren().setAll(byDaySpreadsheet);
        VBox.setVgrow(byDaySpreadsheetHost, javafx.scene.layout.Priority.ALWAYS);
        byDaySpreadsheet.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        byDaySpreadsheet.prefWidthProperty().bind(byDaySpreadsheetHost.widthProperty());
        byDaySpreadsheet.prefHeightProperty().bind(byDaySpreadsheetHost.heightProperty());

        SpreadsheetThemeBridge.install(wideSpreadsheet);
        SpreadsheetThemeBridge.install(byDaySpreadsheet);
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(wideSpreadsheet);
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(byDaySpreadsheet);

        wideSpreadsheet.getSelectionModel().setSelectionMode(javafx.scene.control.SelectionMode.MULTIPLE);
        byDaySpreadsheet.getSelectionModel().setSelectionMode(javafx.scene.control.SelectionMode.MULTIPLE);
        SpreadsheetTabularSupport.installFullRowDataSelection(wideSpreadsheet);
        SpreadsheetTabularSupport.installFullRowDataSelection(byDaySpreadsheet);

        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                wideSpreadsheetHost, WIDE_STATIC_HEADERS::size);
        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                byDaySpreadsheetHost, BY_DAY_STATIC_HEADERS::size);

        installWideDnDHandlers();
        installByDayDoubleClickHandler();

        if (dispatchShortfallTable != null) {
            installDispatchShortfallColumns(dispatchShortfallTable);
            dispatchShortfallTable.setColumnResizePolicy(
                    TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
            wireDispatchShortfallSelectionToWideGrid();
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        Platform.runLater(this::reloadFromDiskQuiet);
        ensureInnerTabPersistenceWired();
    }

    private void ensureInnerTabPersistenceWired() {
        if (innerTabPersistenceWired || innerTabPane == null || shell == null) {
            return;
        }
        innerTabPersistenceWired = true;
        innerTabPane
                .getSelectionModel()
                .selectedIndexProperty()
                .addListener(
                        (obs, a, b) -> {
                            if (suppressInnerTabSessionPersistence.get()) {
                                return;
                            }
                            shell.persistDesktopSessionNow();
                        });
    }

    /** @return 選択中の子タブインデックス。未初期化時は -1 */
    public int snapshotInnerTabSelectedIndex() {
        if (innerTabPane == null) {
            return -1;
        }
        return innerTabPane.getSelectionModel().getSelectedIndex();
    }

    public void applyInnerTabSelectedIndex(int index) {
        if (innerTabPane == null || index < 0) {
            return;
        }
        int n = innerTabPane.getTabs().size();
        if (index >= n) {
            return;
        }
        suppressInnerTabSessionPersistence.set(true);
        try {
            innerTabPane.getSelectionModel().select(index);
        } finally {
            suppressInnerTabSessionPersistence.set(false);
        }
    }

    void clearColumnFiltersAndSort() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(wideSpreadsheet);
        SpreadsheetTabularSupport.clearAllFiltersAndSort(byDaySpreadsheet);
    }

    @FXML
    private void onClearColumnFiltersAction() {
        clearColumnFiltersAndSort();
    }

    /** 子タブに応じて「列の表示」ダイアログを開く（FXML: 列の表示）。 */
    @FXML
    private void onColumnVisibilityAction() {
        int tab = innerTabPane != null ? innerTabPane.getSelectionModel().getSelectedIndex() : 0;
        if (tab <= 0) {
            openWideColumnVisibilityDialog();
        } else {
            openByDayColumnVisibilityDialog();
        }
    }

    private void openWideColumnVisibilityDialog() {
        Stage st = shell != null ? shell.getPrimaryStage() : null;
        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                st,
                TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE,
                wideSpreadsheet,
                () -> new ArrayList<>(buildWideColumnLabelsForAxis(dateAxis)));
    }

    private void openByDayColumnVisibilityDialog() {
        Stage st = shell != null ? shell.getPrimaryStage() : null;
        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                st,
                TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY,
                byDaySpreadsheet,
                () -> new ArrayList<>(buildByDayColumnLabelsForAxis(dateAxis)));
    }

    /**
     * 列の並べ替えダイアログ（先頭固定列の外側＝主に日付列）。ヘッダドラッグ並べ替えと同じく日付軸と JSON 保存用レイアウトを更新する。
     */
    @FXML
    private void onColumnReorderAction() {
        Stage st = shell != null ? shell.getPrimaryStage() : null;
        if (dateAxis.isEmpty()) {
            if (shell != null) {
                shell.appendLog("[dispatch-editor] 列の並べ替え: 表示する列がありません（JSON を読み込んでください）");
            }
            return;
        }
        int tab = innerTabPane != null ? innerTabPane.getSelectionModel().getSelectedIndex() : 0;
        boolean wideMode = tab <= 0;
        List<String> headers =
                wideMode
                        ? new ArrayList<>(buildWideColumnLabelsForAxis(dateAxis))
                        : new ArrayList<>(buildByDayColumnLabelsForAxis(dateAxis));
        int fixed =
                wideMode ? WIDE_STATIC_HEADERS.size() : BY_DAY_STATIC_HEADERS.size();
        TableColumnOrderPersistence.TableId tid =
                wideMode
                        ? TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE
                        : TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY;
        boolean[] vis = TableColumnOrderPersistence.loadColumnVisibility(tid, headers.size());
        Optional<List<Integer>> choice =
                SpreadsheetColumnReorderDialog.showWithFixedLeading(st, headers, vis, fixed);
        if (choice.isEmpty()) {
            boolean anyMovableVisible = false;
            for (int i = fixed; i < headers.size(); i++) {
                if (vis == null || i >= vis.length || vis[i]) {
                    anyMovableVisible = true;
                    break;
                }
            }
            if (!anyMovableVisible && shell != null) {
                shell.appendLog(
                        "[dispatch-editor] 並べ替え対象の列がすべて非表示です。「列の表示」で日付列を表示してください。");
            }
            return;
        }
        applyDispatchInteractiveReorderPermutation(choice.get(), headers, wideMode);
    }

    private void applyDispatchInteractiveReorderPermutation(
            List<Integer> perm, List<String> headersSnapshot, boolean wideMode) {
        if (perm == null || headersSnapshot == null || perm.size() != headersSnapshot.size()) {
            return;
        }
        List<String> titleOrder = new ArrayList<>(perm.size());
        for (Integer idx : perm) {
            if (idx == null || idx < 0 || idx >= headersSnapshot.size()) {
                return;
            }
            titleOrder.add(headersSnapshot.get(idx));
        }
        TableColumnOrderPersistence.TableId tid =
                wideMode
                        ? TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE
                        : TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY;
        boolean[] oldVis =
                TableColumnOrderPersistence.loadColumnVisibility(tid, headersSnapshot.size());
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        headersSnapshot, oldVis, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(tid, newVis);

        List<LocalDate> computed = computeDateAxisList();
        int staticCols = wideMode ? WIDE_STATIC_HEADERS.size() : BY_DAY_STATIC_HEADERS.size();
        List<LocalDate> dates = parseDateTailAsDates(titleOrder, staticCols);
        if (dates == null || !sameMultisetLocalDate(dates, computed)) {
            return;
        }
        if (wideMode && !wideStaticPrefixMatches(titleOrder)) {
            return;
        }
        if (!wideMode && !byDayStaticPrefixMatches(titleOrder)) {
            return;
        }
        if (dates.equals(preferredDateAxisOrder)) {
            persistDispatchColumnLayouts(
                    wideMode ? titleOrder : buildWideColumnLabelsForAxis(dates),
                    wideMode ? buildByDayColumnLabelsForAxis(dates) : titleOrder);
            return;
        }
        preferredDateAxisOrder = new ArrayList<>(dates);
        persistDispatchColumnLayouts(
                wideMode ? titleOrder : buildWideColumnLabelsForAxis(dates),
                wideMode ? buildByDayColumnLabelsForAxis(dates) : titleOrder);
        suppressColumnReorderPersistence.set(true);
        try {
            rebuildGrids();
        } finally {
            suppressColumnReorderPersistence.set(false);
        }
    }

    /**
     * 実行タブと同様、段階1／段階2 実行中は「段階2 実行」を無効化する（{@link MainShellController#applyRunTabGating} から）。
     */
    void setStageRunProgressVisible(boolean stage1Running, boolean stage2Running) {
        stagePipelineBusy = stage1Running || stage2Running;
        applyStage2RunButtonEnabledState();
    }

    private void applyStage2RunButtonEnabledState() {
        if (stage2RunButton == null) {
            return;
        }
        stage2RunButton.setDisable(reloadInteractionDisabled || stagePipelineBusy);
    }

    @FXML
    private void onStage2RunButtonAction() {
        if (shell == null) {
            return;
        }
        shell.triggerStage2();
    }

    @FXML
    private void onLoadAction() {
        reloadFromDiskQuiet();
    }

    @FXML
    private void onSaveAction() {
        if (shell == null) {
            return;
        }
        if (reloadInteractionDisabled) {
            return;
        }
        Path jsonPath = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        ResultDispatchDocument toWrite = doc.copy();
        Path pyExe = resolvePythonExe();
        Path pyDir = AppPaths.resolvePythonScriptDir(shell.snapshotUiEnv());

        statusLabel.setText("保存中…");
        showReloadProgress();

        Task<DispatchSaveOutcome> task =
                new Task<>() {
                    @Override
                    protected DispatchSaveOutcome call() throws Exception {
                        ResultDispatchJsonIo.write(jsonPath, toWrite);
                        String xlsxOut =
                                ResultDispatchPythonExport.exportXlsxNearJson(jsonPath, pyExe, pyDir);
                        return new DispatchSaveOutcome(jsonPath, xlsxOut);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    try {
                        DispatchSaveOutcome r = task.getValue();
                        clearDispatchDocDirty();
                        statusLabel.setText("保存しました");
                        shell.appendLog("[dispatch-editor] saved json: " + r.jsonPath());
                        if (r.xlsxStdoutLine() != null && !r.xlsxStdoutLine().isEmpty()) {
                            shell.appendLog("[dispatch-editor] xlsx: " + r.xlsxStdoutLine());
                        } else {
                            shell.appendLog("[dispatch-editor] xlsx export skipped or failed (Python)");
                        }
                        showDispatchSaveFinishedDialog(r.jsonPath(), r.xlsxStdoutLine());
                    } finally {
                        hideReloadProgress();
                    }
                });
        task.setOnFailed(
                e -> {
                    try {
                        statusLabel.setText("保存エラー");
                        Throwable ex = task.getException();
                        String detail = ex != null ? ex.getMessage() : "";
                        shell.appendLog(
                                "[dispatch-editor] save failed: "
                                        + (detail != null && !detail.isBlank() ? detail : "(不明)"));
                        Alert err = new Alert(AlertType.ERROR);
                        err.setTitle("保存エラー");
                        err.setHeaderText("保存に失敗しました");
                        err.setContentText(
                                detail != null && !detail.isBlank()
                                        ? detail
                                        : (ex != null ? ex.getClass().getSimpleName() : "(不明)"));
                        if (shell.getPrimaryStage() != null) {
                            err.initOwner(shell.getPrimaryStage());
                        }
                        err.showAndWait();
                    } finally {
                        hideReloadProgress();
                    }
                });
        new Thread(task, "dispatch-editor-save").start();
    }

    /** 保存ボタン後: JSON / xlsx の結果をダイアログで通知する。 */
    private void showDispatchSaveFinishedDialog(Path jsonPath, String xlsxStdoutLine) {
        boolean xlsxOk = xlsxStdoutLine != null && !xlsxStdoutLine.isBlank();
        Alert alert = new Alert(xlsxOk ? AlertType.INFORMATION : AlertType.WARNING);
        alert.setTitle("保存");
        alert.setHeaderText(xlsxOk ? "保存が完了しました" : "JSON は保存しました（Excel に注意）");
        StringBuilder text = new StringBuilder();
        text.append("JSON を保存しました。\n").append(jsonPath);
        if (xlsxOk) {
            text.append("\n\n結果配台表の Excel (xlsx) を出力しました。\n").append(xlsxStdoutLine.trim());
        } else {
            text.append(
                    "\n\nExcel (xlsx) は出力されませんでした（スクリプト未配置・タイムアウト・終了コード異常など）。"
                            + " 実行・ログのメッセージも確認してください。");
        }
        alert.setContentText(text.toString());
        if (shell != null && shell.getPrimaryStage() != null) {
            alert.initOwner(shell.getPrimaryStage());
        }
        alert.showAndWait();
    }

    @FXML
    private void onDispatchTrialAction() {
        if (shell == null) {
            return;
        }
        statusLabel.setText("配台試行中...");
        showReloadProgress();
        Path jsonPath = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        final Path trialPythonExe = resolvePythonExe();

        Stage owner = shell.getPrimaryStage();
        Stage logStage = new Stage();
        logStage.initOwner(owner);
        logStage.initModality(Modality.APPLICATION_MODAL);
        logStage.setTitle("配台試行ログ");
        logStage.setMinWidth(560);
        logStage.setMinHeight(360);

        DispatchTrialLogUiSnapshot savedUi = DispatchTrialLogUiStore.load();
        Rectangle2D visual = Screen.getPrimary().getVisualBounds();
        double sceneW = savedUi.width() > 0 ? savedUi.width() : 720;
        double sceneH = savedUi.height() > 0 ? savedUi.height() : 480;
        sceneW = Math.max(560, Math.min(sceneW, visual.getWidth()));
        sceneH = Math.max(360, Math.min(sceneH, visual.getHeight()));

        AtomicBoolean finished = new AtomicBoolean(false);
        AtomicBoolean trialLogWindowReady = new AtomicBoolean(false);

        ObservableList<String> logLines = FXCollections.observableArrayList();
        logLines.add("[配台試行] 処理を開始しました。");
        logLines.add("[配台試行] Python 実行ファイル: " + trialPythonExe.toAbsolutePath().normalize());

        Button copyLogBtn = new Button("ログをコピー");
        copyLogBtn.setTooltip(new Tooltip("ログ一覧の全文をクリップボードにコピーします"));
        copyLogBtn.setOnAction(
                ev -> {
                    String joined = String.join("\n", logLines);
                    ClipboardContent cc = new ClipboardContent();
                    cc.putString(joined);
                    Clipboard.getSystemClipboard().setContent(cc);
                });

        Button closeBtn = new Button("閉じる");
        closeBtn.setDisable(true);
        closeBtn.setOnAction(ev -> logStage.close());

        Runnable releaseTrialModal =
                () -> {
                    finished.set(true);
                    try {
                        closeBtn.setDisable(false);
                    } catch (Throwable ignored) {
                    }
                    try {
                        closeBtn.requestFocus();
                    } catch (Throwable ignored) {
                    }
                    try {
                        hideReloadProgress();
                    } catch (Throwable ignored) {
                    }
                };

        ListView<String> logList = new ListView<>(logLines);
        logList.setEditable(false);

        ObservableList<String> fontFamilies = FXCollections.observableArrayList(Font.getFamilies());
        ComboBox<String> fontFamilyCombo = new ComboBox<>(fontFamilies);
        fontFamilyCombo.setPrefWidth(240);
        fontFamilyCombo.setMaxWidth(Double.MAX_VALUE);
        HBox.setHgrow(fontFamilyCombo, Priority.ALWAYS);
        String defaultLogFontFamily =
                fontFamilies.stream()
                        .filter(f -> "Consolas".equalsIgnoreCase(f))
                        .findFirst()
                        .orElseGet(
                                () ->
                                        fontFamilies.stream()
                                                .filter(
                                                        f ->
                                                                f.contains("ゴシック")
                                                                        || f.toLowerCase()
                                                                                .contains("gothic"))
                                                .findFirst()
                                                .orElse(
                                                        fontFamilies.isEmpty()
                                                                ? Font.getDefault().getFamily()
                                                                : fontFamilies.get(0)));
        String savedFamily = savedUi.fontFamily();
        if (savedFamily != null && !savedFamily.isBlank()) {
            if (!fontFamilies.contains(savedFamily)) {
                fontFamilies.add(0, savedFamily);
            }
            fontFamilyCombo.setValue(savedFamily);
        } else {
            fontFamilyCombo.setValue(defaultLogFontFamily);
        }

        double savedSizePt = savedUi.fontSize();
        double initialSpinner =
                (Double.isFinite(savedSizePt) && savedSizePt >= 6.0 && savedSizePt <= 48.0)
                        ? savedSizePt
                        : 12.0;
        Spinner<Double> fontSizeSpinner =
                new Spinner<>(
                        new SpinnerValueFactory.DoubleSpinnerValueFactory(6.0, 48.0, initialSpinner, 1.0));
        fontSizeSpinner.setEditable(true);
        fontSizeSpinner.setPrefWidth(96);

        ObjectProperty<Font> logFontProp = new SimpleObjectProperty<>(Font.getDefault());

        Runnable syncLogFont =
                () -> {
                    String fam = fontFamilyCombo.getValue();
                    if (fam == null || fam.isBlank()) {
                        fam = Font.getDefault().getFamily();
                    }
                    Double sz = fontSizeSpinner.getValue();
                    if (sz == null || sz < 4.0) {
                        sz = 12.0;
                    }
                    logFontProp.set(Font.font(fam, sz));
                };
        Runnable saveTrialLogUiSnapshot =
                () -> {
                    double w = logStage.getWidth();
                    double h = logStage.getHeight();
                    if (!Double.isFinite(w) || w < 1.0) {
                        w = 720;
                    }
                    if (!Double.isFinite(h) || h < 1.0) {
                        h = 480;
                    }
                    String fam = fontFamilyCombo.getValue();
                    if (fam == null) {
                        fam = "";
                    }
                    Double szObj = fontSizeSpinner.getValue();
                    double sz = (szObj != null && Double.isFinite(szObj)) ? szObj : 12.0;
                    DispatchTrialLogUiStore.save(new DispatchTrialLogUiSnapshot(w, h, fam, sz));
                };
        Runnable persistDispatchTrialLogUi =
                () -> {
                    if (!trialLogWindowReady.get()) {
                        return;
                    }
                    saveTrialLogUiSnapshot.run();
                };
        fontFamilyCombo
                .valueProperty()
                .addListener(
                        (o, a, b) -> {
                            syncLogFont.run();
                            persistDispatchTrialLogUi.run();
                        });
        fontSizeSpinner
                .valueProperty()
                .addListener(
                        (o, a, b) -> {
                            syncLogFont.run();
                            persistDispatchTrialLogUi.run();
                        });
        syncLogFont.run();

        logList.setCellFactory(
                lv ->
                        new ListCell<>() {
                            private final Label lineLabel = new Label();

                            {
                                lineLabel.setWrapText(true);
                                lineLabel.fontProperty().bind(logFontProp);
                                lineLabel
                                        .prefWidthProperty()
                                        .bind(logList.widthProperty().subtract(40));
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setGraphic(null);
                                    setText(null);
                                } else {
                                    lineLabel.setText(item);
                                    setGraphic(lineLabel);
                                }
                            }
                        });

        Label fontCap = new Label("フォント");
        Label sizeCap = new Label("サイズ");
        HBox toolBar = new HBox(8, fontCap, fontFamilyCombo, sizeCap, fontSizeSpinner);
        toolBar.setPadding(new Insets(8, 8, 0, 8));
        toolBar.setAlignment(Pos.CENTER_LEFT);

        BorderPane root = new BorderPane();
        root.setTop(toolBar);
        root.setCenter(logList);
        HBox bottom = new HBox(8);
        bottom.setPadding(new Insets(8));
        bottom.setAlignment(Pos.CENTER_RIGHT);
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        bottom.getChildren().addAll(spacer, copyLogBtn, closeBtn);
        root.setBottom(bottom);

        Scene scene = new Scene(root, sceneW, sceneH);
        shell.registerThemeTrackedScene(scene);
        logStage.setScene(scene);
        logStage.setOnShown(ev -> trialLogWindowReady.set(true));
        logStage.setOnHidden(
                ev -> {
                    shell.unregisterThemeTrackedScene(scene);
                    saveTrialLogUiSnapshot.run();
                });
        logStage.show();

        final ResultDispatchDocument trialInputSnapshot = doc.copy();

        @SuppressWarnings("unchecked")
        final Task<String>[] trialTaskHolder = new Task[] {null};

        Task<String> task =
                new Task<>() {
                    @Override
                    protected String call() throws Exception {
                        Path pyExe = trialPythonExe;
                        Path pyDir = AppPaths.resolvePythonScriptDir(shell.snapshotUiEnv());
                        Map<String, String> pyEnv = shell.snapshotDispatchTrialPythonEnv();
                        return ResultDispatchTrialPython.runTrial(
                                jsonPath,
                                pyExe,
                                pyDir,
                                pyEnv,
                                line ->
                                        Platform.runLater(
                                                () -> {
                                                    logLines.add(line);
                                                    int last = logLines.size() - 1;
                                                    logList.scrollTo(last);
                                                }));
                    }
                };
        trialTaskHolder[0] = task;

        logStage.setOnCloseRequest(
                ev -> {
                    Task<String> t = trialTaskHolder[0];
                    boolean workerStillRunning = t != null && !t.isDone();
                    if (!finished.get() && workerStillRunning) {
                        ev.consume();
                    }
                });

        task.setOnSucceeded(
                e -> {
                    try {
                        String shortagesPath = task.getValue();
                        statusLabel.setText("配台試行完了");
                        shell.refreshRunTabStage2ArtifactLinks();
                        shell.invalidateDeliveryCalendarAfterPipelineRun();
                        shell.refreshEquipmentGanttGraphicAfterPipelineRun();
                        shell.appendLog("[dispatch-editor] trial: " + shortagesPath);
                        logLines.add("");
                        logLines.add("[配台試行] 正常終了しました。");
                        logLines.add("不足情報JSON: " + shortagesPath);
                        logList.scrollTo(logLines.size() - 1);
                        reloadFromDiskQuiet(
                                () -> {
                                    try {
                                        showDispatchQtyShortfallDialogIfNeeded(owner);
                                        showDispatchShortageHintsDialogIfNeeded(owner);
                                        DispatchTrialConsistency.CheckResult cr =
                                                DispatchTrialConsistency.compareDocuments(
                                                        trialInputSnapshot, doc);
                                        if (cr.consistent()) {
                                            logLines.add("");
                                            logLines.add(
                                                    "[整合性] 保存済み表と試行後の成果物（結果_配台表.json）は、"
                                                            + "依頼NO×機械名の当日配台数量合計および配台試行順番（工程別最小値）の観点で一致しました。");
                                            shell.appendLog(
                                                    "[dispatch-editor] trial: 整合性OK（保存表と再読込JSONの数量・試行順）");
                                        } else {
                                            logLines.add("");
                                            logLines.add(
                                                    "[整合性] 保存済み表と試行後の成果物に差異があります（詳細は下記）:");
                                            for (String dl : cr.detailLines()) {
                                                logLines.add(dl);
                                            }
                                            Alert warn = new Alert(AlertType.WARNING);
                                            warn.setTitle("配台試行: 整合性確認");
                                            warn.setHeaderText(
                                                    "試行前の保存内容と、試行後に読み込んだ結果_配台表.json に差異があります。");
                                            warn.setContentText(String.join("\n", cr.detailLines()));
                                            warn.show();
                                            shell.appendLog(
                                                    "[dispatch-editor] trial: 整合性に差異あり（"
                                                            + cr.detailLines().size()
                                                            + " 件）— ログ・ダイアログ参照");
                                        }
                                        int last = logLines.size() - 1;
                                        if (last >= 0) {
                                            logList.scrollTo(last);
                                        }
                                        DispatchTrialUnassignedWizard.showIfNeeded(
                                                owner, shell, Path.of(shortagesPath));
                                    } catch (Throwable upex) {
                                        String em =
                                                upex.getMessage() != null
                                                        ? upex.getMessage()
                                                        : upex.getClass().getSimpleName();
                                        logLines.add("");
                                        logLines.add("[配台試行] 試行後処理で例外: " + em);
                                        shell.appendLog(
                                                "[dispatch-editor] trial post-run: " + em);
                                    }
                                });
                    } catch (Throwable sucEx) {
                        String em =
                                sucEx.getMessage() != null
                                        ? sucEx.getMessage()
                                        : sucEx.getClass().getSimpleName();
                        logLines.add("");
                        logLines.add("[配台試行] 成功ハンドラ内例外: " + em);
                        shell.appendLog("[dispatch-editor] trial onSucceeded: " + em);
                    } finally {
                        releaseTrialModal.run();
                    }
                });
        task.setOnFailed(
                e -> {
                    try {
                        Throwable ex = task.getException();
                        statusLabel.setText("配台試行エラー");
                        String msg = ex != null ? ex.getMessage() : "(不明)";
                        shell.appendLog("[dispatch-editor] trial failed: " + msg);
                        logLines.add("");
                        logLines.add("[配台試行] エラーで終了しました。");
                        logLines.add(msg);
                        if (ex != null) {
                            java.io.StringWriter sw = new java.io.StringWriter();
                            ex.printStackTrace(new java.io.PrintWriter(sw));
                            String stack = sw.toString();
                            int max = 8000;
                            if (stack.length() > max) {
                                stack = stack.substring(0, max) + "\n... (truncated)";
                            }
                            logLines.add(stack);
                        }
                        logList.scrollTo(logLines.size() - 1);
                    } catch (Throwable handlerEx) {
                        shell.appendLog(
                                "[dispatch-editor] trial onFailed handler: "
                                        + handlerEx.getMessage());
                    } finally {
                        releaseTrialModal.run();
                    }
                });
        task.setOnCancelled(
                e -> {
                    try {
                        statusLabel.setText("配台試行キャンセル");
                        logLines.add("");
                        logLines.add("[配台試行] キャンセルされました。");
                        logList.scrollTo(logLines.size() - 1);
                    } finally {
                        releaseTrialModal.run();
                    }
                });
        new Thread(task, "dispatch-trial").start();
    }

    @FXML
    private void onWideRowUpAction() {
        int i = selectedWideProfileIndex();
        if (i <= 0) {
            return;
        }
        int colIdx = wideSpreadsheetFocusedColumnIndex();
        swapWideProfiles(i - 1, i);
        focusWideProfileCellAfterReorder(i - 1, colIdx);
    }

    @FXML
    private void onWideRowDownAction() {
        int i = selectedWideProfileIndex();
        if (i < 0 || i >= wideProfiles.size() - 1) {
            return;
        }
        int colIdx = wideSpreadsheetFocusedColumnIndex();
        swapWideProfiles(i, i + 1);
        focusWideProfileCellAfterReorder(i + 1, colIdx);
    }

    private void reloadFromDiskQuiet() {
        reloadFromDiskQuiet(null);
    }

    /**
     * Reloads JSON from disk asynchronously; runs {@code afterSuccessOnFxThread} on the FX thread after grids are
     * rebuilt (only when load succeeds).
     */
    private void reloadFromDiskQuiet(Runnable afterSuccessOnFxThread) {
        if (shell == null) {
            return;
        }
        Path p = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        jsonPathLabel.setText(p.toString());
        if (!Files.isRegularFile(p)) {
            statusLabel.setText("ファイルなし");
            doc = ResultDispatchDocument.empty();
            clearDispatchShortfallUi();
            rebuildGrids();
            clearDispatchDocDirty();
            return;
        }
        showReloadProgress();
        final MainShellController shellRef = shell;
        final Path jsonPath = p;
        Task<ReloadBundle> task =
                new Task<>() {
                    @Override
                    protected ReloadBundle call() throws Exception {
                        ResultDispatchDocument d = ResultDispatchJsonIo.read(jsonPath);
                        return new ReloadBundle(d);
                    }
                };
        task.setOnSucceeded(
                ev -> {
                    ReloadBundle b = task.getValue();
                    doc = b.doc();
                    statusLabel.setText(doc.rows().size() + " 行");
                    applyDispatchShortfallFromDisk(jsonPath);
                    rebuildGrids();
                    clearDispatchDocDirty();
                    hideReloadProgress();
                    if (afterSuccessOnFxThread != null) {
                        afterSuccessOnFxThread.run();
                    }
                });
        task.setOnFailed(
                ev -> {
                    doc = ResultDispatchDocument.empty();
                    statusLabel.setText("読込エラー");
                    Throwable ex = task.getException();
                    shell.appendLog(
                            "[dispatch-editor] load failed: "
                                    + (ex != null ? ex.getMessage() : ""));
                    clearDispatchShortfallUi();
                    rebuildGrids();
                    clearDispatchDocDirty();
                    hideReloadProgress();
                });
        new Thread(task, "dispatch-editor-reload").start();
    }

    /**
     * 実行・ログから段階2を実行して {@code 結果_配台表.json} が更新されたあと、当タブの表をディスクから再読込する。
     */
    void reloadTableFromDiskAfterExternalUpdate() {
        reloadFromDiskQuiet();
    }

    /** 配台ワークスペース用スナップショット: メモリ上の配台表ドキュメントのコピー（UI スレッド）。 */
    public ResultDispatchDocument copyDispatchDocumentForSnapshot() {
        return doc.copy();
    }

    private void showReloadProgress() {
        if (reloadProgressBar != null) {
            reloadProgressBar.setManaged(true);
            reloadProgressBar.setVisible(true);
            reloadProgressBar.setProgress(ProgressBar.INDETERMINATE_PROGRESS);
        }
        if (busyIndicator != null) {
            busyIndicator.setManaged(true);
            busyIndicator.setVisible(true);
        }
        setReloadInteractionDisabled(true);
    }

    private void hideReloadProgress() {
        if (reloadProgressBar != null) {
            reloadProgressBar.setProgress(0);
            reloadProgressBar.setVisible(false);
            reloadProgressBar.setManaged(false);
        }
        if (busyIndicator != null) {
            busyIndicator.setVisible(false);
            busyIndicator.setManaged(false);
        }
        setReloadInteractionDisabled(false);
    }

    private void setReloadInteractionDisabled(boolean disabled) {
        reloadInteractionDisabled = disabled;
        if (loadButton != null) {
            loadButton.setDisable(disabled);
        }
        if (saveButton != null) {
            saveButton.setDisable(disabled);
        }
        applyStage2RunButtonEnabledState();
        applyDispatchTrialButtonEnabledState();
        if (wideRowUpButton != null) {
            wideRowUpButton.setDisable(disabled);
        }
        if (wideRowDownButton != null) {
            wideRowDownButton.setDisable(disabled);
        }
    }

    private void markDispatchDocDirty() {
        dispatchDocDirtySinceSave = true;
        applyDispatchTrialButtonEnabledState();
    }

    private void clearDispatchDocDirty() {
        dispatchDocDirtySinceSave = false;
        applyDispatchTrialButtonEnabledState();
    }

    /** Dispatch trial reads JSON from disk; disable until「保存」while the editor has unsaved changes. */
    private void applyDispatchTrialButtonEnabledState() {
        if (dispatchTrialButton == null) {
            return;
        }
        boolean block = reloadInteractionDisabled || dispatchDocDirtySinceSave;
        dispatchTrialButton.setDisable(block);
        if (dispatchDocDirtySinceSave && !reloadInteractionDisabled) {
            dispatchTrialButton.setTooltip(
                    new Tooltip("表の変更を段階3に反映するには、先に「保存 (JSON+xlsx)」を押してください。"));
        } else {
            dispatchTrialButton.setTooltip(null);
        }
    }

    /**
     * Wide-grid static columns (editable) push into {@link #doc}; column 0 (試行順) is not edited here (reordered via
     * DnD / buttons).
     */
    private void onWideGridChange(GridChange ev) {
        if (suppressDispatchGridDirty.get()) {
            return;
        }
        int r = ev.getRow();
        int c = ev.getColumn();
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        if (r < firstData) {
            return;
        }
        int staticCols = WIDE_STATIC_HEADERS.size();
        if (c <= 0 || c >= staticCols) {
            return;
        }
        int profileIdx = r - firstData;
        if (profileIdx < 0 || profileIdx >= wideProfiles.size()) {
            return;
        }
        Map<String, String> prof = wideProfiles.get(profileIdx);
        Map<String, String> oldProf = new LinkedHashMap<>(prof);
        String headerKey = WIDE_STATIC_HEADERS.get(c);
        Object nv = ev.getNewValue();
        String s = nv != null ? Objects.toString(nv, "") : "";
        prof.put(headerKey, s);
        List<String> cols = doc.columns();
        for (Map<String, String> row : doc.rows()) {
            if (ResultDispatchPivot.matchesTaskProfileExceptTrialOrder(cols, oldProf, row)) {
                row.put(headerKey, s);
            }
        }
        ResultDispatchNormalizer.normalizeInPlace(cols, doc.rows());
        markDispatchDocDirty();
    }

    private Path resolvePythonExeForShell(MainShellController shellRef) {
        if (shellRef == null) {
            return MainShellController.defaultPythonPathWhenShellMissing();
        }
        return shellRef.resolveStagePythonExecutablePath();
    }

    private Path resolvePythonExe() {
        return resolvePythonExeForShell(shell);
    }

    private void rebuildGrids() {
        FullGridRebuild bundle = buildFullGridRebuild();
        applyFullGridRebuild(bundle);
    }

    private List<LocalDate> computeDateAxisList() {
        List<LocalDate> distinct = ResultDispatchPivot.distinctDates(doc.rows());
        if (distinct.isEmpty()) {
            List<LocalDate> ax = new ArrayList<>();
            LocalDate t = LocalDate.now();
            for (int i = 0; i < 14; i++) {
                ax.add(t.plusDays(i));
            }
            return ax;
        }
        List<LocalDate> range = ResultDispatchPivot.dateRangeInclusive(distinct);
        return range;
    }

    private FullGridRebuild buildFullGridRebuild() {
        ResultDispatchPivot.mergeDispatchRowsByWideIdentity(
                doc.columns(),
                doc.rows(),
                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
        List<LocalDate> axis = axisForRebuild();
        WideGridBundle wide = buildWideGridModel(axis);
        ByDayGridBundle byDay = buildByDayGridModel(axis);
        return new FullGridRebuild(axis, wide, byDay);
    }

    /** データに存在する日付集合は維持しつつ、保存済み／ユーザー設定の列順を日付軸に反映する。 */
    private List<LocalDate> axisForRebuild() {
        List<LocalDate> computed = computeDateAxisList();
        if (preferredDateAxisOrder != null && sameMultisetLocalDate(preferredDateAxisOrder, computed)) {
            return new ArrayList<>(preferredDateAxisOrder);
        }
        preferredDateAxisOrder = null;
        List<LocalDate> fromPersistence = tryLoadPreferredDateOrderFromPersistence(computed);
        if (fromPersistence != null) {
            preferredDateAxisOrder = new ArrayList<>(fromPersistence);
            return new ArrayList<>(fromPersistence);
        }
        return computed;
    }

    private static boolean sameMultisetLocalDate(List<LocalDate> a, List<LocalDate> b) {
        if (a == null || b == null || a.size() != b.size()) {
            return false;
        }
        HashMap<LocalDate, Integer> freq = new HashMap<>();
        for (LocalDate d : a) {
            freq.merge(d, 1, Integer::sum);
        }
        for (LocalDate d : b) {
            Integer n = freq.get(d);
            if (n == null || n <= 0) {
                return false;
            }
            if (n == 1) {
                freq.remove(d);
            } else {
                freq.put(d, n - 1);
            }
        }
        return freq.isEmpty();
    }

    private List<LocalDate> tryLoadPreferredDateOrderFromPersistence(List<LocalDate> computed) {
        List<TableColumnOrderPersistence.ColumnSpec> lay =
                TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE);
        if (lay == null || lay.isEmpty()) {
            return null;
        }
        List<String> titles = lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
        if (!wideStaticPrefixMatches(titles)
                || titles.size() != WIDE_STATIC_HEADERS.size() + computed.size()) {
            return null;
        }
        List<LocalDate> dates = parseDateTailAsDates(titles, WIDE_STATIC_HEADERS.size());
        if (dates == null || !sameMultisetLocalDate(dates, computed)) {
            return null;
        }
        return dates;
    }

    private static boolean wideStaticPrefixMatches(List<String> titles) {
        if (titles == null || titles.size() < WIDE_STATIC_HEADERS.size()) {
            return false;
        }
        for (int i = 0; i < WIDE_STATIC_HEADERS.size(); i++) {
            if (!WIDE_STATIC_HEADERS.get(i).equals(titles.get(i))) {
                return false;
            }
        }
        return true;
    }

    private boolean byDayStaticPrefixMatches(List<String> titles) {
        if (titles == null || titles.size() < BY_DAY_STATIC_HEADERS.size()) {
            return false;
        }
        for (int i = 0; i < BY_DAY_STATIC_HEADERS.size(); i++) {
            if (!BY_DAY_STATIC_HEADERS.get(i).equals(titles.get(i))) {
                return false;
            }
        }
        return true;
    }

    private static List<LocalDate> parseDateTailAsDates(List<String> titles, int staticCount) {
        if (titles == null || titles.size() < staticCount) {
            return null;
        }
        List<LocalDate> dates = new ArrayList<>();
        for (int i = staticCount; i < titles.size(); i++) {
            try {
                dates.add(LocalDate.parse(titles.get(i)));
            } catch (DateTimeParseException e) {
                return null;
            }
        }
        return dates;
    }

    private void persistDispatchColumnLayouts(List<String> wideTitles, List<String> byDayTitles) {
        List<TableColumnOrderPersistence.ColumnSpec> wideLay =
                TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE);
        List<Double> wideW =
                TableColumnOrderPersistence.resolveWidthsForHeaders(wideTitles, wideLay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> wideSpecs = new ArrayList<>();
        for (int i = 0; i < wideTitles.size(); i++) {
            wideSpecs.add(new TableColumnOrderPersistence.ColumnSpec(wideTitles.get(i), wideW.get(i)));
        }
        TableColumnOrderPersistence.saveLayout(TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE, wideSpecs);

        List<TableColumnOrderPersistence.ColumnSpec> byDayLay =
                TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY);
        List<Double> byDayW =
                TableColumnOrderPersistence.resolveWidthsForHeaders(byDayTitles, byDayLay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> byDaySpecs = new ArrayList<>();
        for (int i = 0; i < byDayTitles.size(); i++) {
            byDaySpecs.add(new TableColumnOrderPersistence.ColumnSpec(byDayTitles.get(i), byDayW.get(i)));
        }
        TableColumnOrderPersistence.saveLayout(
                TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY, byDaySpecs);
    }

    private void onWideSpreadsheetVisualColumnOrderChanged(List<String> titles) {
        if (suppressColumnReorderPersistence.get()) {
            return;
        }
        if (!wideStaticPrefixMatches(titles)) {
            return;
        }
        List<LocalDate> computed = computeDateAxisList();
        List<LocalDate> dates = parseDateTailAsDates(titles, WIDE_STATIC_HEADERS.size());
        if (dates == null || !sameMultisetLocalDate(dates, computed)) {
            return;
        }
        if (dates.equals(preferredDateAxisOrder)) {
            return;
        }
        preferredDateAxisOrder = new ArrayList<>(dates);
        persistDispatchColumnLayouts(titles, buildByDayColumnLabelsForAxis(dates));
        suppressColumnReorderPersistence.set(true);
        try {
            rebuildGrids();
        } finally {
            suppressColumnReorderPersistence.set(false);
        }
    }

    private void onByDaySpreadsheetVisualColumnOrderChanged(List<String> titles) {
        if (suppressColumnReorderPersistence.get()) {
            return;
        }
        if (!byDayStaticPrefixMatches(titles)) {
            return;
        }
        List<LocalDate> computed = computeDateAxisList();
        List<LocalDate> dates = parseDateTailAsDates(titles, BY_DAY_STATIC_HEADERS.size());
        if (dates == null || !sameMultisetLocalDate(dates, computed)) {
            return;
        }
        if (dates.equals(preferredDateAxisOrder)) {
            return;
        }
        preferredDateAxisOrder = new ArrayList<>(dates);
        persistDispatchColumnLayouts(buildWideColumnLabelsForAxis(dates), titles);
        suppressColumnReorderPersistence.set(true);
        try {
            rebuildGrids();
        } finally {
            suppressColumnReorderPersistence.set(false);
        }
    }

    private WideGridBundle buildWideGridModel(List<LocalDate> axis) {
        List<Map<String, String>> profiles = new ArrayList<>();
        List<WideRow> rowItems = new ArrayList<>();
        List<String> cols = doc.columns();
        profiles.addAll(
                ResultDispatchPivot.distinctWideTaskProfiles(
                        cols,
                        doc.rows(),
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS));
        profiles.sort(
                Comparator.comparing(DispatchInteractiveTabController::parseTrialOrderKey)
                        .thenComparing(p -> ResultDispatchNormalizer.staticGroupKey(cols, p)));
        assignSequentialTrialOrdersForProfiles(profiles);

        int staticCols = WIDE_STATIC_HEADERS.size();
        int dayCount = axis.size();
        int slotCols = dayCount * DAY_SLOT_COLUMNS;
        int totalCols = staticCols + slotCols;
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int gridRowsTotal = firstData + profiles.size();
        GridBase grid = new GridBase(gridRowsTotal, totalCols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(buildWideColumnLabelsForAxis(axis));

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < totalCols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(
                            SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_FILTER_ROW);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        for (int pr = 0; pr < profiles.size(); pr++) {
            Map<String, String> profile = profiles.get(pr);
            int gridRow = firstData + pr;
            WideRow wr = new WideRow(profile, axis.size());
            for (int j = 0; j < axis.size(); j++) {
                double v =
                        ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                                doc.rows(),
                                profile,
                                axis.get(j),
                                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
                wr.setAmount(j, v);
            }
            rowItems.add(wr);

            ObservableList<SpreadsheetCell> line = FXCollections.observableArrayList();
            for (int c = 0; c < staticCols; c++) {
                String title = WIDE_STATIC_HEADERS.get(c);
                String raw = wr.getStatic(title);
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw != null ? raw : "");
                cell.setEditable(c > 0);
                cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_LEADING_COL);
                line.add(cell);
            }
            for (int di = 0; di < dayCount; di++) {
                double dayAmt = wr.getAmount(di);
                String qtxt = dayAmt > 1e-9 ? ResultDispatchNormalizer.formatQty(dayAmt) : "";
                int col = staticCols + di * DAY_SLOT_COLUMNS;
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, col, 1, 1, qtxt);
                cell.setEditable(false);
                applyWideCellStyle(wr, di, cell);
                line.add(cell);
            }
            gridRows.add(line);
        }
        grid.setRows(gridRows);

        boolean[] wideBlockedCols = computeWideFullyBlockedDateColumns(dayCount);
        return new WideGridBundle(grid, profiles, rowItems, wideBlockedCols, staticCols, dayCount);
    }

    private ByDayGridBundle buildByDayGridModel(List<LocalDate> axis) {
        List<Map.Entry<String, String>> keys = ResultDispatchPivot.sortedProcessMachineKeys(doc.rows());
        int staticCols = BY_DAY_STATIC_HEADERS.size();
        int dayCount = axis.size();
        int slotCols = dayCount * DAY_SLOT_COLUMNS;
        int totalCols = staticCols + slotCols;
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int gridRowsTotal = firstData + keys.size();
        GridBase grid = new GridBase(gridRowsTotal, totalCols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(buildByDayColumnLabelsForAxis(axis));

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < totalCols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(
                            SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_FILTER_ROW);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        List<String> cols = doc.columns();
        List<ByDayRow> byItems = new ArrayList<>();
        for (Map.Entry<String, String> en : keys) {
            String pcSummary =
                    ResultDispatchPivot.processingContentSummaryForProcessMachine(
                            cols, doc.rows(), en.getKey(), en.getValue());
            ByDayRow br = new ByDayRow(en.getKey(), en.getValue(), pcSummary, axis.size());
            for (int j = 0; j < axis.size(); j++) {
                double v =
                        ResultDispatchPivot.sumQuantityForProcessMachineDate(
                                doc.rows(), en.getKey(), en.getValue(), axis.get(j));
                br.setAmount(j, v);
            }
            byItems.add(br);
        }

        for (int ir = 0; ir < byItems.size(); ir++) {
            ByDayRow br = byItems.get(ir);
            int gridRow = firstData + ir;
            ObservableList<SpreadsheetCell> line = FXCollections.observableArrayList();
            SpreadsheetCell c0 =
                    SpreadsheetCellType.STRING.createCell(gridRow, 0, 1, 1, br.process());
            c0.setEditable(false);
            c0.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_LEADING_COL);
            line.add(c0);
            SpreadsheetCell c1 =
                    SpreadsheetCellType.STRING.createCell(gridRow, 1, 1, 1, br.machine());
            c1.setEditable(false);
            c1.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_LEADING_COL);
            line.add(c1);
            SpreadsheetCell c2 =
                    SpreadsheetCellType.STRING.createCell(gridRow, 2, 1, 1, br.processingContent());
            c2.setEditable(false);
            c2.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_LEADING_COL);
            line.add(c2);
            for (int di = 0; di < dayCount; di++) {
                double dayAmt = br.getAmount(di);
                String qtxt = dayAmt > 1e-9 ? ResultDispatchNormalizer.formatQty(dayAmt) : "";
                int col = staticCols + di * DAY_SLOT_COLUMNS;
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, col, 1, 1, qtxt);
                cell.setEditable(false);
                applyByDayCellStyle(br, di, cell);
                line.add(cell);
            }
            gridRows.add(line);
        }
        grid.setRows(gridRows);

        boolean[] byDayBlockedCols = computeByDayFullyBlockedDateColumns(dayCount);
        return new ByDayGridBundle(grid, byDayBlockedCols, staticCols, dayCount);
    }

    private void applyFullGridRebuild(FullGridRebuild bundle) {
        suppressDispatchGridDirty.set(true);
        BitSet wideHiddenSnapshot = SpreadsheetTabularSupport.snapshotHiddenRows(wideSpreadsheet);
        BitSet byDayHiddenSnapshot = SpreadsheetTabularSupport.snapshotHiddenRows(byDaySpreadsheet);
        try {
            dateAxis.clear();
            dateAxis.addAll(bundle.axis());
            wideProfiles.clear();
            wideProfiles.addAll(bundle.wide().profiles());
            wideRowItems.clear();
            wideRowItems.addAll(bundle.wide().rowItems());

            WideGridBundle w = bundle.wide();
            w.grid().addEventHandler(GridChange.GRID_CHANGE_EVENT, this::onWideGridChange);
            wideSpreadsheet.setGrid(w.grid());
            wideSpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);
            SpreadsheetTabularSupport.applyColumnFiltersWithDialog(wideSpreadsheet);
            scheduleWideLayoutAfterColumnSync(w);

            ByDayGridBundle b = bundle.byDay();
            byDaySpreadsheet.setGrid(b.grid());
            byDaySpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);
            SpreadsheetTabularSupport.applyColumnFiltersWithDialog(byDaySpreadsheet);
            scheduleByDayLayoutAfterColumnSync(b);

            SpreadsheetTabularSupport.restoreHiddenRows(wideSpreadsheet, wideHiddenSnapshot);
            SpreadsheetTabularSupport.restoreHiddenRows(byDaySpreadsheet, byDayHiddenSnapshot);
        } finally {
            Platform.runLater(() -> suppressDispatchGridDirty.set(false));
        }
    }

    /**
     * After {@link SpreadsheetView#setGrid}, the inner {@link TableView} may add columns on the next layout pulse.
     * Retrying avoids applying widths while {@link SpreadsheetView#getColumns()} is still shorter than the grid
     * (which skipped date columns and looked like “no date columns”).
     */
    private void scheduleWideLayoutAfterColumnSync(WideGridBundle w) {
        final int expectedCols = w.staticCols() + w.dayCount() * DAY_SLOT_COLUMNS;
        final int[] attempts = {0};
        final Runnable[] job = new Runnable[1];
        job[0] =
                () -> {
                    attempts[0]++;
                    int actual = wideSpreadsheet.getColumns().size();
                    boolean retry = actual < expectedCols && attempts[0] < 48;
                    if (retry) {
                        Platform.runLater(job[0]);
                        return;
                    }
                    SpreadsheetTabularSupport.applyFixedLeadingColumns(
                            wideSpreadsheet, WIDE_STATIC_HEADERS.size());
                    SpreadsheetTabularSupport.pinSpreadsheetFilterRow(wideSpreadsheet);
                    SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(wideSpreadsheet);
                    applyDateColumnWidthsForBlockedDays(
                            wideSpreadsheet,
                            w.staticCols(),
                            w.dayCount(),
                            sanitizeFullyBlockedFlagsForColumnWidth(w.blockedCols()));
                    SpreadsheetColumnDragReorderSupport.refreshAfterGridReady(
                            wideSpreadsheet,
                            suppressColumnReorderPersistence::get,
                            () -> new ArrayList<>(buildWideColumnLabelsForAxis(dateAxis)),
                            WIDE_STATIC_HEADERS.size(),
                            this::onWideSpreadsheetVisualColumnOrderChanged);
                    ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                            wideSpreadsheet,
                            () -> new ArrayList<>(buildWideColumnLabelsForAxis(dateAxis)),
                            () -> {
                                List<String> h = buildWideColumnLabelsForAxis(dateAxis);
                                return TableColumnOrderPersistence.loadColumnVisibility(
                                        TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE,
                                        h.size());
                            });
                };
        Platform.runLater(job[0]);
    }

    private void scheduleByDayLayoutAfterColumnSync(ByDayGridBundle b) {
        final int expectedCols = b.staticCols() + b.dayCount() * DAY_SLOT_COLUMNS;
        final int[] attempts = {0};
        final Runnable[] job = new Runnable[1];
        job[0] =
                () -> {
                    attempts[0]++;
                    if (byDaySpreadsheet.getColumns().size() < expectedCols && attempts[0] < 48) {
                        Platform.runLater(job[0]);
                        return;
                    }
                    SpreadsheetTabularSupport.applyFixedLeadingColumns(
                            byDaySpreadsheet, BY_DAY_STATIC_HEADERS.size());
                    SpreadsheetTabularSupport.pinSpreadsheetFilterRow(byDaySpreadsheet);
                    SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(byDaySpreadsheet);
                    applyDateColumnWidthsForBlockedDays(
                            byDaySpreadsheet,
                            b.staticCols(),
                            b.dayCount(),
                            sanitizeFullyBlockedFlagsForColumnWidth(b.blockedCols()));
                    SpreadsheetColumnDragReorderSupport.refreshAfterGridReady(
                            byDaySpreadsheet,
                            suppressColumnReorderPersistence::get,
                            () -> new ArrayList<>(buildByDayColumnLabelsForAxis(dateAxis)),
                            BY_DAY_STATIC_HEADERS.size(),
                            this::onByDaySpreadsheetVisualColumnOrderChanged);
                    ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                            byDaySpreadsheet,
                            () -> new ArrayList<>(buildByDayColumnLabelsForAxis(dateAxis)),
                            () -> {
                                List<String> h = buildByDayColumnLabelsForAxis(dateAxis);
                                return TableColumnOrderPersistence.loadColumnVisibility(
                                        TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY,
                                        h.size());
                            });
                };
        Platform.runLater(job[0]);
    }

    /**
     * When every date column is “fully blocked”, narrowing all of them to ~5pt makes the timeline disappear. Keep
     * default widths in that case (still gray cells via styles).
     */
    private static boolean[] sanitizeFullyBlockedFlagsForColumnWidth(boolean[] fullyBlocked) {
        if (fullyBlocked == null || fullyBlocked.length == 0) {
            return fullyBlocked;
        }
        int trueCount = 0;
        for (boolean b : fullyBlocked) {
            if (b) {
                trueCount++;
            }
        }
        if (trueCount == fullyBlocked.length) {
            return new boolean[fullyBlocked.length];
        }
        return fullyBlocked;
    }

    private List<String> buildWideColumnLabelsForAxis(List<LocalDate> axis) {
        List<String> headers =
                new ArrayList<>(WIDE_STATIC_HEADERS.size() + axis.size() * DAY_SLOT_COLUMNS);
        headers.addAll(WIDE_STATIC_HEADERS);
        for (LocalDate d : axis) {
            headers.add(d.toString());
        }
        return headers;
    }

    private List<String> buildByDayColumnLabelsForAxis(List<LocalDate> axis) {
        List<String> headers =
                new ArrayList<>(BY_DAY_STATIC_HEADERS.size() + axis.size() * DAY_SLOT_COLUMNS);
        headers.addAll(BY_DAY_STATIC_HEADERS);
        for (LocalDate d : axis) {
            headers.add(d.toString());
        }
        return headers;
    }

    private static double pointsToLocalPixels(double pt) {
        Screen s = Screen.getPrimary();
        if (s == null) {
            return pt * 96.0 / 72.0;
        }
        return pt * s.getDpi() / 72.0;
    }

    /**
     * Date columns where every profile row is blocked on that day get a narrow width; mixed columns stay default.
     */
    private boolean[] computeWideFullyBlockedDateColumns(int dayCount) {
        return new boolean[dayCount];
    }

    private boolean[] computeByDayFullyBlockedDateColumns(int dayCount) {
        return new boolean[dayCount];
    }

    private static void applyDateColumnWidthsForBlockedDays(
            SpreadsheetView view, int staticCols, int dayCount, boolean[] fullyBlocked) {
        if (view == null || fullyBlocked == null || dayCount <= 0) {
            return;
        }
        var cols = view.getColumns();
        if (cols.isEmpty()) {
            return;
        }
        double narrowPt = pointsToLocalPixels(BLOCKED_DATE_COLUMN_PREF_PT);
        double narrow = Math.max(narrowPt, MIN_BLOCKED_DATE_COLUMN_WIDTH_PX);
        for (int di = 0; di < dayCount; di++) {
            int colIndex = staticCols + di * DAY_SLOT_COLUMNS;
            if (colIndex >= cols.size()) {
                continue;
            }
            SpreadsheetColumn sc = cols.get(colIndex);
            if (fullyBlocked[di]) {
                sc.setPrefWidth(narrow);
                sc.setMinWidth(narrow);
                sc.setMaxWidth(narrow);
            } else {
                sc.setMinWidth(MIN_DATE_COLUMN_WIDTH_PX);
                sc.setPrefWidth(Region.USE_COMPUTED_SIZE);
                sc.setMaxWidth(Double.MAX_VALUE);
            }
        }
    }

    private void applyWideCellStyle(WideRow wr, int dateIdx, SpreadsheetCell cell) {
        if (isWideDispatchShortfall(wr, dateIdx)) {
            cell.setStyle(DATE_CELL_STYLE_SHORTFALL);
            return;
        }
        double q = wr.getAmount(dateIdx);
        if (q > 1e-9) {
            cell.setStyle(DATE_CELL_STYLE_POSITIVE_QTY);
        } else {
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_DATA_WHITE);
        }
    }

    private static double parseTrialOrderKey(Map<String, String> profile) {
        String s = profile.get(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER);
        if (s == null || s.isBlank()) {
            return Double.MAX_VALUE;
        }
        try {
            return Double.parseDouble(s.trim().replace(",", ""));
        } catch (NumberFormatException e) {
            return Double.MAX_VALUE;
        }
    }

    /** Ensures each profile row maps to sequential trial order and pushes into doc rows. */
    private void assignSequentialTrialOrdersForProfiles(List<Map<String, String>> profiles) {
        for (int i = 0; i < profiles.size(); i++) {
            String ord = Integer.toString(i + 1);
            Map<String, String> prof = profiles.get(i);
            prof.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, ord);
            for (Map<String, String> row : doc.rows()) {
                if (ResultDispatchPivot.matchesWideMergeIdentity(
                        prof,
                        row,
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS)) {
                    row.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, ord);
                }
            }
        }
    }

    private void swapWideProfiles(int a, int b) {
        Map<String, String> pa = wideProfiles.get(a);
        Map<String, String> pb = wideProfiles.get(b);
        wideProfiles.set(a, pb);
        wideProfiles.set(b, pa);
        assignSequentialTrialOrdersForProfiles(wideProfiles);
        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
        rebuildGrids();
        markDispatchDocDirty();
    }

    /** Column index in {@link SpreadsheetView#getColumns()} for the focused / primary selected cell. */
    private int wideSpreadsheetFocusedColumnIndex() {
        var sm = wideSpreadsheet.getSelectionModel();
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

    /**
     * After row reorder, move selection and focus to the same logical profile row (by index) and column.
     * Runs later so spreadsheet layout is applied before selecting.
     */
    private void focusWideProfileCellAfterReorder(int profileIndex, int columnIndex) {
        if (profileIndex < 0 || profileIndex >= wideProfiles.size()) {
            return;
        }
        var cols = wideSpreadsheet.getColumns();
        if (cols.isEmpty()) {
            return;
        }
        int c = Math.max(0, Math.min(columnIndex, cols.size() - 1));
        SpreadsheetColumn scol = cols.get(c);
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int modelGridRow = firstData + profileIndex;
        Platform.runLater(
                () -> {
                    int viewRow = wideSpreadsheet.getViewRow(modelGridRow);
                    if (viewRow < 0) {
                        return;
                    }
                    var sm = wideSpreadsheet.getSelectionModel();
                    sm.clearSelection();
                    sm.clearAndSelect(viewRow, scol);
                    sm.focus(viewRow, scol);
                    scrollWideSpreadsheetCellIntoView(viewRow, scol);
                });
    }

    /**
     * 未達サマリ表で選択した依頼NO・機械・配台日に対応するワイドグリッドのセルを選択・フォーカスする。
     */
    private void wireDispatchShortfallSelectionToWideGrid() {
        if (dispatchShortfallTable == null) {
            return;
        }
        dispatchShortfallTable
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((obs, prev, row) -> focusWideSpreadsheetOnDispatchShortfallRow(row));
    }

    /**
     * {@link DispatchTrialShortages#wideShortfallKey} と同一規則でプロファイル行を特定する。
     *
     * @return {@link #wideProfiles} インデックス、無ければ -1
     */
    private int findWideProfileIndexMatchingShortfall(DispatchQtyShortfallRow row) {
        if (row == null || wideProfiles.isEmpty()) {
            return -1;
        }
        String expected =
                DispatchTrialShortages.wideShortfallKey(
                        row.taskId(), row.machineName(), row.dispatchDateIso());
        for (int i = 0; i < wideProfiles.size(); i++) {
            Map<String, String> p = wideProfiles.get(i);
            String candidate =
                    DispatchTrialShortages.wideShortfallKey(
                            p.get("依頼NO"),
                            p.get(ResultDispatchSchema.COL_MACHINE),
                            row.dispatchDateIso());
            if (expected.equals(candidate)) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 列同期が直後に終わっていないときは数パルス待ってからフォーカスする（{@link #scheduleWideLayoutAfterColumnSync}
     * と同趣旨）。
     */
    private void scheduleFocusWideCellWhenShortfallReady(int profileIdx, int modelCol) {
        final int[] attempts = {0};
        final Runnable[] job = new Runnable[1];
        job[0] =
                () -> {
                    attempts[0]++;
                    var cols = wideSpreadsheet.getColumns();
                    boolean colsReady =
                            !cols.isEmpty()
                                    && modelCol >= 0
                                    && modelCol < cols.size()
                                    && profileIdx >= 0
                                    && profileIdx < wideProfiles.size();
                    if (!colsReady) {
                        if (attempts[0] < 48) {
                            Platform.runLater(job[0]);
                        }
                        return;
                    }
                    focusWideProfileCellAfterReorder(profileIdx, modelCol);
                };
        Platform.runLater(job[0]);
    }

    private void focusWideSpreadsheetOnDispatchShortfallRow(DispatchQtyShortfallRow row) {
        if (row == null || wideProfiles.isEmpty() || dateAxis.isEmpty()) {
            return;
        }
        int profileIdx = findWideProfileIndexMatchingShortfall(row);
        if (profileIdx < 0) {
            return;
        }
        LocalDate targetDate;
        try {
            targetDate = LocalDate.parse(row.dispatchDateIso().trim());
        } catch (DateTimeParseException e) {
            return;
        }
        int dateIdx = dateAxis.indexOf(targetDate);
        if (dateIdx < 0) {
            return;
        }
        int staticCols = WIDE_STATIC_HEADERS.size();
        int modelCol = staticCols + dateIdx * DAY_SLOT_COLUMNS;
        applyInnerTabSelectedIndex(0);
        scheduleFocusWideCellWhenShortfallReady(profileIdx, modelCol);
    }

    /** ControlsFX の {@link SpreadsheetView#scrollToRow} / {@link SpreadsheetView#scrollToColumn} で見える位置へ寄せる。 */
    private void scrollWideSpreadsheetCellIntoView(int viewRow, SpreadsheetColumn scol) {
        if (scol == null || viewRow < 0) {
            return;
        }
        wideSpreadsheet.scrollToRow(viewRow);
        wideSpreadsheet.scrollToColumn(scol);
    }

    private int selectedWideProfileIndex() {
        var sm = wideSpreadsheet.getSelectionModel();
        TablePosition<?, ?> pos = sm.getFocusedCell();
        if (pos == null || pos.getRow() < 0) {
            var cells = sm.getSelectedCells();
            if (cells == null || cells.isEmpty()) {
                return -1;
            }
            pos = cells.getFirst();
        }
        int viewRow = pos.getRow();
        int gridRow = wideSpreadsheet.getModelRow(viewRow);
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int idx = gridRow - firstData;
        if (idx >= 0 && idx < wideProfiles.size()) {
            return idx;
        }
        return -1;
    }

    /** Maps a SpreadsheetView table/view row index to a {@link #wideProfiles} index. */
    private int wideProfileIndexFromViewRow(int viewRow) {
        if (viewRow < 0) {
            return -1;
        }
        int gridRow = wideSpreadsheet.getModelRow(viewRow);
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int idx = gridRow - firstData;
        if (idx >= 0 && idx < wideProfiles.size()) {
            return idx;
        }
        return -1;
    }

    private int wideProfileIndexFromTableCell(TableCell<?, ?> tc) {
        if (tc == null) {
            return -1;
        }
        return wideProfileIndexFromViewRow(tc.getIndex());
    }

    /**
     * Maps a {@link TableCell}'s column to model column index (accounts for hidden columns / ControlsFX mapping).
     */
    private int wideModelColumnFromTableCell(TableCell<?, ?> tc) {
        if (tc == null || tc.getTableColumn() == null) {
            return -1;
        }
        int viewCol = tc.getTableView().getColumns().indexOf(tc.getTableColumn());
        if (viewCol < 0) {
            return -1;
        }
        return wideSpreadsheet.getModelColumn(viewCol);
    }

    private void applyByDayCellStyle(ByDayRow br, int dateIdx, SpreadsheetCell cell) {
        if (isByDayDispatchShortfall(br, dateIdx)) {
            cell.setStyle(DATE_CELL_STYLE_SHORTFALL);
            return;
        }
        double q = br.getAmount(dateIdx);
        if (q > 1e-9) {
            cell.setStyle(DATE_CELL_STYLE_POSITIVE_QTY);
        } else {
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_DATA_WHITE);
        }
    }

    private boolean isWideDispatchShortfall(WideRow wr, int dateIdx) {
        if (dispatchWideShortfallKeys.isEmpty()
                || dateIdx < 0
                || dateIdx >= dateAxis.size()) {
            return false;
        }
        String tid = wr.getStatic("依頼NO");
        String mach = wr.getStatic(ResultDispatchSchema.COL_MACHINE);
        LocalDate d = dateAxis.get(dateIdx);
        String key =
                DispatchTrialShortages.wideShortfallKey(tid, mach, d.toString());
        return dispatchWideShortfallKeys.contains(key);
    }

    private boolean isByDayDispatchShortfall(ByDayRow br, int dateIdx) {
        if (dispatchByDayShortfallKeys.isEmpty()
                || dateIdx < 0
                || dateIdx >= dateAxis.size()) {
            return false;
        }
        String mach = br.machine();
        LocalDate d = dateAxis.get(dateIdx);
        return dispatchByDayShortfallKeys.contains(
                DispatchTrialShortages.byDayShortfallKey(mach, d.toString()));
    }

    private void installDispatchShortfallColumns(TableView<DispatchQtyShortfallRow> tv) {
        if (tv == null) {
            return;
        }
        tv.getColumns().clear();
        TableColumn<DispatchQtyShortfallRow, String> c0 = new TableColumn<>("依頼NO");
        c0.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(Objects.toString(cd.getValue().taskId(), "")));
        TableColumn<DispatchQtyShortfallRow, String> c1 = new TableColumn<>("機械名");
        c1.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(Objects.toString(cd.getValue().machineName(), "")));
        TableColumn<DispatchQtyShortfallRow, String> c2 = new TableColumn<>("配台日");
        c2.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                Objects.toString(cd.getValue().dispatchDateIso(), "")));
        TableColumn<DispatchQtyShortfallRow, String> c3 = new TableColumn<>("目標(m)");
        c3.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                formatShortfallMeters(cd.getValue().targetM())));
        TableColumn<DispatchQtyShortfallRow, String> c4 = new TableColumn<>("実績(m)");
        c4.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                formatShortfallMeters(cd.getValue().doneM())));
        TableColumn<DispatchQtyShortfallRow, String> c5 = new TableColumn<>("不足(m)");
        c5.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                formatShortfallMeters(cd.getValue().shortfallM())));
        TableColumn<DispatchQtyShortfallRow, String> c6 = new TableColumn<>("補足");
        c6.setPrefWidth(280);
        c6.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(Objects.toString(cd.getValue().note(), "")));
        tv.getColumns().addAll(c0, c1, c2, c3, c4, c5, c6);
    }

    private static String formatShortfallMeters(double m) {
        if (Double.isNaN(m) || Double.isInfinite(m)) {
            return "";
        }
        if (Math.abs(m - Math.rint(m)) < 1e-6) {
            return Long.toString((long) Math.rint(m));
        }
        return String.format("%.3f", m);
    }

    /** {@code 結果_配台表.json} と同じフォルダの {@code dispatch_trial_shortages.json} から未達行を読み UI に反映する。 */
    private void applyDispatchShortfallFromDisk(Path resultDispatchJson) {
        if (resultDispatchJson == null) {
            clearDispatchShortfallUi();
            return;
        }
        Path shortagePath = resultDispatchJson.resolveSibling("dispatch_trial_shortages.json");
        if (!Files.isRegularFile(shortagePath)) {
            clearDispatchShortfallUi();
            return;
        }
        try {
            DispatchTrialShortages.FullBundle fb = DispatchTrialShortages.readFull(shortagePath);
            List<DispatchQtyShortfallRow> rows = fb.dispatchQtyShortfall();
            lastDispatchShortageHints = List.copyOf(fb.shortageHints());
            applyDispatchShortfallRows(rows);
        } catch (IOException e) {
            clearDispatchShortfallUi();
        }
    }

    private void applyDispatchShortfallRows(List<DispatchQtyShortfallRow> rows) {
        lastDispatchShortfallRows = rows != null ? List.copyOf(rows) : List.of();
        dispatchWideShortfallKeys.clear();
        dispatchByDayShortfallKeys.clear();
        for (DispatchQtyShortfallRow r : lastDispatchShortfallRows) {
            dispatchWideShortfallKeys.add(
                    DispatchTrialShortages.wideShortfallKey(
                            r.taskId(), r.machineName(), r.dispatchDateIso()));
            dispatchByDayShortfallKeys.add(
                    DispatchTrialShortages.byDayShortfallKey(
                            r.machineName(), r.dispatchDateIso()));
        }
        if (dispatchShortfallTable != null) {
            dispatchShortfallTable.getItems().setAll(lastDispatchShortfallRows);
        }
        boolean vis = !lastDispatchShortfallRows.isEmpty();
        if (dispatchShortfallPanel != null) {
            dispatchShortfallPanel.setVisible(vis);
            dispatchShortfallPanel.setManaged(vis);
        }
    }

    private void clearDispatchShortfallUi() {
        lastDispatchShortfallRows = List.of();
        lastDispatchShortageHints = List.of();
        dispatchWideShortfallKeys.clear();
        dispatchByDayShortfallKeys.clear();
        if (dispatchShortfallTable != null) {
            dispatchShortfallTable.getItems().clear();
        }
        if (dispatchShortfallPanel != null) {
            dispatchShortfallPanel.setVisible(false);
            dispatchShortfallPanel.setManaged(false);
        }
    }

    /**
     * 配台試行（段階3）成功後、未達があるときにモーダル表示する。{@link DispatchTrialUnassignedWizard} より前。
     *
     * <p>手動確認の目安: 機械カレンダー等でブロックされる暦日に当日配台数量を置いて試行し、該当の日付セルが赤表示になり、
     * ツールバー下サマリ表と本ダイアログに同一内容が並ぶこと。
     */
    private void showDispatchQtyShortfallDialogIfNeeded(Stage owner) {
        if (lastDispatchShortfallRows == null || lastDispatchShortfallRows.isEmpty()) {
            return;
        }
        TableView<DispatchQtyShortfallRow> tv = new TableView<>();
        installDispatchShortfallColumns(tv);
        tv.getItems().setAll(lastDispatchShortfallRows);
        tv.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

        Label head =
                new Label(
                        "次の暦日で、タイムライン上の割付が目標メートルに届きませんでした。"
                                + " 機械カレンダー・人員・その他ブロックでその日に追加できなかった可能性があります。");
        head.setWrapText(true);
        head.setStyle("-fx-font-size: 13px;");
        BorderPane root = new BorderPane();
        root.setTop(head);
        BorderPane.setMargin(head, new Insets(10, 14, 8, 14));
        root.setCenter(tv);
        BorderPane.setMargin(tv, new Insets(0, 14, 14, 14));

        Stage st = new Stage();
        if (owner != null) {
            st.initOwner(owner);
        }
        st.initModality(Modality.APPLICATION_MODAL);
        st.setTitle("配台数量未達（タイムライン実績）");
        Scene sc = new Scene(root, 920, 520);
        if (shell != null) {
            shell.registerThemeTrackedScene(sc);
        }
        st.setScene(sc);
        st.setOnHidden(
                ev -> {
                    if (shell != null) {
                        shell.unregisterThemeTrackedScene(sc);
                    }
                });
        st.showAndWait();
    }

    /**
     * 配台試行後、{@code dispatch_trial_shortages.json} の op_shortage / as_shortage が空でなければモーダルで示す。
     * メートル未達（{@link #showDispatchQtyShortfallDialogIfNeeded}）とは別系統。
     */
    private void showDispatchShortageHintsDialogIfNeeded(Stage owner) {
        if (lastDispatchShortageHints == null || lastDispatchShortageHints.isEmpty()) {
            return;
        }
        TableView<DispatchTrialShortages.ShortageHint> tv = new TableView<>();
        tv.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

        TableColumn<DispatchTrialShortages.ShortageHint, String> h0 = new TableColumn<>("依頼NO");
        h0.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                Objects.toString(cd.getValue().taskId(), "")));
        TableColumn<DispatchTrialShortages.ShortageHint, String> h1 = new TableColumn<>("理由");
        h1.setPrefWidth(220);
        h1.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                Objects.toString(cd.getValue().reason(), "")));
        TableColumn<DispatchTrialShortages.ShortageHint, String> h2 = new TableColumn<>("補足");
        h2.setPrefWidth(380);
        h2.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                Objects.toString(cd.getValue().detail(), "")));
        tv.getColumns().addAll(h0, h1, h2);
        tv.getItems().setAll(lastDispatchShortageHints);

        Label head =
                new Label(
                        "フォーム候補不足（op_shortage）または人数は足りるが割当不可（as_shortage）として記録された件です。"
                                + " メートル目標未達とは別に検出されます。");
        head.setWrapText(true);
        head.setStyle("-fx-font-size: 13px;");
        BorderPane root = new BorderPane();
        root.setTop(head);
        BorderPane.setMargin(head, new Insets(10, 14, 8, 14));
        root.setCenter(tv);
        BorderPane.setMargin(tv, new Insets(0, 14, 14, 14));

        Stage st = new Stage();
        if (owner != null) {
            st.initOwner(owner);
        }
        st.initModality(Modality.APPLICATION_MODAL);
        st.setTitle("人員・割当不足（試行スナップショット）");
        Scene sc = new Scene(root, 920, 480);
        if (shell != null) {
            shell.registerThemeTrackedScene(sc);
        }
        st.setScene(sc);
        st.setOnHidden(
                ev -> {
                    if (shell != null) {
                        shell.unregisterThemeTrackedScene(sc);
                    }
                });
        st.showAndWait();
    }

    private void installWideDnDHandlers() {
        wideSpreadsheet.addEventFilter(
                MouseEvent.DRAG_DETECTED,
                e -> {
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int modelCol = wideModelColumnFromTableCell(tc);
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (modelCol < 0) {
                        return;
                    }

                    // Row reorder: drag from leading (static) model columns ? start gesture from the TableCell node.
                    if (modelCol < staticCols) {
                        int profIdx = wideProfileIndexFromTableCell(tc);
                        if (profIdx < 0 || profIdx >= wideProfiles.size()) {
                            return;
                        }
                        List<String> cols = doc.columns();
                        String gk = ResultDispatchNormalizer.staticGroupKey(cols, wideProfiles.get(profIdx));
                        String b64 =
                                Base64.getUrlEncoder()
                                        .withoutPadding()
                                        .encodeToString(gk.getBytes(StandardCharsets.UTF_8));
                        Dragboard db = tc.startDragAndDrop(TransferMode.MOVE);
                        ClipboardContent cc = new ClipboardContent();
                        cc.putString(DND_ROW_PREFIX + b64);
                        db.setContent(cc);
                        SpreadsheetRowReorderDragGhost.apply(db, tc, e);
                        e.consume();
                        return;
                    }

                    int slot = modelCol - staticCols;
                    int dateIdx = slot / DAY_SLOT_COLUMNS;
                    int profIdx = wideProfileIndexFromTableCell(tc);
                    if (profIdx < 0 || profIdx >= wideRowItems.size() || dateIdx < 0 || dateIdx >= dateAxis.size()) {
                        return;
                    }
                    WideRow wr = wideRowItems.get(profIdx);
                    double qty = wr.getAmount(dateIdx);
                    if (qty <= 1e-9) {
                        return;
                    }
                    Dragboard db = tc.startDragAndDrop(TransferMode.MOVE);
                    ClipboardContent cc = new ClipboardContent();
                    List<String> cols = doc.columns();
                    String gk = ResultDispatchNormalizer.staticGroupKey(cols, wr.profileMap());
                    String b64 =
                            Base64.getUrlEncoder()
                                    .withoutPadding()
                                    .encodeToString(gk.getBytes(StandardCharsets.UTF_8));
                    cc.putString(DND_PREFIX + DND_V2_MARKER + b64 + ":" + dateIdx + ":" + qty);
                    db.setContent(cc);
                    e.consume();
                });

        wideSpreadsheet.addEventFilter(
                DragEvent.DRAG_OVER,
                e -> {
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int modelCol = wideModelColumnFromTableCell(tc);
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (modelCol < 0) {
                        return;
                    }

                    if (modelCol < staticCols) {
                        if (e.getDragboard().hasString()
                                && e.getDragboard().getString().startsWith(DND_ROW_PREFIX)) {
                            int profIdx = wideProfileIndexFromTableCell(tc);
                            if (profIdx >= 0 && profIdx < wideProfiles.size()) {
                                e.acceptTransferModes(TransferMode.MOVE);
                            }
                        }
                        e.consume();
                        return;
                    }

                    int profIdx = wideProfileIndexFromTableCell(tc);
                    if (profIdx < 0 || profIdx >= wideRowItems.size()) {
                        return;
                    }
                    if (e.getDragboard().hasString()
                            && e.getDragboard().getString().startsWith(DND_PREFIX)) {
                        e.acceptTransferModes(TransferMode.MOVE);
                    }
                    e.consume();
                });

        wideSpreadsheet.addEventFilter(
                DragEvent.DRAG_DROPPED,
                e -> {
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int modelCol = wideModelColumnFromTableCell(tc);
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (modelCol < 0) {
                        return;
                    }

                    if (modelCol < staticCols) {
                        String payload = e.getDragboard().getString();
                        if (payload != null && payload.startsWith(DND_ROW_PREFIX)) {
                            boolean ok = handleWideRowReorderDrop(payload, tc);
                            e.setDropCompleted(ok);
                        } else {
                            e.setDropCompleted(false);
                        }
                        e.consume();
                        return;
                    }

                    int slot = modelCol - staticCols;
                    int dateIdx = slot / DAY_SLOT_COLUMNS;
                    int profIdx = wideProfileIndexFromTableCell(tc);
                    if (profIdx < 0 || profIdx >= wideRowItems.size()) {
                        return;
                    }
                    WideRow tgt = wideRowItems.get(profIdx);
                    boolean ok = handleWideDrop(e.getDragboard().getString(), tgt, dateIdx);
                    e.setDropCompleted(ok);
                    e.consume();
                });
    }

    private void installByDayDoubleClickHandler() {
        byDaySpreadsheet.addEventFilter(
                MouseEvent.MOUSE_CLICKED,
                e -> {
                    if (e.getClickCount() != 2) {
                        return;
                    }
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(byDaySpreadsheet, tc)) {
                        return;
                    }
                    int col = tc.getTableView().getColumns().indexOf(tc.getTableColumn());
                    int staticCols = BY_DAY_STATIC_HEADERS.size();
                    if (col < staticCols) {
                        return;
                    }
                    int slot = col - staticCols;
                    int dateIdx = slot / DAY_SLOT_COLUMNS;
                    int row = tc.getIndex();
                    int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
                    int dataIdx = row - firstData;
                    List<Map.Entry<String, String>> keys = ResultDispatchPivot.sortedProcessMachineKeys(doc.rows());
                    if (dataIdx < 0 || dataIdx >= keys.size()) {
                        return;
                    }
                    Map.Entry<String, String> en = keys.get(dataIdx);
                    double cur =
                            ResultDispatchPivot.sumQuantityForProcessMachineDate(
                                    doc.rows(), en.getKey(), en.getValue(), dateAxis.get(dateIdx));
                    TextInputDialog dialog =
                            new TextInputDialog(ResultDispatchNormalizer.formatQty(cur));
                    dialog.initOwner(shell != null ? shell.getPrimaryStage() : null);
                    dialog.setTitle("日別合計");
                    dialog.setHeaderText(
                            en.getKey()
                                    + " / "
                                    + en.getValue()
                                    + " / "
                                    + dateAxis.get(dateIdx));
                    Optional<String> ov = dialog.showAndWait();
                    ov.filter(s -> !s.isBlank())
                            .ifPresent(
                                    s -> {
                                        double newTotal = ResultDispatchNormalizer.parseDouble(s);
                                        ResultDispatchPivot.scaleProcessMachineDateToTotal(
                                                doc.columns(),
                                                doc.rows(),
                                                en.getKey(),
                                                en.getValue(),
                                                dateAxis.get(dateIdx),
                                                newTotal);
                                        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
                                        rebuildGrids();
                                        markDispatchDocDirty();
                                    });
                });
    }

    private static TableCell<?, ?> findTableCell(Node n) {
        while (n != null) {
            if (n instanceof TableCell<?, ?> tc) {
                return tc;
            }
            n = n.getParent();
        }
        return null;
    }

    private static boolean isUnderSpreadsheet(SpreadsheetView spv, Node node) {
        Node n = node;
        while (n != null) {
            if (n == spv) {
                return true;
            }
            n = n.getParent();
        }
        return false;
    }

    private int wideProfileIndexForRow(WideRow row) {
        List<String> cols = doc.columns();
        String gk = ResultDispatchNormalizer.staticGroupKey(cols, row.profileMap());
        return indexOfProfileGroupKey(gk);
    }

    private int indexOfProfileGroupKey(String groupKey) {
        List<String> cols = doc.columns();
        for (int i = 0; i < wideProfiles.size(); i++) {
            if (ResultDispatchNormalizer.staticGroupKey(cols, wideProfiles.get(i)).equals(groupKey)) {
                return i;
            }
        }
        return -1;
    }

    private boolean handleWideRowReorderDrop(String payload, TableCell<?, ?> targetCell) {
        if (payload == null || !payload.startsWith(DND_ROW_PREFIX)) {
            return false;
        }
        String b64 = payload.substring(DND_ROW_PREFIX.length());
        final String gk;
        try {
            gk = new String(Base64.getUrlDecoder().decode(b64), StandardCharsets.UTF_8);
        } catch (IllegalArgumentException ex) {
            return false;
        }
        int fromIdx = indexOfProfileGroupKey(gk);
        int toIdx = wideProfileIndexFromTableCell(targetCell);
        if (fromIdx < 0 || toIdx < 0 || fromIdx >= wideProfiles.size() || toIdx >= wideProfiles.size()) {
            return false;
        }
        if (fromIdx == toIdx) {
            return false;
        }
        wideProfiles.add(toIdx, wideProfiles.remove(fromIdx));
        assignSequentialTrialOrdersForProfiles(wideProfiles);
        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
        rebuildGrids();
        statusLabel.setText("行を移動しました");
        markDispatchDocDirty();
        return true;
    }

    private boolean handleWideDrop(String payload, WideRow targetRow, int targetDateIdx) {
        if (!payload.startsWith(DND_PREFIX)) {
            return false;
        }
        String rest = payload.substring(DND_PREFIX.length());
        boolean payloadIsV2 = rest.startsWith(DND_V2_MARKER);
        List<String> cols = doc.columns();

        int fromRow;
        int fromDateIdx;
        double max;

        if (payloadIsV2) {
            String body = rest.substring(DND_V2_MARKER.length());
            String[] p = body.split(":", 3);
            if (p.length < 3) {
                return false;
            }
            try {
                String gk =
                        new String(Base64.getUrlDecoder().decode(p[0]), StandardCharsets.UTF_8);
                fromDateIdx = Integer.parseInt(p[1]);
                max = Double.parseDouble(p[2]);
                fromRow = indexOfProfileGroupKey(gk);
            } catch (IllegalArgumentException e) {
                return false;
            }
            if (fromRow < 0) {
                return false;
            }
        } else {
            String[] p = rest.split(":");
            if (p.length < 3) {
                return false;
            }
            try {
                fromRow = Integer.parseInt(p[0]);
                fromDateIdx = Integer.parseInt(p[1]);
                max = Double.parseDouble(p[2]);
            } catch (NumberFormatException e) {
                return false;
            }
        }

        int toIdx = wideProfileIndexForRow(targetRow);

        if (fromRow != toIdx) {
            statusLabel.setText(
                    "縦方向への移動はできません（横のみ）");
            return false;
        }
        if (fromRow == toIdx && fromDateIdx == targetDateIdx) {
            return false;
        }
        Optional<Double> moved = pickMoveQuantity(shell != null ? shell.getPrimaryStage() : null, max);
        if (moved.isEmpty()) {
            return false;
        }
        double amt = moved.get();
        if (amt <= 1e-9 || amt > max + 1e-9) {
            return false;
        }
        if (fromRow < 0 || fromRow >= wideProfiles.size() || toIdx < 0 || toIdx >= wideProfiles.size()) {
            return false;
        }
        Map<String, String> fromProfile = wideProfiles.get(fromRow);
        Map<String, String> toProfile = wideProfiles.get(toIdx);
        LocalDate fromDay = dateAxis.get(fromDateIdx);
        LocalDate toDay = dateAxis.get(targetDateIdx);

        double fromSum =
                ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                        doc.rows(),
                        fromProfile,
                        fromDay,
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        double toSum =
                ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                        doc.rows(),
                        toProfile,
                        toDay,
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        ResultDispatchPivot.upsertAllocationForWideMerge(
                cols,
                doc.rows(),
                fromProfile,
                fromDay,
                fromSum - amt,
                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        ResultDispatchPivot.upsertAllocationForWideMerge(
                cols,
                doc.rows(),
                toProfile,
                toDay,
                toSum + amt,
                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        ResultDispatchNormalizer.normalizeInPlace(cols, doc.rows());
        statusLabel.setText("moved");
        rebuildGrids();
        markDispatchDocDirty();
        return true;
    }

    private static Optional<Double> pickMoveQuantity(Window owner, double max) {
        TextInputDialog dialog = new TextInputDialog(ResultDispatchNormalizer.formatQty(max));
        dialog.initOwner(owner);
        dialog.setTitle("Quantity");
        dialog.setHeaderText("Amount to move (max " + ResultDispatchNormalizer.formatQty(max) + ")");
        return dialog.showAndWait()
                .map(ResultDispatchNormalizer::parseDouble)
                .filter(v -> v > 0);
    }

    /** Mutable wide row (amounts indexed by {@link #dateAxis}). */
    public static final class WideRow {
        private final Map<String, String> staticPart;
        private final double[] amounts;

        WideRow(Map<String, String> staticPart, int nDates) {
            this.staticPart = new LinkedHashMap<>(staticPart);
            this.amounts = new double[nDates];
        }

        String getStatic(String col) {
            return staticPart.getOrDefault(col, "");
        }

        Map<String, String> profileMap() {
            return new LinkedHashMap<>(staticPart);
        }

        double getAmount(int di) {
            return amounts[di];
        }

        void setAmount(int di, double v) {
            amounts[di] = v;
        }
    }

    public record ByDayRow(String process, String machine, String processingContent, double[] amounts) {
        ByDayRow(String process, String machine, String processingContent, int n) {
            this(process, machine, processingContent, new double[n]);
        }

        double getAmount(int i) {
            return amounts[i];
        }

        void setAmount(int i, double v) {
            amounts[i] = v;
        }
    }
}
