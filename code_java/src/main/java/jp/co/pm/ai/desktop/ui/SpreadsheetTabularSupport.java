package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.BitSet;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.BooleanSupplier;
import java.util.function.IntSupplier;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.Label;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TableView;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.util.Duration;

import org.controlsfx.control.spreadsheet.Grid;
import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/**
 * Bridges tabular {@link ObservableList} rows to ControlsFX {@link SpreadsheetView} / {@link GridBase}.
 */
public final class SpreadsheetTabularSupport {

    /**
     * {@link #installSpreadsheetChromeRelayoutDebouncerForHost(StackPane, IntSupplier)} を同一 StackPane に二重適用しないためのマーカー。
     */
    private static final String SPREADSHEET_HOST_RELAYOUT_HOOK =
            "jp.co.pm.ai.desktop.spreadsheetHostRelayoutHook";

    /** 納期管理: フィルタ行（列見出しと同じ薄いグレー） */
    private static final String DC_STYLE_HEADER_ROW =
            "-fx-background-color: #e8e8e8; -fx-text-fill: black;";

    /** \u5148\u982d\u306e\u898b\u51fa\u3057\u5217\uff08\u5c5e\u6027\u5217\uff09: \u767d \uff0b \u9ed2 */
    private static final String DC_STYLE_LEADING_COL =
            "-fx-background-color: #ffffff; -fx-text-fill: black;";

    private static final String DC_STYLE_DATA_WHITE =
            "-fx-background-color: #ffffff; -fx-text-fill: black;";

    private static final String DC_STYLE_DATA_GREEN =
            "-fx-background-color: #d4edd4; -fx-text-fill: black;";

    /** 配台シミュレータ実効ロール単位（m）を Excel 側と揃えた強調。Python 着色の {@code FFFF00} 相当。 */
    private static final String DC_STYLE_DATA_YELLOW =
            "-fx-background-color: #ffff00; -fx-text-fill: black;";

    /** {@link #installPmAiReadableSpreadsheetChrome}／グリッド構築用に外部へ公開する固定配色（テーマ非依存）。 */
    public static final String READABLE_STYLE_FILTER_ROW = DC_STYLE_HEADER_ROW;

    public static final String READABLE_STYLE_LEADING_COL = DC_STYLE_LEADING_COL;
    public static final String READABLE_STYLE_DATA_WHITE = DC_STYLE_DATA_WHITE;
    public static final String READABLE_STYLE_DATA_GREEN = DC_STYLE_DATA_GREEN;

    public static final String READABLE_STYLE_DATA_YELLOW = DC_STYLE_DATA_YELLOW;

    /**
     * Date-column triple: lines shown as {@code (prefix)(qty)}. Lines with no value, dash placeholder, or
     * numeric zero are omitted so non-empty cells stand out when scanning.
     */
    private static final String DC_TRIPLE_PREFIX_PLAN = "(\u30a2\u30e9\u8a08\u753b)";

    private static final String DC_TRIPLE_PREFIX_ACTUAL = "(\u5b9f\u7e3e)";

    private static final String DC_TRIPLE_PREFIX_DISPATCH = "(\u30b7\u30b9\u914d\u53f0)";

    /** Grid row index reserved for ControlsFX column filters ({@link SpreadsheetView#setFilteredRow}). */
    public static final int SPREADSHEET_FILTER_ROW = 0;

    /** First grid row index that maps to {@link ObservableList} data row 0. */
    public static int spreadsheetFirstDataRowIndex() {
        return SPREADSHEET_FILTER_ROW + 1;
    }

    private SpreadsheetTabularSupport() {}

    /**
     * 固定配色（薄グレー見出し・白地・黒字）を SpreadsheetView に適用する。{@link SpreadsheetThemeBridge#install}
     * の直後に呼ぶとテーマ由来のセル色を上書きする。
     */
    public static void installPmAiReadableSpreadsheetChrome(SpreadsheetView view) {
        Objects.requireNonNull(view, "view");
        if (!view.getStyleClass().contains("pm-ai-readable-spreadsheet")) {
            view.getStyleClass().add("pm-ai-readable-spreadsheet");
        }
        String url =
                Objects.requireNonNull(
                                SpreadsheetTabularSupport.class.getResource(
                                        "/jp/co/pm/ai/desktop/css/delivery-calendar-spreadsheet.css"),
                                "delivery-calendar-spreadsheet.css")
                        .toExternalForm();
        if (!view.getStylesheets().contains(url)) {
            view.getStylesheets().add(url);
        }
    }

    /** @deprecated {@link #installPmAiReadableSpreadsheetChrome(SpreadsheetView)} を使用。 */
    @Deprecated
    public static void installDeliveryCalendarSpreadsheetChrome(SpreadsheetView view) {
        installPmAiReadableSpreadsheetChrome(view);
    }

    /**
     * Enables ControlsFX per-column filters and row sorting (filter menu on column headers). Requires a filter
     * placeholder row in the {@link GridBase} at {@link #SPREADSHEET_FILTER_ROW}.
     */
    public static void applyColumnFilters(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        SpreadsheetMultiColumnFilterCoordinator.clear(view);
        view.setFilteredRow(SPREADSHEET_FILTER_ROW);
        ObservableList<SpreadsheetColumn> cols = view.getColumns();
        for (int i = 0; i < cols.size(); i++) {
            cols.get(i).setFilter(new ExcelLikeSpreadsheetFilter(view, i));
        }
    }

    /**
     * Same as {@link #applyColumnFilters(SpreadsheetView)} but opens a modal dialog per column (適用 / OK /
     * キャンセル) via {@link DialogExcelLikeSpreadsheetFilter}.
     */
    public static void applyColumnFiltersWithDialog(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        SpreadsheetMultiColumnFilterCoordinator.clear(view);
        view.setFilteredRow(SPREADSHEET_FILTER_ROW);
        ObservableList<SpreadsheetColumn> cols = view.getColumns();
        for (int i = 0; i < cols.size(); i++) {
            cols.get(i).setFilter(new DialogExcelLikeSpreadsheetFilter(view, i));
        }
    }

    /**
     * Clears spreadsheet-wide row sort and per-column row filters (hidden rows), and resets filter menu sort labels.
     */
    public static void clearAllFiltersAndSort(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        SpreadsheetMultiColumnFilterCoordinator.clear(view);
        view.setComparator(null);
        view.setHiddenRows(new BitSet());
        ExcelLikeSpreadsheetFilter.resetAllColumnSortMenus(view);
        DialogExcelLikeSpreadsheetFilter.resetAllColumnSortMenus(view);
    }

    /**
     * 現在の非表示行（列フィルタによる行の非表示）を複製する。{@link SpreadsheetView#setGrid} 後に {@link
     * #restoreHiddenRows} で戻す用途。
     */
    public static BitSet snapshotHiddenRows(SpreadsheetView view) {
        if (view == null) {
            return new BitSet();
        }
        return (BitSet) view.getHiddenRows().clone();
    }

    /**
     * {@link #snapshotHiddenRows} の結果を適用する。グリッド行数より大きいインデックスは無視し、行数が減った再構築でも破綻しないようにする。
     */
    public static void restoreHiddenRows(SpreadsheetView view, BitSet snapshot) {
        if (view == null || snapshot == null) {
            return;
        }
        Grid grid = view.getGrid();
        int rowCount = grid != null && grid.getRows() != null ? grid.getRows().size() : 0;
        if (rowCount <= 0) {
            view.setHiddenRows(new BitSet());
            return;
        }
        BitSet clamped = new BitSet();
        for (int i = snapshot.nextSetBit(0); i >= 0; i = snapshot.nextSetBit(i + 1)) {
            if (i < rowCount) {
                clamped.set(i);
            }
        }
        view.setHiddenRows(clamped);
    }

    /** Applies leading column freeze from persisted 見出し列数 (must run after {@link SpreadsheetView#setGrid}). */
    public static void applyFixedLeadingColumns(SpreadsheetView view, int headerColumnCount) {
        if (view == null) {
            return;
        }
        view.getFixedColumns().clear();
        int n = Math.max(0, headerColumnCount);
        if (n <= 0) {
            return;
        }
        ObservableList<SpreadsheetColumn> cols = view.getColumns();
        int limit = Math.min(n, cols.size());
        for (int i = 0; i < limit; i++) {
            SpreadsheetColumn col = cols.get(i);
            if (col.isColumnFixable()) {
                view.getFixedColumns().add(col);
            }
        }
    }

    public static void applyFixedLeadingColumnsLater(SpreadsheetView view, int headerColumnCount) {
        Platform.runLater(() -> applyFixedLeadingColumns(view, headerColumnCount));
    }

    /**
     * 列フィルタ行（グリッド行 {@link #SPREADSHEET_FILTER_ROW}）を固定行に入れ、縦スクロールしても常に表示する。
     * {@link #applyColumnFilters} / {@link #applyColumnFiltersWithDialog} と {@link SpreadsheetView#setFilteredRow}
     * の後に呼ぶこと。
     */
    public static void pinSpreadsheetFilterRow(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        int r = SPREADSHEET_FILTER_ROW;
        if (!view.isRowFixable(r)) {
            return;
        }
        ObservableList<Integer> fixed = view.getFixedRows();
        if (!fixed.contains(Integer.valueOf(r))) {
            fixed.add(0, r);
        }
    }

    /**
     * 見出し列固定・フィルタ行ピン・UNCONSTRAINED 列幅をまとめて適用する。
     * TitledPane／アコーディオン開閉後など、レイアウトでスキンが組み替わったあとに再度呼ぶこと。
     */
    public static void reapplySpreadsheetColumnChrome(SpreadsheetView view, int headerColumnCount) {
        if (view == null) {
            return;
        }
        applyFixedLeadingColumns(view, headerColumnCount);
        pinSpreadsheetFilterRow(view);
        applyUnconstrainedColumnResizePolicy(view);
    }

    /**
     * StackPane ホストのレイアウトが変わったあと（デバウンス）に {@link #reapplySpreadsheetColumnChrome} を実行する。
     * ホストの先頭子が {@link SpreadsheetView} のときのみ適用する（子が差し替わっても最新の先頭子を参照する）。
     */
    public static void installSpreadsheetChromeRelayoutDebouncerForHost(
            StackPane host, IntSupplier headerColumnCountSupplier) {
        Objects.requireNonNull(host, "host");
        if (Boolean.TRUE.equals(host.getProperties().get(SPREADSHEET_HOST_RELAYOUT_HOOK))) {
            return;
        }
        host.getProperties().put(SPREADSHEET_HOST_RELAYOUT_HOOK, Boolean.TRUE);

        PauseTransition debounce = new PauseTransition(Duration.millis(120));
        Runnable flush =
                () ->
                        Platform.runLater(
                                () -> {
                                    if (host.getChildren().isEmpty()) {
                                        return;
                                    }
                                    Node ch = host.getChildren().get(0);
                                    if (!(ch instanceof SpreadsheetView sv)) {
                                        return;
                                    }
                                    int hc =
                                            headerColumnCountSupplier != null
                                                    ? Math.max(0, headerColumnCountSupplier.getAsInt())
                                                    : 0;
                                    reapplySpreadsheetColumnChrome(sv, hc);
                                });
        debounce.setOnFinished(e -> flush.run());

        host.layoutBoundsProperty()
                .addListener(
                        (obs, o, n) -> {
                            debounce.stop();
                            debounce.playFromStart();
                        });
    }

    /**
     * ControlsFX {@link SpreadsheetView} の内側 {@link TableView}（行＝グリッドの1行）を返す。
     */
    @SuppressWarnings("unchecked")
    public static TableView<ObservableList<SpreadsheetCell>> findInnerTableView(SpreadsheetView view) {
        if (view == null) {
            return null;
        }
        for (Node n : view.getChildrenUnmodifiable()) {
            if (n instanceof TableView<?> tv) {
                return (TableView<ObservableList<SpreadsheetCell>>) tv;
            }
        }
        return null;
    }

    /**
     * 見出し列の固定（{@link #applyFixedLeadingColumns}）や列フィルタ適用後にスキンが分割されても、内側の {@link TableView}
     * すべてに {@link TableView#UNCONSTRAINED_RESIZE_POLICY} を再度適用する必要がある。
     */
    public static void applyUnconstrainedColumnResizePolicy(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        /*
         * SpreadsheetView 配下のスキンに複数の TableView が現れる場合や、列可視性変更後にポリシーが
         * CONSTRAINED に戻る場合がある。サブツリー内の TableView すべてに UNCONSTRAINED を適用する。
         */
        setUnconstrainedOnEmbeddedTableViews(view, 0);
        for (SpreadsheetColumn col : view.getColumns()) {
            col.setResizable(true);
        }
    }

    private static void setUnconstrainedOnEmbeddedTableViews(Node n, int depth) {
        if (n == null || depth > 12) {
            return;
        }
        if (n instanceof TableView<?> tv) {
            tv.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        }
        if (n instanceof Parent p) {
            for (Node c : p.getChildrenUnmodifiable()) {
                setUnconstrainedOnEmbeddedTableViews(c, depth + 1);
            }
        }
    }

    /**
     * After any cell selection change, expands selection to full grid rows (all columns) for each data row that
     * had at least one selected cell, so the active row is visually continuous. Skips the filter row at
     * {@link #SPREADSHEET_FILTER_ROW}.
     *
     * <p>Requires {@link SpreadsheetView#getSelectionModel()}{@code .setSelectionMode(SelectionMode.MULTIPLE)}.
     *
     * @param skipFullRowExpansion when {@code true}, skips expanding selection (e.g. during plan-input row drag).
     */
    public static void installFullRowDataSelection(
            SpreadsheetView view, BooleanSupplier skipFullRowExpansion) {
        if (view == null) {
            return;
        }
        final boolean[] guard = {false};
        var sm = view.getSelectionModel();
        // SpreadsheetView exposes ObservableList<TablePosition> (raw); listener must accept TablePosition.
        sm.getSelectedCells()
                .addListener(
                        (ListChangeListener<? super TablePosition>)
                                change -> {
                                    if (guard[0]) {
                                        return;
                                    }
                                    if (skipFullRowExpansion != null && skipFullRowExpansion.getAsBoolean()) {
                                        return;
                                    }
                                    ObservableList<SpreadsheetColumn> cols = view.getColumns();
                                    if (cols.isEmpty()) {
                                        return;
                                    }
                                    ObservableList<? extends TablePosition> selected =
                                            sm.getSelectedCells();
                                    if (selected.isEmpty()) {
                                        return;
                                    }
                                    int firstData = spreadsheetFirstDataRowIndex();
                                    Set<Integer> rows = new HashSet<>();
                                    for (TablePosition p : selected) {
                                        int r = p.getRow();
                                        if (r >= firstData) {
                                            rows.add(r);
                                        }
                                    }
                                    if (rows.isEmpty()) {
                                        return;
                                    }
                                    int colCount = cols.size();
                                    boolean allFull = true;
                                    for (int r : rows) {
                                        int cnt = 0;
                                        for (TablePosition p : selected) {
                                            if (p.getRow() == r) {
                                                cnt++;
                                            }
                                        }
                                        if (cnt < colCount) {
                                            allFull = false;
                                            break;
                                        }
                                    }
                                    if (allFull) {
                                        return;
                                    }
                                    TablePosition focus = sm.getFocusedCell();
                                    int focusRow;
                                    if (focus != null && focus.getRow() >= firstData) {
                                        focusRow = focus.getRow();
                                    } else {
                                        focusRow = Collections.min(rows);
                                    }
                                    int focusCol =
                                            focus != null && focus.getColumn() >= 0 ? focus.getColumn() : 0;

                                    guard[0] = true;
                                    try {
                                        sm.clearSelection();
                                        SpreadsheetColumn firstCol = cols.get(0);
                                        SpreadsheetColumn lastCol = cols.get(cols.size() - 1);
                                        ArrayList<Integer> sorted = new ArrayList<>(rows);
                                        Collections.sort(sorted);
                                        for (int r : sorted) {
                                            sm.selectRange(r, firstCol, r, lastCol);
                                        }
                                        int fc = Math.min(Math.max(focusCol, 0), cols.size() - 1);
                                        sm.focus(focusRow, cols.get(fc));
                                    } finally {
                                        guard[0] = false;
                                    }
                                });
    }

    /** @see #installFullRowDataSelection(SpreadsheetView, BooleanSupplier) */
    public static void installFullRowDataSelection(SpreadsheetView view) {
        installFullRowDataSelection(view, null);
    }

    private static final String PLAN_INPUT_COL_UNPROCESSED = "未加工";
    private static final String PLAN_INPUT_COL_ROLL_UNIT_M = "ロール単位長さ";
    private static final String PLAN_INPUT_COL_QTY_CONV = "換算数量";
    private static final String PLAN_INPUT_COL_ACTUAL = "実加工数";
    private static final String PLAN_INPUT_COL_PRODUCT = "製品名";
    private static final String PLAN_INPUT_COL_USED_RAW = "使用原反";

    /**
     * Python {@code planning_core._core._effective_roll_unit_m_for_dispatch_task_simulator} に相当。
     *
     * @param qtyM 配台に使う残量 (m) 相当（本 UI では行セルから近似）
     * @param sheetRollUnitM シート上の製品側ロール単位 (m)
     */
    public static double effectiveRollUnitMForDispatchTaskSimulator(double qtyM, double sheetRollUnitM) {
        double q = qtyM;
        double u = sheetRollUnitM;
        if (q <= 1e-12 || u <= 1e-12) {
            return u > 1e-12 ? u : 0.0;
        }
        double nRaw = q / u;
        if (nRaw <= 1e-12) {
            return u;
        }
        if (Math.abs(nRaw - Math.rint(nRaw)) <= 1e-9) {
            return u;
        }
        int nWork = (int) Math.floor(nRaw);
        if (nWork < 1) {
            nWork = 1;
        }
        return q / (double) nWork;
    }

    private static String planInputCellAt(ObservableList<String> src, int col) {
        if (src == null || col < 0 || col >= src.size()) {
            return "";
        }
        String v = src.get(col);
        return v != null ? v : "";
    }

    /**
     * 段階2 {@code build_task_queue_from_planning_df} の {@code qty} に近い残量 (m)。完全な {@code
     * _plan_row_dispatch_qty_metrics} ではなく、実効ロール強調用の近似。
     */
    private static double planInputQtyMForDispatchSimulatorApprox(
            ObservableList<String> src, int idxConv, int idxUnp, int idxAct) {
        double conv =
                idxConv >= 0 ? Stage2RollUnitLengthTables.parseFloatSafe(planInputCellAt(src, idxConv), 0.0) : 0.0;
        double actual =
                idxAct >= 0 ? Stage2RollUnitLengthTables.parseFloatSafe(planInputCellAt(src, idxAct), 0.0) : 0.0;
        if (idxUnp < 0) {
            return Math.max(0.0, conv);
        }
        Optional<Double> unpOpt =
                Stage2RollUnitLengthTables.optionalUnprocessedCell(planInputCellAt(src, idxUnp));
        if (unpOpt.isEmpty()) {
            return Math.max(0.0, conv);
        }
        double unp = unpOpt.get();
        if (conv > 1e-12 && Math.abs(unp) <= 1e-12 && actual <= 1e-12) {
            return conv;
        }
        if (unp > 1e-12) {
            return unp;
        }
        return Math.max(0.0, conv);
    }

    private static double intrinsicProductRollSheetGuess(
            String product,
            String usedRaw,
            Stage2RollUnitLengthTables tablesOrNull,
            double qtyFallback) {
        if (tablesOrNull != null) {
            Optional<Double> byRaw = tablesOrNull.lookupByUsedRaw(usedRaw);
            if (byRaw.isPresent()) {
                return byRaw.get();
            }
            Optional<Double> byProd = tablesOrNull.lookupByProductName(product);
            if (byProd.isPresent()) {
                return byProd.get();
            }
        }
        double fb = qtyFallback > 1e-12 ? qtyFallback : 100.0;
        return Stage2RollUnitLengthTables.inferFromProductDimensions(product, fb);
    }

    private static boolean planInputRollUnitLengthCellIsYellowHighlight(
            ObservableList<String> src,
            int idxConv,
            int idxUnp,
            int idxAct,
            int idxProd,
            int idxUsed,
            String rawRoll,
            Stage2RollUnitLengthTables tablesOrNull) {
        double uDisp = Stage2RollUnitLengthTables.parseFloatSafe(rawRoll, 0.0);
        if (!(uDisp > 1e-12)) {
            return false;
        }
        double q = planInputQtyMForDispatchSimulatorApprox(src, idxConv, idxUnp, idxAct);
        if (!(q > 1e-12)) {
            return false;
        }
        double effFromDisplay = effectiveRollUnitMForDispatchTaskSimulator(q, uDisp);
        if (Math.abs(effFromDisplay - uDisp) > 1e-6) {
            return true;
        }
        String product = idxProd >= 0 ? planInputCellAt(src, idxProd) : "";
        String usedRaw = idxUsed >= 0 ? planInputCellAt(src, idxUsed) : "";
        double uSheet = intrinsicProductRollSheetGuess(product, usedRaw, tablesOrNull, q);
        if (!(uSheet > 1e-12) || Math.abs(uSheet - uDisp) <= 1e-6) {
            return false;
        }
        double effFromSheet = effectiveRollUnitMForDispatchTaskSimulator(q, uSheet);
        return Math.abs(effFromSheet - uDisp) <= 1e-6;
    }

    /**
     * 配台計画タスク入力の編集可能グリッド。先頭固定列は白地、それ以外は既定で白地とし、「未加工」列のみ正の数値で薄緑、
     * 「ロール単位長さ」は実効ロール化が想定されるセルを黄で示す（Excel 側の実効ロール着色と整合）。
     *
     * @param leadingColumnCount 先頭から固定する属性列の本数
     * @param rollUnitLengthTablesOrNull 製品名／使用原反テーブル（黄の「上書き後セル」判定に使用）。{@code null}
     *     のときは製品名寸法推定のみで近似する。
     * @see #buildPlanInputGrid(List, ObservableList, boolean, int)
     */
    public static GridBase buildPlanInputGrid(
            List<String> headersRef,
            ObservableList<ObservableList<String>> rows,
            boolean editable,
            int leadingColumnCount,
            Stage2RollUnitLengthTables rollUnitLengthTablesOrNull) {
        int cols = headersRef.size();
        int rc = rows.size();
        int gridRowsTotal = rc + 1;
        GridBase grid = new GridBase(gridRowsTotal, cols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(headersRef);

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < cols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            cell.setStyle(READABLE_STYLE_FILTER_ROW);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        int lead = Math.max(0, Math.min(leadingColumnCount, cols));
        int firstData = spreadsheetFirstDataRowIndex();
        int idxUnp = headersRef.indexOf(PLAN_INPUT_COL_UNPROCESSED);
        int idxConv = headersRef.indexOf(PLAN_INPUT_COL_QTY_CONV);
        int idxAct = headersRef.indexOf(PLAN_INPUT_COL_ACTUAL);
        int idxProd = headersRef.indexOf(PLAN_INPUT_COL_PRODUCT);
        int idxUsed = headersRef.indexOf(PLAN_INPUT_COL_USED_RAW);
        for (int r = 0; r < rc; r++) {
            int gridRow = firstData + r;
            ObservableList<String> src = rows.get(r);
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            for (int c = 0; c < cols; c++) {
                String raw = c < src.size() && src.get(c) != null ? src.get(c) : "";
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw);
                cell.setEditable(editable);
                String headerTitle = headersRef.get(c);
                if ("配台不要".equals(headerTitle)
                        && TabularCellHighlight.planInputExcludeFromAssignmentIsOn(raw)) {
                    cell.setStyle(TabularCellHighlight.PLAN_INPUT_EXCLUDE_YES_STYLE);
                } else if (c < lead) {
                    cell.setStyle(READABLE_STYLE_LEADING_COL);
                } else if (PLAN_INPUT_COL_UNPROCESSED.equals(headerTitle)
                        && TabularCellHighlight.isStrictPositiveNumber(raw)) {
                    cell.setStyle(READABLE_STYLE_DATA_GREEN);
                } else if (PLAN_INPUT_COL_ROLL_UNIT_M.equals(headerTitle)
                        && planInputRollUnitLengthCellIsYellowHighlight(
                                src,
                                idxConv,
                                idxUnp,
                                idxAct,
                                idxProd,
                                idxUsed,
                                raw,
                                rollUnitLengthTablesOrNull)) {
                    cell.setStyle(READABLE_STYLE_DATA_YELLOW);
                } else {
                    cell.setStyle(READABLE_STYLE_DATA_WHITE);
                }
                rowCells.add(cell);
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        return grid;
    }

    /**
     * @param rollUnitLengthTablesOrNull {@link #buildPlanInputGrid(List, ObservableList, boolean, int,
     *     Stage2RollUnitLengthTables)} 参照
     */
    public static GridBase buildPlanInputGrid(
            List<String> headersRef,
            ObservableList<ObservableList<String>> rows,
            boolean editable,
            int leadingColumnCount) {
        return buildPlanInputGrid(
                headersRef, rows, editable, leadingColumnCount, null);
    }

    /**
     * 段階1成形結果プレビュー。日付系見出しかつ正の数値セルは薄緑で強調（従来の Stage1 ルールを維持）。
     */
    public static GridBase buildStage1PreviewGrid(
            List<String> headersRef, ObservableList<ObservableList<String>> rows, int leadingColumnCount) {
        int cols = headersRef.size();
        int rc = rows.size();
        int gridRowsTotal = rc + 1;
        GridBase grid = new GridBase(gridRowsTotal, cols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(headersRef);

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < cols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            cell.setStyle(READABLE_STYLE_FILTER_ROW);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        int lead = Math.max(0, Math.min(leadingColumnCount, cols));
        int firstData = spreadsheetFirstDataRowIndex();
        for (int r = 0; r < rc; r++) {
            int gridRow = firstData + r;
            ObservableList<String> src = rows.get(r);
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            for (int c = 0; c < cols; c++) {
                String raw = c < src.size() && src.get(c) != null ? src.get(c) : "";
                String headerTitle = headersRef.get(c);
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw);
                cell.setEditable(false);
                if (c < lead) {
                    cell.setStyle(READABLE_STYLE_LEADING_COL);
                } else if (TabularCellHighlight.isStage1DateColumnHeader(headerTitle)
                        && TabularCellHighlight.isStrictPositiveNumber(raw)) {
                    cell.setStyle(READABLE_STYLE_DATA_GREEN);
                } else {
                    cell.setStyle(READABLE_STYLE_DATA_WHITE);
                }
                rowCells.add(cell);
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        return grid;
    }

    /**
     * Read-only string grid with a filter placeholder row; after {@link SpreadsheetView#setGrid} apply
     * {@link #applyColumnFiltersWithDialog} (計画結果ビュー等と同じダイアログ式) または {@link #applyColumnFilters}。行ハイライトなし。
     * Used for JSON-backed viewers such as {@code 結果_配台表.json}.
     */
    public static GridBase buildReadOnlyPlainGrid(
            List<String> headersRef, ObservableList<ObservableList<String>> rows) {
        return buildReadOnlyPlainGrid(headersRef, rows, 0);
    }

    /**
     * Same as {@link #buildReadOnlyPlainGrid(List, ObservableList)}; when {@code deliveryLeadingColumns >= 0},
     * applies fixed delivery-calendar chrome (ignore theme): header/filter row light gray, leading columns white,
     * other data cells white or light green when a non-negative number is present.
     */
    public static GridBase buildReadOnlyPlainGrid(
            List<String> headersRef,
            ObservableList<ObservableList<String>> rows,
            int deliveryLeadingColumns) {
        int cols = headersRef.size();
        int rc = rows.size();
        int gridRowsTotal = rc + 1;
        GridBase grid = new GridBase(gridRowsTotal, cols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(headersRef);

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < cols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            if (deliveryLeadingColumns >= 0) {
                cell.setStyle(DC_STYLE_HEADER_ROW);
            }
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        int firstData = spreadsheetFirstDataRowIndex();
        for (int r = 0; r < rc; r++) {
            int gridRow = firstData + r;
            ObservableList<String> src = rows.get(r);
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            for (int c = 0; c < cols; c++) {
                String raw = c < src.size() && src.get(c) != null ? src.get(c) : "";
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw);
                cell.setEditable(false);
                if (deliveryLeadingColumns >= 0) {
                    if (c < deliveryLeadingColumns) {
                        cell.setStyle(DC_STYLE_LEADING_COL);
                    } else {
                        cell.setStyle(deliveryCalendarDataStyleForDisplayText(raw));
                    }
                }
                rowCells.add(cell);
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        return grid;
    }

    /**
     * Read-only delivery-calendar main grid: date columns use a triple stack (task-input Aladdin / actual
     * detail / dispatch JSON); attribute columns ({@code leadingColumnCount} wide) are plain text.
     *
     * <p>配色（テーマ設定は無視・固定）:
     * <ul>
     *   <li>列見出し・行見出しは {@link #installPmAiReadableSpreadsheetChrome} 由来 CSS で薄グレー＋黒</li>
     *   <li>フィルタ行は {@link #DC_STYLE_HEADER_ROW}（薄グレー＋黒）</li>
     *   <li>先頭固定列（{@code c < leadingColumnCount}）は {@link #DC_STYLE_LEADING_COL}（白＋黒）</li>
     *   <li>日付列で空欄以外は {@link #DC_STYLE_DATA_GREEN}（薄緑＋黒）、空欄は {@link #DC_STYLE_DATA_WHITE}（白＋黒）</li>
     * </ul>
     *
     * @param leadingColumnCount number of left fixed columns (must be {@code >= 0} and {@code <= cols})
     */
    public static GridBase buildReadOnlyDeliveryCalendarMainGrid(
            List<String> headersRef,
            ObservableList<ObservableList<DeliveryCalendarMainCell>> rows,
            int leadingColumnCount) {
        int cols = headersRef.size();
        int rc = rows.size();
        int gridRowsTotal = rc + 1;
        GridBase grid = new GridBase(gridRowsTotal, cols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(headersRef);

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < cols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            cell.setStyle(DC_STYLE_HEADER_ROW);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        int firstData = spreadsheetFirstDataRowIndex();
        int leading = Math.max(0, Math.min(cols, leadingColumnCount));
        for (int r = 0; r < rc; r++) {
            int gridRow = firstData + r;
            ObservableList<DeliveryCalendarMainCell> src = rows.get(r);
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            for (int c = 0; c < cols; c++) {
                DeliveryCalendarMainCell mc =
                        c < src.size() && src.get(c) != null ? src.get(c) : new DeliveryCalendarMainCell.PlainText("");
                boolean isDateColumn = c >= leading;
                if (mc instanceof DeliveryCalendarMainCell.TripleQty t) {
                    List<String> visibleLines = deliveryCalendarTripleVisibleFormattedLines(t);
                    String item = String.join("\n", visibleLines);
                    SpreadsheetCell cell =
                            SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, item);
                    cell.setEditable(false);
                    cell.setWrapText(true);
                    cell.setCellGraphic(true);
                    Node g = deliveryCalendarTripleGraphic(t);
                    cell.setGraphic(g);
                    cell.setStyle(
                            isDateColumn
                                    ? (visibleLines.isEmpty() ? DC_STYLE_DATA_WHITE : DC_STYLE_DATA_GREEN)
                                    : DC_STYLE_LEADING_COL);
                    Tooltip tt =
                            new Tooltip(
                                    "\u30a2\u30e9\u30b8\u30f3\u52a0\u5de5\u8a08\u753b\u53d6\u5f97\u30c7\u30fc\u30bf / "
                                            + "\u52a0\u5de5\u5b9f\u7e3e / "
                                            + "\u914d\u53f0\u7d50\u679c\uff08\u7d50\u679c_\u914d\u53f0\u8868.json\uff09");
                    Tooltip.install(g, tt);
                    rowCells.add(cell);
                } else {
                    String raw =
                            mc instanceof DeliveryCalendarMainCell.PlainText pt ? pt.text() : "";
                    SpreadsheetCell cell =
                            SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw);
                    cell.setEditable(false);
                    cell.setCellGraphic(false);
                    cell.setGraphic(null);
                    cell.setStyle(
                            isDateColumn
                                    ? (raw == null || raw.isBlank() ? DC_STYLE_DATA_WHITE : DC_STYLE_DATA_GREEN)
                                    : DC_STYLE_LEADING_COL);
                    rowCells.add(cell);
                }
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        return grid;
    }

    private static String deliveryCalendarDataStyleForDisplayText(String raw) {
        return deliveryCalendarCellQualifiesGreen(raw) ? DC_STYLE_DATA_GREEN : DC_STYLE_DATA_WHITE;
    }

    /**
     * {@code 0} \u4ee5\u4e0a\u306e\u6570\u5024\u304c\u542b\u307e\u308c\u308b\u30bb\u30eb\u306f\u8584\u3044\u7dd1\uff08\u8907\u6570\u884c\u306f\u884c\u5358\u4f4d\u3067\u5224\u5b9a\uff09\u3002
     */
    private static boolean deliveryCalendarCellQualifiesGreen(String text) {
        if (text == null || text.isBlank()) {
            return false;
        }
        for (String line : text.split("\\R")) {
            String t = line.strip();
            if (t.isEmpty() || "\u2014".equals(t) || "-".equals(t)) {
                continue;
            }
            try {
                double v = Double.parseDouble(t.replace(",", ""));
                if (!Double.isNaN(v) && !Double.isInfinite(v) && v >= 0d) {
                    return true;
                }
            } catch (NumberFormatException ignored) {
                // next line
            }
        }
        return false;
    }

    private static Node deliveryCalendarTripleGraphic(DeliveryCalendarMainCell.TripleQty t) {
        VBox box = new VBox(1);
        box.setPadding(new Insets(2, 4, 2, 4));
        for (String line : deliveryCalendarTripleVisibleFormattedLines(t)) {
            box.getChildren().add(new Label(line));
        }
        return box;
    }

    /**
     * Formatted triple lines (prefix + quantity) for UI only; skips blank / dash / zero quantities.
     *
     * @see #formatDeliveryCalendarTripleLine(String, String) used for inspector APIs without filtering
     */
    private static List<String> deliveryCalendarTripleVisibleFormattedLines(
            DeliveryCalendarMainCell.TripleQty t) {
        List<String> lines = new ArrayList<>(3);
        if (!deliveryCalendarTripleQtyHidden(t.plan())) {
            lines.add(formatDeliveryCalendarTripleLine(DC_TRIPLE_PREFIX_PLAN, t.plan()));
        }
        if (!deliveryCalendarTripleQtyHidden(t.actual())) {
            lines.add(formatDeliveryCalendarTripleLine(DC_TRIPLE_PREFIX_ACTUAL, t.actual()));
        }
        if (!deliveryCalendarTripleQtyHidden(t.dispatch())) {
            lines.add(formatDeliveryCalendarTripleLine(DC_TRIPLE_PREFIX_DISPATCH, t.dispatch()));
        }
        return lines;
    }

    /**
     * {@code true} when the quantity string should not produce a visible line (null/blank, em dash / hyphen
     * alone, or parses as numeric zero).
     */
    private static boolean deliveryCalendarTripleQtyHidden(String qty) {
        if (qty == null || qty.isBlank()) {
            return true;
        }
        String s = qty.strip();
        if (s.isEmpty()) {
            return true;
        }
        if ("\u2014".equals(s) || "-".equals(s)) {
            return true;
        }
        try {
            double v = Double.parseDouble(s.replace(",", ""));
            if (!Double.isNaN(v) && !Double.isInfinite(v)) {
                return v == 0d;
            }
        } catch (NumberFormatException ignored) {
            return false;
        }
        return false;
    }

    /**
     * Same strings as rendered in date columns (for inspectors / logs). Quantity is the JSON triple field
     * string (already formatted).
     */
    public static String deliveryCalendarPlanLineForInspector(String qty) {
        return formatDeliveryCalendarTripleLine(DC_TRIPLE_PREFIX_PLAN, qty);
    }

    public static String deliveryCalendarActualLineForInspector(String qty) {
        return formatDeliveryCalendarTripleLine(DC_TRIPLE_PREFIX_ACTUAL, qty);
    }

    public static String deliveryCalendarDispatchLineForInspector(String qty) {
        return formatDeliveryCalendarTripleLine(DC_TRIPLE_PREFIX_DISPATCH, qty);
    }

    private static String formatDeliveryCalendarTripleLine(String prefix, String qty) {
        if (qty == null || qty.isBlank()) {
            return prefix + "\u2014";
        }
        return prefix + qty.strip();
    }

    /**
     * Read-only grid with timeline / Gantt-style cell coloring (see {@link GanttScheduleStyle}).
     *
     * @param kind 設備ガント Excel 風など表現の選択
     */
    public static GridBase buildReadOnlyGanttGrid(
            List<String> headersRef,
            ObservableList<ObservableList<String>> rows,
            GanttSheetKind kind) {
        int cols = headersRef.size();
        int rc = rows.size();
        int gridRowsTotal = rc + 1;
        GridBase grid = new GridBase(gridRowsTotal, cols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(headersRef);

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < cols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            cell.setStyle(READABLE_STYLE_FILTER_ROW);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        int firstData = spreadsheetFirstDataRowIndex();
        for (int r = 0; r < rc; r++) {
            int gridRow = firstData + r;
            ObservableList<String> src = rows.get(r);
            boolean sectionRow = false;
            if (!src.isEmpty() && src.get(0) != null) {
                String head = src.get(0);
                sectionRow = head.contains("■") || head.contains("▪");
            }
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            for (int c = 0; c < cols; c++) {
                String raw = c < src.size() && src.get(c) != null ? src.get(c) : "";
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw);
                cell.setEditable(false);
                GanttScheduleStyle.applyTimelineCell(
                        cell, c, headersRef.get(c), raw, sectionRow, r, kind);
                rowCells.add(cell);
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        return grid;
    }

    public static javafx.event.EventHandler<org.controlsfx.control.spreadsheet.GridChange> newRowsSyncHandler(
            ObservableList<ObservableList<String>> rows, List<String> headersRef, int firstDataGridRow) {
        return ev -> {
            int r = ev.getRow();
            int c = ev.getColumn();
            if (r < firstDataGridRow || c < 0 || c >= headersRef.size()) {
                return;
            }
            int dataIndex = r - firstDataGridRow;
            if (dataIndex < 0 || dataIndex >= rows.size()) {
                return;
            }
            Object nv = ev.getNewValue();
            String s = nv != null ? Objects.toString(nv, "") : "";
            ObservableList<String> row = rows.get(dataIndex);
            while (row.size() <= c) {
                row.add("");
            }
            row.set(c, s);
        };
    }

    public static void applyColumnWidths(SpreadsheetView view, List<Double> widths, double defaultWidth) {
        if (view == null || widths == null) {
            return;
        }
        var cols = view.getColumns();
        for (int i = 0; i < cols.size(); i++) {
            double w = i < widths.size() ? widths.get(i) : defaultWidth;
            cols.get(i).setPrefWidth(w);
        }
    }

    /** 計画結果スプレッドシートの行高さ％の下限（既定行高に対する倍率）。 */
    public static final double PLAN_RESULT_ROW_HEIGHT_PCT_MIN = 50.0;

    /** 計画結果スプレッドシートの行高さ％の上限（既定行高に対する倍率）。 */
    public static final double PLAN_RESULT_ROW_HEIGHT_PCT_MAX = 2000.0;

    /**
     * グリッド物理行数がこの値以上のとき、{@link #refreshSpreadsheetAfterRowPresentationChange} で {@link
     * SpreadsheetView#resizeRowsToDefault()} を呼ばない（大量行で Prism/SW のオフスクリーン確保がヒープを押し上げるため）。
     */
    public static final int PLAN_RESULT_REFRESH_SKIP_RESIZE_ROWS = 1200;

    /**
     * 旧実装では「この行数未満なら第 2 パルスを積む」に使った。post-fix-v3 で単発 flush に統一したためコードからは参照しないが、
     * 調整履歴として公開のまま残す。
     */
    public static final int PLAN_RESULT_REFRESH_SINGLE_FLUSH_ROWS = 400;

    /**
     * 列フィルタ行（グリッド行 {@link #SPREADSHEET_FILTER_ROW}）の高さ上限（px）。データ行の行高スライダー倍率に
     * 引きずられて巨大化しないようにする。
     */
    private static final double PLAN_RESULT_FILTER_ROW_MAX_HEIGHT_PX = 34.0;

    /**
     * 列フィルタ行付き {@link GridBase} 共通：データ行の高さをスケールし、フィルタ行は上限 px で抑える。データ行のセルに折り返しを適用する。
     *
     * @param cellWrapText {@code true} で折り返し、{@code false} で単行（見切れ）
     * @param rowHeightPercent {@link #PLAN_RESULT_ROW_HEIGHT_PCT_MIN}〜{@link #PLAN_RESULT_ROW_HEIGHT_PCT_MAX}
     *         （ControlsFX 既定行高に対する倍率％、100＝既定）
     */
    public static void applySpreadsheetGridRowHeightsAndWrap(
            GridBase grid, boolean cellWrapText, double rowHeightPercent) {
        applySpreadsheetGridRowHeightsAndWrap(grid, cellWrapText, rowHeightPercent, 24.0, 0.0);
    }

    /**
     * 同上。{@code baseDataRowHeightPx} は 100% 時のデータ行基準高さ、{@code minDataRowHeightPx} {@code > 0} でデータ行の下限 px。
     */
    public static void applySpreadsheetGridRowHeightsAndWrap(
            GridBase grid,
            boolean cellWrapText,
            double rowHeightPercent,
            double baseDataRowHeightPx,
            double minDataRowHeightPx) {
        applySpreadsheetGridRowHeightsAndWrap(
                grid,
                cellWrapText,
                rowHeightPercent,
                baseDataRowHeightPx,
                minDataRowHeightPx,
                PLAN_RESULT_ROW_HEIGHT_PCT_MIN,
                PLAN_RESULT_ROW_HEIGHT_PCT_MAX);
    }

    /**
     * \u540c\u4e0a\u3002{@code pctMin}/{@code pctMax} \u3067\u300c\u8a08\u753b\u7d50\u679c\u300d\u3068\u300c\u7d0d\u671f\u30ab\u30ec\u30f3\u30c0\u30fc\u300d\u3067\u7570\u306a\u308b\u30b9\u30e9\u30a4\u30c0\u30fc\u7bc4\u56f2\u3092\u6307\u5b9a\u3059\u308b\u3002
     */
    public static void applySpreadsheetGridRowHeightsAndWrap(
            GridBase grid,
            boolean cellWrapText,
            double rowHeightPercent,
            double baseDataRowHeightPx,
            double minDataRowHeightPx,
            double pctMin,
            double pctMax) {
        if (grid == null) {
            return;
        }
        double lo = pctMin > 0 ? pctMin : PLAN_RESULT_ROW_HEIGHT_PCT_MIN;
        double hi = pctMax > lo ? pctMax : PLAN_RESULT_ROW_HEIGHT_PCT_MAX;
        double pct = rowHeightPercent;
        if (Double.isNaN(pct) || pct <= 0) {
            pct = 100.0;
        }
        pct = Math.min(hi, Math.max(lo, pct));
        final double basePx = baseDataRowHeightPx > 0 ? baseDataRowHeightPx : 24.0;
        final double scaled = basePx * (pct / 100.0);
        final double rowPx =
                minDataRowHeightPx > 0 ? Math.max(minDataRowHeightPx, scaled) : scaled;
        final double filterRowPx = Math.min(rowPx, PLAN_RESULT_FILTER_ROW_MAX_HEIGHT_PX);
        grid.setRowHeightCallback(
                row -> row == SPREADSHEET_FILTER_ROW ? filterRowPx : rowPx);
        List<ObservableList<SpreadsheetCell>> rows = grid.getRows();
        if (rows == null) {
            return;
        }
        int firstData = spreadsheetFirstDataRowIndex();
        for (int r = 0; r < rows.size(); r++) {
            ObservableList<SpreadsheetCell> line = rows.get(r);
            boolean wrap = cellWrapText && r >= firstData;
            for (SpreadsheetCell cell : line) {
                cell.setWrapText(wrap);
            }
        }
    }

    /**
     * {@link #applySpreadsheetGridRowHeightsAndWrap} のエイリアス（計画結果ビューア向け名称）。
     */
    public static void applyPlanResultGridPresentation(
            GridBase grid, boolean cellWrapText, double rowHeightPercent) {
        applySpreadsheetGridRowHeightsAndWrap(grid, cellWrapText, rowHeightPercent);
    }

    /**
     * {@link #applySpreadsheetGridRowHeightsAndWrap} 後、内側 {@link TableView} が古い行高を保持することがあるため、
     * グリッド由来の行高を再適用し、表示を更新する（スクロールしなくても反映されるようにする）。
     *
     * <p>行数が極端に多いグリッドでは {@link SpreadsheetView#resizeRowsToDefault()} や内側 {@link TableView#refresh()} の再帰が
     * ヒープを急増させるため、{@link #PLAN_RESULT_REFRESH_SKIP_RESIZE_ROWS} 以上のときは軽い経路に切り替える。
     * 小表向けの二重 {@link Platform#runLater} は NDJSON 上で追加の {@code resize_then_layout} 連打と相関したため廃止。
     */
    public static void refreshSpreadsheetAfterRowPresentationChange(SpreadsheetView view) {
        refreshSpreadsheetAfterRowPresentationChange(view, false);
    }

    /**
     * @param skipResizeRowsToDefault {@code true} のとき {@link SpreadsheetView#resizeRowsToDefault()} を呼ばない。
     *         {@link GridBase#setRowHeightCallback} で行高を決めているビュー（納期管理カレンダー等）では既定の再適用で高さが潰れるため。
     */
    public static void refreshSpreadsheetAfterRowPresentationChange(
            SpreadsheetView view, boolean skipResizeRowsToDefault) {
        if (view == null) {
            return;
        }
        /*
         * 呼び出し時点では Grid の行がまだ増えていないことがある（後から数千行になる）。
         * 閾値は flush 実行時に再評価する。1 パルスのみ（二重 runLater 無し）。
         */
        Platform.runLater(
                () ->
                        presentationFlushAfterRowPresentationChangeOnce(
                                view, skipResizeRowsToDefault));
    }

    private static int spreadsheetPhysicalRowCount(SpreadsheetView view) {
        if (view == null || !(view.getGrid() instanceof GridBase gb)) {
            return -1;
        }
        var rows = gb.getRows();
        return rows != null ? rows.size() : -1;
    }

    /**
     * 行高・折り返し変更後の 1 パルス分（実行時点の行数で {@link SpreadsheetView#resizeRowsToDefault()} を抑止する）。
     */
    private static void presentationFlushAfterRowPresentationChangeOnce(
            SpreadsheetView view, boolean skipResizeRowsToDefault) {
        if (view == null) {
            return;
        }
        int physicalRows = spreadsheetPhysicalRowCount(view);
        boolean skipResize = physicalRows >= PLAN_RESULT_REFRESH_SKIP_RESIZE_ROWS;
        if (!skipResize) {
            if (!skipResizeRowsToDefault) {
                view.resizeRowsToDefault();
            }
            refreshEmbeddedTableViewsRecursive(view, 0);
        }
        view.requestLayout();
    }

    private static void refreshEmbeddedTableViewsRecursive(Node n, int depth) {
        if (n == null || depth > 24) {
            return;
        }
        if (n instanceof TableView<?> tv) {
            tv.refresh();
            tv.requestLayout();
        }
        if (n instanceof Parent p) {
            for (Node c : p.getChildrenUnmodifiable()) {
                refreshEmbeddedTableViewsRecursive(c, depth + 1);
            }
        }
    }
}
