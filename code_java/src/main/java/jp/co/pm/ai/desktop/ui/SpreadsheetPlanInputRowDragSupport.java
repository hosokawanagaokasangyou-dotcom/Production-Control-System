package jp.co.pm.ai.desktop.ui;

import java.util.BitSet;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.collections.ObservableList;
import javafx.scene.Node;
import javafx.scene.control.TableCell;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.DataFormat;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.MouseEvent;
import javafx.scene.input.TransferMode;

import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * Plan-input {@link SpreadsheetView} row reorder via drag-and-drop on data rows.
 *
 * <p>ControlsFX delivers {@code DRAG_DETECTED} to {@link TableCell}, not {@code TableRow}, so drag starts with
 * {@link TableCell#startDragAndDrop}. Hover/drop hit {@link javafx.scene.control.Cell} skins ({@code CellView}), so
 * {@code DRAG_OVER} / {@code DRAG_DROPPED} are handled on {@link SpreadsheetView} via event filters (same pattern as
 * dispatch interactive wide grid).
 */
public final class SpreadsheetPlanInputRowDragSupport {

    private static final DataFormat MODEL_ROW =
            new DataFormat("application/x-pm-ai-plan-input-model-row");

    /** While true, {@link SpreadsheetTabularSupport#installFullRowDataSelection} should not expand selection. */
    private static final AtomicBoolean PLAN_INPUT_ROW_DRAG_ACTIVE = new AtomicBoolean(false);

    private SpreadsheetPlanInputRowDragSupport() {}

    /** Predicate for {@link SpreadsheetTabularSupport#installFullRowDataSelection(SpreadsheetView, java.util.function.BooleanSupplier)}. */
    public static boolean skipFullRowExpansionDuringPlanInputRowDrag() {
        return PLAN_INPUT_ROW_DRAG_ACTIVE.get();
    }

    /**
     * Installs spreadsheet-level filters for drag source (cell), drag-over acceptance, and drop.
     *
     * @param firstDataGridRow {@link SpreadsheetTabularSupport#spreadsheetFirstDataRowIndex()} (first data grid row)
     * @param afterReorder callback after model list + grid rebuild
     */
    public static void install(
            SpreadsheetView spreadsheetView,
            int firstDataGridRow,
            ObservableList<ObservableList<String>> dataRows,
            Runnable afterReorder) {
        spreadsheetView.addEventFilter(DragEvent.DRAG_DONE, e -> PLAN_INPUT_ROW_DRAG_ACTIVE.set(false));

        spreadsheetView.addEventFilter(
                MouseEvent.DRAG_DETECTED,
                e -> {
                    Node n =
                            e.getPickResult() != null
                                    ? e.getPickResult().getIntersectedNode()
                                    : null;
                    TableCell<?, ?> tc = findTableCell(n);
                    if (tc == null || !isUnderSpreadsheet(spreadsheetView, tc)) {
                        return;
                    }
                    int viewRow = tc.getIndex();
                    if (viewRow < 0 || tc.isEmpty()) {
                        return;
                    }
                    int modelRow = spreadsheetView.getModelRow(viewRow);
                    if (!isDataRowAcceptable(spreadsheetView, firstDataGridRow, modelRow)) {
                        return;
                    }
                    PLAN_INPUT_ROW_DRAG_ACTIVE.set(true);
                    try {
                        Dragboard db = tc.startDragAndDrop(TransferMode.MOVE);
                        ClipboardContent cc = new ClipboardContent();
                        cc.put(MODEL_ROW, modelRow);
                        db.setContent(cc);
                        SpreadsheetRowReorderDragGhost.apply(db, tc, e);
                        e.consume();
                    } catch (RuntimeException ex) {
                        PLAN_INPUT_ROW_DRAG_ACTIVE.set(false);
                        throw ex;
                    }
                });

        spreadsheetView.addEventFilter(
                DragEvent.DRAG_OVER,
                e -> {
                    Dragboard dragboard = e.getDragboard();
                    if (!dragboard.hasContent(MODEL_ROW)) {
                        return;
                    }
                    Node n =
                            e.getPickResult() != null
                                    ? e.getPickResult().getIntersectedNode()
                                    : null;
                    TableCell<?, ?> tc = findTableCell(n);
                    if (tc == null || !isUnderSpreadsheet(spreadsheetView, tc)) {
                        return;
                    }
                    int viewRow = tc.getIndex();
                    if (viewRow < 0 || tc.isEmpty()) {
                        return;
                    }
                    int tgtModel = spreadsheetView.getModelRow(viewRow);
                    if (!isDataRowAcceptable(spreadsheetView, firstDataGridRow, tgtModel)) {
                        return;
                    }
                    e.acceptTransferModes(TransferMode.MOVE);
                    e.consume();
                });

        spreadsheetView.addEventFilter(
                DragEvent.DRAG_DROPPED,
                e -> {
                    Dragboard dragboard = e.getDragboard();
                    boolean success = false;
                    if (!dragboard.hasContent(MODEL_ROW)) {
                        return;
                    }
                    Node n =
                            e.getPickResult() != null
                                    ? e.getPickResult().getIntersectedNode()
                                    : null;
                    TableCell<?, ?> tc = findTableCell(n);
                    if (tc == null || !isUnderSpreadsheet(spreadsheetView, tc)) {
                        e.setDropCompleted(false);
                        e.consume();
                        return;
                    }
                    int viewRow = tc.getIndex();
                    if (viewRow < 0 || tc.isEmpty()) {
                        e.setDropCompleted(false);
                        e.consume();
                        return;
                    }
                    Object payload = dragboard.getContent(MODEL_ROW);
                    Integer srcModel = payload instanceof Integer ? (Integer) payload : null;
                    int tgtModel = spreadsheetView.getModelRow(viewRow);
                    if (srcModel != null
                            && isDataRowAcceptable(spreadsheetView, firstDataGridRow, srcModel)
                            && isDataRowAcceptable(spreadsheetView, firstDataGridRow, tgtModel)) {
                        int srcData = srcModel - firstDataGridRow;
                        int tgtData = tgtModel - firstDataGridRow;
                        if (srcData >= 0
                                && tgtData >= 0
                                && srcData < dataRows.size()
                                && tgtData < dataRows.size()) {
                            if (srcData != tgtData) {
                                moveDataRow(dataRows, srcData, tgtData);
                                spreadsheetView.setComparator(null);
                                ExcelLikeSpreadsheetFilter.resetAllColumnSortMenus(spreadsheetView);
                                afterReorder.run();
                            }
                            success = true;
                        }
                    }
                    e.setDropCompleted(success);
                    e.consume();
                });
    }

    /** Resolves the hosting {@link TableCell} for spreadsheet hit-testing (plan-input edit dialog, row drag, etc.). */
    public static TableCell<?, ?> findTableCell(Node n) {
        while (n != null) {
            if (n instanceof TableCell<?, ?> tc) {
                return tc;
            }
            n = n.getParent();
        }
        return null;
    }

    public static boolean isUnderSpreadsheet(SpreadsheetView spv, Node node) {
        for (Node x = node; x != null; x = x.getParent()) {
            if (x == spv) {
                return true;
            }
        }
        return false;
    }

    private static boolean isHidden(SpreadsheetView spv, int modelRow) {
        BitSet h = spv.getHiddenRows();
        return modelRow >= 0 && modelRow < h.length() && h.get(modelRow);
    }

    private static boolean isDataRowAcceptable(
            SpreadsheetView spv, int firstDataGridRow, int modelRow) {
        return modelRow >= firstDataGridRow && !isHidden(spv, modelRow);
    }

    private static void moveDataRow(ObservableList<ObservableList<String>> rows, int from, int to) {
        if (from == to || from < 0 || to < 0 || from >= rows.size() || to >= rows.size()) {
            return;
        }
        ObservableList<String> moved = rows.remove(from);
        rows.add(to, moved);
    }
}
