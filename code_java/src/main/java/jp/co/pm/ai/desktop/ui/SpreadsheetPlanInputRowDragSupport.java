package jp.co.pm.ai.desktop.ui;

import java.util.BitSet;

import javafx.collections.ObservableList;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.DataFormat;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;

import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * ?z??v??^?X?N??????? {@link SpreadsheetView} ??A?f?[?^?s???h???b?O???h???b?v???????????B
 *
 * <p>??t?B???^???????????????? {@link SpreadsheetView#getModelRow(int)} ????f???s????????B
 */
public final class SpreadsheetPlanInputRowDragSupport {

    private static final DataFormat MODEL_ROW =
            new DataFormat("application/x-pm-ai-plan-input-model-row");

    private SpreadsheetPlanInputRowDragSupport() {}

    /**
     * ???? {@link TableView} ?? {@link TableView#setRowFactory(javafx.util.Callback)} ?? DnD ??t?^????B
     *
     * @param firstDataGridRow {@link SpreadsheetTabularSupport#spreadsheetFirstDataRowIndex()}?i?t?B???^?s????j
     * @param afterReorder ???f???s???X?g?X?V???O???b?h??\?z?????s???R?[???o?b?N
     */
    public static void install(
            SpreadsheetView spreadsheetView,
            int firstDataGridRow,
            ObservableList<ObservableList<String>> dataRows,
            Runnable afterReorder) {
        TableView<ObservableList<SpreadsheetCell>> tv =
                SpreadsheetTabularSupport.findInnerTableView(spreadsheetView);
        if (tv == null) {
            return;
        }
        tv.setRowFactory(
                table -> {
                    TableRow<ObservableList<SpreadsheetCell>> row = new TableRow<>();
                    row.setOnDragDetected(
                            event -> {
                                if (row.isEmpty()) {
                                    return;
                                }
                                int modelRow = spreadsheetView.getModelRow(row.getIndex());
                                if (!isDataRowAcceptable(spreadsheetView, firstDataGridRow, modelRow)) {
                                    return;
                                }
                                Dragboard db = row.startDragAndDrop(TransferMode.MOVE);
                                ClipboardContent cc = new ClipboardContent();
                                cc.put(MODEL_ROW, modelRow);
                                db.setContent(cc);
                                event.consume();
                            });
                    row.setOnDragOver(
                            event -> {
                                Dragboard dragboard = event.getDragboard();
                                if (dragboard.hasContent(MODEL_ROW)
                                        && row.getIndex() >= 0
                                        && !row.isEmpty()) {
                                    int modelRow = spreadsheetView.getModelRow(row.getIndex());
                                    if (isDataRowAcceptable(spreadsheetView, firstDataGridRow, modelRow)) {
                                        event.acceptTransferModes(TransferMode.MOVE);
                                    }
                                }
                                event.consume();
                            });
                    row.setOnDragDropped(
                            event -> {
                                Dragboard dragboard = event.getDragboard();
                                boolean success = false;
                                if (dragboard.hasContent(MODEL_ROW)
                                        && row.getIndex() >= 0
                                        && !row.isEmpty()) {
                                    Object payload = dragboard.getContent(MODEL_ROW);
                                    Integer srcModel =
                                            payload instanceof Integer ? (Integer) payload : null;
                                    int tgtModel = spreadsheetView.getModelRow(row.getIndex());
                                    if (srcModel != null
                                            && isDataRowAcceptable(
                                                    spreadsheetView, firstDataGridRow, srcModel)
                                            && isDataRowAcceptable(
                                                    spreadsheetView, firstDataGridRow, tgtModel)) {
                                        int srcData = srcModel - firstDataGridRow;
                                        int tgtData = tgtModel - firstDataGridRow;
                                        if (srcData >= 0
                                                && tgtData >= 0
                                                && srcData < dataRows.size()
                                                && tgtData < dataRows.size()) {
                                            if (srcData != tgtData) {
                                                moveDataRow(dataRows, srcData, tgtData);
                                                spreadsheetView.setComparator(null);
                                                ExcelLikeSpreadsheetFilter.resetAllColumnSortMenus(
                                                        spreadsheetView);
                                                afterReorder.run();
                                            }
                                            success = true;
                                        }
                                    }
                                }
                                event.setDropCompleted(success);
                                event.consume();
                            });
                    return row;
                });
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
