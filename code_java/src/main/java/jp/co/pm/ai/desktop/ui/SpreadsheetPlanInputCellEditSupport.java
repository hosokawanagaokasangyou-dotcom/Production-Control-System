package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.function.Consumer;

import javafx.collections.ObservableList;
import javafx.scene.Node;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextInputControl;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.input.MouseButton;
import javafx.scene.input.MouseEvent;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * Plan-input {@link SpreadsheetView}: cell editing via double-click dialog (no inline editor) and DISABLE of DELETE
 * clearing/removing full rows (ControlsFX default).
 */
public final class SpreadsheetPlanInputCellEditSupport {

    @FunctionalInterface
    public interface LimitedOperatorEditor {
        void edit(
                ObservableList<String> row,
                String currentValue,
                double anchorScreenX,
                double anchorScreenY,
                Consumer<String> onConfirmed);
    }

    private SpreadsheetPlanInputCellEditSupport() {}

    /**
     * @param firstDataGridRow {@link SpreadsheetTabularSupport#spreadsheetFirstDataRowIndex()}
     * @param headersRef column titles aligned with {@code rows.get(r).get(c)}
     */
    public static void install(
            SpreadsheetView spreadsheetView,
            Stage owner,
            int firstDataGridRow,
            List<String> headersRef,
            ObservableList<ObservableList<String>> rows,
            Runnable rebuildSpreadsheet) {
        install(
                spreadsheetView,
                owner,
                firstDataGridRow,
                headersRef,
                rows,
                rebuildSpreadsheet,
                null,
                null);
    }

    /**
     * @param afterColumnEdit セル確定後・{@code rebuildSpreadsheet} 直前（列見出しを渡す。{@code null} 可）
     */
    public static void install(
            SpreadsheetView spreadsheetView,
            Stage owner,
            int firstDataGridRow,
            List<String> headersRef,
            ObservableList<ObservableList<String>> rows,
            Runnable rebuildSpreadsheet,
            Consumer<String> afterColumnEdit) {
        install(
                spreadsheetView,
                owner,
                firstDataGridRow,
                headersRef,
                rows,
                rebuildSpreadsheet,
                afterColumnEdit,
                null);
    }

    /**
     * @param limitedOperatorEditor 「担当OP_限定」専用編集処理（{@code null} の場合は編集しない）
     */
    public static void install(
            SpreadsheetView spreadsheetView,
            Stage owner,
            int firstDataGridRow,
            List<String> headersRef,
            ObservableList<ObservableList<String>> rows,
            Runnable rebuildSpreadsheet,
            Consumer<String> afterColumnEdit,
            LimitedOperatorEditor limitedOperatorEditor) {
        if (spreadsheetView == null || owner == null || rebuildSpreadsheet == null) {
            return;
        }

        spreadsheetView.addEventFilter(
                MouseEvent.MOUSE_CLICKED,
                e -> {
                    if (e.getClickCount() != 2 || e.getButton() != MouseButton.PRIMARY) {
                        return;
                    }
                    Node n =
                            e.getPickResult() != null
                                    ? e.getPickResult().getIntersectedNode()
                                    : null;
                    TableCell<?, ?> tc = SpreadsheetPlanInputRowDragSupport.findTableCell(n);
                    if (tc == null || !SpreadsheetPlanInputRowDragSupport.isUnderSpreadsheet(spreadsheetView, tc)) {
                        return;
                    }
                    int viewRow = tc.getIndex();
                    if (viewRow < 0 || tc.isEmpty()) {
                        return;
                    }
                    int modelRow = spreadsheetView.getModelRow(viewRow);
                    if (modelRow < firstDataGridRow) {
                        return;
                    }
                    int dataIndex = modelRow - firstDataGridRow;
                    if (dataIndex < 0 || dataIndex >= rows.size()) {
                        return;
                    }

                    int colIndex = resolveColumnIndex(tc);
                    if (colIndex < 0 || colIndex >= headersRef.size()) {
                        return;
                    }

                    String columnTitle = headersRef.get(colIndex);
                    ObservableList<String> row = rows.get(dataIndex);
                    while (row.size() <= colIndex) {
                        row.add("");
                    }
                    String cur = row.get(colIndex) != null ? row.get(colIndex) : "";

                    if ("配台不要".equals(columnTitle)) {
                        if (TabularCellHighlight.planInputExcludeFromAssignmentIsOn(cur)) {
                            row.set(colIndex, "");
                        } else {
                            row.set(colIndex, "yes");
                        }
                        rebuildSpreadsheet.run();
                        e.consume();
                        return;
                    }

                    @SuppressWarnings("rawtypes")
                    TableColumn col = tc.getTableColumn();
                    double colW = col != null ? col.getWidth() : 0;

                    if (PlanInputCellEditRouting.editorFor(columnTitle)
                            == PlanInputCellEditRouting.Editor.READ_ONLY) {
                        SpreadsheetPlanInputCellEditDialog.viewReadOnly(
                                owner,
                                columnTitle,
                                cur,
                                "段階2 が「"
                                        + PlanInputAiSpecialParseColumn.resolveSourceColumnTitle(
                                                headersRef)
                                        + "」を解析した結果です。表示専用で、手入力しても配台には反映されません。",
                                colW,
                                e.getScreenX(),
                                e.getScreenY());
                        e.consume();
                        return;
                    }

                    if (PlanInputCellEditRouting.editorFor(columnTitle)
                            == PlanInputCellEditRouting.Editor.LIMITED_OPERATOR_CHECKLIST) {
                        if (limitedOperatorEditor != null) {
                            limitedOperatorEditor.edit(
                                    row,
                                    cur,
                                    e.getScreenX(),
                                    e.getScreenY(),
                                    newVal ->
                                            applyCellEditValue(
                                                    columnTitle,
                                                    colIndex,
                                                    headersRef,
                                                    row,
                                                    newVal,
                                                    afterColumnEdit,
                                                    rebuildSpreadsheet));
                        }
                        e.consume();
                        return;
                    }

                    if (PlanInputDateColumnSupport.isEditableDateTimeColumn(columnTitle)) {
                        SpreadsheetPlanInputCellEditDialog.editDateTime(
                                        owner,
                                        columnTitle,
                                        cur,
                                        e.getScreenX(),
                                        e.getScreenY())
                                .ifPresent(
                                        newVal -> {
                                            row.set(colIndex, newVal);
                                            rebuildSpreadsheet.run();
                                        });
                        e.consume();
                        return;
                    }

                    if (PlanInputDateColumnSupport.isEditableDateColumn(columnTitle)) {
                        SpreadsheetPlanInputCellEditDialog.editDate(
                                        owner,
                                        columnTitle,
                                        cur,
                                        e.getScreenX(),
                                        e.getScreenY())
                                .ifPresent(
                                        newVal -> {
                                            if (PlanInputRawInputDateShift.COL_RAW_INPUT_DATE.equals(
                                                    columnTitle)) {
                                                PlanInputRawInputDateShift
                                                        .propagateRawInputDateToSameTaskIdRows(
                                                                headersRef,
                                                                rows,
                                                                dataIndex,
                                                                newVal);
                                            } else {
                                                row.set(colIndex, newVal);
                                            }
                                            rebuildSpreadsheet.run();
                                        });
                        e.consume();
                        return;
                    }

                    SpreadsheetPlanInputCellEditDialog.edit(
                                    owner,
                                    columnTitle,
                                    cur,
                                    colW,
                                    e.getScreenX(),
                                    e.getScreenY())
                            .ifPresent(
                                    newVal -> applyCellEditValue(
                                            columnTitle,
                                            colIndex,
                                            headersRef,
                                            row,
                                            newVal,
                                            afterColumnEdit,
                                            rebuildSpreadsheet));
                    e.consume();
                });

        spreadsheetView.addEventFilter(
                KeyEvent.KEY_PRESSED,
                e -> {
                    if (e.getCode() != KeyCode.DELETE) {
                        return;
                    }
                    Node target = e.getTarget() instanceof Node ? (Node) e.getTarget() : null;
                    if (!SpreadsheetPlanInputRowDragSupport.isUnderSpreadsheet(spreadsheetView, target)) {
                        return;
                    }
                    if (isInsideTextInput(target)) {
                        return;
                    }
                    e.consume();
                });
    }

    private static void applyCellEditValue(
            String columnTitle,
            int colIndex,
            List<String> headersRef,
            ObservableList<String> row,
            String newVal,
            Consumer<String> afterColumnEdit,
            Runnable rebuildSpreadsheet) {
        if (PlanInputProcessSequenceRowOrder.COL_DISPATCH_TRIAL_ORDER.equals(columnTitle)) {
            String trimmed = newVal != null ? newVal.strip() : "";
            if (!trimmed.isEmpty()
                    && PlanInputProcessSequenceRowOrder.parsePositiveTrialOrderSortKey(trimmed)
                            == null) {
                return;
            }
            newVal = trimmed;
        }
        row.set(colIndex, newVal);
        PlanInputAiSpecialParseColumn.clearStaleParseAfterRemarkEdit(headersRef, row, columnTitle);
        if (afterColumnEdit != null) {
            afterColumnEdit.accept(columnTitle);
        }
        rebuildSpreadsheet.run();
    }

    private static boolean isInsideTextInput(Node n) {
        while (n != null) {
            if (n instanceof TextInputControl) {
                return true;
            }
            n = n.getParent();
        }
        return false;
    }

    private static int resolveColumnIndex(TableCell<?, ?> tc) {
        TableColumn<?, ?> tcol = tc.getTableColumn();
        if (tcol == null) {
            return -1;
        }
        TableView<?> tv = tc.getTableView();
        if (tv == null) {
            return -1;
        }
        return tv.getColumns().indexOf(tcol);
    }
}
