package jp.co.pm.ai.desktop.ui;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;

import javafx.collections.ObservableList;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;

import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/**
 * 配台計画タスク入力: 「未加工」が 0 より大きく、列「配台使用残数量」と数値が一致しない行を検出する。
 *
 * <p>スリット等で段階1の配台使用残数量が未加工と乖離するケース（例: JR260602）向け。
 */
public final class PlanInputUnprocessedDispatchRemainingMismatchSupport {

    private static final double EPS = 1e-6;

    public static final String COL_UNPROCESSED = "未加工";

    public static final String COL_DISPATCH_REMAINING = "配台使用残数量";

    public static final String COL_TASK_ID = "依頼NO";

    private PlanInputUnprocessedDispatchRemainingMismatchSupport() {}

    /**
     * 未加工 &gt; 0 かつ配台使用残数量と数値不一致（空・非数値も不一致扱い）。
     */
    public static boolean isMismatch(List<String> headers, List<String> row) {
        if (headers == null || row == null) {
            return false;
        }
        int idxUnp = headers.indexOf(COL_UNPROCESSED);
        int idxRem = headers.indexOf(COL_DISPATCH_REMAINING);
        if (idxUnp < 0 || idxRem < 0) {
            return false;
        }
        Optional<Double> unpOpt =
                Stage2RollUnitLengthTables.optionalUnprocessedCell(cellAt(row, idxUnp));
        if (unpOpt.isEmpty() || unpOpt.get() <= EPS) {
            return false;
        }
        double unp = unpOpt.get();
        Optional<Double> remOpt = optionalNumericCell(cellAt(row, idxRem));
        if (remOpt.isEmpty()) {
            return true;
        }
        return Math.abs(unp - remOpt.get()) > EPS;
    }

    public static List<String> collectMismatchTaskIds(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        if (headers == null || rows == null || rows.isEmpty()) {
            return List.of();
        }
        int idxTask = headers.indexOf(COL_TASK_ID);
        Set<String> ids = new LinkedHashSet<>();
        int rowNum = 0;
        for (ObservableList<String> row : rows) {
            if (isMismatch(headers, row)) {
                String id = taskIdForRow(headers, row, idxTask, rowNum);
                if (!id.isBlank()) {
                    ids.add(id);
                }
            }
            rowNum++;
        }
        return List.copyOf(ids);
    }

    public static String warningMessage(List<String> mismatchTaskIds) {
        if (mismatchTaskIds == null || mismatchTaskIds.isEmpty()) {
            return "";
        }
        if (mismatchTaskIds.size() == 1) {
            return mismatchTaskIds.get(0)
                    + "の未加工と配台使用残数量が異なります、手動修正してください。";
        }
        return String.join("、", mismatchTaskIds)
                + " の未加工と配台使用残数量が異なります、手動修正してください。";
    }

    /** 違反行の {@code 配台使用残数量} セル背景を赤系で強調する。 */
    public static void applyViolationHighlights(
            GridBase grid,
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            int firstDataRowIndex) {
        if (grid == null || headers == null || rows == null || rows.isEmpty()) {
            return;
        }
        int idxRem = headers.indexOf(COL_DISPATCH_REMAINING);
        if (idxRem < 0) {
            return;
        }
        var gridRows = grid.getRows();
        for (int r = 0; r < rows.size(); r++) {
            int gridRow = firstDataRowIndex + r;
            if (gridRow < 0 || gridRow >= gridRows.size()) {
                continue;
            }
            var rowCells = gridRows.get(gridRow);
            if (idxRem >= rowCells.size()) {
                continue;
            }
            SpreadsheetCell cell = rowCells.get(idxRem);
            if (cell == null) {
                continue;
            }
            if (isMismatch(headers, rows.get(r))) {
                cell.setStyle(TabularCellHighlight.PLAN_INPUT_DISPATCH_REMAINING_MISMATCH_STYLE);
            }
        }
    }

    private static String taskIdForRow(
            List<String> headers, List<String> row, int idxTask, int rowIndex) {
        if (idxTask >= 0) {
            String id = cellAt(row, idxTask);
            if (!id.isBlank()) {
                return id;
            }
        }
        return "行" + (rowIndex + 1);
    }

    private static Optional<Double> optionalNumericCell(String cell) {
        if (cell == null) {
            return Optional.empty();
        }
        String s = cell.strip();
        if (s.isEmpty()
                || s.equalsIgnoreCase("nan")
                || s.equalsIgnoreCase("none")
                || s.equals("-")
                || s.equals("—")
                || s.equals("―")) {
            return Optional.empty();
        }
        try {
            return Optional.of(Double.parseDouble(s.replace(",", ".")));
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }

    private static String cellAt(List<String> row, int colIndex) {
        if (row == null || colIndex < 0 || colIndex >= row.size()) {
            return "";
        }
        String v = row.get(colIndex);
        return v != null ? v.strip() : "";
    }
}
