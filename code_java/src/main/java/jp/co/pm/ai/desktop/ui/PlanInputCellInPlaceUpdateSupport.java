package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Objects;

import javafx.collections.ObservableList;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;

/**
 * 配台計画タスク入力: セル編集後に {@code setGrid} せず、既存グリッドの item だけモデルへ揃える。
 *
 * <p>全表再構築はホスト {@code layoutBounds} が跳ねてウィンドウ揺れに見える（配台不要トグルと同趣旨）。
 */
public final class PlanInputCellInPlaceUpdateSupport {

    private PlanInputCellInPlaceUpdateSupport() {}

    /**
     * データ行1セルの表示値を更新する。
     *
     * @return 更新できたとき {@code true}
     */
    public static boolean applyValueToGrid(
            GridBase grid,
            int firstDataRowIndex,
            int dataIndex,
            int colIndex,
            String newValue) {
        SpreadsheetCell cell = cellAt(grid, firstDataRowIndex, dataIndex, colIndex);
        if (cell == null) {
            return false;
        }
        SpreadsheetTabularSupport.setSpreadsheetCellDisplayValue(
                cell, newValue != null ? newValue : "");
        return true;
    }

    /**
     * 全データ行のセル item を {@code rows} の文字列へ揃える（列数不一致なら失敗）。
     *
     * @return 同期できたとき {@code true}
     */
    public static boolean syncAllRowsFromModel(
            GridBase grid,
            int firstDataRowIndex,
            List<String> headers,
            ObservableList<ObservableList<String>> rows) {
        if (grid == null || headers == null || rows == null) {
            return false;
        }
        var gridRows = grid.getRows();
        if (gridRows == null) {
            return false;
        }
        int expectedGridRows = firstDataRowIndex + rows.size();
        if (gridRows.size() < expectedGridRows) {
            return false;
        }
        for (int r = 0; r < rows.size(); r++) {
            int gridRow = firstDataRowIndex + r;
            var rowCells = gridRows.get(gridRow);
            if (rowCells == null || rowCells.size() < headers.size()) {
                return false;
            }
            ObservableList<String> modelRow = rows.get(r);
            for (int c = 0; c < headers.size(); c++) {
                SpreadsheetCell cell = rowCells.get(c);
                if (cell == null) {
                    return false;
                }
                String want = cellAtModel(modelRow, c);
                String have = cell.getItem() != null ? Objects.toString(cell.getItem(), "") : "";
                if (!want.equals(have)) {
                    SpreadsheetTabularSupport.setSpreadsheetCellDisplayValue(cell, want);
                }
            }
        }
        return true;
    }

    private static SpreadsheetCell cellAt(
            GridBase grid, int firstDataRowIndex, int dataIndex, int colIndex) {
        if (grid == null || dataIndex < 0 || colIndex < 0) {
            return null;
        }
        var gridRows = grid.getRows();
        if (gridRows == null) {
            return null;
        }
        int gridRow = firstDataRowIndex + dataIndex;
        if (gridRow < 0 || gridRow >= gridRows.size()) {
            return null;
        }
        var rowCells = gridRows.get(gridRow);
        if (rowCells == null || colIndex >= rowCells.size()) {
            return null;
        }
        return rowCells.get(colIndex);
    }

    private static String cellAtModel(ObservableList<String> row, int colIndex) {
        if (row == null || colIndex < 0 || colIndex >= row.size()) {
            return "";
        }
        String v = row.get(colIndex);
        return v != null ? v : "";
    }
}
