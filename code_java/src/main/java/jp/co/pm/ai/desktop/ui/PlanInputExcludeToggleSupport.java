package jp.co.pm.ai.desktop.ui;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;

/**
 * 配台計画タスク入力の「配台不要」列: ダブルクリックトグルの値と、全表再構築なしのセル表示更新。
 *
 * <p>{@link org.controlsfx.control.spreadsheet.SpreadsheetView#setGrid} を伴う再構築はホスト
 * {@code layoutBounds} が跳ねてウィンドウ揺れに見えるため、トグル時は既存セルへ値・スタイルだけ反映する。
 */
public final class PlanInputExcludeToggleSupport {

    private PlanInputExcludeToggleSupport() {}

    /** 現在値からトグル後のセル文字列（オン→空、オフ→{@code yes}）。 */
    public static String toggledValue(String current) {
        if (TabularCellHighlight.planInputExcludeFromAssignmentIsOn(current)) {
            return "";
        }
        return "yes";
    }

    /**
     * 非編集セルでも値を入れ、配台不要オンなら赤スタイル、オフなら先頭列／通常白スタイルにする。
     *
     * @param leadingColumn オフ時に先頭固定列スタイルを使うか
     */
    public static void applyVisual(SpreadsheetCell cell, String newValue, boolean leadingColumn) {
        if (cell == null) {
            return;
        }
        String v = newValue != null ? newValue : "";
        SpreadsheetTabularSupport.setSpreadsheetCellDisplayValue(cell, v);
        if (TabularCellHighlight.planInputExcludeFromAssignmentIsOn(v)) {
            cell.setStyle(TabularCellHighlight.PLAN_INPUT_EXCLUDE_YES_STYLE);
        } else if (leadingColumn) {
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_LEADING_COL);
        } else {
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_DATA_WHITE);
        }
    }

    /**
     * データ行 index（フィルタ行を除く）のセルだけ更新する。
     *
     * @return 更新できたとき {@code true}
     */
    public static boolean applyToGrid(
            GridBase grid,
            int firstDataRowIndex,
            int dataIndex,
            int colIndex,
            String newValue,
            boolean leadingColumn) {
        if (grid == null || dataIndex < 0 || colIndex < 0) {
            return false;
        }
        var gridRows = grid.getRows();
        if (gridRows == null) {
            return false;
        }
        int gridRow = firstDataRowIndex + dataIndex;
        if (gridRow < 0 || gridRow >= gridRows.size()) {
            return false;
        }
        var rowCells = gridRows.get(gridRow);
        if (rowCells == null || colIndex >= rowCells.size()) {
            return false;
        }
        SpreadsheetCell cell = rowCells.get(colIndex);
        if (cell == null) {
            return false;
        }
        applyVisual(cell, newValue, leadingColumn);
        return true;
    }
}
