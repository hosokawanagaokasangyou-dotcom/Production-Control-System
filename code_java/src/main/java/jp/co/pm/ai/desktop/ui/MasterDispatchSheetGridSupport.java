package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import org.controlsfx.control.spreadsheet.Grid;
import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.io.MasterDispatchSheetsDocument;

/** MASTER 4 シート用の編集可能格子。 */
public final class MasterDispatchSheetGridSupport {

    static final int EXTRA_ROWS = 40;
    static final int EXTRA_COLS = 16;

    private MasterDispatchSheetGridSupport() {}

    public static GridBase buildEditable(List<List<String>> rows) {
        List<List<String>> src = rows != null ? rows : List.of();
        int dataCols = 1;
        for (List<String> row : src) {
            if (row != null && row.size() > dataCols) {
                dataCols = row.size();
            }
        }
        int cols = dataCols + EXTRA_COLS;
        int rc = Math.max(src.size(), 1) + EXTRA_ROWS;
        GridBase grid = new GridBase(rc, cols);
        grid.getColumnHeaders().clear();
        for (int c = 0; c < cols; c++) {
            grid.getColumnHeaders().add(excelColumnLabel(c));
        }
        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(rc);
        for (int r = 0; r < rc; r++) {
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            List<String> srcRow = r < src.size() ? src.get(r) : List.of();
            for (int c = 0; c < cols; c++) {
                String raw = srcRow != null && c < srcRow.size() && srcRow.get(c) != null ? srcRow.get(c) : "";
                SpreadsheetCell cell = SpreadsheetCellType.STRING.createCell(r, c, 1, 1, raw);
                cell.setEditable(true);
                cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_DATA_WHITE);
                rowCells.add(cell);
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        return grid;
    }

    public static List<List<String>> extract(SpreadsheetView view) {
        if (view == null || view.getGrid() == null) {
            return List.of();
        }
        Grid grid = view.getGrid();
        int rows = grid.getRowCount();
        int cols = grid.getColumnCount();
        List<List<String>> raw = new ArrayList<>(rows);
        for (int r = 0; r < rows; r++) {
            ObservableList<SpreadsheetCell> row = grid.getRows().get(r);
            List<String> cells = new ArrayList<>(cols);
            for (int c = 0; c < cols; c++) {
                String v = "";
                if (row != null && c < row.size() && row.get(c) != null) {
                    String t = row.get(c).getText();
                    v = t != null ? t : "";
                }
                cells.add(v);
            }
            raw.add(cells);
        }
        return MasterDispatchSheetsDocument.trimTrailingEmpty(raw);
    }

    static String excelColumnLabel(int col0) {
        StringBuilder sb = new StringBuilder();
        int n = col0 + 1;
        while (n > 0) {
            n--;
            sb.append((char) ('A' + (n % 26)));
            n /= 26;
        }
        return sb.reverse().toString();
    }
}
