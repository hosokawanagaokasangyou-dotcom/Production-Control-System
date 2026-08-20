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

    private MasterDispatchSheetGridSupport() {}

    public static GridBase buildEditable(
            MasterDispatchSheetEditRules.SheetKind kind, List<List<String>> rows) {
        MasterDispatchSheetEditRules.SheetKind sheet =
                kind != null ? kind : MasterDispatchSheetEditRules.SheetKind.SKILLS;
        List<List<String>> src = rows != null ? rows : List.of();
        int dataCols = 1;
        for (List<String> row : src) {
            if (row != null && row.size() > dataCols) {
                dataCols = row.size();
            }
        }
        int cols = dataCols + MasterDispatchSheetEditRules.EXTRA_COLS;
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int rc = firstData + Math.max(src.size(), 1) + MasterDispatchSheetEditRules.EXTRA_ROWS;
        GridBase grid = new GridBase(rc, cols);
        List<String> titles = MasterDispatchSheetEditRules.columnTitles(sheet, src, cols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(titles);
        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(rc);
        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < cols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(
                            SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_FILTER_ROW);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);
        for (int dataRow = 0; dataRow < rc - firstData; dataRow++) {
            int gridRow = firstData + dataRow;
            ObservableList<SpreadsheetCell> rowCells = FXCollections.observableArrayList();
            List<String> srcRow = dataRow < src.size() ? src.get(dataRow) : List.of();
            for (int c = 0; c < cols; c++) {
                String raw =
                        srcRow != null && c < srcRow.size() && srcRow.get(c) != null
                                ? srcRow.get(c)
                                : "";
                SpreadsheetCell cell = createCell(sheet, src, dataRow, gridRow, c, raw);
                rowCells.add(cell);
            }
            gridRows.add(rowCells);
        }
        grid.setRows(gridRows);
        applyTitleRowHeights(grid, sheet);
        if (sheet == MasterDispatchSheetEditRules.SheetKind.COMBINATIONS) {
            wireCombinationLiveSync(grid);
        } else {
            wireEquipmentColumnTitles(grid, sheet);
        }
        return grid;
    }

    public static List<List<String>> extract(
            SpreadsheetView view, MasterDispatchSheetEditRules.SheetKind kind) {
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
        List<List<String>> data = MasterDispatchSheetEditRules.skipFilterRow(raw);
        List<List<String>> trimmed = MasterDispatchSheetsDocument.trimTrailingEmpty(data);
        return MasterDispatchSheetEditRules.normalizeOnExtract(kind, trimmed);
    }

    private static SpreadsheetCell createCell(
            MasterDispatchSheetEditRules.SheetKind kind,
            List<List<String>> src,
            int dataRow,
            int gridRow,
            int col,
            String raw) {
        SpreadsheetCell cell;
        if (kind == MasterDispatchSheetEditRules.SheetKind.SKILLS
                && dataRow >= 2
                && col >= 1) {
            cell =
                    SpreadsheetCellType.LIST(skillChoices(raw))
                            .createCell(gridRow, col, 1, 1, raw);
        } else {
            cell = SpreadsheetCellType.STRING.createCell(gridRow, col, 1, 1, raw);
        }
        cell.setEditable(MasterDispatchSheetEditRules.isEditable(kind, dataRow, col, src));
        cell.setStyle(cellStyle(kind, src, dataRow, col, raw));
        if (MasterDispatchSheetEditRules.isColumnTitleSourceRow(kind, dataRow, src)) {
            cell.setWrapText(true);
        }
        return cell;
    }

    private static List<String> skillChoices(String current) {
        List<String> items = new ArrayList<>();
        items.add("");
        for (int i = 1; i <= 15; i++) {
            items.add("OP" + i);
            items.add("AS" + i);
        }
        String cur = current != null ? current.strip() : "";
        if (!cur.isEmpty() && !items.contains(cur)) {
            items.add(cur);
        }
        return items;
    }

    private static String cellStyle(
            MasterDispatchSheetEditRules.SheetKind kind,
            List<List<String>> src,
            int dataRow,
            int col,
            String raw) {
        if (MasterDispatchSheetEditRules.isInvalidValue(kind, dataRow, col, src)) {
            return TabularCellHighlight.PLAN_INPUT_DISPATCHABLE_DATETIME_VIOLATION_STYLE;
        }
        if (kind == MasterDispatchSheetEditRules.SheetKind.COMBINATIONS && dataRow > 0) {
            int procCol = MasterDispatchSheetEditRules.headerIndex(header(src), "工程名");
            int machCol = MasterDispatchSheetEditRules.headerIndex(header(src), "機械名");
            int comboCol = MasterDispatchSheetEditRules.headerIndex(header(src), "工程+機械", "工程＋機械");
            String proc = MasterDispatchSheetEditRules.cell(src, dataRow, procCol);
            String mach = MasterDispatchSheetEditRules.cell(src, dataRow, machCol);
            String comboCell = MasterDispatchSheetEditRules.cell(src, dataRow, comboCol);
            String combo = MasterDispatchSheetEditRules.comboRowStyle(proc, mach, comboCell);
            if (!combo.isEmpty()) {
                return combo;
            }
        }
        if (!MasterDispatchSheetEditRules.isEditable(kind, dataRow, col, src)) {
            return SpreadsheetTabularSupport.READABLE_STYLE_FILTER_ROW;
        }
        if (col == 0) {
            return SpreadsheetTabularSupport.READABLE_STYLE_LEADING_COL;
        }
        if ((kind == MasterDispatchSheetEditRules.SheetKind.SKILLS
                        || kind == MasterDispatchSheetEditRules.SheetKind.NEED
                        || kind == MasterDispatchSheetEditRules.SheetKind.SPEED)
                && dataRow <= 1) {
            return SpreadsheetTabularSupport.READABLE_STYLE_FILTER_ROW;
        }
        return SpreadsheetTabularSupport.READABLE_STYLE_DATA_WHITE;
    }

    private static List<String> header(List<List<String>> src) {
        if (src == null || src.isEmpty() || src.get(0) == null) {
            return List.of();
        }
        return src.get(0);
    }

    private static void applyTitleRowHeights(
            GridBase grid, MasterDispatchSheetEditRules.SheetKind kind) {
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        List<List<String>> snap = snapshotSkippingFilter(grid);
        grid.setRowHeightCallback(
                row -> {
                    if (row == SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW) {
                        return 26.0;
                    }
                    int dataRow = row - firstData;
                    if (MasterDispatchSheetEditRules.isColumnTitleSourceRow(kind, dataRow, snap)) {
                        return 42.0;
                    }
                    return 24.0;
                });
    }

    private static void wireEquipmentColumnTitles(
            GridBase grid, MasterDispatchSheetEditRules.SheetKind kind) {
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        if (grid == null || grid.getRows() == null) {
            return;
        }
        Runnable refresh = () -> refreshColumnTitles(grid, kind);
        int limit = Math.min(grid.getRowCount(), firstData + 8);
        for (int gridRow = firstData; gridRow < limit; gridRow++) {
            ObservableList<SpreadsheetCell> row = grid.getRows().get(gridRow);
            if (row == null || row.isEmpty()) {
                continue;
            }
            String a = cellText(row, 0);
            if (!"工程名".equals(a) && !"機械名".equals(a)) {
                continue;
            }
            for (SpreadsheetCell cell : row) {
                if (cell != null) {
                    cell.itemProperty().addListener((obs, o, n) -> refresh.run());
                }
            }
        }
    }

    private static void refreshColumnTitles(
            GridBase grid, MasterDispatchSheetEditRules.SheetKind kind) {
        List<String> titles =
                MasterDispatchSheetEditRules.columnTitles(
                        kind, snapshotSkippingFilter(grid), grid.getColumnCount());
        ObservableList<String> headers = grid.getColumnHeaders();
        for (int c = 0; c < titles.size() && c < headers.size(); c++) {
            if (!titles.get(c).equals(headers.get(c))) {
                headers.set(c, titles.get(c));
            }
        }
    }

    private static List<List<String>> snapshotSkippingFilter(GridBase grid) {
        if (grid == null || grid.getRows() == null) {
            return List.of();
        }
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        List<List<String>> out = new ArrayList<>();
        for (int r = firstData; r < grid.getRowCount(); r++) {
            ObservableList<SpreadsheetCell> row = grid.getRows().get(r);
            List<String> cells = new ArrayList<>();
            int cols = grid.getColumnCount();
            for (int c = 0; c < cols; c++) {
                cells.add(cellText(row, c));
            }
            out.add(cells);
        }
        return out;
    }

    private static void wireCombinationLiveSync(GridBase grid) {
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        if (grid == null || grid.getRows() == null || grid.getRowCount() <= firstData) {
            return;
        }
        List<String> header = new ArrayList<>();
        ObservableList<SpreadsheetCell> headerRow = grid.getRows().get(firstData);
        if (headerRow == null) {
            return;
        }
        for (SpreadsheetCell c : headerRow) {
            header.add(c != null && c.getText() != null ? c.getText() : "");
        }
        int procCol = MasterDispatchSheetEditRules.headerIndex(header, "工程名");
        int machCol = MasterDispatchSheetEditRules.headerIndex(header, "機械名");
        int comboCol = MasterDispatchSheetEditRules.headerIndex(header, "工程+機械", "工程＋機械");
        if (procCol < 0 || machCol < 0) {
            return;
        }
        for (int gridRow = firstData + 1; gridRow < grid.getRowCount(); gridRow++) {
            ObservableList<SpreadsheetCell> row = grid.getRows().get(gridRow);
            if (row == null) {
                continue;
            }
            Runnable refresh = () -> applyCombinationRowLive(row, procCol, machCol, comboCol);
            if (procCol < row.size() && row.get(procCol) != null) {
                row.get(procCol).itemProperty().addListener((obs, o, n) -> refresh.run());
            }
            if (machCol < row.size() && row.get(machCol) != null) {
                row.get(machCol).itemProperty().addListener((obs, o, n) -> refresh.run());
            }
        }
    }

    private static void applyCombinationRowLive(
            ObservableList<SpreadsheetCell> row, int procCol, int machCol, int comboCol) {
        String proc = cellText(row, procCol);
        String mach = cellText(row, machCol);
        String display =
                !proc.isBlank() && !mach.isBlank() ? proc.strip() + "+" + mach.strip() : "";
        if (comboCol >= 0 && comboCol < row.size() && row.get(comboCol) != null) {
            String cur = cellText(row, comboCol);
            if (!display.equals(cur)) {
                row.get(comboCol).setItem(display);
            }
        }
        String style = MasterDispatchSheetEditRules.comboRowStyle(proc, mach, display);
        if (style.isEmpty()) {
            style = SpreadsheetTabularSupport.READABLE_STYLE_DATA_WHITE;
        }
        for (int c = 0; c < row.size(); c++) {
            SpreadsheetCell cell = row.get(c);
            if (cell == null) {
                continue;
            }
            cell.setStyle(style);
        }
    }

    private static String cellText(ObservableList<SpreadsheetCell> row, int col) {
        if (row == null || col < 0 || col >= row.size() || row.get(col) == null) {
            return "";
        }
        String t = row.get(col).getText();
        return t != null ? t : "";
    }
}
