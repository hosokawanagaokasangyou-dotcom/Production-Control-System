package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.junit.jupiter.api.Test;

class PlanInputCellInPlaceUpdateSupportTest {

    @Test
    void syncAllRowsFromModel_updatesChangedCellsOnlyKeepsStructure() {
        GridBase grid = new GridBase(3, 2);
        ObservableList<ObservableList<SpreadsheetCell>> gridRows =
                FXCollections.observableArrayList();
        for (int r = 0; r < 3; r++) {
            ObservableList<SpreadsheetCell> row = FXCollections.observableArrayList();
            for (int c = 0; c < 2; c++) {
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(
                                r, c, 1, 1, r == 0 ? "hdr" : "old");
                cell.setEditable(false);
                row.add(cell);
            }
            gridRows.add(row);
        }
        grid.setRows(gridRows);

        List<String> headers = List.of("依頼NO", "在庫場所");
        ObservableList<ObservableList<String>> model = FXCollections.observableArrayList();
        model.add(FXCollections.observableArrayList("W9-4", "K"));
        model.add(FXCollections.observableArrayList("W9-4", "S"));

        assertTrue(
                PlanInputCellInPlaceUpdateSupport.syncAllRowsFromModel(grid, 1, headers, model));
        assertEquals("W9-4", String.valueOf(grid.getRows().get(1).get(0).getItem()));
        assertEquals("K", String.valueOf(grid.getRows().get(1).get(1).getItem()));
        assertEquals("S", String.valueOf(grid.getRows().get(2).get(1).getItem()));
        assertEquals("hdr", String.valueOf(grid.getRows().get(0).get(0).getItem()));
    }

    @Test
    void syncAllRowsFromModel_failsWhenGridTooShort() {
        GridBase grid = new GridBase(1, 1);
        ObservableList<ObservableList<SpreadsheetCell>> gridRows =
                FXCollections.observableArrayList();
        ObservableList<SpreadsheetCell> row = FXCollections.observableArrayList();
        row.add(SpreadsheetCellType.STRING.createCell(0, 0, 1, 1, "f"));
        gridRows.add(row);
        grid.setRows(gridRows);

        ObservableList<ObservableList<String>> model = FXCollections.observableArrayList();
        model.add(FXCollections.observableArrayList("A"));
        assertFalse(
                PlanInputCellInPlaceUpdateSupport.syncAllRowsFromModel(
                        grid, 1, List.of("依頼NO"), model));
    }
}
