package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.junit.jupiter.api.Test;

class PlanInputExcludeToggleSupportTest {

    @Test
    void toggledValue_emptyBecomesYes() {
        assertEquals("yes", PlanInputExcludeToggleSupport.toggledValue(""));
        assertEquals("yes", PlanInputExcludeToggleSupport.toggledValue(null));
    }

    @Test
    void toggledValue_yesBecomesEmpty() {
        assertEquals("", PlanInputExcludeToggleSupport.toggledValue("yes"));
        assertEquals("", PlanInputExcludeToggleSupport.toggledValue("はい"));
    }

    @Test
    void applyVisual_setsYesWithExcludeStyleOnNonEditableCell() {
        SpreadsheetCell cell = SpreadsheetCellType.STRING.createCell(1, 0, 1, 1, "");
        cell.setEditable(false);

        PlanInputExcludeToggleSupport.applyVisual(cell, "yes", false);

        assertEquals("yes", String.valueOf(cell.getItem()));
        assertEquals(TabularCellHighlight.PLAN_INPUT_EXCLUDE_YES_STYLE, cell.getStyle());
        assertFalse(cell.isEditable());
    }

    @Test
    void applyVisual_clearsYesToLeadingStyle() {
        SpreadsheetCell cell =
                SpreadsheetCellType.STRING.createCell(
                        1, 0, 1, 1, "yes");
        cell.setEditable(false);
        cell.setStyle(TabularCellHighlight.PLAN_INPUT_EXCLUDE_YES_STYLE);

        PlanInputExcludeToggleSupport.applyVisual(cell, "", true);

        assertEquals("", String.valueOf(cell.getItem() != null ? cell.getItem() : ""));
        assertEquals(SpreadsheetTabularSupport.READABLE_STYLE_LEADING_COL, cell.getStyle());
    }

    @Test
    void applyToGrid_updatesTargetCellWithoutChangingOthers() {
        GridBase grid = new GridBase(3, 2);
        ObservableList<ObservableList<SpreadsheetCell>> rows = FXCollections.observableArrayList();
        for (int r = 0; r < 3; r++) {
            ObservableList<SpreadsheetCell> row = FXCollections.observableArrayList();
            for (int c = 0; c < 2; c++) {
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(r, c, 1, 1, r == 0 ? "f" : "");
                cell.setEditable(false);
                row.add(cell);
            }
            rows.add(row);
        }
        grid.setRows(rows);

        boolean ok = PlanInputExcludeToggleSupport.applyToGrid(grid, 1, 0, 0, "yes", true);
        assertTrue(ok);
        assertEquals("yes", String.valueOf(grid.getRows().get(1).get(0).getItem()));
        assertEquals(
                TabularCellHighlight.PLAN_INPUT_EXCLUDE_YES_STYLE,
                grid.getRows().get(1).get(0).getStyle());
        assertEquals("", String.valueOf(grid.getRows().get(2).get(0).getItem()));
    }
}
