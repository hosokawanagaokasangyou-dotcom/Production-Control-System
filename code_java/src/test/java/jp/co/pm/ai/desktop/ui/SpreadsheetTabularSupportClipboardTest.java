package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TablePosition;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class SpreadsheetTabularSupportClipboardTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void pastePlainTextIntoSpreadsheetSelection_writesListCellValue() throws Exception {
        CountDownLatch done = new CountDownLatch(1);
        AtomicInteger applied = new AtomicInteger(-1);
        Platform.runLater(
                () -> {
                    GridBase grid = new GridBase(3, 3);
                    java.util.List<javafx.collections.ObservableList<SpreadsheetCell>> gridRows =
                            new java.util.ArrayList<>(3);
                    for (int r = 0; r < 3; r++) {
                        javafx.collections.ObservableList<SpreadsheetCell> rowCells =
                                FXCollections.observableArrayList();
                        for (int c = 0; c < 3; c++) {
                            SpreadsheetCell cell =
                                    new MasterDispatchListCellType(List.of("", "OP", "AS"))
                                            .createCell(r, c, 1, 1, "");
                            cell.setEditable(true);
                            rowCells.add(cell);
                        }
                        gridRows.add(rowCells);
                    }
                    grid.setRows(gridRows);
                    SpreadsheetView view = new SpreadsheetView(grid);
                    view.setEditable(true);
                    SpreadsheetColumn column = view.getColumns().get(2);
                    view.getSelectionModel().clearAndSelect(1, column);
                    SpreadsheetCell beforePaste = grid.getRows().get(1).get(2);
                    applied.set(
                            SpreadsheetTabularSupport.pastePlainTextIntoSpreadsheetSelection(
                                    view, "AS"));
                    SpreadsheetCell target = grid.getRows().get(1).get(2);
                    assertEquals("AS", target.getItem());
                    assertEquals("AS", target.getText());
                    // itemProperty リスナー（工程+機械の連動など）を失わないよう、セル実体は差し替えない
                    assertTrue(beforePaste == target);
                    done.countDown();
                });
        assertTrue(done.await(8, TimeUnit.SECONDS));
        assertEquals(1, applied.get());
    }

    @Test
    void modelColumnIndexFromTablePosition_usesSpreadsheetColumnIndex() throws Exception {
        CountDownLatch done = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    GridBase grid = new GridBase(2, 4);
                    SpreadsheetView view = new SpreadsheetView(grid);
                    SpreadsheetColumn column = view.getColumns().get(3);
                    TableColumn<?, ?> inner = SpreadsheetTabularSupport.innerTableColumnOf(column);
                    TablePosition<?, ?> pos = new TablePosition<>(null, 1, inner);
                    assertEquals(3, SpreadsheetTabularSupport.modelColumnIndexFromTablePosition(view, pos));
                    done.countDown();
                });
        assertTrue(done.await(5, TimeUnit.SECONDS));
    }

    @Test
    void parseSpreadsheetTsv_splitsRowsAndColumns() {
        List<List<String>> table =
                SpreadsheetTabularSupport.parseSpreadsheetTsv("OP\nAS\n\nOP");
        assertEquals(List.of(List.of("OP"), List.of("AS"), List.of(""), List.of("OP")), table);
    }

    @Test
    void parseSpreadsheetTsv_handlesTabsAndCrlf() {
        List<List<String>> table =
                SpreadsheetTabularSupport.parseSpreadsheetTsv("a\tb\r\nc\td");
        assertEquals(List.of(List.of("a", "b"), List.of("c", "d")), table);
    }

    @Test
    void parseSpreadsheetTsv_emptyYieldsEmpty() {
        assertTrue(SpreadsheetTabularSupport.parseSpreadsheetTsv(null).isEmpty());
        assertTrue(SpreadsheetTabularSupport.parseSpreadsheetTsv("").isEmpty());
    }

    @Test
    void spreadsheetNativeDataFormatMime_isSpreadsheetView() {
        assertEquals("SpreadsheetView", SpreadsheetTabularSupport.SPREADSHEET_NATIVE_CLIPBOARD_MIME);
    }
}
