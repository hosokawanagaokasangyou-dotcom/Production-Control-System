package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

class PlanInputRawInputDateShiftTest {

    @Test
    void applyMinusOneDay_shiftsBaseColumn() {
        List<String> headers =
                List.of("依頼NO", PlanInputRawInputDateShift.COL_RAW_INPUT_DATE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("A", "2026/6/15"));
        int n = PlanInputRawInputDateShift.applyMinusOneDayToAllRows(headers, rows);
        assertEquals(1, n);
        assertEquals("2026/6/14", rows.get(0).get(1));
    }

    @Test
    void applyMinusOneDay_skipsUnparseableRows() {
        List<String> headers = List.of(PlanInputRawInputDateShift.COL_RAW_INPUT_DATE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList(""));
        rows.add(FXCollections.observableArrayList("2026/6/2"));
        assertEquals(1, PlanInputRawInputDateShift.applyMinusOneDayToAllRows(headers, rows));
        assertEquals("2026/6/1", rows.get(1).get(0));
    }

    @Test
    void applyMinusOneDay_missingBaseColumn() {
        List<String> headers = List.of("依頼NO");
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("A"));
        assertEquals(
                PlanInputRawInputDateShift.MISSING_RAW_INPUT_DATE_COLUMN,
                PlanInputRawInputDateShift.applyMinusOneDayToAllRows(headers, rows));
    }
}
