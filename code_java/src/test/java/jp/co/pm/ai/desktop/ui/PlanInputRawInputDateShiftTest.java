package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

class PlanInputRawInputDateShiftTest {

    @Test
    void applyMinusOneDay_writesOverrideFromBaseWhenOverrideEmpty() {
        List<String> headers =
                List.of(
                        "依頼NO",
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE,
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE_OVERRIDE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("A", "2026/6/15", ""));
        int n = PlanInputRawInputDateShift.applyMinusOneDayToAllOverrides(headers, rows);
        assertEquals(1, n);
        assertEquals("2026/6/14", rows.get(0).get(2));
    }

    @Test
    void applyMinusOneDay_shiftsExistingOverride() {
        List<String> headers =
                List.of(
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE,
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE_OVERRIDE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("2026/6/15", "2026/6/13"));
        int n = PlanInputRawInputDateShift.applyMinusOneDayToAllOverrides(headers, rows);
        assertEquals(1, n);
        assertEquals("2026/6/12", rows.get(0).get(1));
    }

    @Test
    void applyMinusOneDay_skipsUnparseableRows() {
        List<String> headers =
                List.of(
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE,
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE_OVERRIDE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("", ""));
        rows.add(FXCollections.observableArrayList("2026/6/2", ""));
        assertEquals(1, PlanInputRawInputDateShift.applyMinusOneDayToAllOverrides(headers, rows));
        assertEquals("2026/6/1", rows.get(1).get(1));
    }

    @Test
    void applyMinusOneDay_missingOverrideColumn() {
        List<String> headers = List.of(PlanInputRawInputDateShift.COL_RAW_INPUT_DATE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("2026/6/2"));
        assertEquals(
                PlanInputRawInputDateShift.MISSING_OVERRIDE_COLUMN,
                PlanInputRawInputDateShift.applyMinusOneDayToAllOverrides(headers, rows));
    }

    @Test
    void clearAllOverrides_clearsNonEmptyCellsOnly() {
        List<String> headers =
                List.of(
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE,
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE_OVERRIDE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("2026/6/15", "2026/6/14"));
        rows.add(FXCollections.observableArrayList("2026/6/2", ""));
        assertEquals(1, PlanInputRawInputDateShift.clearAllOverrides(headers, rows));
        assertEquals("", rows.get(0).get(1));
        assertEquals("", rows.get(1).get(1));
    }

    @Test
    void clearAllOverrides_missingOverrideColumn() {
        List<String> headers = List.of(PlanInputRawInputDateShift.COL_RAW_INPUT_DATE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("2026/6/2", "2026/6/1"));
        assertEquals(
                PlanInputRawInputDateShift.MISSING_OVERRIDE_COLUMN,
                PlanInputRawInputDateShift.clearAllOverrides(headers, rows));
    }
}
