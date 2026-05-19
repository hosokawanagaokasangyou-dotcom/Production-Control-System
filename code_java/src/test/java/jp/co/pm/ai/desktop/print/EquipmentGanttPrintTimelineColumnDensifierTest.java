package jp.co.pm.ai.desktop.print;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalTime;
import java.util.List;

import org.junit.jupiter.api.Test;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

class EquipmentGanttPrintTimelineColumnDensifierTest {

    @Test
    void densify_keepsHalfOpenRegularShiftRange() {
        List<String> cols =
                List.of("日付", "機械名", "工程名", "8:00", "8:10", "8:20", "17:00");
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(
                FXCollections.observableArrayList(
                        "", "M1", "EC", "a", "b", "c", "z"));

        EquipmentGanttPrintTableData out =
                EquipmentGanttPrintTimelineColumnDensifier.densify(
                        cols,
                        rows,
                        List.of(List.of("A", "B", "C", "Z")),
                        LocalTime.of(8, 10),
                        LocalTime.of(8, 30));

        assertEquals(
                List.of("日付", "機械名", "工程名", "8:10", "8:20"),
                out.columns());
        assertEquals(List.of("", "M1", "EC", "b", "c"), out.rows().get(0));
        assertEquals(List.of("B", "C"), out.badgeSlotRows().get(0));
    }

    @Test
    void densify_invalidRange_returnsInputUnchanged() {
        List<String> cols = List.of("日付", "8:00", "8:10");
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("", "x", "y"));
        EquipmentGanttPrintTableData out =
                EquipmentGanttPrintTimelineColumnDensifier.densify(
                        cols, rows, null, LocalTime.of(9, 0), LocalTime.of(8, 0));
        assertEquals(cols, out.columns());
        assertEquals(1, out.rows().size());
        assertEquals(rows.get(0), out.rows().get(0));
    }

    @Test
    void slotInHalfOpenRange_excludesEndBoundary() {
        LocalTime start = LocalTime.of(8, 25);
        LocalTime end = LocalTime.of(17, 0);
        assertTrue(
                EquipmentGanttPrintTimelineColumnDensifier.slotInHalfOpenRange(
                        LocalTime.of(16, 50), start, end));
        assertTrue(
                !EquipmentGanttPrintTimelineColumnDensifier.slotInHalfOpenRange(
                        LocalTime.of(17, 0), start, end));
    }
}
