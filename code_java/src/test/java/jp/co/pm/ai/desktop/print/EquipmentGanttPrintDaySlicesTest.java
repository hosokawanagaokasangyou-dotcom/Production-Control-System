package jp.co.pm.ai.desktop.print;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

class EquipmentGanttPrintDaySlicesTest {

    @Test
    void isPrintDayBoundaryRow_bracketedDateInDateColumn() {
        List<String> cols = List.of("日付", "機械名", "工程名", "8:00");
        ObservableList<String> row =
                FXCollections.observableArrayList("【2026/5/19】", "", "", "");
        assertTrue(EquipmentGanttPrintDaySlices.isPrintDayBoundaryRow(cols, row));
    }

    @Test
    void rowIndexGroups_splitsOnBracketedBanner() {
        List<String> cols = List.of("日付", "機械名", "工程名", "8:00", "8:10");
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("【2026/5/18】", "", "", "", ""));
        rows.add(FXCollections.observableArrayList("", "M1", "EC", "a", ""));
        rows.add(FXCollections.observableArrayList("【2026/5/19】", "", "", "", ""));
        rows.add(FXCollections.observableArrayList("", "M2", "EC", "", "b"));

        List<List<Integer>> groups =
                EquipmentGanttPrintDaySlices.rowIndexGroupsOnePagePerDay(cols, rows);
        assertEquals(2, groups.size());
        assertEquals(List.of(0, 1), groups.get(0));
        assertEquals(List.of(2, 3), groups.get(1));
    }

    @Test
    void isSectionLikeRow_detectsMachineSectionMarker() {
        ObservableList<String> row =
                FXCollections.observableArrayList("■EC機 湖南", "", "", "");
        assertTrue(EquipmentGanttPrintDaySlices.isSectionLikeRow(row));
    }

    @Test
    void isSectionLikeRow_plainBracketedDateIsNotSection() {
        ObservableList<String> row =
                FXCollections.observableArrayList("【2026/5/19】", "", "", "");
        assertFalse(EquipmentGanttPrintDaySlices.isSectionLikeRow(row));
    }
}
