package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class MachineCalendarCellValuesTest {

    @Test
    void toggleEmptyToOccupied() {
        assertEquals("*", MachineCalendarCellValues.toggle(""));
        assertEquals("*", MachineCalendarCellValues.toggle(null));
    }

    @Test
    void toggleOccupiedToEmpty() {
        assertEquals("", MachineCalendarCellValues.toggle("*"));
    }

    @Test
    void shortLabel() {
        assertEquals("·", MachineCalendarCellValues.shortLabel(""));
        assertEquals("*", MachineCalendarCellValues.shortLabel("*"));
        assertEquals("abc", MachineCalendarCellValues.shortLabel("abcdef"));
    }

    @Test
    void resolvePaintModeFromAnchor() {
        assertEquals(
                MachineCalendarCellValues.OccupancyMode.OCCUPIED,
                MachineCalendarCellValues.resolvePaintModeFromAnchor(""));
        assertEquals(
                MachineCalendarCellValues.OccupancyMode.AVAILABLE,
                MachineCalendarCellValues.resolvePaintModeFromAnchor("*"));
    }

    @Test
    void isOccupied() {
        assertFalse(MachineCalendarCellValues.isOccupied(""));
        assertTrue(MachineCalendarCellValues.isOccupied("*"));
    }
}
