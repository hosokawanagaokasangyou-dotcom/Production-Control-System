package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.Optional;

import org.junit.jupiter.api.Test;

class PlanInputDateColumnSupportTest {

    @Test
    void editableDateColumn_detectsOverrideAndBaseColumns() {
        assertTrue(PlanInputDateColumnSupport.isEditableDateColumn("原反投入日_上書き"));
        assertTrue(PlanInputDateColumnSupport.isEditableDateColumn("回答納期"));
        assertTrue(PlanInputDateColumnSupport.isEditableDateColumn("指定納期"));
        assertFalse(PlanInputDateColumnSupport.isEditableDateColumn("（元）原反投入日"));
        assertFalse(PlanInputDateColumnSupport.isEditableDateColumn("原反投入日_試行前"));
        assertFalse(PlanInputDateColumnSupport.isEditableDateColumn("加工速度_上書き"));
    }

    @Test
    void parseAndFormat_roundTrip() {
        Optional<LocalDate> d = PlanInputDateColumnSupport.parseCellValue("2026/6/15 00:00:00");
        assertTrue(d.isPresent());
        assertEquals(LocalDate.of(2026, 6, 15), d.get());
        assertEquals("2026/6/15", PlanInputDateColumnSupport.formatCellValue(d.get()));
    }

    @Test
    void formatCellValue_emptyWhenNull() {
        assertEquals("", PlanInputDateColumnSupport.formatCellValue(null));
    }
}
