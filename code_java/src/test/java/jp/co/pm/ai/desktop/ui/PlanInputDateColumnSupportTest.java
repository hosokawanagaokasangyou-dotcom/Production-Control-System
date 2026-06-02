package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class PlanInputDateColumnSupportTest {

    @Test
    void editableDateColumn_includesRawInputDateOnly() {
        assertTrue(PlanInputDateColumnSupport.isEditableDateColumn("原反投入日"));
        assertFalse(PlanInputDateColumnSupport.isEditableDateColumn("原反投入日_上書き"));
    }

    @Test
    void editableDateColumn_rejectsSpeedOverride() {
        assertFalse(PlanInputDateColumnSupport.isEditableDateColumn("加工速度_上書き"));
    }

    @Test
    void editableDateTimeColumn_includesDispatchableOnly() {
        assertTrue(PlanInputDateColumnSupport.isEditableDateTimeColumn("配台可能日時"));
        assertFalse(PlanInputDateColumnSupport.isEditableDateTimeColumn("配台可能日時_上書き"));
    }
}
