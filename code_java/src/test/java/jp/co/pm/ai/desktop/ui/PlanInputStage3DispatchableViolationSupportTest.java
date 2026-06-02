package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class PlanInputStage3DispatchableViolationSupportTest {

    @Test
    void violation_whenRawInput1245AfterDispatchableShiftStart() {
        List<String> headers =
                List.of(
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE,
                        PlanInputStage3DispatchableViolationSupport.COL_DISPATCHABLE_DATETIME);
        List<String> row = List.of("2026/6/10", "2026/6/10 8:45");
        assertTrue(
                PlanInputStage3DispatchableViolationSupport.isDispatchableBeforeRawInputLimit(
                        headers, row));
    }

    @Test
    void noViolation_whenDispatchableAfter1245SameDay() {
        List<String> headers =
                List.of(
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE,
                        PlanInputStage3DispatchableViolationSupport.COL_DISPATCHABLE_DATETIME);
        List<String> row = List.of("2026/6/10", "2026/6/10 13:00");
        assertFalse(
                PlanInputStage3DispatchableViolationSupport.isDispatchableBeforeRawInputLimit(
                        headers, row));
    }

    @Test
    void noViolation_whenDispatchableEquals1245() {
        List<String> headers =
                List.of(
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE,
                        PlanInputStage3DispatchableViolationSupport.COL_DISPATCHABLE_DATETIME);
        List<String> row = List.of("2026/6/10", "2026/6/10 12:45");
        assertFalse(
                PlanInputStage3DispatchableViolationSupport.isDispatchableBeforeRawInputLimit(
                        headers, row));
    }

    @Test
    void usesBaseRawInputDateColumn() {
        List<String> headers =
                List.of(
                        PlanInputRawInputDateShift.COL_RAW_INPUT_DATE,
                        PlanInputStage3DispatchableViolationSupport.COL_DISPATCHABLE_DATETIME);
        List<String> row = List.of("2026/6/10", "2026/6/10 8:45");
        assertTrue(
                PlanInputStage3DispatchableViolationSupport.isDispatchableBeforeRawInputLimit(
                        headers, row));
    }

    @Test
    void warningMessage_mentionsReviewRawInputDate() {
        String msg = PlanInputStage3DispatchableViolationSupport.warningMessage(2);
        assertTrue(msg.contains("2 件"));
        assertTrue(msg.contains("原反投入日"));
        assertTrue(msg.contains("12:45"));
    }
}
