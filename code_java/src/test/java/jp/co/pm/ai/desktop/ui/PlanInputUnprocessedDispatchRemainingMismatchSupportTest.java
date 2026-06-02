package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class PlanInputUnprocessedDispatchRemainingMismatchSupportTest {

    @Test
    void mismatch_whenUnprocessedPositiveAndDispatchRemainingDiffers() {
        List<String> headers =
                List.of(
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.COL_TASK_ID,
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.COL_UNPROCESSED,
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.COL_DISPATCH_REMAINING);
        List<String> row = List.of("JR260602", "6090", "5000");
        assertTrue(
                PlanInputUnprocessedDispatchRemainingMismatchSupport.isMismatch(headers, row));
    }

    @Test
    void noMismatch_whenUnprocessedZero() {
        List<String> headers =
                List.of(
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.COL_UNPROCESSED,
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.COL_DISPATCH_REMAINING);
        List<String> row = List.of("0", "5000");
        assertFalse(
                PlanInputUnprocessedDispatchRemainingMismatchSupport.isMismatch(headers, row));
    }

    @Test
    void noMismatch_whenValuesEqual() {
        List<String> headers =
                List.of(
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.COL_UNPROCESSED,
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.COL_DISPATCH_REMAINING);
        List<String> row = List.of("800", "800");
        assertFalse(
                PlanInputUnprocessedDispatchRemainingMismatchSupport.isMismatch(headers, row));
    }

    @Test
    void mismatch_whenDispatchRemainingEmpty() {
        List<String> headers =
                List.of(
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.COL_UNPROCESSED,
                        PlanInputUnprocessedDispatchRemainingMismatchSupport.COL_DISPATCH_REMAINING);
        List<String> row = List.of("800", "");
        assertTrue(
                PlanInputUnprocessedDispatchRemainingMismatchSupport.isMismatch(headers, row));
    }

    @Test
    void warningMessage_usesTaskId() {
        String msg =
                PlanInputUnprocessedDispatchRemainingMismatchSupport.warningMessage(
                        List.of("JR260602"));
        assertTrue(msg.contains("JR260602"));
        assertTrue(msg.contains("未加工と配台使用残数量が異なります"));
        assertTrue(msg.contains("手動修正"));
    }

    @Test
    void warningMessage_joinsMultipleTaskIds() {
        String msg =
                PlanInputUnprocessedDispatchRemainingMismatchSupport.warningMessage(
                        List.of("JR260602", "JR260603"));
        assertEquals(
                "JR260602、JR260603 の未加工と配台使用残数量が異なります、手動修正してください。",
                msg);
    }
}
