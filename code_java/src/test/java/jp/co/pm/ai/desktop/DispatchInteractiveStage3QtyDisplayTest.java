package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class DispatchInteractiveStage3QtyDisplayTest {

    @Test
    void format_twoLinesWhenPlanAndActualAfterStage3() {
        String s =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        100, 80, true, 1e-3, false);
        assertTrue(s.contains(DispatchInteractiveTabController.LABEL_STAGE3_PLAN + "100"));
        assertTrue(s.contains(DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL + "80"));
        assertTrue(
                s.indexOf(DispatchInteractiveTabController.LABEL_STAGE3_PLAN)
                        < s.indexOf(DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL));
        assertTrue(s.contains("\n"));
    }

    @Test
    void format_singleLineWhenFlagEnabled() {
        String s =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        100, 80, true, 1e-3, true);
        assertFalse(s.contains("\n"));
        assertTrue(
                s.contains(
                        DispatchInteractiveTabController.LABEL_STAGE3_PLAN
                                + "100 "
                                + DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL
                                + "80"));
    }

    @Test
    void format_plainQtyBeforeStage3() {
        assertEquals(
                "50",
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        50, 0, false, 1e-3, false));
    }

    @Test
    void format_actualOnlyAfterStage3() {
        String s =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        0, 40, true, 1e-3, false);
        assertFalse(s.contains(DispatchInteractiveTabController.LABEL_STAGE3_PLAN));
        assertEquals(
                DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL + "40", s.trim());
    }
}
