package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class DispatchInteractiveStage3QtyDisplayTest {

    @Test
    void format_twoLinesWhenPlanAndActualAfterStage3() {
        String s =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        0, 100, 80, true, 1e-3, false);
        String[] lines = s.split("\n", -1);
        assertEquals(3, lines.length);
        assertEquals("", lines[0]);
        assertEquals(DispatchInteractiveTabController.LABEL_STAGE3_PLAN + "100", lines[1]);
        assertEquals(DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL + "80", lines[2]);
    }

    @Test
    void format_singleLineWhenFlagEnabled() {
        String s =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        0, 100, 80, true, 1e-3, true);
        assertFalse(s.contains("\n"));
        assertTrue(
                s.contains(
                        DispatchInteractiveTabController.LABEL_STAGE3_PLAN
                                + "100 "
                                + DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL
                                + "80"));
    }

    @Test
    void format_stage2LabelBeforeStage3() {
        assertEquals(
                DispatchInteractiveTabController.LABEL_STAGE2_PLAN + "50",
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        0, 50, 0, false, 1e-3, false));
    }

    @Test
    void format_actualOnlyAfterStage3() {
        String s =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        0, 0, 40, true, 1e-3, false);
        String[] lines = s.split("\n", -1);
        assertEquals(3, lines.length);
        assertEquals("", lines[0]);
        assertEquals("", lines[1]);
        assertEquals(DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL + "40", lines[2]);
    }

    @Test
    void format_aladdinAndDispatchBeforeStage3() {
        String s =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        900, 50, 0, false, 1e-3, false);
        assertTrue(s.startsWith(DispatchInteractiveTabController.LABEL_ALADDIN_PLAN + "900"));
        assertTrue(s.contains(DispatchInteractiveTabController.LABEL_STAGE2_PLAN + "50"));
    }

    @Test
    void format_stage3Revised_fixedThreeRows() {
        String s =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        4000, 7600, 7600, true, 1e-3, false, true);
        String[] lines = s.split("\n", -1);
        assertEquals(3, lines.length);
        assertEquals(DispatchInteractiveTabController.LABEL_ALADDIN_PLAN + "4000", lines[0]);
        assertEquals("", lines[1]);
        assertEquals(DispatchInteractiveTabController.LABEL_STAGE3_REVISED + "7600", lines[2]);
    }

    @Test
    void format_aladdinDispatchAndActualAfterStage3() {
        String s =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        900, 100, 80, true, 1e-3, false);
        String[] lines = s.split("\n", -1);
        assertEquals(3, lines.length);
        assertEquals(DispatchInteractiveTabController.LABEL_ALADDIN_PLAN + "900", lines[0]);
        assertEquals(DispatchInteractiveTabController.LABEL_STAGE3_PLAN + "100", lines[1]);
        assertEquals(DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL + "80", lines[2]);
    }

    @Test
    void fixedSlots_orderAndVisibility() {
        List<DispatchInteractiveTabController.Stage3QtyLineSlot> slots =
                DispatchInteractiveTabController.buildStage3QtyFixedLineSlots(
                        4000, 7600, 7600, false, 1e-3);
        assertEquals(3, slots.size());
        assertTrue(slots.get(0).lineText().startsWith(DispatchInteractiveTabController.LABEL_ALADDIN_PLAN));
        assertTrue(slots.get(1).lineText().startsWith(DispatchInteractiveTabController.LABEL_STAGE3_PLAN));
        assertTrue(slots.get(2).lineText().startsWith(DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL));
    }

    @Test
    void dateQtyLineFilter_hidesSelectedTypes() {
        List<DispatchInteractiveTabController.Stage3QtyLineSlot> slots =
                DispatchInteractiveTabController.buildStage3QtyFixedLineSlots(
                        4000, 7600, 7600, false, 1e-3);
        var filter =
                new jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence
                        .DispatchInteractiveDateQtyLineFilterPrefs(true, false, true);
        List<DispatchInteractiveTabController.Stage3QtyLineSlot> filtered =
                DispatchInteractiveTabController.applyDateQtyLineFilterToSlots(slots, filter);
        assertTrue(filtered.get(0).visible());
        assertFalse(filtered.get(1).visible());
        assertTrue(filtered.get(2).visible());
        assertTrue(filtered.get(2).lineText().startsWith(DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL));
    }

    @Test
    void dateQtyLineFilter_textMultiline() {
        String raw =
                DispatchInteractiveTabController.formatDispatchPlanActualQtyDisplay(
                        900, 100, 80, true, 1e-3, false);
        var filter =
                new jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence
                        .DispatchInteractiveDateQtyLineFilterPrefs(false, true, false);
        String filtered =
                DispatchInteractiveTabController.filterDispatchQtyDisplayText(raw, filter, false);
        assertFalse(filtered.contains(DispatchInteractiveTabController.LABEL_ALADDIN_PLAN));
        assertTrue(filtered.contains(DispatchInteractiveTabController.LABEL_STAGE3_PLAN + "100"));
        assertFalse(filtered.contains(DispatchInteractiveTabController.LABEL_STAGE3_ACTUAL));
    }
}
