package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.ui.DeliveryCalendarMainCellTripleQty;

class SpreadsheetTabularSupportDeliveryCalendarTripleAlignTest {

    @Test
    void stage3Mode_keepsStage3AfterOnSameSlotWhenPlanMissing() {
        DeliveryCalendarMainCellTripleQty withPlan =
                new DeliveryCalendarMainCellTripleQty("4400", "", "", "3000");
        DeliveryCalendarMainCellTripleQty stage3Only =
                new DeliveryCalendarMainCellTripleQty("", "", "", "1400");

        List<Boolean> planVis =
                SpreadsheetTabularSupport.deliveryCalendarTripleAlignedSlotVisibleForTest(
                        withPlan, true);
        List<Boolean> onlyVis =
                SpreadsheetTabularSupport.deliveryCalendarTripleAlignedSlotVisibleForTest(
                        stage3Only, true);

        assertEquals(3, planVis.size());
        assertEquals(3, onlyVis.size());
        assertTrue(planVis.get(0));
        assertFalse(planVis.get(1));
        assertTrue(planVis.get(2));
        assertFalse(onlyVis.get(0));
        assertFalse(onlyVis.get(1));
        assertTrue(onlyVis.get(2));
    }

    @Test
    void stage3Mode_includesDispatchSlotWhenPlanLineVisible() {
        DeliveryCalendarMainCellTripleQty t = new DeliveryCalendarMainCellTripleQty("10", "20", "30", "40");
        List<String> texts =
                SpreadsheetTabularSupport.deliveryCalendarTripleAlignedSlotTextsForTest(t, false);
        assertEquals(4, texts.size());
        assertTrue(texts.get(0).startsWith("(アラ計画)"));
        assertTrue(texts.get(3).startsWith("(段階3後)"));
    }
}
