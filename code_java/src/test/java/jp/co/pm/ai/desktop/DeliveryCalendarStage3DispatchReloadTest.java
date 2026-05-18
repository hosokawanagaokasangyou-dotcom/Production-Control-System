package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.ui.DeliveryCalendarMainCell;

class DeliveryCalendarStage3DispatchReloadTest {

    @Test
    void overlayTripleQty_setsPlanActualDispatchAndStage3() {
        DeliveryCalendarMainCell cell =
                DeliveryCalendarViewTabController.overlayTripleQty("10", "20", "30", "40");
        assertEquals(
                new DeliveryCalendarMainCell.TripleQty("10", "20", "30", "40"), cell);
    }

    @Test
    void mergeTripleDispatchAndStage3Qty_preservesPlanAndActual() {
        DeliveryCalendarMainCell existing =
                new DeliveryCalendarMainCell.TripleQty("10", "20", "5", "7");
        DeliveryCalendarMainCell merged =
                DeliveryCalendarViewTabController.mergeTripleDispatchAndStage3Qty(
                        existing, "99", "88");
        assertEquals(
                new DeliveryCalendarMainCell.TripleQty("10", "20", "99", "88"), merged);
    }
}
