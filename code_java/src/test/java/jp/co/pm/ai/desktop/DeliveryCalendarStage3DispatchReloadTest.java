package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.ui.DeliveryCalendarMainCell;

class DeliveryCalendarStage3DispatchReloadTest {

    @Test
    void overlayTripleQty_setsPlanActualAndDispatch() {
        DeliveryCalendarMainCell cell =
                DeliveryCalendarViewTabController.overlayTripleQty("10", "20", "30");
        assertEquals(
                new DeliveryCalendarMainCell.TripleQty("10", "20", "30"), cell);
    }

    @Test
    void mergeTripleDispatchQty_preservesPlanAndActual() {
        DeliveryCalendarMainCell existing =
                new DeliveryCalendarMainCell.TripleQty("10", "20", "5");
        DeliveryCalendarMainCell merged =
                DeliveryCalendarViewTabController.mergeTripleDispatchQty(existing, "99");
        assertEquals(
                new DeliveryCalendarMainCell.TripleQty("10", "20", "99"), merged);
    }
}
