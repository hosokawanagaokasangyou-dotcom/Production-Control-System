package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class AladdinSystemDispatchDisplayQtyTest {

    @Test
    void displayQtyForDay_whenConvBelowRawRoll_usesConvertedQty() {
        assertEquals(20.0, AladdinSystemDispatchDisplayQty.displayQtyForDay(400.0, 20.0, 400.0), 1e-9);
    }

    @Test
    void displayQtyForDay_whenConvNotBelowRawRoll_keepsDispatchQty() {
        assertEquals(400.0, AladdinSystemDispatchDisplayQty.displayQtyForDay(400.0, 500.0, 400.0), 1e-9);
    }

    @Test
    void allocateDay_capsTotalAcrossDays() {
        var ctx = new AladdinSystemDispatchDisplayQty.TaskQtyContext(20.0, 400.0);
        assertTrue(ctx.usesConvertedQtyForAladdinDisplay());
        Double cap = ctx.qtyConvM();
        var d1 = AladdinSystemDispatchDisplayQty.allocateDay(15.0, 20.0, 400.0, cap);
        assertEquals(15.0, d1.displayM(), 1e-9);
        var d2 = AladdinSystemDispatchDisplayQty.allocateDay(10.0, 20.0, 400.0, d1.remainingConvCap());
        assertEquals(5.0, d2.displayM(), 1e-9);
        var d3 = AladdinSystemDispatchDisplayQty.allocateDay(100.0, 20.0, 400.0, d2.remainingConvCap());
        assertEquals(0.0, d3.displayM(), 1e-9);
    }

    @Test
    void taskQtyContext_detectsWhenAladdinShouldShowConvertedQty() {
        assertTrue(
                new AladdinSystemDispatchDisplayQty.TaskQtyContext(20.0, 400.0)
                        .usesConvertedQtyForAladdinDisplay());
        assertFalse(
                new AladdinSystemDispatchDisplayQty.TaskQtyContext(500.0, 400.0)
                        .usesConvertedQtyForAladdinDisplay());
    }
}
