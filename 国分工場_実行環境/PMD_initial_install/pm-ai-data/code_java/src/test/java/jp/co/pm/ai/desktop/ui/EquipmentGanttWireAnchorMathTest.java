package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class EquipmentGanttWireAnchorMathTest {

    @Test
    void barAnchorCenter_singleSlot_matchesMidpoint() {
        double sw = 9.0;
        double zoom = 1.0;
        assertEquals(4.5, EquipmentGanttWireAnchorMath.barAnchorCenterX(sw, zoom, 0, 0), 1e-9);
    }

    @Test
    void barAnchorCenter_twoSlots_midBetweenSlots() {
        double sw = 10.0;
        double zoom = 1.0;
        assertEquals(10.0, EquipmentGanttWireAnchorMath.barAnchorCenterX(sw, zoom, 0, 1), 1e-9);
    }

    @Test
    void barAnchorCenterY_matchesBandMid() {
        double pad = 8.0;
        double rowH = 26.0;
        double zoom = 1.0;
        double innerTop = 3 * zoom;
        double barTop = pad + innerTop;
        double barH = rowH - 2 * innerTop;
        double expected = barTop + barH / 2;
        assertEquals(
                expected,
                EquipmentGanttWireAnchorMath.barAnchorCenterY(pad, rowH, zoom),
                1e-9);
    }
}
