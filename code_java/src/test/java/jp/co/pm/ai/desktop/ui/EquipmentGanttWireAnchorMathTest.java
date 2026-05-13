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

    @Test
    void personBadgeRadialAngles_fullCircle_evenStep() {
        double[] a1 = EquipmentGanttWireAnchorMath.personBadgeRadialAnglesRad(1);
        assertEquals(1, a1.length);
        assertEquals(-Math.PI / 2, a1[0], 1e-9);

        double[] a4 = EquipmentGanttWireAnchorMath.personBadgeRadialAnglesRad(4);
        assertEquals(4, a4.length);
        assertEquals(-Math.PI / 2, a4[0], 1e-9);
        double step = Math.PI / 2;
        for (int i = 1; i < 4; i++) {
            assertEquals(a4[i - 1] + step, a4[i], 1e-9);
        }

        double[] a2 = EquipmentGanttWireAnchorMath.personBadgeRadialAnglesRad(2);
        assertEquals(-Math.PI / 2, a2[0], 1e-9);
        assertEquals(Math.PI / 2, a2[1], 1e-9);
    }
}
