package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class DispatchAladdinPlanAlignerTest {

    @Test
    void alignRow_movesToAladdinDaysWithRollUnit() {
        var result =
                DispatchAladdinPlanAligner.alignRow(
                        new DispatchAladdinPlanAligner.RowInput(
                                new double[] {300, 300, 0},
                                new double[] {500, 0, 500},
                                300,
                                false));
        assertTrue(result.changed());
        assertArrayEquals(new double[] {300, 0, 300}, result.newByDayIndex(), 1e-9);
        assertEquals(1, result.rollMoves());
    }

    @Test
    void alignRow_usesRollUnitWhenAladdinShowsConvertedQty() {
        var result =
                DispatchAladdinPlanAligner.alignRow(
                        new DispatchAladdinPlanAligner.RowInput(
                                new double[] {300, 300, 0},
                                new double[] {20, 0, 0},
                                300,
                                true));
        assertTrue(result.changed());
        assertArrayEquals(new double[] {600, 0, 0}, result.newByDayIndex(), 1e-9);
    }

    @Test
    void alignRow_distributesRollsAcrossConvertedQtyDays() {
        var result =
                DispatchAladdinPlanAligner.alignRow(
                        new DispatchAladdinPlanAligner.RowInput(
                                new double[] {600, 0, 0},
                                new double[] {20, 20, 0},
                                300,
                                true));
        assertTrue(result.changed());
        assertArrayEquals(new double[] {300, 300, 0}, result.newByDayIndex(), 1e-9);
    }

    @Test
    void alignRow_skipsWhenNoAladdinTarget() {
        var result =
                DispatchAladdinPlanAligner.alignRow(
                        new DispatchAladdinPlanAligner.RowInput(
                                new double[] {300, 0}, new double[] {0, 0}, 300, false));
        assertFalse(result.changed());
        assertArrayEquals(new double[] {300, 0}, result.newByDayIndex(), 1e-9);
    }

    @Test
    void allocateRollsByWeight_usesLargestRemainder() {
        int[] rolls =
                DispatchAladdinPlanAligner.allocateRollsByWeight(
                        2, new double[] {500, 100, 0});
        assertArrayEquals(new int[] {2, 0, 0}, rolls);
    }
}
