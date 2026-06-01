package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.time.LocalTime;
import java.util.Optional;

import org.junit.jupiter.api.Test;

class DispatchAladdinPlanAlignerTest {

    @Test
    void resolveAlignFromDate_beforeRegularShiftStart_includesToday() {
        LocalDate op = LocalDate.of(2026, 6, 2);
        assertEquals(
                op,
                DispatchAladdinPlanAligner.resolveAlignFromDate(
                        op, LocalTime.of(7, 30), Optional.of(LocalTime.of(8, 0))));
    }

    @Test
    void resolveAlignFromDate_atOrAfterRegularShiftStart_startsTomorrow() {
        LocalDate op = LocalDate.of(2026, 6, 2);
        assertEquals(
                op.plusDays(1),
                DispatchAladdinPlanAligner.resolveAlignFromDate(
                        op, LocalTime.of(8, 0), Optional.of(LocalTime.of(8, 0))));
        assertEquals(
                op.plusDays(1),
                DispatchAladdinPlanAligner.resolveAlignFromDate(
                        op, LocalTime.of(9, 0), Optional.of(LocalTime.of(8, 0))));
    }

    @Test
    void resolveAlignFromDate_withoutRegularShiftStart_startsTomorrow() {
        LocalDate op = LocalDate.of(2026, 6, 2);
        assertEquals(
                op.plusDays(1),
                DispatchAladdinPlanAligner.resolveAlignFromDate(
                        op, LocalTime.of(7, 0), Optional.empty()));
    }

    @Test
    void alignRowFromDayIndex_preservesPrefixAndAlignsSuffix() {
        var result =
                DispatchAladdinPlanAligner.alignRowFromDayIndex(
                        new DispatchAladdinPlanAligner.RowInput(
                                new double[] {300, 300, 0},
                                new double[] {500, 0, 500},
                                300,
                                false),
                        1);
        assertTrue(result.changed());
        assertArrayEquals(new double[] {300, 0, 300}, result.newByDayIndex(), 1e-9);
        assertEquals(1, result.rollMoves());
    }

    @Test
    void alignRowFromDayIndex_whenNoFutureDays_unchanged() {
        var result =
                DispatchAladdinPlanAligner.alignRowFromDayIndex(
                        new DispatchAladdinPlanAligner.RowInput(
                                new double[] {300, 300}, new double[] {0, 500}, 300, false),
                        2);
        assertFalse(result.changed());
        assertArrayEquals(new double[] {300, 300}, result.newByDayIndex(), 1e-9);
    }

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
