package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

final class SpreadsheetTabularSupportEffectiveRollUnitTest {

    @Test
    void effectiveRollUnit_800_over_95_becomes_100() {
        assertEquals(
                100.0,
                SpreadsheetTabularSupport.effectiveRollUnitMForDispatchTaskSimulator(800.0, 95.0),
                1e-9);
    }

    @Test
    void effectiveRollUnit_integer_rolls_returns_sheet_unit() {
        assertEquals(
                100.0,
                SpreadsheetTabularSupport.effectiveRollUnitMForDispatchTaskSimulator(800.0, 100.0),
                1e-9);
    }
}
