package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class FactorySiteSwitchTimingTest {

    @Test
    void formatsPhaseAndTotalWithNonNegativeMilliseconds() {
        FactorySiteSwitchTiming timing = new FactorySiteSwitchTiming(1_000_000_000L);

        assertEquals(
                "[factory-timing] phase=restore-env workMs=250 elapsedMs=500",
                timing.phaseLine("restore-env", 1_250_000_000L, 1_500_000_000L));
        assertEquals(
                "[factory-timing] totalMs=750",
                timing.totalLine(1_750_000_000L));
    }

    @Test
    void rejectsBlankPhaseAndNegativeMeasurement() {
        FactorySiteSwitchTiming timing = new FactorySiteSwitchTiming(1_000_000_000L);

        assertTrue(timing.phaseLine(" ", 1_000_000_000L, 1_100_000_000L).isEmpty());
        assertTrue(timing.phaseLine("connect", 1_100_000_000L, 1_000_000_000L).isEmpty());
        assertTrue(timing.totalLine(900_000_000L).isEmpty());
    }
}
