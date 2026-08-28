package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class FactorySiteSwitchTimingIntegrationTest {

    @Test
    void exposesStableCoreAndPostPhaseOrder() {
        assertEquals(
                List.of(
                        "connect",
                        "save-old-workspace",
                        "load-new-workspace",
                        "restore-env-session",
                        "refresh-request-form",
                        "refresh-pipeline",
                        "refresh-remote-toolbar",
                        "stabilize-env",
                        "match-env",
                        "finish"),
                FactorySiteSwitchTiming.corePhaseNames());
        assertEquals(
                List.of(
                        "attendance-company",
                        "attendance-member",
                        "attendance-machine",
                        "attendance-master",
                        "background-load"),
                FactorySiteSwitchTiming.postPhaseNames());
    }

    @Test
    void phaseElapsedTimeIncludesSchedulingGapButWorkTimeDoesNot() {
        FactorySiteSwitchTiming timing = new FactorySiteSwitchTiming(1_000_000_000L);

        assertEquals(
                "[factory-timing] phase=connect workMs=10 elapsedMs=110",
                timing.phaseLine("connect", 1_100_000_000L, 1_110_000_000L));
        assertEquals(
                "[factory-timing] phase=save-old-workspace workMs=20 elapsedMs=150",
                timing.phaseLine("save-old-workspace", 1_130_000_000L, 1_150_000_000L));
        assertTrue(
                timing.phaseLine("save-old-workspace", 1_150_000_000L, 1_140_000_000L)
                        .isEmpty());
    }

    @Test
    void hasOneHookForFinalTotalTiming() {
        assertDoesNotThrow(
                () ->
                        MainShellController.class.getDeclaredMethod(
                                "appendFactorySiteSwitchTotal"));
    }
}
