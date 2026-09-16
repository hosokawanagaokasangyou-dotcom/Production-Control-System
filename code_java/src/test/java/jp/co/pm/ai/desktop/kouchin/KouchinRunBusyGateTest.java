package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class KouchinRunBusyGateTest {

    @Test
    void overlayIsOrOfStageDispatchAndKouchin() {
        assertFalse(KouchinRunBusyGate.overlayVisible(false, false, false));
        assertTrue(KouchinRunBusyGate.overlayVisible(false, false, true));
        assertTrue(KouchinRunBusyGate.overlayVisible(true, false, false));
        assertTrue(KouchinRunBusyGate.overlayVisible(false, true, false));
    }

    @Test
    void cannotStartWhenStageOrKouchinBusy() {
        assertTrue(KouchinRunBusyGate.canStart(false, false));
        assertFalse(KouchinRunBusyGate.canStart(true, false));
        assertFalse(KouchinRunBusyGate.canStart(false, true));
    }

    @Test
    void cancelTargetsKouchinOnlyWhenKouchinBusy() {
        assertTrue(KouchinRunBusyGate.cancelTargetsKouchin(true));
        assertFalse(KouchinRunBusyGate.cancelTargetsKouchin(false));
    }
}
