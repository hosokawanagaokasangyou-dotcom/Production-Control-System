package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class EnvVarsAutoInitializationTest {

    @Test
    void shouldRun_whenCheckCompletedAndPending() {
        assertTrue(EnvVarsAutoInitialization.shouldRun(true, true, false, false));
    }

    @Test
    void shouldRun_falseUntilStartupCheckCompletes() {
        assertFalse(
                EnvVarsAutoInitialization.shouldRun(false, true, false, false),
                "起動時照合が終わる前は自動初期化しない");
    }

    @Test
    void shouldRun_falseWhenNotPending() {
        assertFalse(EnvVarsAutoInitialization.shouldRun(true, false, false, false));
    }

    @Test
    void shouldRun_falseForGuestOrResetInProgress() {
        assertFalse(EnvVarsAutoInitialization.shouldRun(true, true, true, false));
        assertFalse(EnvVarsAutoInitialization.shouldRun(true, true, false, true));
    }
}
