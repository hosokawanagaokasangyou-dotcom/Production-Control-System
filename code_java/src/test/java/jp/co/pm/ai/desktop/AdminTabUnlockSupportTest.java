package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;

class AdminTabUnlockSupportTest {

    @Test
    void verifyFallbackPassword_acceptsConfiguredPasswordAndStrippedInput() {
        assertTrue(AdminTabUnlockSupport.verifyFallbackPassword(FactoryOperatorUserStore.ADMIN_TAB_PASSWORD));
        assertTrue(AdminTabUnlockSupport.verifyFallbackPassword(" nagaoka1 "));
    }

    @Test
    void verifyFallbackPassword_rejectsMismatchEmptyAndNull() {
        assertFalse(AdminTabUnlockSupport.verifyFallbackPassword("nagaoka123"));
        assertFalse(AdminTabUnlockSupport.verifyFallbackPassword("nagaoka"));
        assertFalse(AdminTabUnlockSupport.verifyFallbackPassword(""));
        assertFalse(AdminTabUnlockSupport.verifyFallbackPassword(null));
        assertFalse(AdminTabUnlockSupport.verifyFallbackPassword("NAGAOKA1"));
    }
}
