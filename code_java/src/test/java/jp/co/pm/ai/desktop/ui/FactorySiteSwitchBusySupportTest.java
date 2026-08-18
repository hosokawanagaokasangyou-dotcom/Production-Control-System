package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class FactorySiteSwitchBusySupportTest {

    @Test
    void keepBusyDialogForPostSwitchTabLoad_onlyWhenFactorySwitchLoadStarted() {
        assertTrue(FactorySiteSwitchBusySupport.keepBusyDialogForPostSwitchTabLoad(false, true));
        assertFalse(FactorySiteSwitchBusySupport.keepBusyDialogForPostSwitchTabLoad(true, true));
        assertFalse(FactorySiteSwitchBusySupport.keepBusyDialogForPostSwitchTabLoad(false, false));
        assertFalse(FactorySiteSwitchBusySupport.keepBusyDialogForPostSwitchTabLoad(true, false));
    }

    @Test
    void resolveTabLoadStatus_usesMessageOrDefault() {
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD,
                FactorySiteSwitchBusySupport.resolveTabLoadStatus(""));
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD,
                FactorySiteSwitchBusySupport.resolveTabLoadStatus(null));
        assertEquals(
                "起動後読込 (1/6): リモートデスクトップ…",
                FactorySiteSwitchBusySupport.resolveTabLoadStatus("起動後読込 (1/6): リモートデスクトップ…"));
    }

    @Test
    void centerX_centersChildOverOwner() {
        assertEquals(400.0, FactorySiteSwitchBusySupport.centerX(100.0, 800.0, 200.0), 0.001);
        assertEquals(240.0, FactorySiteSwitchBusySupport.centerY(40.0, 600.0, 200.0), 0.001);
    }
}
