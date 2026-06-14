package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

class RdpPreviewSettingsTest {

    @Test
    void defaultEnabled() {
        assertTrue(RdpPreviewSettings.isPreviewInTabEnabled(Map.of()));
        assertTrue(RdpPreviewSettings.isPreviewInTabEnabled(null));
    }

    @Test
    void explicitOff() {
        assertFalse(
                RdpPreviewSettings.isPreviewInTabEnabled(
                        Map.of(RdpPreviewSettings.KEY_PM_AI_RDP_PREVIEW_IN_TAB, "0")));
    }
}
