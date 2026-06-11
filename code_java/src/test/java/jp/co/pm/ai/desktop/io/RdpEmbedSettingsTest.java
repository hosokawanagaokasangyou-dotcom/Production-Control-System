package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

class RdpEmbedSettingsTest {

    @Test
    void embedEnabledByDefault() {
        assertTrue(RdpEmbedSettings.isEmbedInTabEnabled(Map.of()));
        assertTrue(RdpEmbedSettings.isEmbedInTabEnabled(null));
    }

    @Test
    void embedDisabledWhenZero() {
        assertFalse(
                RdpEmbedSettings.isEmbedInTabEnabled(
                        Map.of(RdpEmbedSettings.KEY_PM_AI_RDP_EMBED_IN_TAB, "0")));
    }
}
