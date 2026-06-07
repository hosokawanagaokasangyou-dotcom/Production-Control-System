package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class RdpLaunchDisplaySettingsTest {

    @Test
    void resolveDefaults_windowed1280x800() {
        assertFalse(RdpLaunchDisplaySettings.resolveFullScreen(Map.of()));
        assertEquals(1280, RdpLaunchDisplaySettings.resolveWidth(Map.of()));
        assertEquals(800, RdpLaunchDisplaySettings.resolveHeight(Map.of()));
        assertEquals("1280 x 800（ウィンドウ）", RdpLaunchDisplaySettings.formatSummary(Map.of()));
    }

    @Test
    void resolveFullScreen_fromEnvValues() {
        assertTrue(
                RdpLaunchDisplaySettings.resolveFullScreen(
                        Map.of(AppPaths.KEY_PM_AI_RDP_FULLSCREEN, "1")));
        assertFalse(
                RdpLaunchDisplaySettings.resolveFullScreen(
                        Map.of(AppPaths.KEY_PM_AI_RDP_FULLSCREEN, "0")));
    }

    @Test
    void applyToProfile_writesWindowedSize(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("factory.rdp");
        Files.writeString(rdp, "screen mode id:i:2\r\n", StandardCharsets.UTF_16LE);

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_RDP_FULLSCREEN,
                        "0",
                        AppPaths.KEY_PM_AI_RDP_DESKTOP_WIDTH,
                        "1366",
                        AppPaths.KEY_PM_AI_RDP_DESKTOP_HEIGHT,
                        "768");

        RdpLaunchDisplaySettings.applyToProfile(rdp, ui);

        String text = Files.readString(rdp, StandardCharsets.UTF_16LE);
        assertTrue(text.contains("screen mode id:i:1"));
        assertTrue(text.contains("desktopwidth:i:1366"));
        assertTrue(text.contains("desktopheight:i:768"));
    }
}
