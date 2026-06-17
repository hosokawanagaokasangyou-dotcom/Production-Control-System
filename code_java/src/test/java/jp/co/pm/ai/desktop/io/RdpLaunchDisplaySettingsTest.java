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
import jp.co.pm.ai.desktop.io.RdpLaunchDisplaySettings.LaunchDisplay;

class RdpLaunchDisplaySettingsTest {

    @Test
    void resolveDefaults_windowed1920x1080() {
        assertFalse(RdpLaunchDisplaySettings.resolveFullScreen(Map.of()));
        assertEquals(1920, RdpLaunchDisplaySettings.resolveWidth(Map.of()));
        assertEquals(1080, RdpLaunchDisplaySettings.resolveHeight(Map.of()));
        assertEquals("1920 x 1080（ウィンドウ）", RdpLaunchDisplaySettings.formatSummary(Map.of()));
    }

    @Test
    void clampMinimum270x200() {
        assertEquals(270, RdpLaunchDisplaySettings.clampWidth(100));
        assertEquals(200, RdpLaunchDisplaySettings.clampHeight(100));
        LaunchDisplay display =
                RdpLaunchDisplaySettings.resolveLaunchDisplay(
                        new RdpLaunchProfile(
                                1, "", "", "", null, null, false, 100, 100, null, null),
                        Map.of());
        assertEquals(270, display.width());
        assertEquals(200, display.height());
    }

    @Test
    void resolveLaunchDisplay_profileOverridesEnv() {
        RdpLaunchProfile profile =
                new RdpLaunchProfile(2, "t", "", "", null, null, false, 640, 480, null, null);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_RDP_DESKTOP_WIDTH,
                        "1280",
                        AppPaths.KEY_PM_AI_RDP_DESKTOP_HEIGHT,
                        "800");
        LaunchDisplay display =
                RdpLaunchDisplaySettings.resolveLaunchDisplay(profile, ui);
        assertEquals(640, display.width());
        assertEquals(480, display.height());
    }

    @Test
    void resolveLaunchDisplay_profileFullScreen() {
        RdpLaunchProfile profile =
                new RdpLaunchProfile(1, "", "", "", null, null, true, 1280, 800, null, null);
        LaunchDisplay display =
                RdpLaunchDisplaySettings.resolveLaunchDisplay(profile, Map.of());
        assertTrue(display.fullScreen());
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
