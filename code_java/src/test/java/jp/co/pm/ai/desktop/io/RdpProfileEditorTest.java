package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RdpProfileEditorTest {

    @Test
    void applyDesktopDisplay_windowedMode(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("factory.rdp");
        Files.writeString(rdp, "screen mode id:i:2\r\ndesktopwidth:i:800\r\n", StandardCharsets.UTF_16LE);

        assertFalse(RdpProfileEditor.applyDesktopDisplay(rdp, 1280, 800, false));

        String text = Files.readString(rdp, StandardCharsets.UTF_16LE);
        assertTrue(text.contains("desktopwidth:i:1280"));
        assertTrue(text.contains("desktopheight:i:800"));
        assertTrue(text.contains("screen mode id:i:1"));
    }

    @Test
    void applyDesktopSize_updatesUtf16LeProfile(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("factory.rdp");
        Files.writeString(rdp, "screen mode id:i:1\r\ndesktopwidth:i:800\r\n", StandardCharsets.UTF_16LE);

        assertFalse(RdpProfileEditor.applyDesktopSize(rdp, 1650, 1050));

        String text = Files.readString(rdp, StandardCharsets.UTF_16LE);
        assertTrue(text.contains("desktopwidth:i:1650"));
        assertTrue(text.contains("desktopheight:i:1050"));
        assertTrue(text.contains("screen mode id:i:2"));
    }

    @Test
    void applyDesktopSize_removesSignatureLine(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("signed.rdp");
        Files.writeString(rdp, "desktopwidth:i:800\r\nsignature:s:ABC\r\n", StandardCharsets.UTF_16LE);

        assertTrue(RdpProfileEditor.applyDesktopSize(rdp, 1650, 1050));

        String text = Files.readString(rdp, StandardCharsets.UTF_16LE);
        assertFalse(text.toLowerCase().contains("signature:s:"));
        assertTrue(text.contains("desktopwidth:i:1650"));
    }

    @Test
    void applyRemoteStartupProgram_writesAlternateShell(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("factory.rdp");
        Files.writeString(
                rdp,
                "screen mode id:i:2\r\nremoteapplicationmode:i:0\r\nremoteapplicationprogram:s:\r\n"
                        + "remoteapplicationcmdline:s:\r\nalternate shell:s:\r\n",
                StandardCharsets.UTF_16LE);

        assertFalse(
                RdpProfileEditor.applyRemoteStartupProgram(
                        rdp, "C:\\Windows\\System32\\notepad.exe", "test.txt"));

        String text = Files.readString(rdp, StandardCharsets.UTF_16LE);
        assertTrue(text.contains("remoteapplicationmode:i:0"));
        assertTrue(text.contains("alternate shell:s:C:\\Windows\\System32\\notepad.exe test.txt"));
        assertTrue(text.contains("shell working directory:s:C:\\Windows\\System32"));
        assertTrue(text.contains("remoteapplicationprogram:s:"));
    }

    @Test
    void applyRemoteStartupProgram_clearsWhenBlank(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("signed.rdp");
        Files.writeString(
                rdp,
                "alternate shell:s:C:\\app.exe\r\nsignature:s:ABC\r\n",
                StandardCharsets.UTF_16LE);

        assertTrue(RdpProfileEditor.applyRemoteStartupProgram(rdp, "", ""));

        String text = Files.readString(rdp, StandardCharsets.UTF_16LE);
        assertTrue(text.contains("alternate shell:s:"));
        assertFalse(text.toLowerCase().contains("signature:s:"));
    }
}
