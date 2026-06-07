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
    void applyRemoteStartupProgram_writesRemoteAppSettings(@TempDir Path tmp) throws Exception {
        Path rdp =
                tmp.resolve("factory.rdp");
        Files.writeString(
                rdp,
                "screen mode id:i:2\r\nremoteapplicationmode:i:0\r\nremoteapplicationprogram:s:\r\n"
                        + "remoteapplicationcmdline:s:\r\nalternate shell:s:\r\n",
                StandardCharsets.UTF_16LE);

        assertFalse(
                RdpProfileEditor.applyRemoteStartupProgram(
                        rdp, "C:\\Windows\\System32\\notepad.exe", "test.txt"));

        String text = Files.readString(rdp, StandardCharsets.UTF_16LE);
        assertTrue(text.contains("remoteapplicationmode:i:1"));
        assertTrue(text.contains("remoteapplicationprogram:s:C:\\Windows\\System32\\notepad.exe"));
        assertTrue(text.contains("remoteapplicationcmdline:s:test.txt"));
        assertTrue(text.contains("alternate shell:s:rdpinit.exe"));
        assertTrue(text.contains("disableremoteappcheck:i:1"));
    }

    @Test
    void applyRemoteStartupProgram_clearsWhenBlank(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("signed.rdp");
        Files.writeString(
                rdp,
                "remoteapplicationmode:i:1\r\nremoteapplicationprogram:s:C:\\app.exe\r\n"
                        + "signature:s:ABC\r\n",
                StandardCharsets.UTF_16LE);

        assertTrue(RdpProfileEditor.applyRemoteStartupProgram(rdp, "", ""));

        String text = Files.readString(rdp, StandardCharsets.UTF_16LE);
        assertTrue(text.contains("remoteapplicationmode:i:0"));
        assertTrue(text.contains("remoteapplicationprogram:s:"));
        assertFalse(text.toLowerCase().contains("signature:s:"));
    }

}
