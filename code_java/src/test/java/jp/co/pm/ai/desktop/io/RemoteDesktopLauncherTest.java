package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RemoteDesktopLauncherTest {

    @Test
    void validateRdpProfile_rejectsMissingFile(@TempDir Path tmp) {
        Path missing = tmp.resolve("missing.rdp");
        IOException ex =
                assertThrows(IOException.class, () -> RemoteDesktopLauncher.validateRdpProfile(missing));
        assertTrue(ex.getMessage().contains("見つかりません"));
    }

    @Test
    void validateRdpProfile_rejectsNonRdpExtension(@TempDir Path tmp) throws Exception {
        Path txt = tmp.resolve("profile.txt");
        Files.writeString(txt, "dummy");
        IOException ex =
                assertThrows(IOException.class, () -> RemoteDesktopLauncher.validateRdpProfile(txt));
        assertTrue(ex.getMessage().contains(".rdp"));
    }

    @Test
    void validateRdpProfile_acceptsRdpFile(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("factory.rdp");
        Files.writeString(rdp, "screen mode id:i:2");
        Path validated = RemoteDesktopLauncher.validateRdpProfile(rdp);
        assertEquals(rdp.toAbsolutePath().normalize(), validated);
    }

    @Test
    void ensureLaunchableRdpProfile_rejectsDefaultRdp(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("Default.rdp");
        Files.writeString(rdp, "screen mode id:i:2");
        IOException ex =
                assertThrows(
                        IOException.class, () -> RemoteDesktopLauncher.ensureLaunchableRdpProfile(rdp));
        assertTrue(ex.getMessage().contains("Default.rdp"));
    }

    @Test
    void ensureLaunchableRdpProfile_acceptsSignedProfile(@TempDir Path tmp) throws Exception {
        Path rdp = tmp.resolve("Default.pm-ai-signed.rdp");
        Files.writeString(rdp, "screen mode id:i:2");
        RemoteDesktopLauncher.ensureLaunchableRdpProfile(rdp);
    }
}
