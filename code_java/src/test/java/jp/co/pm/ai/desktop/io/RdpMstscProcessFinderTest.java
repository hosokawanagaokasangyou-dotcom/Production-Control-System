package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RdpMstscProcessFinderTest {

    @Test
    void commandLineRefersToProfile_matchesFullPath() {
        Path profile = Path.of("C:\\Users\\0585\\OneDrive\\ドキュメント\\Default.rdp");
        String cmd =
                "\"C:\\Windows\\System32\\mstsc.exe\" \"C:\\Users\\0585\\OneDrive\\ドキュメント\\Default.rdp\"";
        assertTrue(RdpMstscProcessFinder.commandLineRefersToProfile(cmd, profile));
    }

    @Test
    void commandLineRefersToProfile_matchesFileNameOnly() {
        Path profile = Path.of("C:\\Users\\0585\\OneDrive\\ドキュメント\\Default.rdp");
        String cmd = "mstsc.exe Default.rdp";
        assertTrue(RdpMstscProcessFinder.commandLineRefersToProfile(cmd, profile));
    }

    @Test
    void commandLineRefersToProfile_rejectsOtherProfile() {
        Path profile = Path.of("C:\\data\\Default.rdp");
        String cmd = "mstsc.exe C:\\data\\Other.rdp";
        assertFalse(RdpMstscProcessFinder.commandLineRefersToProfile(cmd, profile));
    }

    @Test
    void readPidMarkerFile_readsAsciiPid(@TempDir Path temp) throws Exception {
        Path marker = temp.resolve("rdp-mstsc.pid");
        Files.writeString(marker, "12345", StandardCharsets.US_ASCII);
        assertTrue(RdpMstscProcessFinder.readPidMarkerFile(marker).isPresent());
        assertTrue(RdpMstscProcessFinder.readPidMarkerFile(marker).getAsLong() == 12345L);
    }
}
