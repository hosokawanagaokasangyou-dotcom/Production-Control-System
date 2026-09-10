package jp.co.pm.ai.desktop.io.conflict;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class FileContentFingerprintTest {

    @TempDir Path dir;

    @Test
    void absentPath_returnsAbsentSentinel() throws Exception {
        Path missing = dir.resolve("no-such.json");
        assertEquals(FileContentFingerprint.ABSENT, FileContentFingerprint.sha256Hex(missing));
    }

    @Test
    void sameBytes_sameHash() throws Exception {
        Path f = dir.resolve("a.json");
        Files.writeString(f, "{\"x\":1}", StandardCharsets.UTF_8);
        String h1 = FileContentFingerprint.sha256Hex(f);
        String h2 = FileContentFingerprint.sha256Hex(f);
        assertEquals(h1, h2);
        assertTrue(h1.matches("[0-9a-f]{64}"));
    }

    @Test
    void changedBytes_differentHash() throws Exception {
        Path f = dir.resolve("a.json");
        Files.writeString(f, "{\"x\":1}", StandardCharsets.UTF_8);
        String h1 = FileContentFingerprint.sha256Hex(f);
        Files.writeString(f, "{\"x\":2}", StandardCharsets.UTF_8);
        assertNotEquals(h1, FileContentFingerprint.sha256Hex(f));
    }

    @Test
    void hashOfBytes_matchesFile() throws Exception {
        byte[] bytes = "{\"y\":3}".getBytes(StandardCharsets.UTF_8);
        Path f = dir.resolve("b.json");
        Files.write(f, bytes);
        assertEquals(FileContentFingerprint.sha256Hex(bytes), FileContentFingerprint.sha256Hex(f));
    }
}
