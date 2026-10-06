package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class FileChangeStampTest {

    @TempDir
    Path temp;

    @Test
    void missingFileIsAbsentAndStable() {
        Path p = temp.resolve("attendance-data.json");
        FileChangeStamp a = FileChangeStamp.read(p);
        assertFalse(a.exists());
        assertEquals(a, FileChangeStamp.read(p));
    }

    @Test
    void unchangedFileYieldsEqualStamp() throws Exception {
        Path p = temp.resolve("attendance-data.json");
        Files.writeString(p, "{\"a\":1}", StandardCharsets.UTF_8);
        assertEquals(FileChangeStamp.read(p), FileChangeStamp.read(p));
    }

    @Test
    void contentWriteWithNewMtimeChangesStamp() throws Exception {
        Path p = temp.resolve("attendance-data.json");
        Files.writeString(p, "{\"a\":1}", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(p, FileTime.fromMillis(1_000_000L));
        FileChangeStamp before = FileChangeStamp.read(p);
        Files.writeString(p, "{\"a\":2}", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(p, FileTime.fromMillis(2_000_000L));
        assertNotEquals(before, FileChangeStamp.read(p));
    }

    @Test
    void sizeChangeWithSameMtimeChangesStamp() throws Exception {
        Path p = temp.resolve("attendance-data.json");
        Files.writeString(p, "{}", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(p, FileTime.fromMillis(1_000_000L));
        FileChangeStamp before = FileChangeStamp.read(p);
        Files.writeString(p, "{\"members\":[]}", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(p, FileTime.fromMillis(1_000_000L));
        assertNotEquals(before, FileChangeStamp.read(p));
    }

    @Test
    void creationAndDeletionChangeStamp() throws Exception {
        Path p = temp.resolve("attendance-data.json");
        FileChangeStamp absent = FileChangeStamp.read(p);
        Files.writeString(p, "{}", StandardCharsets.UTF_8);
        FileChangeStamp present = FileChangeStamp.read(p);
        assertNotEquals(absent, present);
        Files.delete(p);
        assertEquals(absent, FileChangeStamp.read(p));
    }

    @Test
    void differentPathChangesStamp() throws Exception {
        Path a = temp.resolve("a.json");
        Path b = temp.resolve("b.json");
        assertNotEquals(FileChangeStamp.read(a), FileChangeStamp.read(b));
    }
}
