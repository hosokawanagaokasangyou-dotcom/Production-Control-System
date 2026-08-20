package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class MasterDispatchSheetsGenerationBackupTest {

    @TempDir Path tmp;

    @Test
    void backupExisting_copiesFileIntoSiblingHistoryAndKeepsOriginal() throws Exception {
        Path json = tmp.resolve("master-dispatch-sheets.json");
        Files.writeString(json, "{\"v\":1}\n", StandardCharsets.UTF_8);

        Optional<Path> backup = MasterDispatchSheetsGenerationBackup.backupExisting(json);

        assertTrue(backup.isPresent());
        assertTrue(Files.isRegularFile(backup.get()));
        assertEquals("{\"v\":1}\n", Files.readString(backup.get(), StandardCharsets.UTF_8));
        assertEquals("{\"v\":1}\n", Files.readString(json, StandardCharsets.UTF_8));
        assertEquals(
                MasterDispatchSheetsGenerationBackup.historyDirFor(json),
                backup.get().getParent());
        assertTrue(backup.get().getFileName().toString().endsWith("master-dispatch-sheets.json"));
    }

    @Test
    void backupExisting_missingFile_returnsEmpty() throws Exception {
        Path missing = tmp.resolve("master.xlsm");
        assertTrue(MasterDispatchSheetsGenerationBackup.backupExisting(missing).isEmpty());
        assertFalse(Files.exists(MasterDispatchSheetsGenerationBackup.historyDirFor(missing)));
    }

    @Test
    void backupExisting_prunesOldestBeyondMaxGenerations() throws Exception {
        Path xlsm = tmp.resolve("master.xlsm");
        Files.writeString(xlsm, "current", StandardCharsets.UTF_8);
        Path history = MasterDispatchSheetsGenerationBackup.historyDirFor(xlsm);
        Files.createDirectories(history);
        Instant base = Instant.parse("2026-01-01T00:00:00Z");
        for (int i = 0; i < MasterDispatchSheetsGenerationBackup.MAX_GENERATIONS + 2; i++) {
            Path old = history.resolve(String.format("20260101-00000%02d_master.xlsm", i));
            Files.writeString(old, "old-" + i, StandardCharsets.UTF_8);
            Files.setLastModifiedTime(old, FileTime.from(base.plusSeconds(i)));
        }

        MasterDispatchSheetsGenerationBackup.backupExisting(xlsm);

        List<Path> kept;
        try (Stream<Path> stream = Files.list(history)) {
            kept =
                    stream.filter(Files::isRegularFile)
                            .sorted()
                            .toList();
        }
        assertEquals(MasterDispatchSheetsGenerationBackup.MAX_GENERATIONS, kept.size());
        assertFalse(kept.get(0).getFileName().toString().startsWith("20260101-0000000_"));
    }
}
