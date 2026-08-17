package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class OperatorActionLogStoreTest {

    @Test
    void resolveDailyFile_usesSharedDataOperatorAndDate(@TempDir Path tempDir) {
        Path shared = tempDir.resolve("shared");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, shared.toString());
        LocalDate day = LocalDate.of(2026, 8, 17);

        Path file = OperatorActionLogStore.resolveDailyFile(ui, "テスト太郎", day);

        assertEquals(
                AppPaths.resolveOperatorActionLogRoot(ui)
                        .resolve(OperatorUserPaths.sanitizeOperatorDirName("テスト太郎"))
                        .resolve("2026-08-17.ndjson"),
                file);
    }

    @Test
    void append_writesNdjsonAndReadReturnsNewestFirst(@TempDir Path tempDir) throws Exception {
        Map<String, String> ui = testUi(tempDir);

        assertTrue(
                OperatorActionLogStore.append(
                        ui, "テスト太郎", "stage2_complete", "ok", "段階2完了"));
        assertTrue(
                OperatorActionLogStore.append(
                        ui, "テスト太郎", "close_warning", "shown", "差異 1件"));

        List<OperatorActionLogStore.Entry> entries =
                OperatorActionLogStore.readOperator(ui, "テスト太郎", Instant.now());

        assertEquals(2, entries.size());
        assertEquals("close_warning", entries.get(0).action());
        assertEquals("shown", entries.get(0).result());
        assertEquals("stage2_complete", entries.get(1).action());
        assertEquals("テスト太郎", entries.get(0).operator());
        assertFalse(entries.get(0).ts().isBlank());
    }

    @Test
    void prune_deletesNdjsonOlderThan90Days(@TempDir Path tempDir) throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Path dir =
                OperatorActionLogStore.resolveOperatorDir(ui, "tester");
        Files.createDirectories(dir);
        Path oldFile = dir.resolve("2026-01-01.ndjson");
        Path keepFile = dir.resolve("2026-08-01.ndjson");
        Files.writeString(oldFile, "{}\n");
        Files.writeString(keepFile, "{}\n");
        Instant now = Instant.parse("2026-08-17T00:00:00Z");
        Files.setLastModifiedTime(
                oldFile,
                java.nio.file.attribute.FileTime.from(now.minus(java.time.Duration.ofDays(91))));
        Files.setLastModifiedTime(
                keepFile,
                java.nio.file.attribute.FileTime.from(now.minus(java.time.Duration.ofDays(10))));

        int removed = OperatorActionLogStore.pruneOlderThan(dir, now, 90);

        assertEquals(1, removed);
        assertFalse(Files.exists(oldFile));
        assertTrue(Files.exists(keepFile));
    }

    @Test
    void listOperators_returnsDirectoryNames(@TempDir Path tempDir) throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Files.createDirectories(OperatorActionLogStore.resolveOperatorDir(ui, "alpha"));
        Files.createDirectories(OperatorActionLogStore.resolveOperatorDir(ui, "beta"));

        List<String> names = OperatorActionLogStore.listOperators(ui);

        assertEquals(List.of("alpha", "beta"), names);
    }

    private static Map<String, String> testUi(Path tempDir) {
        return Map.of(
                AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                tempDir.resolve("shared").toString());
    }
}
