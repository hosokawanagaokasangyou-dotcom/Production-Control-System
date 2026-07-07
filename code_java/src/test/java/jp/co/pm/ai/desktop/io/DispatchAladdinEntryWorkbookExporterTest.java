package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class DispatchAladdinEntryWorkbookExporterTest {

    @TempDir
    Path tempDir;

    @Test
    void pruneKeepsAtMostMaxGenerationsDeletingOldest() throws IOException {
        int total = DispatchAladdinEntryWorkbookExporter.MAX_GENERATIONS_PER_USER + 5;
        Instant base = Instant.parse("2026-07-01T00:00:00Z");
        for (int i = 0; i < total; i++) {
            Path f = tempDir.resolve(String.format("アラジン入力用_配台計画_%03d.xlsx", i));
            Files.writeString(f, "x");
            Files.setLastModifiedTime(f, FileTime.from(base.plusSeconds(i * 60L)));
        }

        DispatchAladdinEntryWorkbookExporter.pruneGenerations(tempDir);

        try (var stream = Files.list(tempDir)) {
            assertEquals(
                    DispatchAladdinEntryWorkbookExporter.MAX_GENERATIONS_PER_USER,
                    stream.filter(Files::isRegularFile).count());
        }
        // 最古の5件が削除され、新しい側が残る
        assertFalse(Files.exists(tempDir.resolve("アラジン入力用_配台計画_000.xlsx")));
        assertFalse(Files.exists(tempDir.resolve("アラジン入力用_配台計画_004.xlsx")));
        assertTrue(Files.exists(tempDir.resolve("アラジン入力用_配台計画_005.xlsx")));
        assertTrue(
                Files.exists(
                        tempDir.resolve(
                                String.format("アラジン入力用_配台計画_%03d.xlsx", total - 1))));
    }

    @Test
    void pruneIgnoresNonXlsxFiles() throws IOException {
        for (int i = 0; i < DispatchAladdinEntryWorkbookExporter.MAX_GENERATIONS_PER_USER; i++) {
            Files.writeString(tempDir.resolve("gen_" + i + ".xlsx"), "x");
        }
        Path other = tempDir.resolve("readme.txt");
        Files.writeString(other, "keep");

        DispatchAladdinEntryWorkbookExporter.pruneGenerations(tempDir);

        assertTrue(Files.exists(other));
        try (var stream = Files.list(tempDir)) {
            assertEquals(
                    DispatchAladdinEntryWorkbookExporter.MAX_GENERATIONS_PER_USER + 1,
                    stream.filter(Files::isRegularFile).count());
        }
    }
}
