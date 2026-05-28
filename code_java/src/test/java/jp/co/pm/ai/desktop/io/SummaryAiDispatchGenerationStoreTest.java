package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class SummaryAiDispatchGenerationStoreTest {

    @TempDir
    Path tmp;

    @BeforeEach
    void isolateStoreRoot() {
        System.setProperty("pm.ai.test.summaryGenerationRoot", tmp.resolve("generations").toString());
    }

    @AfterEach
    void clearStoreRootProperty() {
        System.clearProperty("pm.ai.test.summaryGenerationRoot");
    }

    @Test
    void archiveBeforeOverwrite_createsIndexAndWorkbook(@TempDir Path repo) throws Exception {
        Path shared = repo.resolve("shared");
        Files.createDirectories(shared);
        Path current = shared.resolve("summary.xlsx");
        Files.writeString(current, "old-content");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        repo.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        current.toString());

        var archived =
                SummaryAiDispatchGenerationStore.archiveBeforeOverwrite(
                                current, ui, "delivery-reload")
                        .orElseThrow();
        assertEquals("delivery-reload", archived.reason());

        List<SummaryAiDispatchGenerationStore.SummaryAiDispatchGenerationEntry> index =
                SummaryAiDispatchGenerationStore.loadIndex(ui);
        assertEquals(1, index.size());
        Path archivedBook =
                archived.resolveWorkbookPath(
                        SummaryAiDispatchGenerationStore.resolveGenerationsRoot(ui));
        assertTrue(Files.isRegularFile(archivedBook));
        assertEquals("old-content", Files.readString(archivedBook));
    }

    @Test
    void restoreToCurrentWorkbook_replacesCurrentAndArchivesPrior(@TempDir Path repo) throws Exception {
        Path shared = repo.resolve("shared");
        Files.createDirectories(shared);
        Path current = shared.resolve("summary.xlsx");
        Files.writeString(current, "generation-a");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        repo.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        current.toString());

        var genA =
                SummaryAiDispatchGenerationStore.archiveCurrent(ui, "snapshot-A").orElseThrow();
        Files.writeString(current, "generation-b");

        SummaryAiDispatchGenerationStore.restoreToCurrentWorkbook(genA, ui);
        assertEquals("generation-a", Files.readString(current));

        List<SummaryAiDispatchGenerationStore.SummaryAiDispatchGenerationEntry> index =
                SummaryAiDispatchGenerationStore.loadIndex(ui);
        assertTrue(index.size() >= 2);
        assertTrue(index.stream().anyMatch(e -> "restore-backup".equals(e.reason())));
    }

    @Test
    void trimRemovesOldestWhenExceedingMax() throws Exception {
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, tmp.toString());
        Path current = tmp.resolve("book.xlsx");
        for (int i = 0; i < SummaryAiDispatchGenerationStore.DEFAULT_MAX_GENERATIONS + 3; i++) {
            Files.writeString(current, "v" + i);
            SummaryAiDispatchGenerationStore.archiveBeforeOverwrite(
                    current, ui, "export-" + i);
        }
        List<SummaryAiDispatchGenerationStore.SummaryAiDispatchGenerationEntry> index =
                SummaryAiDispatchGenerationStore.loadIndex(ui);
        assertEquals(SummaryAiDispatchGenerationStore.DEFAULT_MAX_GENERATIONS, index.size());
        assertFalse(index.stream().anyMatch(e -> "export-0".equals(e.reason())));
        assertTrue(index.stream().anyMatch(e -> "export-32".equals(e.reason())));
    }
}
