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
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;

class SummaryAiDispatchGenerationStoreTest {

    private static final Map<String, String> USER_SUNADA =
            Map.of(AppPaths.KEY_PM_AI_OPERATOR_USER, "砂田");

    private static final Map<String, String> USER_FURUYA =
            Map.of(AppPaths.KEY_PM_AI_OPERATOR_USER, "古家");

    @TempDir
    Path tmp;

    @BeforeEach
    void isolateStoreRoot() throws Exception {
        System.setProperty("pm.ai.test.summaryGenerationRoot", tmp.resolve("generations").toString());
        System.setProperty(
                "pm.ai.test.factoryOperatorUserStore", tmp.resolve("operators.json").toString());
        FactoryOperatorUserStore.resetStoreForTests();
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
    }

    @AfterEach
    void clearStoreRootProperty() {
        System.clearProperty("pm.ai.test.summaryGenerationRoot");
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        FactoryOperatorUserStore.clearSessionOperatorName();
    }

    private static Map<String, String> ui(Path repoRoot, Map<String, String> user) {
        Map<String, String> base =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        repoRoot.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        repoRoot.resolve("shared").resolve("summary.xlsx").toString());
        if (user == null || user.isEmpty()) {
            return base;
        }
        var merged = new java.util.HashMap<>(base);
        merged.putAll(user);
        return Map.copyOf(merged);
    }

    @Test
    void archiveBeforeOverwrite_createsIndexAndWorkbook(@TempDir Path repo) throws Exception {
        Path shared = repo.resolve("shared");
        Files.createDirectories(shared);
        Path current = shared.resolve("summary.xlsx");
        Files.writeString(current, "old-content");

        Map<String, String> env = ui(repo, USER_SUNADA);
        var archived =
                SummaryAiDispatchGenerationStore.archiveBeforeOverwrite(
                                current, env, "delivery-reload")
                        .orElseThrow();
        assertEquals("delivery-reload", archived.reason());
        assertEquals("砂田", archived.operatorUser());

        List<SummaryAiDispatchGenerationStore.SummaryAiDispatchGenerationEntry> index =
                SummaryAiDispatchGenerationStore.loadIndex(env);
        assertEquals(1, index.size());
        Path archivedBook =
                archived.resolveWorkbookPath(
                        SummaryAiDispatchGenerationStore.resolveUserGenerationsRoot(env));
        assertTrue(Files.isRegularFile(archivedBook));
        assertEquals("old-content", Files.readString(archivedBook));
    }

    @Test
    void restoreToCurrentWorkbook_replacesCurrentAndArchivesPrior(@TempDir Path repo) throws Exception {
        Path shared = repo.resolve("shared");
        Files.createDirectories(shared);
        Path current = shared.resolve("summary.xlsx");
        Files.writeString(current, "generation-a");

        Map<String, String> env = ui(repo, USER_SUNADA);
        var genA = SummaryAiDispatchGenerationStore.archiveCurrent(env, "snapshot-A").orElseThrow();
        Files.writeString(current, "generation-b");

        SummaryAiDispatchGenerationStore.restoreToCurrentWorkbook(genA, env);
        assertEquals("generation-a", Files.readString(current));

        List<SummaryAiDispatchGenerationStore.SummaryAiDispatchGenerationEntry> index =
                SummaryAiDispatchGenerationStore.loadIndex(env);
        assertTrue(index.size() >= 2);
        assertTrue(index.stream().anyMatch(e -> "restore-backup".equals(e.reason())));
    }

    @Test
    void trimRemovesOldestWhenExceedingMaxPerUser() throws Exception {
        Map<String, String> env = ui(tmp, USER_SUNADA);
        Path current = tmp.resolve("book.xlsx");
        for (int i = 0; i < SummaryAiDispatchGenerationStore.MAX_GENERATIONS_PER_USER + 3; i++) {
            Files.writeString(current, "v" + i);
            SummaryAiDispatchGenerationStore.archiveBeforeOverwrite(
                    current, env, "export-" + i);
        }
        List<SummaryAiDispatchGenerationStore.SummaryAiDispatchGenerationEntry> index =
                SummaryAiDispatchGenerationStore.loadIndex(env);
        assertEquals(SummaryAiDispatchGenerationStore.MAX_GENERATIONS_PER_USER, index.size());
        assertFalse(index.stream().anyMatch(e -> "export-0".equals(e.reason())));
        assertTrue(index.stream().anyMatch(e -> "export-12".equals(e.reason())));
    }

    @Test
    void generationsAreScopedPerOperator(@TempDir Path repo) throws Exception {
        Path shared = repo.resolve("shared");
        Files.createDirectories(shared);
        Path current = shared.resolve("summary.xlsx");
        Files.writeString(current, "sunada-old");
        SummaryAiDispatchGenerationStore.archiveBeforeOverwrite(
                current, ui(repo, USER_SUNADA), "sunada");

        Files.writeString(current, "furuya-old");
        SummaryAiDispatchGenerationStore.archiveBeforeOverwrite(
                current, ui(repo, USER_FURUYA), "furuya");

        assertEquals(1, SummaryAiDispatchGenerationStore.loadIndex(ui(repo, USER_SUNADA)).size());
        assertEquals(1, SummaryAiDispatchGenerationStore.loadIndex(ui(repo, USER_FURUYA)).size());
        assertEquals(
                "sunada",
                SummaryAiDispatchGenerationStore.loadIndex(ui(repo, USER_SUNADA))
                        .get(0)
                        .reason());
        assertEquals(
                "furuya",
                SummaryAiDispatchGenerationStore.loadIndex(ui(repo, USER_FURUYA))
                        .get(0)
                        .reason());
    }
}
