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

class RequestFormJuchuFileBackupStoreTest {

    @TempDir
    Path tmp;

    @BeforeEach
    void isolateStoreRoot() {
        System.setProperty(
                "pm.ai.test.requestFormJuchuBackupRoot", tmp.resolve("juchu-backups").toString());
    }

    @AfterEach
    void clearStoreRootProperty() {
        System.clearProperty("pm.ai.test.requestFormJuchuBackupRoot");
    }

    private static Map<String, String> ui(Path repoRoot) {
        return Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repoRoot.toString());
    }

    @Test
    void maybeBackupBeforeWrite_createsGeneration(@TempDir Path repo) throws Exception {
        Path juchu = repo.resolve("加工依頼書入力.xlsm");
        Files.writeString(juchu, "content-v1");
        Map<String, String> env = ui(repo);

        var backed =
                RequestFormJuchuFileBackupStore.maybeBackupBeforeWrite(
                                juchu, env, "single-transfer")
                        .orElseThrow();
        assertEquals("single-transfer", backed.reason());

        List<RequestFormJuchuFileBackupStore.RequestFormJuchuFileBackupEntry> index =
                RequestFormJuchuFileBackupStore.loadIndexForSource(env, juchu);
        assertEquals(1, index.size());
        Path archive =
                backed.resolveArchivePath(
                        RequestFormJuchuFileBackupStore.resolveSourceBackupsRoot(env, juchu));
        assertTrue(Files.isRegularFile(archive));
        assertEquals("content-v1", Files.readString(archive));
    }

    @Test
    void maybeBackupBeforeWrite_skipsWithinFifteenMinutes(@TempDir Path repo) throws Exception {
        Path juchu = repo.resolve("juchu.xlsm");
        Files.writeString(juchu, "content-v1");
        Map<String, String> env = ui(repo);

        assertTrue(
                RequestFormJuchuFileBackupStore.maybeBackupBeforeWrite(
                                juchu, env, "single-transfer")
                        .isPresent());
        Files.writeString(juchu, "content-v2");
        assertFalse(
                RequestFormJuchuFileBackupStore.maybeBackupBeforeWrite(
                                juchu, env, "bulk-transfer")
                        .isPresent());

        List<RequestFormJuchuFileBackupStore.RequestFormJuchuFileBackupEntry> index =
                RequestFormJuchuFileBackupStore.loadIndexForSource(env, juchu);
        assertEquals(1, index.size());
        assertEquals("single-transfer", index.get(0).reason());
    }

    @Test
    void restoreToSourceWorkbook_replacesTargetAndBacksUpPrior(@TempDir Path repo) throws Exception {
        Path juchu = repo.resolve("juchu.xlsm");
        Files.writeString(juchu, "content-v1");
        Map<String, String> env = ui(repo);

        var backed =
                RequestFormJuchuFileBackupStore.maybeBackupBeforeWrite(
                                juchu, env, "single-transfer")
                        .orElseThrow();

        Files.writeString(juchu, "content-broken");

        RequestFormJuchuFileBackupStore.restoreToSourceWorkbook(backed, env, juchu);

        assertEquals("content-v1", Files.readString(juchu));

        List<RequestFormJuchuFileBackupStore.RequestFormJuchuFileBackupEntry> index =
                RequestFormJuchuFileBackupStore.loadIndexForSource(env, juchu);
        assertEquals(2, index.size());
        assertEquals("pre-restore", index.get(0).reason());
        assertEquals("single-transfer", index.get(1).reason());
    }
}
