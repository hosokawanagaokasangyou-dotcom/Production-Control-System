package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class WorkspaceCacheArchiveStoreTest {

    @TempDir
    Path temp;

    @org.junit.jupiter.api.BeforeEach
    void useTempArchiveRoot() {
        System.setProperty(
                "pm.ai.test.workspaceCacheArchiveRoot", temp.resolve("archives").toString());
    }

    @org.junit.jupiter.api.AfterEach
    void clearTempArchiveRoot() {
        System.clearProperty("pm.ai.test.workspaceCacheArchiveRoot");
    }

    @Test
    void archiveAndRestore_roundTrip() throws Exception {
        Path repo = temp.resolve("repo");
        Path code = repo.resolve("code");
        Path python = code.resolve("python");
        Path json = code.resolve("json");
        Path output = code.resolve("output");
        Files.createDirectories(python);
        Files.createDirectories(json);
        Files.createDirectories(output);
        Files.writeString(python.resolve("task_extract_stage1.py"), "# stub\n");

        Path cache = json.resolve(Stage1AiCacheClearer.AI_REMARKS_CACHE_FILENAME);
        Files.writeString(cache, "{\"k\":1}");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        repo.toString(),
                        AppPaths.KEY_PM_AI_CODE_PYTHON_DIR,
                        python.toString(),
                        AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
                        output.toString());

        WorkspaceCacheArchiveStore.WorkspaceCacheArchiveEntry entry =
                WorkspaceCacheArchiveStore.archiveDiskCaches(ui, "test", "manual-archive");
        assertNotNull(entry);
        assertTrue(entry.fileCount() >= 1);

        Files.delete(cache);
        assertFalse(Files.exists(cache));

        List<String> logs = WorkspaceCacheArchiveStore.restoreToWorkspace(entry, ui);
        assertTrue(logs.stream().anyMatch(l -> l.contains("復元:")));
        assertTrue(Files.isRegularFile(cache));
    }
}
