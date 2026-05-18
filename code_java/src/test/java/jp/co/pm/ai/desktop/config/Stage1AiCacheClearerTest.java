package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage1AiCacheClearerTest {

    @TempDir
    Path temp;

    @Test
    void clearBeforeStage1Run_deletesFileUnderCodeJsonDir() throws Exception {
        Path repo = temp.resolve("repo");
        Path code = repo.resolve("code");
        Path python = code.resolve("python");
        Path json = code.resolve("json");
        Files.createDirectories(python);
        Files.createDirectories(json);
        Files.writeString(python.resolve("task_extract_stage1.py"), "# stub\n");
        Path cache = json.resolve(Stage1AiCacheClearer.AI_REMARKS_CACHE_FILENAME);
        Files.writeString(cache, "{}");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        repo.toString(),
                        AppPaths.KEY_PM_AI_CODE_PYTHON_DIR,
                        python.toString());

        assertTrue(Files.isRegularFile(cache));
        List<String> logs = Stage1AiCacheClearer.clearBeforeStage1Run(ui);
        assertTrue(logs.stream().anyMatch(l -> l.contains("削除:")));
        assertTrue(!Files.exists(cache));
    }
}
