package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage1AiCacheClearerTest {

    @TempDir
    Path temp;

    @Test
    void clearBeforeStage1Run_deletesAiAndWorkspaceCaches() throws Exception {
        Path repo = temp.resolve("repo");
        Path code = repo.resolve("code");
        Path python = code.resolve("python");
        Path json = code.resolve("json");
        Path output = code.resolve("output");
        Files.createDirectories(python);
        Files.createDirectories(json);
        Files.createDirectories(output);
        Files.writeString(python.resolve("task_extract_stage1.py"), "# stub\n");

        Path aiCache = json.resolve(Stage1AiCacheClearer.AI_REMARKS_CACHE_FILENAME);
        Path dispatchJson = output.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME);
        Path aladdinShaped = output.resolve(AppPaths.SHAPED_ALADDIN_PLAN_JSON_BASENAME);
        Files.writeString(aiCache, "{}");
        Files.writeString(dispatchJson, "{}");
        Files.writeString(aladdinShaped, "{}");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        repo.toString(),
                        AppPaths.KEY_PM_AI_CODE_PYTHON_DIR,
                        python.toString());

        assertTrue(Stage1AiCacheClearer.hasAnyExistingDiskCache(ui));
        assertTrue(Stage1AiCacheClearer.hasExistingWorkspaceShapedCache(ui));

        Stage1AiCacheClearer.ClearResult result = Stage1AiCacheClearer.clearBeforeStage1Run(ui);
        assertTrue(result.anyDeleted());
        assertTrue(result.detailLines().stream().anyMatch(l -> l.contains("削除:")));
        assertFalse(Files.exists(aiCache));
        assertFalse(Files.exists(dispatchJson));
        assertFalse(Files.exists(aladdinShaped));
        assertFalse(Stage1AiCacheClearer.hasAnyExistingDiskCache(ui));
    }
}
