package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class NetworkSourceDirResolverTest {

    @Test
    void taskInputSourceDir_reachable_whenUnderRepo(@TempDir Path fakeRepo) throws Exception {
        Path src = fakeRepo.resolve("in").resolve("src");
        Files.createDirectories(src);
        Files.writeString(src.resolve("a.csv"), "x");
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        src.toString());
        assertTrue(NetworkSourceDirResolver.isTaskInputSourceDirReachable(ui));
    }

    @Test
    void taskInputSourceDir_unreachable_whenDirMissing() {
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, "/no/such/path/pm_ai_task_input_" + System.nanoTime());
        assertFalse(NetworkSourceDirResolver.isTaskInputSourceDirReachable(ui));
    }
}
