package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PipelineLocalResultsPolicyTest {

    @Test
    void isSharedOrUncPath_detectsUncAndFactoryShared() {
        assertTrue(PipelineLocalResultsPolicy.isSharedOrUncPathText("\\\\192.168.0.101\\共有フォルダ\\x"));
        assertTrue(
                PipelineLocalResultsPolicy.isSharedOrUncPathText(
                        AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR));
        assertTrue(
                PipelineLocalResultsPolicy.isSharedOrUncPathText(
                        AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR_M));
        assertFalse(PipelineLocalResultsPolicy.isSharedOrUncPathText("C:\\Users\\me\\output"));
    }

    @Test
    void rewritePipelineOutputEnvToLocal_rewritesSharedKeys(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("code").resolve("python");
        Files.createDirectories(code);
        Files.createFile(code.resolve("task_extract_stage1.py"));
        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        ui.put(AppPaths.KEY_PM_AI_OUTPUT_DIR, AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR);
        ui.put(
                AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR + "\\plan_input_tasks.xlsx");
        assertTrue(PipelineLocalResultsPolicy.rewritePipelineOutputEnvToLocal(ui));
        Path local = fakeRepo.resolve("output").toAbsolutePath().normalize();
        assertEquals(local.toString(), ui.get(AppPaths.KEY_PM_AI_OUTPUT_DIR));
        assertEquals(
                local.resolve(AppPaths.STAGE1_PLAN_TASKS_FILENAME).toString(),
                ui.get(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH));
    }
}
