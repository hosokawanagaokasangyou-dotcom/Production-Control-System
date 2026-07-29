package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RequestFormPipelineCheckAladdinPlanRevisionTest {

    @Test
    void registerAladdinPlanSourceReload_detectsPathOrTimestampChange(@TempDir Path tempDir)
            throws Exception {
        Path file = tempDir.resolve("plan.xlsx");
        Files.writeString(file, "v1");

        RequestFormPipelineCheckTabController controller =
                new RequestFormPipelineCheckTabController();
        assertTrue(controller.registerAladdinPlanSourceReload(file));
        assertFalse(controller.registerAladdinPlanSourceReload(file));

        Thread.sleep(5);
        Files.writeString(file, "v2");
        assertTrue(controller.registerAladdinPlanSourceReload(file));
    }

    @Test
    void isAladdinPlanSourceNewerThanLastScan_whenRevisionChanges(@TempDir Path tempDir)
            throws Exception {
        Path sourceDir = tempDir.resolve("task-input");
        Files.createDirectories(sourceDir);
        Path plan = sourceDir.resolve("plan.xlsx");
        Files.writeString(plan, "v1");

        Map<String, String> ui = Map.of("PM_AI_TASK_INPUT_SOURCE_DIR", sourceDir.toString());
        RequestFormPipelineCheckTabController controller =
                new RequestFormPipelineCheckTabController();
        assertFalse(controller.isAladdinPlanSourceNewerThanLastScan(ui));

        controller.captureLastScannedAladdinPlanRevisionForTest(ui);
        assertFalse(controller.isAladdinPlanSourceNewerThanLastScan(ui));

        Thread.sleep(5);
        Files.writeString(plan, "v2");
        assertTrue(controller.isAladdinPlanSourceNewerThanLastScan(ui));
    }

    @Test
    void aladdinPlanSourceRevisionKey_usesPathAndLastModified(@TempDir Path tempDir)
            throws Exception {
        Path file = tempDir.resolve("plan.xlsx");
        Files.writeString(file, "data");
        String key = RequestFormPipelineCheckTabController.aladdinPlanSourceRevisionKey(file);
        assertTrue(key.startsWith(file.toAbsolutePath().normalize() + "|"));
        assertEquals(key, RequestFormPipelineCheckTabController.aladdinPlanSourceRevisionKey(file));
    }
}
