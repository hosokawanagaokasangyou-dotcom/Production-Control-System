package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class AladdinProcessingPlanSourceReloaderTest {

    @Test
    void reloadNewestFromDiskAndSaveShapedJson_writesShapedJson(@TempDir Path tempDir) throws Exception {
        Path sourceDir = tempDir.resolve("task-input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(sourceDir);
        Files.createDirectories(outputDir);

        Path csv =
                sourceDir.resolve("aladdin-plan.csv");
        Files.writeString(
                csv,
                "機械名,依頼NO,工程名,2026/07/07\n"
                        + "M1,T001,工程A,10\n");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        sourceDir.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        outputDir.toString());

        AladdinProcessingPlanSourceReloader.ReloadResult result =
                AladdinProcessingPlanSourceReloader.reloadNewestFromDiskAndSaveShapedJson(ui);

        assertEquals(csv, result.sourceFile());
        assertTrue(Files.isRegularFile(result.shapedJsonPath()));
        assertEquals(1, result.rowCount());
        assertTrue(result.columnCount() >= 4);

        JsonTableIo.ArrayTable saved = JsonTableIo.loadArrayTable(result.shapedJsonPath());
        assertTrue(saved.columns().contains("機械名"));
        assertEquals("T001", saved.rows().getFirst().get(saved.columns().indexOf("依頼NO")));
    }

    @Test
    void reloadNewestFromDiskAndSaveShapedJson_missingDirThrows(@TempDir Path tempDir) {
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        tempDir.resolve("missing").toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        tempDir.resolve("output").toString());

        assertThrows(
                java.io.IOException.class,
                () -> AladdinProcessingPlanSourceReloader.reloadNewestFromDiskAndSaveShapedJson(ui));
    }
}
