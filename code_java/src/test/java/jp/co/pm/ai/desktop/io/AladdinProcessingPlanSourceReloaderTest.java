package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class AladdinProcessingPlanSourceReloaderTest {

    @Test
    void reloadNewestFromDiskAndSaveShapedJson_writesShapedJson(@TempDir Path tempDir) throws Exception {
        Path sourceDir = tempDir.resolve("task-input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(outputDir);

        Path xlsx = TestAladdinPlanXlsx.writeMinimal(sourceDir, "aladdin-plan.xlsx");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        sourceDir.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        outputDir.toString());

        AladdinProcessingPlanSourceReloader.ReloadResult result =
                AladdinProcessingPlanSourceReloader.reloadNewestFromDiskAndSaveShapedJson(ui);

        assertEquals(xlsx, result.sourceFile());
        assertTrue(Files.isRegularFile(result.shapedJsonPath()));
        assertEquals(1, result.rowCount());
        assertTrue(result.columnCount() >= 4);

        JsonTableIo.ArrayTable saved = JsonTableIo.loadArrayTable(result.shapedJsonPath());
        assertTrue(saved.columns().contains("機械名"));
        assertEquals("T001", saved.rows().getFirst().get(saved.columns().indexOf("依頼NO")));
    }

    @Test
    void reloadNewestFromDiskAndSaveShapedJson_rejectsNewerCsv(@TempDir Path tempDir) throws Exception {
        Path sourceDir = tempDir.resolve("task-input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(sourceDir);
        Files.createDirectories(outputDir);
        TestAladdinPlanXlsx.writeMinimal(sourceDir, "aladdin-plan.xlsx");
        Path csv = sourceDir.resolve("newer.csv");
        Files.writeString(csv, "a,b\n1,2");
        Files.setLastModifiedTime(
                csv,
                java.nio.file.attribute.FileTime.fromMillis(
                        Files.getLastModifiedTime(sourceDir.resolve("aladdin-plan.xlsx")).toMillis()
                                + 60_000L));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        sourceDir.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        outputDir.toString());

        IOException ex =
                assertThrows(
                        IOException.class,
                        () ->
                                AladdinProcessingPlanSourceReloader
                                        .reloadNewestFromDiskAndSaveShapedJson(ui));
        assertTrue(ex.getMessage().contains("拡張子が不正"));
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
