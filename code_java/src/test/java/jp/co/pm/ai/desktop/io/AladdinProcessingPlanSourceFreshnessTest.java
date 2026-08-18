package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class AladdinProcessingPlanSourceFreshnessTest {

    @Test
    void isSavedShapedPlanIdenticalToNewestSource_trueWhenShapedMatchesSource(@TempDir Path tempDir)
            throws Exception {
        Path sourceDir = tempDir.resolve("task-input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(outputDir);
        TestAladdinPlanXlsx.writeMinimal(sourceDir, "aladdin-plan.xlsx");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        sourceDir.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        outputDir.toString());

        AladdinProcessingPlanSourceReloader.reloadNewestFromDiskAndSaveShapedJson(ui);

        assertTrue(AladdinProcessingPlanSourceFreshness.isSavedShapedPlanIdenticalToNewestSource(ui));
    }

    @Test
    void isSavedShapedPlanIdenticalToNewestSource_falseWhenSourceChanged(@TempDir Path tempDir)
            throws Exception {
        Path sourceDir = tempDir.resolve("task-input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(outputDir);
        TestAladdinPlanXlsx.writeMinimal(sourceDir, "aladdin-plan.xlsx");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        sourceDir.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        outputDir.toString());

        AladdinProcessingPlanSourceReloader.reloadNewestFromDiskAndSaveShapedJson(ui);

        TestAladdinPlanXlsx.writeWithQty(sourceDir, "aladdin-plan.xlsx", "99");

        assertFalse(
                AladdinProcessingPlanSourceFreshness.isSavedShapedPlanIdenticalToNewestSource(ui));
    }

    @Test
    void isSavedShapedPlanIdenticalToNewestSource_falseWhenShapedMissing(@TempDir Path tempDir) {
        Path sourceDir = tempDir.resolve("task-input");
        Path outputDir = tempDir.resolve("output");
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        sourceDir.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        outputDir.toString());

        assertFalse(
                AladdinProcessingPlanSourceFreshness.isSavedShapedPlanIdenticalToNewestSource(ui));
    }

    @Test
    void tabularSheetsEqual_ignoresNullCells() {
        PlanInputTabularIo.TabularSheet a =
                new PlanInputTabularIo.TabularSheet(List.of("A"), List.of(List.of("1")));
        PlanInputTabularIo.TabularSheet b =
                new PlanInputTabularIo.TabularSheet(List.of("A"), List.of(List.of("1")));

        assertTrue(AladdinProcessingPlanSourceFreshness.tabularSheetsEqual(a, b));
    }
}
