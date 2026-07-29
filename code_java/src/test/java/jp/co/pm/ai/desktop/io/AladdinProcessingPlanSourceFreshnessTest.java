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

    private static final String PLAN_CSV =
            "列1,列2,列3,列4\n"
                    + "上段1,,,\n"
                    + "上段2,,,\n"
                    + "上段3,,,\n"
                    + "機械名,依頼NO,工程名,2026/07/07\n"
                    + ",,,\n"
                    + "M1,T001,工程A,10\n";

    @Test
    void isSavedShapedPlanIdenticalToNewestSource_trueWhenShapedMatchesSource(@TempDir Path tempDir)
            throws Exception {
        Path sourceDir = tempDir.resolve("task-input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(sourceDir);
        Files.createDirectories(outputDir);
        Files.writeString(sourceDir.resolve("aladdin-plan.csv"), PLAN_CSV);

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
        Files.createDirectories(sourceDir);
        Files.createDirectories(outputDir);
        Path csv = sourceDir.resolve("aladdin-plan.csv");
        Files.writeString(csv, PLAN_CSV);

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        sourceDir.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        outputDir.toString());

        AladdinProcessingPlanSourceReloader.reloadNewestFromDiskAndSaveShapedJson(ui);

        Files.writeString(
                csv, PLAN_CSV.replace("M1,T001,工程A,10", "M1,T001,工程A,99"));

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
