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

    @Test
    void refreshButtonNeedsAttention_whenUsedFilesDifferFromRemoteDesktopLatest(
            @TempDir Path tempDir) throws Exception {
        Path planDir = tempDir.resolve("plan");
        Path dailyDir = tempDir.resolve("daily");
        Files.createDirectories(planDir);
        Files.createDirectories(dailyDir);
        Path planOld = planDir.resolve("plan-old.xlsx");
        Path planNew = planDir.resolve("plan-new.xlsx");
        Path dailyOld = dailyDir.resolve("加工日報発行問合せ_old.csv");
        Path dailyNew = dailyDir.resolve("加工日報発行問合せ_new.csv");
        Files.writeString(planOld, "old");
        Files.writeString(planNew, "new");
        Files.writeString(dailyOld, "a,b\n1,2");
        Files.writeString(dailyNew, "a,b\n3,4");
        Files.setLastModifiedTime(
                planNew,
                java.nio.file.attribute.FileTime.fromMillis(
                        Files.getLastModifiedTime(planOld).toMillis() + 60_000));
        Files.setLastModifiedTime(
                dailyNew,
                java.nio.file.attribute.FileTime.fromMillis(
                        Files.getLastModifiedTime(dailyOld).toMillis() + 60_000));

        Map<String, String> ui =
                Map.of(
                        "PM_AI_TASK_INPUT_SOURCE_DIR",
                        planDir.toString(),
                        jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup
                                .KEY_DAILY_REPORT_SOURCE_DIR,
                        dailyDir.toString());
        RequestFormPipelineCheckTabController controller =
                new RequestFormPipelineCheckTabController();
        controller.captureLastScannedAladdinPlanRevisionForTest(ui);

        String latestPlan = planNew.toAbsolutePath().normalize().toString();
        String latestDaily = dailyNew.toAbsolutePath().normalize().toString();
        String usedPlan = planOld.toAbsolutePath().normalize().toString();
        String usedDaily = dailyOld.toAbsolutePath().normalize().toString();

        assertFalse(controller.refreshButtonNeedsAttention(ui, latestPlan, latestDaily));
        assertTrue(controller.refreshButtonNeedsAttention(ui, usedPlan, latestDaily));
        assertTrue(controller.refreshButtonNeedsAttention(ui, latestPlan, usedDaily));
    }
}
