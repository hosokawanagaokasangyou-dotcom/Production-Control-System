package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class RemoteDesktopLatestSourceFilesTest {

    @TempDir
    Path temp;

    @Test
    void resolveAll_picksNewestPerCategory() throws Exception {
        Path root = temp;
        Path planDir = root.resolve("plan");
        Path dailyDir = root.resolve("daily");
        Path actualDir = root.resolve("actual");
        Path orderDir = root.resolve("order");
        Files.createDirectories(planDir);
        Files.createDirectories(dailyDir);
        Files.createDirectories(actualDir);
        Files.createDirectories(orderDir);

        Path planOld = planDir.resolve("工程別生産計画問合せ_20260101.xlsx");
        Path planNew = planDir.resolve("工程別生産計画問合せ_20260201.xlsx");
        Files.writeString(planOld, "old");
        Files.writeString(planNew, "new");
        bumpMtime(planNew, planOld);

        Path dailyOld = dailyDir.resolve("加工日報発行問合せ_20260101.csv");
        Path dailyNew = dailyDir.resolve("加工日報発行問合せ_20260201.csv");
        Files.writeString(dailyOld, "a,b\n1,2");
        Files.writeString(dailyNew, "a,b\n3,4");
        bumpMtime(dailyNew, dailyOld);

        Path actualOld = actualDir.resolve("NO(ロット)別問合せ_20260101.xlsx");
        Path actualNew = actualDir.resolve("NO(ロット)別問合せ_20260201.xlsx");
        Files.writeString(actualOld, "old");
        Files.writeString(actualNew, "new");
        bumpMtime(actualNew, actualOld);

        Path orderOld = orderDir.resolve("受注明細表_20260101.xlsx");
        Path orderNew = orderDir.resolve("受注明細表問合せ_20260201.xlsx");
        Files.writeString(orderOld, "old");
        Files.writeString(orderNew, "new");
        bumpMtime(orderNew, orderOld);

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        planDir.toString(),
                        KonanDailyReportLookup.KEY_DAILY_REPORT_SOURCE_DIR,
                        dailyDir.toString(),
                        AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
                        actualDir.toString(),
                        RemoteDesktopLatestSourceFiles.KEY_ORDER_DETAIL_SOURCE_DIR,
                        orderDir.toString());

        List<RemoteDesktopLatestSourceFiles.Row> rows = RemoteDesktopLatestSourceFiles.resolveAll(ui);
        assertEquals(4, rows.size());
        assertEquals(planNew.toAbsolutePath().normalize().toString(), pathFor(rows, "加工計画"));
        assertEquals(dailyNew.toAbsolutePath().normalize().toString(), pathFor(rows, "加工日報"));
        assertEquals(actualNew.toAbsolutePath().normalize().toString(), pathFor(rows, "加工実績"));
        assertEquals(orderNew.toAbsolutePath().normalize().toString(), pathFor(rows, "受注明細表"));
        assertTrue(acquiredAtFor(rows, "加工計画") > 0L);
        assertEquals(
                RemoteDesktopLatestSourceFiles.acquiredAtEpochMillis(planNew),
                acquiredAtFor(rows, "加工計画"));
    }

    @Test
    void resolveAll_missingFilesShowEmptyPath() {
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        temp.resolve("missing-plan").toString(),
                        KonanDailyReportLookup.KEY_DAILY_REPORT_SOURCE_DIR,
                        temp.resolve("missing-daily").toString(),
                        AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
                        temp.resolve("missing-actual").toString(),
                        RemoteDesktopLatestSourceFiles.KEY_ORDER_DETAIL_SOURCE_DIR,
                        temp.resolve("missing-order").toString());

        List<RemoteDesktopLatestSourceFiles.Row> rows = RemoteDesktopLatestSourceFiles.resolveAll(ui);
        assertEquals(4, rows.size());
        for (RemoteDesktopLatestSourceFiles.Row row : rows) {
            assertTrue(row.fullPath().isBlank());
            assertEquals("（該当ファイルなし）", row.displayPath());
            assertEquals("—", row.displayAcquiredAt());
        }
    }

    @Test
    void pipelineCheckSourcesOlderThan_detectsStalePlanAndDaily() throws Exception {
        Path planDir = temp.resolve("plan");
        Path dailyDir = temp.resolve("daily");
        Files.createDirectories(planDir);
        Files.createDirectories(dailyDir);

        Path planFile = planDir.resolve("plan.xlsx");
        Path dailyFile = dailyDir.resolve("加工日報発行問合せ_1.csv");
        Files.writeString(planFile, "p");
        Files.writeString(dailyFile, "a,b\n1,2");

        long staleMillis = System.currentTimeMillis() - (2L * 60L * 60L * 1000L);
        Files.setLastModifiedTime(planFile, FileTime.fromMillis(staleMillis));
        Files.setLastModifiedTime(dailyFile, FileTime.fromMillis(staleMillis));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        planDir.toString(),
                        KonanDailyReportLookup.KEY_DAILY_REPORT_SOURCE_DIR,
                        dailyDir.toString());

        List<RemoteDesktopLatestSourceFiles.Row> stale =
                RemoteDesktopLatestSourceFiles.pipelineCheckSourcesOlderThan(
                        ui,
                        RemoteDesktopLatestSourceFiles.PIPELINE_CHECK_MAX_SOURCE_AGE_MS,
                        System.currentTimeMillis());
        assertEquals(2, stale.size());
    }

    @Test
    void pipelineCheckSourcesOlderThan_ignoresFreshAndMissing() throws Exception {
        Path planDir = temp.resolve("plan-fresh");
        Files.createDirectories(planDir);
        Path planFile = planDir.resolve("plan.xlsx");
        Files.writeString(planFile, "p");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        planDir.toString(),
                        KonanDailyReportLookup.KEY_DAILY_REPORT_SOURCE_DIR,
                        temp.resolve("missing-daily").toString());

        List<RemoteDesktopLatestSourceFiles.Row> stale =
                RemoteDesktopLatestSourceFiles.pipelineCheckSourcesOlderThan(
                        ui,
                        RemoteDesktopLatestSourceFiles.PIPELINE_CHECK_MAX_SOURCE_AGE_MS,
                        System.currentTimeMillis());
        assertTrue(stale.isEmpty());
    }

    private static void bumpMtime(Path newer, Path older) throws Exception {
        long olderMillis = Files.getLastModifiedTime(older).toMillis();
        Files.setLastModifiedTime(newer, FileTime.fromMillis(olderMillis + 60_000));
    }

    private static String pathFor(List<RemoteDesktopLatestSourceFiles.Row> rows, String label) {
        return rows.stream()
                .filter(r -> r.category().label().equals(label))
                .map(RemoteDesktopLatestSourceFiles.Row::fullPath)
                .findFirst()
                .orElseThrow();
    }

    private static long acquiredAtFor(List<RemoteDesktopLatestSourceFiles.Row> rows, String label) {
        return rows.stream()
                .filter(r -> r.category().label().equals(label))
                .map(RemoteDesktopLatestSourceFiles.Row::acquiredAtEpochMillis)
                .findFirst()
                .orElseThrow();
    }
}
