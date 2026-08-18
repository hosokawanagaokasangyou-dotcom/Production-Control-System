package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class SourceFileExtensionPolicyTest {

    @TempDir
    Path temp;

    @Test
    void processingPlan_okWhenNewestIsXlsx() throws Exception {
        Path olderCsv = temp.resolve("plan.csv");
        Path newerXlsx = temp.resolve("plan.xlsx");
        Files.writeString(olderCsv, "csv");
        Files.writeString(newerXlsx, "xlsx");
        bumpNewer(newerXlsx, olderCsv);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkProcessingPlanDirectory(temp);

        assertTrue(r.ok());
        assertEquals(newerXlsx.toAbsolutePath().normalize(), r.loadablePath().orElseThrow());
        assertTrue(r.errorMessage().isEmpty());
    }

    @Test
    void processingPlan_errorWhenNewestIsCsvEvenIfOlderXlsxExists() throws Exception {
        Path olderXlsx = temp.resolve("plan.xlsx");
        Path newerCsv = temp.resolve("受注明細表.csv");
        Files.writeString(olderXlsx, "xlsx");
        Files.writeString(newerCsv, "csv");
        bumpNewer(newerCsv, olderXlsx);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkProcessingPlanDirectory(temp);

        assertFalse(r.ok());
        assertTrue(r.loadablePath().isEmpty());
        assertEquals(newerCsv.toAbsolutePath().normalize(), r.newestCandidatePath().orElseThrow());
        assertTrue(r.errorMessage().contains(".xlsx"));
        assertTrue(r.errorMessage().contains(".csv") || r.errorMessage().contains("csv"));
    }

    @Test
    void processingPlan_ignoresXlsmForLoadAndErrorsIfNewest() throws Exception {
        Path xlsm = temp.resolve("plan.xlsm");
        Files.writeString(xlsm, "xlsm");

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkProcessingPlanDirectory(temp);

        assertFalse(r.ok());
        assertTrue(r.loadablePath().isEmpty());
    }

    @Test
    void processingPlan_respectsRequiredExtFromUi() throws Exception {
        Path olderXlsx = temp.resolve("plan.xlsx");
        Path newerCsv = temp.resolve("plan.csv");
        Files.writeString(olderXlsx, "xlsx");
        Files.writeString(newerCsv, "csv");
        bumpNewer(newerCsv, olderXlsx);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkProcessingPlanDirectory(
                        temp,
                        Map.of(AppPaths.KEY_PM_AI_PROCESSING_PLAN_REQUIRED_EXT, "csv"));

        assertTrue(r.ok());
        assertEquals(newerCsv.toAbsolutePath().normalize(), r.loadablePath().orElseThrow());
    }

    @Test
    void dailyReport_okWhenNewestMatchingPrefixIsCsv() throws Exception {
        Path older = temp.resolve("加工日報発行問合せ_20260101.csv");
        Path newer = temp.resolve("加工日報発行問合せ_20260201.csv");
        Files.writeString(older, "a");
        Files.writeString(newer, "b");
        bumpNewer(newer, older);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkDailyReportDirectory(temp);

        assertTrue(r.ok());
        assertEquals(newer.toAbsolutePath().normalize(), r.loadablePath().orElseThrow());
    }

    @Test
    void dailyReport_errorWhenNewestMatchingPrefixIsXlsx() throws Exception {
        Path olderCsv = temp.resolve("加工日報発行問合せ_20260101.csv");
        Path newerXlsx = temp.resolve("加工日報発行問合せ_20260201.xlsx");
        Files.writeString(olderCsv, "csv");
        Files.writeString(newerXlsx, "xlsx");
        bumpNewer(newerXlsx, olderCsv);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkDailyReportDirectory(temp);

        assertFalse(r.ok());
        assertTrue(r.loadablePath().isEmpty());
        assertEquals(newerXlsx.toAbsolutePath().normalize(), r.newestCandidatePath().orElseThrow());
        assertTrue(r.errorMessage().contains(".csv"));
    }

    @Test
    void dailyReport_respectsRequiredExtFromUi() throws Exception {
        Path olderCsv = temp.resolve("加工日報発行問合せ_20260101.csv");
        Path newerXlsx = temp.resolve("加工日報発行問合せ_20260201.xlsx");
        Files.writeString(olderCsv, "csv");
        Files.writeString(newerXlsx, "xlsx");
        bumpNewer(newerXlsx, olderCsv);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkDailyReportDirectory(
                        temp, Map.of(AppPaths.KEY_PM_AI_DAILY_REPORT_REQUIRED_EXT, ".xlsx"));

        assertTrue(r.ok());
        assertEquals(newerXlsx.toAbsolutePath().normalize(), r.loadablePath().orElseThrow());
    }

    @Test
    void blockingMismatchMessages_includesPlanWhenNewestIsCsv() throws Exception {
        Path planDir = temp.resolve("plan");
        Path dailyDir = temp.resolve("daily");
        Files.createDirectories(planDir);
        Files.createDirectories(dailyDir);
        Files.writeString(planDir.resolve("old.xlsx"), "x");
        Path newerCsv = planDir.resolve("new.csv");
        Files.writeString(newerCsv, "c");
        bumpNewer(newerCsv, planDir.resolve("old.xlsx"));
        Files.writeString(dailyDir.resolve("加工日報発行問合せ_1.csv"), "d");

        List<String> msgs =
                SourceFileExtensionPolicy.blockingMismatchMessages(
                        Map.of(
                                AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                                planDir.toString(),
                                AppPaths.KEY_PM_AI_DAILY_REPORT_SOURCE_DIR,
                                dailyDir.toString()));

        assertEquals(1, msgs.size());
        assertTrue(msgs.getFirst().contains("加工計画"));
        assertTrue(msgs.getFirst().contains("拡張子が不正"));
    }

    @Test
    void blockingMismatchPaths_returnsNewestBadPlan() throws Exception {
        Path planDir = temp.resolve("plan");
        Path dailyDir = temp.resolve("daily");
        Files.createDirectories(planDir);
        Files.createDirectories(dailyDir);
        Files.writeString(planDir.resolve("old.xlsx"), "x");
        Path newerCsv = planDir.resolve("new.csv");
        Files.writeString(newerCsv, "c");
        bumpNewer(newerCsv, planDir.resolve("old.xlsx"));
        Files.writeString(dailyDir.resolve("加工日報発行問合せ_1.csv"), "d");

        List<Path> paths =
                SourceFileExtensionPolicy.blockingMismatchPaths(
                        Map.of(
                                AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                                planDir.toString(),
                                AppPaths.KEY_PM_AI_DAILY_REPORT_SOURCE_DIR,
                                dailyDir.toString()));

        assertEquals(1, paths.size());
        assertEquals(newerCsv.toAbsolutePath().normalize(), paths.getFirst());
    }

    @Test
    void dailyReport_fileRequiresFilenamePrefix() throws Exception {
        Path badName = temp.resolve("任意名.csv");
        Files.writeString(badName, "x");

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkDailyReportFile(badName);

        assertFalse(r.ok());
        assertTrue(r.errorMessage().contains("接頭辞"));
    }

    @Test
    void normalizeFileExtension_addsDotAndLowercases() {
        assertEquals(".xlsx", AppPaths.normalizeFileExtension("XLSX"));
        assertEquals(".csv", AppPaths.normalizeFileExtension(".CSV"));
        assertEquals(".pq", AppPaths.normalizeFileExtension("*.pq"));
        assertEquals("", AppPaths.normalizeFileExtension(""));
    }

    private static void bumpNewer(Path newer, Path older) throws Exception {
        long base = Instant.now().toEpochMilli();
        Files.setLastModifiedTime(older, FileTime.fromMillis(base));
        Files.setLastModifiedTime(newer, FileTime.fromMillis(base + 60_000L));
    }
}
