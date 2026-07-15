package jp.co.pm.ai.planning.stage2.source;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.Map;
import java.util.HashMap;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup;

class Stage2SourceConsistencyGuardTest {

    @TempDir Path temp;

    @Test
    void verify_allowsMatchingPaths() throws Exception {
        Path plan = temp.resolve("plan.xlsx");
        Path daily = temp.resolve("加工日報発行問合せ_20260710_071800.csv");
        java.nio.file.Files.writeString(plan, "x");
        java.nio.file.Files.writeString(daily, "a\nb\nc\nh\n");

        Stage1SourceBundle bundle =
                new Stage1SourceBundle(
                        LocalDateTime.of(2026, 7, 10, 7, 5),
                        LocalDateTime.of(2026, 7, 10, 7, 18),
                        13L,
                        plan.toString(),
                        daily.toString(),
                        plan.toString(),
                        1L);

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH,
                        plan.toString(),
                        KonanDailyReportLookup.KEY_DAILY_REPORT_CSV_PATH,
                        daily.toString());

        Stage2SourceConsistencyGuard.Result result = Stage2SourceConsistencyGuard.verify(ui, bundle);
        assertTrue(result.allowed());
    }

    @Test
    void verify_blocksMismatchedPlanPath() throws Exception {
        Path plan = temp.resolve("plan.xlsx");
        Path otherPlan = temp.resolve("other.xlsx");
        Path daily = temp.resolve("daily.csv");
        java.nio.file.Files.writeString(plan, "x");
        java.nio.file.Files.writeString(otherPlan, "y");
        java.nio.file.Files.writeString(daily, "a\nb\nc\nh\n");

        Stage1SourceBundle bundle =
                new Stage1SourceBundle(
                        LocalDateTime.of(2026, 7, 10, 7, 5),
                        LocalDateTime.of(2026, 7, 10, 7, 18),
                        13L,
                        plan.toString(),
                        daily.toString(),
                        plan.toString(),
                        1L);

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH,
                        otherPlan.toString(),
                        KonanDailyReportLookup.KEY_DAILY_REPORT_CSV_PATH,
                        daily.toString());

        Stage2SourceConsistencyGuard.Result result = Stage2SourceConsistencyGuard.verify(ui, bundle);
        assertFalse(result.allowed());
    }

    @Test
    void prepareOffAllowsWithoutBundle() {
        assertTrue(Stage2SourceConsistencyGuard.prepareForStageRun(false, new HashMap<>(), null).allowed());
    }

    @Test
    void prepareBlocksInvalidExplicitPath() throws Exception {
        Path plan = temp.resolve("plan-invalid.xlsx");
        Path daily = temp.resolve("daily-invalid.csv");
        java.nio.file.Files.writeString(plan, "x");
        java.nio.file.Files.writeString(daily, "x");
        Stage1SourceBundle bundle = bundle(plan, daily, plan);
        Map<String, String> env = new HashMap<>();
        env.put(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH, "bad\u0000path");
        assertFalse(Stage2SourceConsistencyGuard.prepareForStageRun(true, env, bundle).allowed());
    }

    @Test
    void prepareBlocksMissingAnyFixedFileBeforeOverlay() throws Exception {
        Path plan = temp.resolve("plan-missing-extract.xlsx");
        Path daily = temp.resolve("daily-missing-extract.csv");
        java.nio.file.Files.writeString(plan, "x");
        java.nio.file.Files.writeString(daily, "x");
        Map<String, String> env = new HashMap<>();
        assertFalse(Stage2SourceConsistencyGuard.prepareForStageRun(true, env, bundle(plan, daily, temp.resolve("missing.xlsx"))).allowed());
        assertFalse(env.containsKey(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH));
    }

    private static Stage1SourceBundle bundle(Path plan, Path daily, Path extraction) {
        return new Stage1SourceBundle(LocalDateTime.of(2026,7,10,7,5), LocalDateTime.of(2026,7,10,7,18), 13L, plan.toString(), daily.toString(), extraction.toString(), 1L);
    }
}
