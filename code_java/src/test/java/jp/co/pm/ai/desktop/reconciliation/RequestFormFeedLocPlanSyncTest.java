package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

class RequestFormFeedLocPlanSyncTest {

    @Test
    void collectDistinctFeedLocations_readsPlanInputTasks(@TempDir Path tmp) throws Exception {
        Path plan = tmp.resolve(AppPaths.STAGE1_PLAN_TASKS_FILENAME);
        List<String> headers =
                List.of("製品名", RequestFormFeedLocPlanSync.COL_RAW_FEED_LOCATION, "加工内容");
        List<List<String>> rows =
                List.of(
                        List.of("A", "EC機　湖南", "x"),
                        List.of("B", "SEC機　湖南", "y"),
                        List.of("C", "EC機　湖南", "z"),
                        List.of("D", "", "w"),
                        List.of("E", "  スライス機1　湖南  ", "v"));
        PlanInputTabularIo.write(
                plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET, new PlanInputTabularIo.TabularSheet(headers, rows));
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH, plan.toString());
        List<String> values = RequestFormFeedLocPlanSync.collectDistinctFeedLocations(ui);
        assertEquals(
                List.of("EC機　湖南", "SEC機　湖南", "スライス機1　湖南"),
                values);
    }

    @Test
    void mergeDistinctFeedLocations_preservesExistingAndAppendsNew() {
        List<String> before = List.of("EC", "SEC");
        List<String> fromPlan = List.of("EC機　湖南", "SEC", "スリット機1　湖南");
        List<String> merged = RequestFormFeedLocPlanSync.mergeDistinctFeedLocations(before, fromPlan);
        assertEquals(List.of("EC", "SEC", "EC機　湖南", "スリット機1　湖南"), merged);
        assertEquals(2, RequestFormFeedLocPlanSync.countNewValues(before, merged));
    }

    @Test
    void collectDistinctFeedLocations_returnsEmptyWhenColumnMissing(@TempDir Path tmp) throws Exception {
        Path plan = tmp.resolve("plan.xlsx");
        PlanInputTabularIo.write(
                plan,
                AppPaths.STAGE1_PLAN_OUTPUT_SHEET,
                new PlanInputTabularIo.TabularSheet(List.of("製品名"), List.of(List.of("A"))));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH, plan.toString());
        assertTrue(RequestFormFeedLocPlanSync.collectDistinctFeedLocations(ui).isEmpty());
        assertTrue(Files.isRegularFile(plan));
    }
}
