package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class DispatchPlanInputInteractiveCoverageCheckTest {

    @Test
    void isExcluded_whenDispatchNotRequiredOn() {
        Map<String, String> row = new LinkedHashMap<>();
        row.put("配台不要", "はい");
        row.put("依頼NO", "W5-13");
        assertTrue(DispatchPlanInputInteractiveCoverageCheck.isExcludedFromDispatchCoverage(row));
    }

    @Test
    void isEligible_whenDispatchNotRequiredOff() {
        Map<String, String> row = new LinkedHashMap<>();
        row.put("配台不要", "いいえ");
        row.put("依頼NO", "Y5-16");
        row.put("工程名", "スライス");
        row.put("機械名", "スライス機1　湖南");
        assertTrue(DispatchPlanInputInteractiveCoverageCheck.isEligiblePlanInputRow(row));
    }

    @Test
    void findMissing_detectsAbsentTaskKey() {
        var expected =
                List.of(
                        new DispatchPlanInputInteractiveCoverageCheck.TaskKey(
                                "Y5-27", "SEC", "SEC機　湖南"));
        Map<String, String> present = new LinkedHashMap<>();
        present.put("依頼NO", "Y5-16");
        present.put("工程名", "スライス");
        present.put("機械名", "スライス機1　湖南");
        var missing =
                DispatchPlanInputInteractiveCoverageCheck.findMissingInDispatchTable(
                        expected, List.of(present));
        assertEquals(1, missing.size());
        assertEquals("Y5-27", missing.getFirst().requestNo());
    }

    @Test
    void excludeInProgressZeroNextDay_noZeroEntries_keepsAllExpected() throws Exception {
        java.nio.file.Path json =
                java.nio.file.Files.createTempFile("stage2-in-progress-empty", ".json");
        java.nio.file.Files.writeString(json, "{\"version\":1,\"entries\":[]}");
        var expected =
                List.of(
                        new DispatchPlanInputInteractiveCoverageCheck.TaskKey(
                                "V6-3", "スライス", "スライス機1　湖南"),
                        new DispatchPlanInputInteractiveCoverageCheck.TaskKey(
                                "W6-8", "EC", "EC機　湖南"));
        Map<String, String> ui =
                Map.of("PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON", json.toString());
        var filtered =
                DispatchPlanInputInteractiveCoverageCheck.excludeInProgressZeroNextDayFromExpected(
                        expected, ui);
        assertEquals(2, filtered.size());
        java.nio.file.Files.deleteIfExists(json);
    }

    @Test
    void excludeInProgressZeroNextDay_filtersZeroMeterEntryFromJson() throws Exception {
        java.nio.file.Path json =
                java.nio.file.Files.createTempFile("stage2-in-progress-next-day", ".json");
        java.nio.file.Files.writeString(
                json,
                """
                {
                  "version": 1,
                  "entries": [
                    {
                      "task_id": "V6-3",
                      "process": "スライス",
                      "machine_name": "スライス機1　湖南",
                      "next_day_dispatch_m": 0.0
                    },
                    {
                      "task_id": "W6-8",
                      "process": "EC",
                      "machine_name": "EC機　湖南",
                      "next_day_dispatch_m": 1500.0
                    }
                  ]
                }
                """);
        var expected =
                List.of(
                        new DispatchPlanInputInteractiveCoverageCheck.TaskKey(
                                "V6-3", "スライス", "スライス機1　湖南"),
                        new DispatchPlanInputInteractiveCoverageCheck.TaskKey(
                                "W6-8", "EC", "EC機　湖南"));
        Map<String, String> ui =
                Map.of("PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON", json.toString());
        var filtered =
                DispatchPlanInputInteractiveCoverageCheck.excludeInProgressZeroNextDayFromExpected(
                        expected, ui);
        assertEquals(1, filtered.size());
        assertEquals("W6-8", filtered.getFirst().requestNo());
        java.nio.file.Files.deleteIfExists(json);
    }
}
