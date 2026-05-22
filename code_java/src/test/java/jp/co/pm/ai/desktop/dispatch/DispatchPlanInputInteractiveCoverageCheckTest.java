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
}
