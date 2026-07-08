package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

class AladdinShapedPlanRawInputDateLookupTest {

    @Test
    void resolvesRawInputDate_forMatchingTaskId() {
        List<String> headers = List.of("機械名", "依頼NO", "原反投入日", "2026/07/07");
        List<List<String>> rows =
                List.of(
                        List.of("M1", "T001", "2026/7/5", "10"),
                        List.of("M2", "T002", "2026/7/9", "20"));
        assertEquals(
                "2026/7/5",
                AladdinShapedPlanQtyLookup.resolveRawInputDateDisplayForTaskId(
                        headers, rows, "T001"));
    }

    @Test
    void acceptsLegacyHeader_genpanTonyubi() {
        List<String> headers = List.of("機械名", "依頼NO", "原板投入日");
        List<List<String>> rows = List.of(List.of("M1", "T001", "2026/7/5"));
        assertEquals(
                "2026/7/5",
                AladdinShapedPlanQtyLookup.resolveRawInputDateDisplayForTaskId(
                        headers, rows, "T001"));
    }

    @Test
    void joinsDistinctValues_forMultipleRows() {
        List<String> headers = List.of("機械名", "依頼NO", "原反投入日");
        List<List<String>> rows =
                List.of(
                        List.of("M1", "T001", "2026/7/5"),
                        List.of("M2", "T001", "2026/7/6"),
                        List.of("M3", "T001", "2026/7/5"));
        assertEquals(
                "2026/7/5\n2026/7/6",
                AladdinShapedPlanQtyLookup.resolveRawInputDateDisplayForTaskId(
                        headers, rows, "T001"));
    }

    @Test
    void returnsEmpty_whenColumnMissing() {
        List<String> headers = List.of("機械名", "依頼NO");
        List<List<String>> rows = List.of(List.of("M1", "T001"));
        assertEquals(
                "",
                AladdinShapedPlanQtyLookup.resolveRawInputDateDisplayForTaskId(
                        headers, rows, "T001"));
    }
}
