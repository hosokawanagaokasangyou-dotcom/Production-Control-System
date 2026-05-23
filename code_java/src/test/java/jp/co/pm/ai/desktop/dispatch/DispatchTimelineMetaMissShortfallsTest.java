package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

class DispatchTimelineMetaMissShortfallsTest {

    @Test
    void detectsPlanRowWithEmptyMachiningStart() {
        List<String> cols =
                List.of(
                        "依頼NO",
                        ResultDispatchSchema.COL_MACHINE,
                        ResultDispatchSchema.COL_DISPATCH_DATE,
                        ResultDispatchSchema.COL_DISPATCH_QTY,
                        "加工開始日時");
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("V5-7", "スライス機1　湖南", "2026-05-26T00:00:00.000", "400", ""));

        var found = DispatchTimelineMetaMissShortfalls.detectFromRows(cols, rows);
        assertEquals(1, found.size());
        assertEquals("V5-7", found.getFirst().taskId());
        assertEquals(400.0, found.getFirst().targetM(), 1e-6);
        assertEquals(400.0, found.getFirst().shortfallM(), 1e-6);
    }

    @Test
    void skipsWhenMachiningStartPresent() {
        List<String> cols =
                List.of(
                        "依頼NO",
                        ResultDispatchSchema.COL_MACHINE,
                        ResultDispatchSchema.COL_DISPATCH_DATE,
                        ResultDispatchSchema.COL_DISPATCH_QTY,
                        "加工開始日時");
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(
                row(
                        "V5-7",
                        "スライス機1　湖南",
                        "2026-05-26",
                        "400",
                        "2026/05/26 15:05"));

        assertTrue(DispatchTimelineMetaMissShortfalls.detectFromRows(cols, rows).isEmpty());
    }

    @Test
    void detectsWhenActualQtyPositiveButMachiningStartEmpty() {
        List<String> cols =
                List.of(
                        "依頼NO",
                        ResultDispatchSchema.COL_MACHINE,
                        ResultDispatchSchema.COL_DISPATCH_DATE,
                        ResultDispatchSchema.COL_DISPATCH_QTY,
                        ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL,
                        "加工開始日時");
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> r =
                row("V5-7", "スライス機1　湖南", "2026-05-26", "400", "");
        r.put(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL, "400");
        rows.add(r);

        var found = DispatchTimelineMetaMissShortfalls.detectFromRows(cols, rows);
        assertEquals(1, found.size());
        assertEquals("V5-7", found.getFirst().taskId());
    }

    private static Map<String, String> row(
            String tid, String mach, String day, String plan, String start) {
        Map<String, String> m = new LinkedHashMap<>();
        m.put("依頼NO", tid);
        m.put(ResultDispatchSchema.COL_MACHINE, mach);
        m.put(ResultDispatchSchema.COL_DISPATCH_DATE, day);
        m.put(ResultDispatchSchema.COL_DISPATCH_QTY, plan);
        m.put("加工開始日時", start);
        return m;
    }
}
