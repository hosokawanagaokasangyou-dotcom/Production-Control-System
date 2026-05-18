package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class ResultDispatchInteractiveConsolidatorTest {

    @Test
    void consolidatesOrphanPlanRowsIntoTimelineRows() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        if (!cols.contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL)) {
            cols.add(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        }
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row(cols, "W5-13", "EC", "EC機　湖南", "2026/06/03", "900", "0", ""));
        rows.add(row(cols, "W5-13", "EC", "EC機　湖南", "2026/06/08", "6300", "0", ""));
        rows.add(row(cols, "W5-13", "EC", "EC機　湖南", "2026/06/09", "0", "3000", "2026/06/09 14:08"));
        rows.add(row(cols, "W5-13", "EC", "EC機　湖南", "2026/06/10", "0", "4200", "2026/06/10 09:15"));

        ResultDispatchInteractiveConsolidator.consolidatePlanAndTimelineRowsInPlace(cols, rows);

        assertEquals(2, rows.size());
        assertEquals("3000", rows.get(0).get(ResultDispatchSchema.COL_DISPATCH_QTY));
        assertEquals("2026/06/09 14:08", rows.get(0).get("加工開始日時"));
        assertEquals("4200", rows.get(1).get(ResultDispatchSchema.COL_DISPATCH_QTY));
        assertEquals("2026/06/10 09:15", rows.get(1).get("加工開始日時"));
    }

    @Test
    void skipsWhenNoActualColumn() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row(cols, "T1", "P", "M1", "2026/06/03", "100", "", ""));
        int before = rows.size();
        ResultDispatchInteractiveConsolidator.consolidatePlanAndTimelineRowsInPlace(cols, rows);
        assertEquals(before, rows.size());
    }

    @Test
    void syncsPlanQtyFromActualWhenNoOrphanPlanRows() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        cols.add(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(
                row(
                        cols,
                        "Y5-15",
                        "スライス",
                        "スライス機1　湖南",
                        "2026/05/21",
                        "4400",
                        "3000",
                        "2026/05/21 15:00"));
        rows.add(
                row(
                        cols,
                        "Y5-15",
                        "スライス",
                        "スライス機1　湖南",
                        "2026/05/22",
                        "3600",
                        "3600",
                        "2026/05/22 11:38"));
        rows.add(
                row(
                        cols,
                        "Y5-15",
                        "スライス",
                        "スライス機1　湖南",
                        "2026/05/26",
                        "0",
                        "1400",
                        "2026/05/26 09:27"));

        ResultDispatchInteractiveConsolidator.consolidatePlanAndTimelineRowsInPlace(cols, rows);

        assertEquals(2, rows.size());
        assertEquals("4400", rows.get(0).get(ResultDispatchSchema.COL_DISPATCH_QTY));
        assertEquals("3600", rows.get(1).get(ResultDispatchSchema.COL_DISPATCH_QTY));
        assertEquals("2026/05/21 15:00", rows.get(0).get("加工開始日時"));
        assertEquals("2026/05/22 11:38", rows.get(1).get("加工開始日時"));
    }

    @Test
    void keepsOrphanPlanWhenNoTimelineRowsInGroup() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        cols.add(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row(cols, "T2", "P", "M1", "2026/06/03", "500", "0", ""));

        ResultDispatchInteractiveConsolidator.consolidatePlanAndTimelineRowsInPlace(cols, rows);

        assertEquals(1, rows.size());
        assertEquals("500", rows.getFirst().get(ResultDispatchSchema.COL_DISPATCH_QTY));
        assertEquals("", rows.getFirst().get("加工開始日時"));
    }

    private static Map<String, String> row(
            List<String> cols,
            String tid,
            String proc,
            String machine,
            String dispatchDate,
            String planQty,
            String actualQty,
            String startDt) {
        LinkedHashMap<String, String> m = new LinkedHashMap<>();
        for (String c : cols) {
            m.put(c, "");
        }
        m.put("依頼NO", tid);
        m.put(ResultDispatchSchema.COL_PROCESS, proc);
        m.put(ResultDispatchSchema.COL_MACHINE, machine);
        m.put(ResultDispatchSchema.COL_DISPATCH_DATE, dispatchDate);
        m.put(ResultDispatchSchema.COL_DISPATCH_QTY, planQty);
        m.put(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL, actualQty);
        m.put("加工開始日時", startDt);
        m.put("加工終了日時", startDt.isEmpty() ? "" : startDt.substring(0, 10) + " 17:00");
        return m;
    }
}
