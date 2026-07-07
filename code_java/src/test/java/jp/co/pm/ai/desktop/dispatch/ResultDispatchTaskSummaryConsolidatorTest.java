package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class ResultDispatchTaskSummaryConsolidatorTest {

    @Test
    void mergesSameTaskAcrossDaysWithMinStartMaxEndAndQtySum() {
        List<String> cols =
                List.of(
                        "依頼NO",
                        ResultDispatchSchema.COL_PROCESS,
                        ResultDispatchSchema.COL_MACHINE,
                        "加工開始日時",
                        "加工終了日時",
                        ResultDispatchSchema.COL_DISPATCH_DATE,
                        ResultDispatchSchema.COL_DISPATCH_QTY);

        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("W5-13", "巻返し", "M1", "2026/07/01 08:00", "2026/07/01 12:00", "2026/07/01", "100"));
        rows.add(row("W5-13", "巻返し", "M1", "2026/07/02 09:00", "2026/07/02 17:00", "2026/07/02", "50"));

        List<Map<String, String>> out = ResultDispatchTaskSummaryConsolidator.consolidate(cols, rows);

        assertEquals(1, out.size());
        Map<String, String> merged = out.getFirst();
        assertEquals("W5-13", merged.get("依頼NO"));
        assertEquals("2026/07/01 08:00", merged.get("加工開始日時"));
        assertEquals("2026/07/02 17:00", merged.get("加工終了日時"));
        assertEquals("", merged.get(ResultDispatchSchema.COL_DISPATCH_DATE));
        assertTrue(
                ResultDispatchNormalizer.parseDouble(merged.get(ResultDispatchSchema.COL_DISPATCH_QTY))
                        > 149.9);
    }

    @Test
    void keepsDistinctTasksSeparate() {
        List<String> cols =
                List.of(
                        "依頼NO",
                        ResultDispatchSchema.COL_PROCESS,
                        ResultDispatchSchema.COL_MACHINE,
                        "加工開始日時",
                        "加工終了日時",
                        ResultDispatchSchema.COL_DISPATCH_DATE,
                        ResultDispatchSchema.COL_DISPATCH_QTY);

        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("JR1", "工程A", "M1", "2026/07/01 08:00", "2026/07/01 12:00", "2026/07/01", "10"));
        rows.add(row("JR2", "工程A", "M1", "2026/07/01 08:00", "2026/07/01 12:00", "2026/07/01", "10"));

        assertEquals(2, ResultDispatchTaskSummaryConsolidator.consolidate(cols, rows).size());
    }

    @Test
    void indexDailyRowsByTaskGroupGroupsSameTask() {
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("W5-13", "巻返し", "M1", "2026/07/01 08:00", "2026/07/01 12:00", "2026/07/01", "100"));
        rows.add(row("W5-13", "巻返し", "M1", "2026/07/02 09:00", "2026/07/02 17:00", "2026/07/02", "50"));
        rows.add(row("JR1", "工程A", "M1", "2026/07/01 08:00", "2026/07/01 12:00", "2026/07/01", "10"));

        Map<String, List<Map<String, String>>> indexed =
                ResultDispatchTaskSummaryConsolidator.indexDailyRowsByTaskGroup(rows);

        assertEquals(2, indexed.size());
        String wKey = ResultDispatchTaskSummaryConsolidator.taskGroupKey(rows.getFirst());
        assertEquals(2, indexed.get(wKey).size());
    }

    @Test
    void sortedDailyScheduleRowsOrdersByDispatchDate() {
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("W5-13", "巻返し", "M1", "2026/07/02 09:00", "2026/07/02 17:00", "2026/07/02", "50"));
        rows.add(row("W5-13", "巻返し", "M1", "2026/07/01 08:00", "2026/07/01 12:00", "2026/07/01", "100"));

        List<Map<String, String>> sorted =
                ResultDispatchTaskSummaryConsolidator.sortedDailyScheduleRows(rows);

        assertEquals("2026/07/01", sorted.getFirst().get(ResultDispatchSchema.COL_DISPATCH_DATE));
        assertEquals("2026/07/02", sorted.get(1).get(ResultDispatchSchema.COL_DISPATCH_DATE));
    }

    private static Map<String, String> row(
            String tid,
            String proc,
            String mach,
            String start,
            String end,
            String dispatchDate,
            String qty) {
        Map<String, String> m = new LinkedHashMap<>();
        m.put("依頼NO", tid);
        m.put(ResultDispatchSchema.COL_PROCESS, proc);
        m.put(ResultDispatchSchema.COL_MACHINE, mach);
        m.put("加工開始日時", start);
        m.put("加工終了日時", end);
        m.put(ResultDispatchSchema.COL_DISPATCH_DATE, dispatchDate);
        m.put(ResultDispatchSchema.COL_DISPATCH_QTY, qty);
        return m;
    }
}
