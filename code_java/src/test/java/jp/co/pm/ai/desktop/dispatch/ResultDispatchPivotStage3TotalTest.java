package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

class ResultDispatchPivotStage3TotalTest {

    @Test
    void sumActualForProfileIgnoresExtraCalendarDays() {
        Map<String, String> profile =
                Map.of(
                        ResultDispatchSchema.COL_PROCESS,
                        "分割",
                        ResultDispatchSchema.COL_MACHINE,
                        "スリット機1　湖南",
                        "依頼NO",
                        "V6-2",
                        "換算数量",
                        "10000");
        List<Map<String, String>> rows =
                List.of(
                        row("V6-2", "2026-06-12", "7200", "7200"),
                        row("V6-2", "2026-06-15", "2800", "2800"));
        double total =
                ResultDispatchPivot.sumActualQuantityForProfileForWideMerge(
                        rows,
                        profile,
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        assertEquals(10000.0, total, 1e-6);
        assertEquals(
                7200.0,
                ResultDispatchPivot.sumActualQuantityForProfileAndDateForWideMerge(
                        rows,
                        profile,
                        LocalDate.parse("2026-06-12"),
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS),
                1e-6);
    }

    private static Map<String, String> row(
            String taskId, String dispatchDate, String planQty, String actualQty) {
        Map<String, String> m = new LinkedHashMap<>();
        m.put("依頼NO", taskId);
        m.put(ResultDispatchSchema.COL_PROCESS, "分割");
        m.put(ResultDispatchSchema.COL_MACHINE, "スリット機1　湖南");
        m.put("配台日", dispatchDate);
        m.put(ResultDispatchSchema.COL_DISPATCH_QTY, planQty);
        m.put(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL, actualQty);
        return m;
    }
}
