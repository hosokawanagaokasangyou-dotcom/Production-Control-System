package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class DispatchInteractiveDateAxisTest {

    @Test
    void computeInclusiveRange_extendsPastJsonMaxWhenMetaMissOnLastDay() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("E6-1", "スリット機1\u3000湖南", "2026/06/10", 4200, ""));
        rows.add(row("E6-1", "スリット機1\u3000湖南", "2026/06/12", 4200, ""));

        ResultDispatchDocument doc = new ResultDispatchDocument(cols, rows);
        List<LocalDate> axis =
                DispatchInteractiveDateAxis.computeInclusiveRange(doc, Map.of(), List.of());

        assertTrue(axis.contains(LocalDate.of(2026, 6, 13)));
        assertEquals(
                LocalDate.of(2026, 6, 19),
                axis.getLast(),
                "6/12 meta_miss から 7 暦日の余白");
    }

    @Test
    void computeInclusiveRange_includesShortfallAndAladdinDates() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        List<Map<String, String>> rows = List.of(row("E6-1", "スリット機1\u3000湖南", "2026/06/08", 4200, "2026/06/08 08:55"));
        ResultDispatchDocument doc = new ResultDispatchDocument(cols, rows);

        List<String> headers = List.of("機械名", "依頼NO", "工程名", "2026/06/10", "2026/06/11");
        List<List<String>> aladdinRows =
                List.of(List.of("スリット機1\u3000湖南", "E6-1", "スリット", "4200", "4200"));
        var lookup = AladdinShapedPlanQtyLookup.buildLookup(headers, aladdinRows);

        List<DispatchTrialShortages.DispatchQtyShortfallRow> shortfalls =
                List.of(
                        new DispatchTrialShortages.DispatchQtyShortfallRow(
                                "E6-1",
                                "スリット機1\u3000湖南",
                                "2026-06-12",
                                4200,
                                0,
                                4200,
                                "test"));

        List<LocalDate> axis =
                DispatchInteractiveDateAxis.computeInclusiveRange(doc, lookup, shortfalls);

        assertTrue(axis.contains(LocalDate.of(2026, 6, 10)));
        assertTrue(axis.contains(LocalDate.of(2026, 6, 11)));
        assertTrue(axis.contains(LocalDate.of(2026, 6, 12)));
        assertTrue(axis.contains(LocalDate.of(2026, 6, 13)));
    }

    private static Map<String, String> row(
            String tid, String machine, String dispatchDate, double planQty, String startDt) {
        Map<String, String> m = new LinkedHashMap<>();
        m.put("依頼NO", tid);
        m.put(ResultDispatchSchema.COL_MACHINE, machine);
        m.put(ResultDispatchSchema.COL_DISPATCH_DATE, dispatchDate);
        m.put(ResultDispatchSchema.COL_DISPATCH_QTY, Double.toString(planQty));
        m.put("加工開始日時", startDt);
        m.put("加工完了日", "2026/06/11");
        m.put("指定納期", "2026/06/12");
        return m;
    }
}
