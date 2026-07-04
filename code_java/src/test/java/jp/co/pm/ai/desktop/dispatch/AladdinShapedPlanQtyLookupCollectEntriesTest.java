package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class AladdinShapedPlanQtyLookupCollectEntriesTest {

    private static final List<String> HEADERS =
            List.of(
                    "機械名",
                    "依頼NO",
                    "工程名",
                    "2026/06/09",
                    "2026/06/10",
                    "2026/06/11");

    private static final List<List<String>> ROWS =
            List.of(
                    List.of("スリット機1\u3000湖南", "E6-1", "スリット", "4200", "4200", "4200"),
                    List.of("カレンダー1\u3000湖南", "E6-1", "カレンダー", "1000", "0", "1000"),
                    List.of("スリット機1\u3000湖南", "E6-2", "スリット", "500", "0", "0"));

    @Test
    void collectEntriesForTaskId_returnsAllMachinesAndDates() {
        Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                AladdinShapedPlanQtyLookup.buildLookup(HEADERS, ROWS);

        List<AladdinShapedPlanQtyLookup.PlanEntry> entries =
                AladdinShapedPlanQtyLookup.collectEntriesForTaskId(lookup, "E6-1");

        assertEquals(5, entries.size());
        assertTrue(entries.stream().anyMatch(e -> e.dateYmd().equals("2026/06/09") && e.planMeters() == 4200));
        assertTrue(entries.stream().anyMatch(e -> e.dateYmd().equals("2026/06/09") && e.planMeters() == 1000));
    }

    @Test
    void collectEntriesForTaskIdFromTable_preservesDisplayNames() {
        List<AladdinShapedPlanQtyLookup.PlanEntry> entries =
                AladdinShapedPlanQtyLookup.collectEntriesForTaskIdFromTable(HEADERS, ROWS, "E6-1");

        assertEquals(5, entries.size());
        assertTrue(entries.stream().anyMatch(e -> "スリット".equals(e.processName())));
        assertTrue(entries.stream().anyMatch(e -> e.machineName().contains("カレンダー1")));
    }

    @Test
    void collectEntriesForTaskId_emptyWhenNotFound() {
        Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                AladdinShapedPlanQtyLookup.buildLookup(HEADERS, ROWS);

        assertTrue(AladdinShapedPlanQtyLookup.collectEntriesForTaskId(lookup, "NO-SUCH").isEmpty());
    }

    @Test
    void extractSortedDateColumnHeaders_returnsFirstSevenInOrder() {
        List<String> headers =
                List.of(
                        "機械名",
                        "依頼NO",
                        "2026/06/11",
                        "2026/06/09",
                        "2026/06/10",
                        "換算数量");
        List<String> dates =
                AladdinShapedPlanQtyLookup.extractSortedDateColumnHeaders(headers, 7);
        assertEquals(3, dates.size());
        assertEquals("2026/06/09", dates.get(0));
        assertEquals("2026/06/11", dates.get(2));
    }

    @Test
    void circledSlotColumnLabel_usesCircledDigits() {
        assertEquals("①", AladdinShapedPlanQtyLookup.circledSlotColumnLabel(0));
        assertEquals("⑦", AladdinShapedPlanQtyLookup.circledSlotColumnLabel(6));
    }

    @Test
    void formatPlanDateMetersCell_includesDateAndUnit() {
        assertEquals("7/3 100m", AladdinShapedPlanQtyLookup.formatPlanDateMetersCell("2026/07/03", 100));
        assertEquals("", AladdinShapedPlanQtyLookup.formatPlanDateMetersCell("2026/07/03", 0));
    }

    @Test
    void aggregatePlanMetersByEntryDates_ordersByDatePerTask() {
        List<AladdinShapedPlanQtyLookup.PlanEntry> entries =
                List.of(
                        new AladdinShapedPlanQtyLookup.PlanEntry("M1", "スリット", "2026/06/11", 1000),
                        new AladdinShapedPlanQtyLookup.PlanEntry("M2", "カレンダー", "2026/06/09", 4200));
        List<String> slots =
                AladdinShapedPlanQtyLookup.aggregatePlanMetersByEntryDates(entries, 7);
        assertEquals("6/9 4200m", slots.get(0));
        assertEquals("6/11 1000m", slots.get(1));
        assertEquals("", slots.get(2));
    }

    @Test
    void isPlanDateColumnHeader_acceptsFlexibleMonthDay() {
        assertTrue(AladdinShapedPlanQtyLookup.isPlanDateColumnHeader("2026/6/9"));
        assertTrue(AladdinShapedPlanQtyLookup.isPlanDateColumnHeader("2026/06/09"));
        assertFalse(AladdinShapedPlanQtyLookup.isPlanDateColumnHeader("依頼NO"));
    }

    @Test
    void collectEntries_matchesTaskIdWithNormalization() {
        List<String> headers =
                List.of("機械名", "依頼No", "工程名", "2026/6/9", "2026/6/10");
        List<List<String>> rows =
                List.of(List.of("スリット機1", "e6-1", "スリット", "100", "200"));
        List<AladdinShapedPlanQtyLookup.PlanEntry> entries =
                AladdinShapedPlanQtyLookup.collectEntriesForTaskIdFromTable(headers, rows, "E6-1");
        assertEquals(2, entries.size());
    }

    @Test
    void aggregatePlanMetersByDateSlots_sumsSameDayAcrossMachines() {
        List<AladdinShapedPlanQtyLookup.PlanEntry> entries =
                AladdinShapedPlanQtyLookup.collectEntriesForTaskIdFromTable(HEADERS, ROWS, "E6-1");
        List<String> slots =
                AladdinShapedPlanQtyLookup.aggregatePlanMetersByDateSlots(
                        entries, List.of("2026/06/09", "2026/06/10", "2026/06/11"), 7);
        assertEquals("6/9 5200m", slots.get(0));
        assertEquals("6/10 4200m", slots.get(1));
    }
}
