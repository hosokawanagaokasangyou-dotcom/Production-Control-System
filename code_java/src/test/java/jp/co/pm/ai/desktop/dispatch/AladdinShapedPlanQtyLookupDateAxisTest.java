package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class AladdinShapedPlanQtyLookupDateAxisTest {

    @Test
    void distinctPlanDatesFor_returnsPositivePlanCalendarDays() {
        List<String> headers =
                List.of("機械名", "依頼NO", "工程名", "2026/06/09", "2026/06/10", "2026/06/11");
        List<List<String>> rows =
                List.of(List.of("スリット機1\u3000湖南", "E6-1", "スリット", "4200", "4200", "4200"));
        Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                AladdinShapedPlanQtyLookup.buildLookup(headers, rows);

        List<LocalDate> dates =
                AladdinShapedPlanQtyLookup.distinctPlanDatesFor(lookup, "スリット機1\u3000湖南", "E6-1");

        assertEquals(3, dates.size());
        assertEquals(LocalDate.of(2026, 6, 9), dates.get(0));
        assertEquals(LocalDate.of(2026, 6, 11), dates.get(2));
    }

    @Test
    void parsePlanDateColumn_acceptsSlashAndHyphen() {
        assertEquals(LocalDate.of(2026, 6, 11), AladdinShapedPlanQtyLookup.parsePlanDateColumn("2026/06/11"));
        assertEquals(LocalDate.of(2026, 6, 11), AladdinShapedPlanQtyLookup.parsePlanDateColumn("2026-06-11"));
    }
}
