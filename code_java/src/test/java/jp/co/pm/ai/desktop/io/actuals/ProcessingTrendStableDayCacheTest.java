package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;
import java.util.TreeMap;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.ActualSource;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Filter;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.PlanSource;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Result;

class ProcessingTrendStableDayCacheTest {

    private ProcessingTrendStableDayCache cache;

    @BeforeEach
    void setUp() {
        cache = new ProcessingTrendStableDayCache();
        cache.clearAll();
    }

    @Test
    void stableEndInclusive_isTodayMinus30() {
        LocalDate today = LocalDate.of(2026, 9, 7);
        assertEquals(LocalDate.of(2026, 8, 8), ProcessingTrendStableDayCache.stableEndInclusive(today));
    }

    @Test
    void secondAggregate_skipsStableDaysButMatchesFirst() {
        LocalDate today = LocalDate.of(2026, 9, 7);
        LocalDate from = LocalDate.of(2026, 6, 1);
        LocalDate to = LocalDate.of(2026, 9, 7);
        Filter filter =
                new Filter(from, to, ActualSource.DAILY_REPORT, PlanSource.ALADDIN, null, null, 30);

        ActualsSnapshot daily = dailyRows();
        ActualsSnapshot detail = detailRows();
        AladdinSnapshot aladdin = emptyAladdin();
        DispatchSnapshot dispatch = new DispatchSnapshot(List.of(), List.of());
        String pathId = "test-path-identity";

        Result first =
                ProcessingTrendAggregator.aggregate(
                        daily, detail, aladdin, dispatch, filter, today, cache, pathId);
        Result second =
                ProcessingTrendAggregator.aggregate(
                        daily, detail, aladdin, dispatch, filter, today, cache, pathId);

        assertFalse(first.isEmpty());
        assertEquals(first.days().size(), second.days().size());
        assertEquals(first.actualTotalM(), second.actualTotalM(), 1e-6);
        assertEquals(first.planTotalM(), second.planTotalM(), 1e-6);
        assertEquals(first.compareActualTotalM(), second.compareActualTotalM(), 1e-6);
        for (int i = 0; i < first.days().size(); i++) {
            assertEquals(first.days().get(i).actualM(), second.days().get(i).actualM(), 1e-6);
            assertEquals(first.days().get(i).planM(), second.days().get(i).planM(), 1e-6);
        }
    }

    @Test
    void tryFill_failsWhenGapInStableRange() {
        LocalDate today = LocalDate.of(2026, 9, 7);
        LocalDate from = LocalDate.of(2026, 7, 1);
        LocalDate to = LocalDate.of(2026, 9, 7);
        TreeMap<LocalDate, double[]> byDay = new TreeMap<>();
        for (LocalDate d = from; !d.isAfter(to); d = d.plusDays(1)) {
            byDay.put(d, new double[3]);
        }
        var key =
                new ProcessingTrendStableDayCache.SeriesKey(
                        "p", ActualSource.DAILY_REPORT, PlanSource.ALADDIN, "", "");
        // only one day cached
        TreeMap<LocalDate, double[]> one = new TreeMap<>();
        one.put(LocalDate.of(2026, 7, 1), new double[] {10, 0, 1});
        cache.putStableDays(key, one, today);
        assertFalse(cache.tryFillStableDays(key, from, to, today, byDay));
    }

    @Test
    void tryFill_succeedsWhenAllStableDaysPresentIncludingZeros() {
        LocalDate today = LocalDate.of(2026, 9, 7);
        LocalDate stableEnd = ProcessingTrendStableDayCache.stableEndInclusive(today);
        LocalDate from = stableEnd.minusDays(2);
        LocalDate to = today;
        TreeMap<LocalDate, double[]> byDay = new TreeMap<>();
        for (LocalDate d = from; !d.isAfter(to); d = d.plusDays(1)) {
            byDay.put(d, new double[3]);
        }
        var key =
                new ProcessingTrendStableDayCache.SeriesKey(
                        "p", ActualSource.DETAIL, PlanSource.DISPATCH, "", "");
        TreeMap<LocalDate, double[]> seed = new TreeMap<>();
        for (LocalDate d = from; !d.isAfter(stableEnd); d = d.plusDays(1)) {
            seed.put(d, new double[] {d.equals(from) ? 5.0 : 0.0, 0, 0});
        }
        cache.putStableDays(key, seed, today);
        assertTrue(cache.tryFillStableDays(key, from, to, today, byDay));
        assertEquals(5.0, byDay.get(from)[0], 1e-9);
        assertEquals(0.0, byDay.get(stableEnd)[0], 1e-9);
        // recent days untouched (still 0)
        assertEquals(0.0, byDay.get(today)[0], 1e-9);
    }

    private static ActualsSnapshot dailyRows() {
        return new ActualsSnapshot(
                List.of("加工日付", "機械名", "工程名", "実加工量"),
                List.of(
                        List.of("2026/06/15", "M1", "P1", "100"),
                        List.of("2026/08/01", "M1", "P1", "200"),
                        List.of("2026/09/01", "M1", "P1", "50"),
                        List.of("2026/09/07", "M1", "P1", "10")));
    }

    private static ActualsSnapshot detailRows() {
        return new ActualsSnapshot(
                List.of("加工日", "機械名", "工程名", "実加工数"),
                List.of(
                        List.of("2026/06/15", "M1", "P1", "90"),
                        List.of("2026/08/01", "M1", "P1", "180"),
                        List.of("2026/09/01", "M1", "P1", "40"),
                        List.of("2026/09/07", "M1", "P1", "8")));
    }

    private static AladdinSnapshot emptyAladdin() {
        return new AladdinSnapshot(List.of("機械名", "工程名", "依頼NO"), List.of());
    }
}
