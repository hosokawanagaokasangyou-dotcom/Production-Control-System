package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.ProcessingTrendChartSupport.NiceRange;

class ProcessingTrendChartSupportTest {

    @Test
    void niceRangeZeroOrNegativeFallsBackToTen() {
        assertEquals(new NiceRange(10, 2), ProcessingTrendChartSupport.niceRange(0));
        assertEquals(new NiceRange(10, 2), ProcessingTrendChartSupport.niceRange(-5));
        assertEquals(new NiceRange(10, 2), ProcessingTrendChartSupport.niceRange(Double.NaN));
    }

    @Test
    void niceRangeUsesOneTwoFiveSteps() {
        NiceRange r = ProcessingTrendChartSupport.niceRange(4_750);
        assertEquals(1_000, r.tickUnit());
        assertTrue(r.upperBound() > 4_750);
        assertEquals(0, r.upperBound() % r.tickUnit());

        r = ProcessingTrendChartSupport.niceRange(84_000);
        assertEquals(20_000, r.tickUnit());
        assertEquals(100_000, r.upperBound());

        r = ProcessingTrendChartSupport.niceRange(12);
        assertEquals(2, r.tickUnit());
        assertEquals(14, r.upperBound());
    }

    @Test
    void niceRangeKeepsHeadroomAboveMax() {
        // 最大値がちょうど目盛に乗るときも 1 目盛ぶん余白を残す
        NiceRange r = ProcessingTrendChartSupport.niceRange(6_000);
        assertEquals(1_000, r.tickUnit());
        assertEquals(7_000, r.upperBound());
    }

    @Test
    void labelledDatesAllWithinAMonth() {
        List<LocalDate> dates = days(LocalDate.of(2026, 9, 1), 30);
        assertEquals(30, ProcessingTrendChartSupport.labelledDates(dates).size());
    }

    @Test
    void labelledDatesWeeklyBeyondAMonth() {
        LocalDate start = LocalDate.of(2026, 7, 5);
        List<LocalDate> dates = days(start, 90);
        Set<LocalDate> labelled = ProcessingTrendChartSupport.labelledDates(dates);
        // 先頭 7/5(日) は翌日の月曜 7/6 と近接するので月曜側だけ残る
        assertFalse(labelled.contains(start));
        assertTrue(labelled.contains(LocalDate.of(2026, 7, 6)), "先頭直後の月曜");
        assertTrue(labelled.contains(LocalDate.of(2026, 8, 1)), "月初（土）は隣の月曜 8/3 より優先");
        assertFalse(labelled.contains(LocalDate.of(2026, 8, 3)));
        assertTrue(labelled.contains(LocalDate.of(2026, 9, 1)), "月初（火）は前日の月曜 8/31 より優先");
        assertFalse(labelled.contains(LocalDate.of(2026, 8, 31)));
        for (LocalDate d : labelled) {
            assertTrue(
                    d.getDayOfMonth() == 1 || d.getDayOfWeek() == DayOfWeek.MONDAY,
                    "月曜・月初以外にラベル: " + d);
        }
        assertMinGap(labelled, ProcessingTrendChartSupport.LABEL_MIN_GAP_DAYS);
        assertTrue(labelled.size() <= 90 / 7 + 6, "ラベル数: " + labelled.size());
        assertFalse(labelled.contains(LocalDate.of(2026, 7, 8)), "水曜はラベル無し");
    }

    @Test
    void labelledDatesKeepsRangeStartWhenNothingNearby() {
        // 7/8(水) 開始: 次の月曜 7/13 まで 5 日空くので先頭日ラベルが残る
        LocalDate start = LocalDate.of(2026, 7, 8);
        Set<LocalDate> labelled = ProcessingTrendChartSupport.labelledDates(days(start, 60));
        assertTrue(labelled.contains(start));
        assertTrue(labelled.contains(LocalDate.of(2026, 7, 13)));
    }

    private static void assertMinGap(Set<LocalDate> labelled, int minGapDays) {
        LocalDate prev = null;
        for (LocalDate d : labelled) {
            if (prev != null) {
                assertTrue(
                        java.time.temporal.ChronoUnit.DAYS.between(prev, d) >= minGapDays,
                        prev + " と " + d + " が近接");
            }
            prev = d;
        }
    }

    @Test
    void labelledDatesFirstAndFifteenthForLongRanges() {
        LocalDate start = LocalDate.of(2026, 1, 3);
        List<LocalDate> dates = days(start, 365);
        Set<LocalDate> labelled = ProcessingTrendChartSupport.labelledDates(dates);
        for (LocalDate d : labelled) {
            assertTrue(d.equals(start) || d.getDayOfMonth() == 1 || d.getDayOfMonth() == 15, "1 日・15 日以外: " + d);
        }
        assertMinGap(labelled, ProcessingTrendChartSupport.LABEL_MIN_GAP_DAYS_SEMIMONTHLY);
        assertTrue(labelled.size() <= 26);

        // 先頭日が 15 日の直前（11/7 と 11/15）: 3px/日では重なるので先頭日を落として 15 日を残す
        LocalDate nearFifteenth = LocalDate.of(2025, 11, 7);
        Set<LocalDate> l2 = ProcessingTrendChartSupport.labelledDates(days(nearFifteenth, 365));
        assertFalse(l2.contains(nearFifteenth));
        assertTrue(l2.contains(LocalDate.of(2025, 11, 15)));
        assertMinGap(l2, ProcessingTrendChartSupport.LABEL_MIN_GAP_DAYS_SEMIMONTHLY);
    }

    @Test
    void gapsShrinkWithMoreDays() {
        assertEquals(8, ProcessingTrendChartSupport.categoryGapFor(31));
        assertEquals(4, ProcessingTrendChartSupport.categoryGapFor(62));
        assertEquals(2, ProcessingTrendChartSupport.categoryGapFor(90));
        assertEquals(1, ProcessingTrendChartSupport.categoryGapFor(365));
        assertEquals(1, ProcessingTrendChartSupport.barGapFor(31));
        assertEquals(0, ProcessingTrendChartSupport.barGapFor(90));
    }

    @Test
    void categoryLabel_zeroPadsSingleMonthDaysForLexicographicOrder() {
        assertEquals("01", ProcessingTrendChartSupport.categoryLabel(LocalDate.of(2026, 9, 1), true, true));
        assertEquals("07", ProcessingTrendChartSupport.categoryLabel(LocalDate.of(2026, 9, 7), true, true));
        assertEquals("10", ProcessingTrendChartSupport.categoryLabel(LocalDate.of(2026, 9, 10), true, true));
        // ゼロ埋めなら "07" < "10"（文字列比較でも日付順）
        assertTrue(
                ProcessingTrendChartSupport.categoryLabel(LocalDate.of(2026, 9, 7), true, true)
                        .compareTo(
                                ProcessingTrendChartSupport.categoryLabel(LocalDate.of(2026, 9, 10), true, true))
                        < 0);
    }

    @Test
    void cumPointInclusion_actualThroughToday_projectedFromToday() {
        LocalDate today = LocalDate.of(2026, 9, 8);
        assertTrue(ProcessingTrendChartSupport.includeActualCumPoint(LocalDate.of(2026, 9, 7), today));
        assertTrue(ProcessingTrendChartSupport.includeActualCumPoint(today, today));
        assertFalse(ProcessingTrendChartSupport.includeActualCumPoint(LocalDate.of(2026, 9, 9), today));

        assertFalse(ProcessingTrendChartSupport.includeProjectedCumPoint(LocalDate.of(2026, 9, 6), today));
        assertFalse(ProcessingTrendChartSupport.includeProjectedCumPoint(LocalDate.of(2026, 9, 7), today));
        assertTrue(ProcessingTrendChartSupport.includeProjectedCumPoint(today, today));
        assertTrue(ProcessingTrendChartSupport.includeProjectedCumPoint(LocalDate.of(2026, 9, 9), today));
    }

    @Test
    void monthBands_splitsContiguousDatesByYearMonth() {
        List<LocalDate> dates =
                List.of(
                        LocalDate.of(2026, 8, 30),
                        LocalDate.of(2026, 8, 31),
                        LocalDate.of(2026, 9, 1),
                        LocalDate.of(2026, 9, 2),
                        LocalDate.of(2026, 10, 1));
        var bands = ProcessingTrendChartSupport.monthBands(dates);
        assertEquals(3, bands.size());
        assertEquals(YearMonth.of(2026, 8), bands.get(0).month());
        assertEquals(0, bands.get(0).fromIndexInclusive());
        assertEquals(1, bands.get(0).toIndexInclusive());
        assertEquals(YearMonth.of(2026, 9), bands.get(1).month());
        assertEquals(2, bands.get(1).fromIndexInclusive());
        assertEquals(3, bands.get(1).toIndexInclusive());
        assertEquals(YearMonth.of(2026, 10), bands.get(2).month());
        assertEquals(4, bands.get(2).fromIndexInclusive());
        assertEquals(4, bands.get(2).toIndexInclusive());
        assertEquals(List.of(), ProcessingTrendChartSupport.monthBands(List.of()));
        assertEquals("9月", ProcessingTrendChartSupport.monthBandLabel(YearMonth.of(2026, 9), false));
        assertEquals("2026/9", ProcessingTrendChartSupport.monthBandLabel(YearMonth.of(2026, 9), true));
    }

    private static List<LocalDate> days(LocalDate start, int n) {
        List<LocalDate> out = new ArrayList<>(n);
        for (int i = 0; i < n; i++) {
            out.add(start.plusDays(i));
        }
        return out;
    }
}
