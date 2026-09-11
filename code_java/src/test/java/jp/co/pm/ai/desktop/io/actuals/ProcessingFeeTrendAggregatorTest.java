package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.actuals.JuchuProcessingFeeRateLoader.FeeInfo;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.DayPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.QuantityLine;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.RequestPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.Result;

class ProcessingFeeTrendAggregatorTest {

    private static FeeInfo ah(double rate) {
        return new FeeInfo(rate, null, "", null, null, null);
    }

    private static FeeInfo ao(double aoYen, double orderFinalM, String processContent) {
        return new FeeInfo(null, aoYen, processContent, null, null, orderFinalM);
    }

    /** 日報完了行（加工日付＋終了時間）。 */
    private static QuantityLine done(
            LocalDate day, String req, double meters, String process, int hour, int minute) {
        return new QuantityLine(
                day, req, meters, process, LocalDateTime.of(day, java.time.LocalTime.of(hour, minute)));
    }

    @Test
    void aoIsAuthority_actualPlusRemainEqualsAo() {
        // AO=1000, 受注100m → 10円/m。実績60 → 600、未了40 → 400
        LocalDate d = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees = Map.of("R", ao(1_000.0, 100.0, "最終"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(done(d, "R", 60, "最終", 15, 0)),
                        List.of(),
                        fees,
                        d,
                        d.plusDays(10),
                        d);
        RequestPoint row = r.requests().get(0);
        assertEquals(10.0, row.rateYenPerM(), 1e-9);
        assertEquals(60.0, row.actualMeters(), 1e-9);
        assertEquals(40.0, row.remainMeters(), 1e-9);
        assertEquals(600.0, row.actualYen(), 1e-9);
        assertEquals(400.0, row.planYen(), 1e-9);
        assertEquals(1_000.0, row.aoYen(), 1e-9);
        assertEquals(1_000.0, row.actualYen() + row.planYen(), 1e-9);
    }

    @Test
    void multiProcessUsesOnlyLatestFinishedProcessMeters() {
        LocalDate d0 = LocalDate.of(2026, 9, 1);
        LocalDate d1 = LocalDate.of(2026, 9, 2);
        // 工程A より B の終了が遅い → B のみ。受注50m・AO10000 → 200円/m
        Map<String, FeeInfo> fees = Map.of("R", ao(10_000.0, 50.0, "A,B"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                done(d0, "R", 100, "A", 10, 0),
                                done(d0, "R", 30, "B", 12, 0),
                                done(d1, "R", 20, "B", 11, 0)),
                        List.of(),
                        fees,
                        d0,
                        d1,
                        d1);
        RequestPoint row = r.requests().get(0);
        assertEquals(200.0, row.rateYenPerM(), 1e-6);
        assertEquals(50.0, row.actualMeters(), 1e-6);
        assertEquals(0.0, row.remainMeters(), 1e-6);
        assertEquals(10_000.0, row.actualYen(), 1e-6);
        assertEquals(0.0, row.planYen(), 1e-6);
        assertEquals(6_000.0, r.days().get(0).actualYen(), 1e-6);
        assertEquals(4_000.0, r.days().get(1).actualYen(), 1e-6);
    }

    @Test
    void missingEndTimeMeansIncomplete_notCountedEvenWithProcessContent() {
        LocalDate d = LocalDate.of(2026, 7, 15);
        LocalDate from = LocalDate.of(2026, 7, 1);
        LocalDate to = LocalDate.of(2026, 7, 31);
        Map<String, FeeInfo> fees =
                Map.of("C7-10", new FeeInfo(null, 112_000.0, "SEC,増刷", 2026, 7, 4_000.0));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                new QuantityLine(d, "C7-10", 2_000, "SEC"),
                                new QuantityLine(d, "C7-10", 2_000, "SEC"),
                                new QuantityLine(d, "C7-10", 0, "増刷")),
                        List.of(),
                        fees,
                        from,
                        to,
                        LocalDate.of(2026, 9, 11));
        // 終了時間なし＝未完了 → 実績0。受注月があるので未了に AO 全額
        assertEquals(1, r.requests().size());
        assertEquals(0.0, r.requests().get(0).actualMeters(), 1e-6);
        assertEquals(0.0, r.requests().get(0).actualYen(), 1e-6);
        assertEquals(112_000.0, r.requests().get(0).planYen(), 1e-6);
    }

    @Test
    void incompleteRowsExcluded_evenWhenSameProcessHasCompletedRows() {
        LocalDate day = LocalDate.of(2026, 7, 15);
        Map<String, FeeInfo> fees = Map.of("R", ao(1_000.0, 100.0, "SEC"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                done(day, "R", 40, "SEC", 14, 0),
                                new QuantityLine(day, "R", 60, "SEC"), // 終了時間なし＝未完了
                                done(day, "R", 10, "裁断", 10, 0)),
                        List.of(),
                        fees,
                        day,
                        day,
                        day);
        assertEquals(40.0, r.requests().get(0).actualMeters(), 1e-6);
        assertEquals(400.0, r.requests().get(0).actualYen(), 1e-6);
    }

    @Test
    void emptyProcessContent_multiProcess_usesLatestFinishedAt_withinPeriod() {
        // C7-10 相当: 期間内・加工内容空・SEC+裁断 → 終了時間が遅い工程（SEC）の完了出来高のみ
        LocalDate day = LocalDate.of(2026, 7, 15);
        LocalDate from = LocalDate.of(2026, 7, 1);
        LocalDate to = LocalDate.of(2026, 7, 31);
        Map<String, FeeInfo> fees =
                Map.of("C7-10", new FeeInfo(null, 112_000.0, "", 2026, 7, 4_000.0));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                done(day, "C7-10", 2_000, "SEC", 14, 30),
                                done(day, "C7-10", 2_000, "SEC", 16, 0),
                                done(day, "C7-10", 2_000, "裁断", 10, 0)),
                        List.of(),
                        fees,
                        from,
                        to,
                        LocalDate.of(2026, 9, 11));
        RequestPoint row = r.requests().get(0);
        assertEquals(4_000.0, row.actualMeters(), 1e-6);
        assertEquals(0.0, row.remainMeters(), 1e-6);
        assertEquals(112_000.0, row.actualYen(), 1e-6);
        assertEquals(112_000.0, r.days().stream().mapToDouble(DayPoint::actualYen).sum(), 1e-6);
    }

    @Test
    void finishedAtOverridesProcessContentLastToken() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees = Map.of("R", ao(1_000.0, 100.0, "A,B"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(done(d, "R", 100, "A", 17, 0), done(d, "R", 50, "B", 12, 0)),
                        List.of(),
                        fees,
                        d,
                        d,
                        d);
        assertEquals(100.0, r.requests().get(0).actualMeters(), 1e-6);
        assertEquals(1_000.0, r.requests().get(0).actualYen(), 1e-6);
        assertEquals(0.0, r.requests().get(0).planYen(), 1e-6);
    }

    @Test
    void laterCalendarDayWinsEvenIfClockIsEarlier() {
        LocalDate d0 = LocalDate.of(2026, 7, 14);
        LocalDate d1 = LocalDate.of(2026, 7, 15);
        Map<String, FeeInfo> fees = Map.of("R", ao(2_000.0, 100.0, ""));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(done(d0, "R", 80, "裁断", 20, 0), done(d1, "R", 100, "SEC", 9, 0)),
                        List.of(),
                        fees,
                        LocalDate.of(2026, 7, 1),
                        LocalDate.of(2026, 7, 31),
                        LocalDate.of(2026, 9, 11));
        assertEquals(100.0, r.requests().get(0).actualMeters(), 1e-6);
    }

    @Test
    void requestActualUsesOutsidePeriodMeters_dailyBarsStayInPeriod() {
        LocalDate from = LocalDate.of(2026, 7, 1);
        LocalDate to = LocalDate.of(2026, 7, 31);
        LocalDate today = LocalDate.of(2026, 9, 11);
        LocalDate outDay = LocalDate.of(2024, 7, 15);
        Map<String, FeeInfo> fees = Map.of("C7-10", ao(112_000.0, 4_000.0, "SEC"));
        List<QuantityLine> actual =
                List.of(
                        done(outDay, "C7-10", 2_000, "SEC", 10, 0),
                        done(outDay, "C7-10", 2_000, "SEC", 11, 0));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        actual, List.of(), fees, from, to, today);
        RequestPoint row = r.requests().get(0);
        assertEquals(4_000.0, row.actualMeters(), 1e-6);
        assertEquals(0.0, row.remainMeters(), 1e-6);
        assertEquals(112_000.0, row.actualYen(), 1e-6);
        assertEquals(0.0, r.days().stream().mapToDouble(DayPoint::actualYen).sum(), 1e-9);
    }

    @Test
    void pastPeriod_noDailyRemainBars_kpiRemainFromAlloc() {
        LocalDate from = LocalDate.of(2026, 7, 1);
        LocalDate to = LocalDate.of(2026, 7, 31);
        LocalDate today = LocalDate.of(2026, 9, 11);
        Map<String, FeeInfo> fees = Map.of("R", ao(1_000.0, 100.0, "最終"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(done(LocalDate.of(2026, 7, 10), "R", 60, "最終", 15, 0)),
                        List.of(),
                        fees,
                        from,
                        to,
                        today);
        assertEquals(31, r.days().size());
        assertEquals(0.0, r.days().stream().mapToDouble(DayPoint::planYen).sum(), 1e-9);
        DayPoint last = r.days().get(30);
        assertEquals(600.0, last.actualCumYen(), 1e-6);
        assertEquals(0.0, last.planCumYen(), 1e-6);
        assertEquals(600.0, last.projectedCumYen(), 1e-6);
        assertEquals(400.0, r.planTotalYen(), 1e-6);
        assertEquals(400.0, r.requests().get(0).planYen(), 1e-6);
    }

    @Test
    void idleOrderMonthRow_allAoGoesToRemain() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 30);
        Map<String, FeeInfo> fees =
                Map.of("IDLE", new FeeInfo(null, 4_000.0, "X", 2026, 9, 200.0));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(), List.of(), fees, from, to, from);
        RequestPoint idle = r.requests().get(0);
        assertEquals(0.0, idle.actualYen(), 1e-9);
        assertEquals(4_000.0, idle.planYen(), 1e-9);
        assertEquals(4_000.0, idle.aoYen(), 1e-9);
        assertEquals(200.0, idle.remainMeters(), 1e-9);
    }

    @Test
    void currentMonth_withoutPlan_stillAnchorsRemainOnTomorrow() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 30);
        LocalDate today = LocalDate.of(2026, 9, 11);
        Map<String, FeeInfo> fees = Map.of("R", ao(1_000.0, 100.0, "最終"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(done(LocalDate.of(2026, 9, 5), "R", 60, "最終", 12, 0)),
                        List.of(),
                        fees,
                        from,
                        to,
                        today);
        DayPoint tomorrow =
                r.days().stream()
                        .filter(d -> d.date().equals(today.plusDays(1)))
                        .findFirst()
                        .orElseThrow();
        assertEquals(400.0, tomorrow.planYen(), 1e-6);
    }

    @Test
    void aggregatesDailyAndCumulativeYen_ahFallback() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 3);
        LocalDate today = LocalDate.of(2026, 9, 2);
        Map<String, FeeInfo> fees =
                Map.of(
                        "A", new FeeInfo(10.0, null, "", null, null, 9.0),
                        "B", new FeeInfo(20.0, null, "", null, null, 3.0));
        List<QuantityLine> actual =
                List.of(
                        done(from, "A", 5, "P", 10, 0),
                        done(today, "B", 2, "P", 11, 0),
                        done(to, "A", 1, "P", 12, 0));
        List<QuantityLine> plan = List.of(new QuantityLine(to, "B", 1, "P"));

        Result r = ProcessingFeeTrendAggregator.aggregate(actual, plan, fees, from, to, today);
        assertEquals(3, r.days().size());
        DayPoint d0 = r.days().get(0);
        assertEquals(50.0, d0.actualYen(), 1e-6);
        RequestPoint a =
                r.requests().stream().filter(x -> "A".equals(x.requestNo())).findFirst().orElseThrow();
        assertEquals(6.0, a.actualMeters(), 1e-9);
        assertEquals(3.0, a.remainMeters(), 1e-9);
        assertEquals(60.0, a.actualYen(), 1e-9);
        assertEquals(30.0, a.planYen(), 1e-9);
    }

    @Test
    void missingRateCountsAsZeroYen() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(done(d, "NO-RATE", 100, "P", 10, 0)),
                        List.of(),
                        Map.of("OTHER", ah(1.0)),
                        d,
                        d,
                        d);
        assertEquals(0.0, r.days().get(0).actualYen(), 1e-9);
        assertEquals(1, r.missingRateLines());
        assertEquals(0, r.actualLinesCounted());
        assertEquals(true, r.requests().get(0).rateMissing());
    }

    @Test
    void withLeadingTotalRow_aoEqualsActualPlusRemain() {
        List<RequestPoint> rows =
                List.of(
                        new RequestPoint("A", 10.0, 1_000.0, false, 60, 40, 600, 400),
                        new RequestPoint("B", 20.0, 2_000.0, false, 0, 100, 0, 2_000));
        List<RequestPoint> withTotal = ProcessingFeeTrendAggregator.withLeadingTotalRow(rows);
        RequestPoint tot = withTotal.get(0);
        assertEquals(3_000.0, tot.aoYen(), 1e-9);
        assertEquals(3_000.0, tot.actualYen() + tot.planYen(), 1e-9);
        assertEquals(3_000.0, ProcessingFeeTrendAggregator.sumOrderAoYen(rows), 1e-9);
    }

    @Test
    void includeOrderMonth_actualPlusRemainMatchesAoSum() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 30);
        Map<String, FeeInfo> fees =
                Map.of(
                        "ACT",
                        new FeeInfo(null, 1_000.0, "A", 2026, 9, 100.0),
                        "IDLE",
                        new FeeInfo(null, 4_796_700.0, "B", 2026, 9, 100.0),
                        "OTHER",
                        new FeeInfo(null, 9_999.0, "C", 2026, 8, 50.0));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(done(from, "ACT", 10, "A", 10, 0)),
                        List.of(),
                        fees,
                        from,
                        to,
                        from);
        assertEquals(2, r.requests().size());
        double orderAo = ProcessingFeeTrendAggregator.sumOrderAoYen(r.requests());
        assertEquals(4_797_700.0, orderAo, 1e-6);
        double allocated =
                r.requests().stream().mapToDouble(x -> x.actualYen() + x.planYen()).sum();
        assertEquals(4_797_700.0, allocated, 1e-6);
        RequestPoint act =
                r.requests().stream().filter(x -> "ACT".equals(x.requestNo())).findFirst().orElseThrow();
        assertEquals(100.0, act.actualYen(), 1e-6);
        assertEquals(900.0, act.planYen(), 1e-6);
    }

    @Test
    void latestFinishedProcessWins_overProcessContentLastToken() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees = Map.of("E9-2", ao(5_000.0, 25.0, "スリット,E9-2"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                done(d, "E9-2", 40, "スリット", 10, 0),
                                done(d, "E9-2", 25, "E9-2", 16, 0)),
                        List.of(),
                        fees,
                        d,
                        d,
                        d);
        assertEquals(25.0, r.requests().get(0).actualMeters(), 1e-6);
        assertEquals(200.0, r.requests().get(0).rateYenPerM(), 1e-6);
        assertEquals(5_000.0, r.requests().get(0).actualYen(), 1e-6);
        assertEquals(0.0, r.requests().get(0).planYen(), 1e-6);
    }

    @Test
    void missingEndTimeWithMultipleProcesses_countsAsIncomplete() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees = Map.of("R", ao(9_000.0, 30.0, ""));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                new QuantityLine(d, "R", 10, "A"),
                                new QuantityLine(d, "R", 20, "B")),
                        List.of(),
                        fees,
                        d,
                        d,
                        d);
        assertTrue(
                r.requests().isEmpty()
                        || (r.requests().get(0).actualMeters() == 0.0
                                && r.requests().get(0).actualYen() == 0.0));
    }
}
