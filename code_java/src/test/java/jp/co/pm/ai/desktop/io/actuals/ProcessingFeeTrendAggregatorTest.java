package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
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
        return new FeeInfo(rate, null, "");
    }

    private static FeeInfo ao(double aoYen, String processContent) {
        return new FeeInfo(null, aoYen, processContent);
    }

    @Test
    void aggregatesDailyAndCumulativeYen() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 3);
        LocalDate today = LocalDate.of(2026, 9, 2);
        Map<String, FeeInfo> fees = Map.of("A", ah(10.0), "B", ah(20.0));
        List<QuantityLine> actual =
                List.of(
                        new QuantityLine(from, "A", 5), // 50
                        new QuantityLine(today, "B", 2), // 40
                        new QuantityLine(to, "A", 1)); // future actual still counted in day but cum stops
        List<QuantityLine> plan =
                List.of(
                        new QuantityLine(from, "A", 3), // 30
                        new QuantityLine(to, "B", 1)); // 20

        Result r = ProcessingFeeTrendAggregator.aggregate(actual, plan, fees, from, to, today);
        assertEquals(3, r.days().size());
        DayPoint d0 = r.days().get(0);
        assertEquals(50.0, d0.actualYen(), 1e-6);
        assertEquals(30.0, d0.planYen(), 1e-6);
        assertEquals(50.0, d0.actualCumYen(), 1e-6);
        assertEquals(30.0, d0.planCumYen(), 1e-6);
        // 見込累計: 当日まで実績
        assertEquals(50.0, d0.projectedCumYen(), 1e-6);

        DayPoint d1 = r.days().get(1);
        assertEquals(40.0, d1.actualYen(), 1e-6);
        assertEquals(0.0, d1.planYen(), 1e-6);
        assertEquals(90.0, d1.actualCumYen(), 1e-6);
        assertEquals(30.0, d1.planCumYen(), 1e-6);
        // 当日: 実績を採用 → 実績累計と接続
        assertEquals(90.0, d1.projectedCumYen(), 1e-6);
        assertEquals(d1.actualCumYen(), d1.projectedCumYen(), 1e-6);

        DayPoint d2 = r.days().get(2);
        assertEquals(10.0, d2.actualYen(), 1e-6);
        // 翌日以降: 同日実績分を予定から差し引き → 20-10=10
        assertEquals(10.0, d2.planYen(), 1e-6);
        // 実績累計は today まで（9/2）で止まる
        assertEquals(90.0, d2.actualCumYen(), 1e-6);
        assertEquals(40.0, d2.planCumYen(), 1e-6);
        assertEquals(r.planTotalYen(), d2.planCumYen(), 1e-6);
        // 見込累計: 実績先端(90) + 差し引き後予定(10) → 100
        assertEquals(100.0, d2.projectedCumYen(), 1e-6);

        assertEquals(100.0, r.actualTotalYen(), 1e-6);
        assertEquals(40.0, r.planTotalYen(), 1e-6);
        assertEquals(3, r.actualLinesCounted());
        assertEquals(2, r.planLinesCounted());
        assertEquals(0, r.missingRateLines());
    }

    @Test
    void cumulativeYenResetsAtMonthBoundary() {
        LocalDate from = LocalDate.of(2026, 8, 31);
        LocalDate to = LocalDate.of(2026, 9, 2);
        LocalDate today = LocalDate.of(2026, 9, 2);
        Map<String, FeeInfo> fees = Map.of("A", ah(10.0));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                new QuantityLine(from, "A", 5), // 50
                                new QuantityLine(LocalDate.of(2026, 9, 1), "A", 3), // 30
                                new QuantityLine(today, "A", 2)), // 20
                        List.of(
                                new QuantityLine(from, "A", 1), // 10
                                new QuantityLine(LocalDate.of(2026, 9, 1), "A", 4), // 40
                                new QuantityLine(today, "A", 6)), // 60
                        fees,
                        from,
                        to,
                        today);
        DayPoint aug = r.days().get(0);
        DayPoint sep1 = r.days().get(1);
        DayPoint sep2 = r.days().get(2);
        assertEquals(50.0, aug.actualCumYen(), 1e-6);
        assertEquals(10.0, aug.planCumYen(), 1e-6);
        assertEquals(50.0, aug.projectedCumYen(), 1e-6);
        // 9/1 でリセット
        assertEquals(30.0, sep1.actualCumYen(), 1e-6);
        assertEquals(40.0, sep1.planCumYen(), 1e-6);
        assertEquals(30.0, sep1.projectedCumYen(), 1e-6);
        assertEquals(50.0, sep2.actualCumYen(), 1e-6);
        assertEquals(100.0, sep2.planCumYen(), 1e-6);
        assertEquals(50.0, sep2.projectedCumYen(), 1e-6); // 当日まで実績
        assertEquals(100.0, r.actualTotalYen(), 1e-6);
        assertEquals(110.0, r.planTotalYen(), 1e-6);
    }

    @Test
    void projectedCumConnectsAtTodayTipEvenWhenPlanExceedsActual() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 3);
        LocalDate today = LocalDate.of(2026, 9, 2);
        Map<String, FeeInfo> fees = Map.of("A", ah(10.0));
        // 当日: 予定 50 > 実績 10 → 見込は実績先端で接続し、翌日予定のみ積む
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                new QuantityLine(from, "A", 5), // 50
                                new QuantityLine(today, "A", 1)), // 10
                        List.of(
                                new QuantityLine(today, "A", 5), // 50
                                new QuantityLine(to, "A", 2)), // 20
                        fees,
                        from,
                        to,
                        today);
        DayPoint tip = r.days().get(1);
        assertEquals(60.0, tip.actualCumYen(), 1e-6);
        assertEquals(tip.actualCumYen(), tip.projectedCumYen(), 1e-6);
        DayPoint after = r.days().get(2);
        assertEquals(80.0, after.projectedCumYen(), 1e-6); // 60 + 20
    }

    @Test
    void futureDayPlanIsReducedBySameDayActual() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 3);
        LocalDate today = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees = Map.of("A", ah(10.0));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                new QuantityLine(from, "A", 1), // 10
                                new QuantityLine(to, "A", 3)), // 30 future actual
                        List.of(new QuantityLine(to, "A", 5)), // 50 future plan
                        fees,
                        from,
                        to,
                        today);
        DayPoint future = r.days().get(2);
        assertEquals(30.0, future.actualYen(), 1e-6);
        assertEquals(20.0, future.planYen(), 1e-6); // 50-30
        assertEquals(10.0, future.actualCumYen(), 1e-6); // tip at today
        assertEquals(30.0, future.projectedCumYen(), 1e-6); // 10 + 20
    }

    @Test
    void missingRateCountsAsZeroYen() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(new QuantityLine(d, "NO-RATE", 100)),
                        List.of(),
                        Map.of("OTHER", ah(1.0)),
                        d,
                        d,
                        d);
        assertEquals(0.0, r.days().get(0).actualYen(), 1e-9);
        assertEquals(1, r.missingRateLines());
        assertEquals(0, r.actualLinesCounted());
        assertEquals(1, r.requests().size());
        assertEquals("NO-RATE", r.requests().get(0).requestNo());
        assertEquals(true, r.requests().get(0).rateMissing());
        assertEquals(100.0, r.requests().get(0).actualMeters(), 1e-9);
        assertEquals(0.0, r.requests().get(0).actualYen(), 1e-9);
    }

    @Test
    void aggregatesByRequestNo() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 2);
        Map<String, FeeInfo> fees = Map.of("A", ah(10.0), "B", ah(20.0));
        List<QuantityLine> actual =
                List.of(
                        new QuantityLine(from, "B", 2), // 40
                        new QuantityLine(from, "A", 5), // 50
                        new QuantityLine(to, "A", 1)); // 10
        List<QuantityLine> plan =
                List.of(
                        new QuantityLine(from, "A", 3), // 30
                        new QuantityLine(to, "B", 4)); // 80

        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        actual, plan, fees, from, to, to);
        assertEquals(2, r.requests().size());
        // 依頼NO昇順
        assertEquals("A", r.requests().get(0).requestNo());
        assertEquals(10.0, r.requests().get(0).rateYenPerM(), 1e-9);
        assertEquals(6.0, r.requests().get(0).actualMeters(), 1e-9);
        assertEquals(3.0, r.requests().get(0).planMeters(), 1e-9);
        assertEquals(60.0, r.requests().get(0).actualYen(), 1e-9);
        assertEquals(30.0, r.requests().get(0).planYen(), 1e-9);
        assertEquals(false, r.requests().get(0).rateMissing());

        assertEquals("B", r.requests().get(1).requestNo());
        assertEquals(20.0, r.requests().get(1).rateYenPerM(), 1e-9);
        assertEquals(2.0, r.requests().get(1).actualMeters(), 1e-9);
        assertEquals(4.0, r.requests().get(1).planMeters(), 1e-9);
        assertEquals(40.0, r.requests().get(1).actualYen(), 1e-9);
        assertEquals(80.0, r.requests().get(1).planYen(), 1e-9);
    }

    @Test
    void multiProcessUsesOnlyFinalProcessMetersAndAllocatesAo() {
        LocalDate d0 = LocalDate.of(2026, 9, 1);
        LocalDate d1 = LocalDate.of(2026, 9, 2);
        // 工程A 100m + 最終工程B 50m、AO=10,000 → 円/m=200。Bの日次だけ円が立つ
        Map<String, FeeInfo> fees = Map.of("R", ao(10_000.0, "A,B"));
        List<QuantityLine> actual =
                List.of(
                        new QuantityLine(d0, "R", 100, "A"),
                        new QuantityLine(d0, "R", 30, "B"),
                        new QuantityLine(d1, "R", 20, "B"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        actual, List.of(), fees, d0, d1, d1);
        assertEquals(200.0, r.requests().get(0).rateYenPerM(), 1e-6);
        assertEquals(50.0, r.requests().get(0).actualMeters(), 1e-6);
        assertEquals(10_000.0, r.requests().get(0).actualYen(), 1e-6);
        assertEquals(10_000.0, r.requests().get(0).aoYen(), 1e-6);
        assertEquals(6000.0, r.days().get(0).actualYen(), 1e-6); // 30*200
        assertEquals(4000.0, r.days().get(1).actualYen(), 1e-6); // 20*200
        assertEquals(10_000.0, r.actualTotalYen(), 1e-6);
    }

    @Test
    void processContentLastTokenMatchesNormalizedProcessName() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees = Map.of("E9-2", ao(5_000.0, "スリット,E9-2"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                new QuantityLine(d, "E9-2", 40, "スリット"),
                                new QuantityLine(d, "E9-2", 25, "E9-2")),
                        List.of(),
                        fees,
                        d,
                        d,
                        d);
        assertEquals(25.0, r.requests().get(0).actualMeters(), 1e-6);
        assertEquals(200.0, r.requests().get(0).rateYenPerM(), 1e-6); // 5000/25
        assertEquals(5_000.0, r.requests().get(0).actualYen(), 1e-6);
    }

    @Test
    void aoMissingFallsBackToAhTimesFinalProcessMeters() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees =
                Map.of("R", new FeeInfo(50.0, null, "スリット,最終"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                new QuantityLine(d, "R", 10, "スリット"),
                                new QuantityLine(d, "R", 4, "最終")),
                        List.of(),
                        fees,
                        d,
                        d,
                        d);
        assertEquals(4.0, r.requests().get(0).actualMeters(), 1e-6);
        assertEquals(50.0, r.requests().get(0).rateYenPerM(), 1e-6);
        assertEquals(200.0, r.requests().get(0).actualYen(), 1e-6);
    }

    @Test
    void emptyProcessContentWithMultipleProcessesDropsAll() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees = Map.of("R", ao(9_000.0, ""));
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
        assertTrue(r.requests().isEmpty() || r.requests().get(0).actualMeters() == 0.0);
        assertEquals(0.0, r.actualTotalYen(), 1e-9);
    }

    @Test
    void withLeadingTotalRow_sumsMetersAndYen_rateMissing() {
        List<RequestPoint> rows =
                List.of(
                        new RequestPoint("A", 10.0, 1_000.0, false, 5, 3, 50, 30),
                        new RequestPoint("B", 20.0, 2_000.0, false, 2, 4, 40, 80));
        List<RequestPoint> withTotal = ProcessingFeeTrendAggregator.withLeadingTotalRow(rows);
        assertEquals(3, withTotal.size());
        RequestPoint tot = withTotal.get(0);
        assertTrue(tot.isTotalRow());
        assertEquals("合計", tot.requestNo());
        assertEquals(3_000.0, tot.aoYen(), 1e-9);
        assertEquals(7.0, tot.actualMeters(), 1e-9);
        assertEquals(7.0, tot.planMeters(), 1e-9);
        assertEquals(90.0, tot.actualYen(), 1e-9);
        assertEquals(110.0, tot.planYen(), 1e-9);
        assertTrue(tot.rateMissing());
        assertEquals("A", withTotal.get(1).requestNo());
        assertEquals("B", withTotal.get(2).requestNo());
        assertEquals(3_000.0, ProcessingFeeTrendAggregator.sumOrderAoYen(rows), 1e-9);
        assertTrue(ProcessingFeeTrendAggregator.withLeadingTotalRow(List.of()).isEmpty());
    }

    @Test
    void includeOrderMonthRequests_addsInactiveOrdersInPeriod() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 30);
        Map<String, FeeInfo> fees =
                Map.of(
                        "ACT",
                        new FeeInfo(10.0, 1_000.0, "A", 2026, 9),
                        "IDLE",
                        new FeeInfo(20.0, 4_796_700.0, "B", 2026, 9),
                        "OTHER",
                        new FeeInfo(30.0, 9_999.0, "C", 2026, 8));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(new QuantityLine(from, "ACT", 10, "A")),
                        List.of(),
                        fees,
                        from,
                        to,
                        from);
        assertEquals(2, r.requests().size());
        double orderAo = ProcessingFeeTrendAggregator.sumOrderAoYen(r.requests());
        assertEquals(4_797_700.0, orderAo, 1e-6);
        RequestPoint idle =
                r.requests().stream().filter(x -> "IDLE".equals(x.requestNo())).findFirst().orElseThrow();
        assertEquals(0.0, idle.actualMeters(), 1e-9);
        assertEquals(4_796_700.0, idle.aoYen(), 1e-6);
        List<RequestPoint> withTotal = ProcessingFeeTrendAggregator.withLeadingTotalRow(r.requests());
        assertEquals(4_797_700.0, withTotal.get(0).aoYen(), 1e-6);
        assertEquals(1_000.0, withTotal.get(0).actualYen() + withTotal.get(0).planYen(), 1e-6);
    }
}
