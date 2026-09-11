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
        return new FeeInfo(rate, null, "", null, null, null);
    }

    private static FeeInfo ao(double aoYen, double orderFinalM, String processContent) {
        return new FeeInfo(null, aoYen, processContent, null, null, orderFinalM);
    }

    @Test
    void aoIsAuthority_actualPlusRemainEqualsAo() {
        // AO=1000, 受注100m → 10円/m。実績60 → 600、未了40 → 400
        LocalDate d = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees = Map.of("R", ao(1_000.0, 100.0, "最終"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(new QuantityLine(d, "R", 60, "最終")),
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
    void multiProcessUsesOnlyFinalProcessMetersForActualAndOrderM() {
        LocalDate d0 = LocalDate.of(2026, 9, 1);
        LocalDate d1 = LocalDate.of(2026, 9, 2);
        // 工程A は捨て、最終B の受注50m・AO10000 → 200円/m。実績50で未了0
        Map<String, FeeInfo> fees = Map.of("R", ao(10_000.0, 50.0, "A,B"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(
                                new QuantityLine(d0, "R", 100, "A"),
                                new QuantityLine(d0, "R", 30, "B"),
                                new QuantityLine(d1, "R", 20, "B")),
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
    void pastPeriod_noDailyRemainBars_kpiRemainFromAlloc() {
        // 7月を9月に見る: 日次未了棒は出さない。KPI・依頼の未了は残す
        LocalDate from = LocalDate.of(2026, 7, 1);
        LocalDate to = LocalDate.of(2026, 7, 31);
        LocalDate today = LocalDate.of(2026, 9, 11);
        Map<String, FeeInfo> fees = Map.of("R", ao(1_000.0, 100.0, "最終"));
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(new QuantityLine(LocalDate.of(2026, 7, 10), "R", 60, "最終")),
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
                        List.of(new QuantityLine(LocalDate.of(2026, 9, 5), "R", 60, "最終")),
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
        // AH のみ・受注 m = 実績+予定相当は allocateRequest が act+planSched を orderM に使う
        Map<String, FeeInfo> fees =
                Map.of(
                        "A", new FeeInfo(10.0, null, "", null, null, 9.0),
                        "B", new FeeInfo(20.0, null, "", null, null, 3.0));
        List<QuantityLine> actual =
                List.of(
                        new QuantityLine(from, "A", 5),
                        new QuantityLine(today, "B", 2),
                        new QuantityLine(to, "A", 1));
        List<QuantityLine> plan = List.of(new QuantityLine(to, "B", 1));

        Result r = ProcessingFeeTrendAggregator.aggregate(actual, plan, fees, from, to, today);
        assertEquals(3, r.days().size());
        DayPoint d0 = r.days().get(0);
        assertEquals(50.0, d0.actualYen(), 1e-6);
        RequestPoint a =
                r.requests().stream().filter(x -> "A".equals(x.requestNo())).findFirst().orElseThrow();
        // A: orderM=9, act=6 → remain 3, yen 60+30
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
                        List.of(new QuantityLine(d, "NO-RATE", 100)),
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
                        List.of(new QuantityLine(from, "ACT", 10, "A")),
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
    void processContentLastTokenMatchesNormalizedProcessName() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Map<String, FeeInfo> fees = Map.of("E9-2", ao(5_000.0, 25.0, "スリット,E9-2"));
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
        assertEquals(200.0, r.requests().get(0).rateYenPerM(), 1e-6);
        assertEquals(5_000.0, r.requests().get(0).actualYen(), 1e-6);
        assertEquals(0.0, r.requests().get(0).planYen(), 1e-6);
    }

    @Test
    void emptyProcessContentWithMultipleProcessesDropsAll() {
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
        // 加工内容空・複数工程 → 実績採用なし。受注月が無いので依頼自体も出ないか 0
        assertTrue(
                r.requests().isEmpty()
                        || (r.requests().get(0).actualMeters() == 0.0
                                && r.requests().get(0).actualYen() == 0.0));
    }
}
