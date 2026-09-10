package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.DayPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.QuantityLine;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.Result;

class ProcessingFeeTrendAggregatorTest {

    @Test
    void aggregatesDailyAndCumulativeYen() {
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 3);
        LocalDate today = LocalDate.of(2026, 9, 2);
        Map<String, Double> rates = Map.of("A", 10.0, "B", 20.0);
        List<QuantityLine> actual =
                List.of(
                        new QuantityLine(from, "A", 5), // 50
                        new QuantityLine(today, "B", 2), // 40
                        new QuantityLine(to, "A", 1)); // future actual still counted in day but cum stops
        List<QuantityLine> plan =
                List.of(
                        new QuantityLine(from, "A", 3), // 30
                        new QuantityLine(to, "B", 1)); // 20

        Result r = ProcessingFeeTrendAggregator.aggregate(actual, plan, rates, from, to, today);
        assertEquals(3, r.days().size());
        DayPoint d0 = r.days().get(0);
        assertEquals(50.0, d0.actualYen(), 1e-6);
        assertEquals(30.0, d0.planYen(), 1e-6);
        assertEquals(50.0, d0.actualCumYen(), 1e-6);
        assertEquals(30.0, d0.planCumYen(), 1e-6);
        // 見込累計: 当日より前は実績
        assertEquals(50.0, d0.projectedCumYen(), 1e-6);

        DayPoint d1 = r.days().get(1);
        assertEquals(40.0, d1.actualYen(), 1e-6);
        assertEquals(0.0, d1.planYen(), 1e-6);
        assertEquals(90.0, d1.actualCumYen(), 1e-6);
        assertEquals(30.0, d1.planCumYen(), 1e-6);
        // 当日: max(実績,予定)=40 → 実績累計と接続
        assertEquals(90.0, d1.projectedCumYen(), 1e-6);

        DayPoint d2 = r.days().get(2);
        assertEquals(10.0, d2.actualYen(), 1e-6);
        assertEquals(20.0, d2.planYen(), 1e-6);
        // 実績累計は today まで（9/2）で止まる
        assertEquals(90.0, d2.actualCumYen(), 1e-6);
        // 予定累計は予定のみの累計（表・KPI用）
        assertEquals(50.0, d2.planCumYen(), 1e-6);
        assertEquals(r.planTotalYen(), d2.planCumYen(), 1e-6);
        // 見込累計: 実績先端(90)から予定日次を積む → 110
        assertEquals(110.0, d2.projectedCumYen(), 1e-6);

        assertEquals(100.0, r.actualTotalYen(), 1e-6);
        assertEquals(50.0, r.planTotalYen(), 1e-6);
        assertEquals(3, r.actualLinesCounted());
        assertEquals(2, r.planLinesCounted());
        assertEquals(0, r.missingRateLines());
    }

    @Test
    void missingRateCountsAsZeroYen() {
        LocalDate d = LocalDate.of(2026, 9, 1);
        Result r =
                ProcessingFeeTrendAggregator.aggregate(
                        List.of(new QuantityLine(d, "NO-RATE", 100)),
                        List.of(),
                        Map.of("OTHER", 1.0),
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
        Map<String, Double> rates = Map.of("A", 10.0, "B", 20.0);
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
                        actual, plan, rates, from, to, to);
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
}
