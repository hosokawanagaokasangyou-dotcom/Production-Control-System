package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.QuantityLine;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Filter;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.PlanSource;

class ProcessingFeeTrendQuantityExtractorTest {

    @Test
    void extractDispatchSkipsOnOrBeforeToday() {
        LocalDate today = LocalDate.of(2026, 9, 11);
        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 30);
        DispatchSnapshot dispatch =
                new DispatchSnapshot(
                        List.of(
                                ResultDispatchSchema.COL_MACHINE,
                                ResultDispatchSchema.COL_PROCESS,
                                ResultDispatchSchema.COL_DISPATCH_DATE,
                                ResultDispatchSchema.COL_DISPATCH_QTY,
                                "依頼NO"),
                        List.of(
                                List.of("M1", "P1", "2026/09/10", "100", "A"),
                                List.of("M1", "P1", "2026/09/11", "200", "A"),
                                List.of("M1", "P1", "2026/09/12", "300", "A")));
        Filter filter =
                new Filter(
                        from,
                        to,
                        ProcessingTrendAggregator.ActualSource.DAILY_REPORT,
                        PlanSource.DISPATCH,
                        null,
                        null,
                        7);
        List<QuantityLine> plan =
                ProcessingFeeTrendQuantityExtractor.extractPlan(null, dispatch, filter, today);
        assertEquals(1, plan.size());
        assertEquals(LocalDate.of(2026, 9, 12), plan.get(0).date());
        assertEquals(300.0, plan.get(0).meters(), 1e-9);
        assertEquals("P1", plan.get(0).processName());
    }

    @Test
    void extractActual_dailyReportPrefersProductOutputOverProcessingQty() {
        LocalDate d = LocalDate.of(2026, 9, 10);
        ActualsSnapshot daily =
                new ActualsSnapshot(
                        List.of(
                                "機械名",
                                "依頼NO",
                                "工程名",
                                "加工日付",
                                "実加工量",
                                "実製品出来高",
                                "製品加工終了時間分"),
                        List.of(
                                List.of(
                                        "M1",
                                        "Y7-2",
                                        "最終",
                                        "2026/09/10",
                                        "400",
                                        "1600",
                                        "15:30")));
        Filter filter =
                new Filter(
                        LocalDate.of(2026, 9, 1),
                        LocalDate.of(2026, 9, 30),
                        ProcessingTrendAggregator.ActualSource.DAILY_REPORT,
                        PlanSource.ALADDIN,
                        null,
                        null,
                        7);
        List<QuantityLine> actual =
                ProcessingFeeTrendQuantityExtractor.extractActual(daily, null, filter);
        assertEquals(1, actual.size());
        assertEquals(1_600.0, actual.get(0).meters(), 1e-9);
        assertEquals("Y7-2", actual.get(0).requestNo());
        assertEquals(LocalDateTime.of(2026, 9, 10, 15, 30), actual.get(0).finishedAt());
    }

    @Test
    void extractActual_usesProductEndTimeColumnFromDailyReportCsv() {
        // 加工日報発行問合せ CSV は「終了時間」ではなく「製品加工終了時間分」
        LocalDate d = LocalDate.of(2026, 7, 15);
        ActualsSnapshot daily =
                new ActualsSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "加工日付", "実製品出来高", "製品加工終了時間分"),
                        List.of(
                                List.of("SEC機", "C7-10", "SEC", "2026/07/15", "2000", "14:27"),
                                List.of("SEC機", "C7-10", "SEC", "2026/07/15", "2000", "15:20"),
                                List.of("接続機", "C7-10", "接続", "2026/07/15", "2000", "11:47")));
        Filter filter =
                new Filter(
                        LocalDate.of(2026, 7, 1),
                        LocalDate.of(2026, 7, 31),
                        ProcessingTrendAggregator.ActualSource.DAILY_REPORT,
                        PlanSource.ALADDIN,
                        null,
                        null,
                        7);
        List<QuantityLine> actual =
                ProcessingFeeTrendQuantityExtractor.extractActual(daily, null, filter);
        assertEquals(3, actual.size());
        assertEquals(LocalDateTime.of(2026, 7, 15, 15, 20), actual.get(1).finishedAt());
        assertEquals(LocalDateTime.of(2026, 7, 15, 11, 47), actual.get(2).finishedAt());
    }

    @Test
    void extractActual_skipsCanceledZeroOrderQty() {
        // W7-23: 受注数量0はキャンセル。W7-23-1: 300は残す
        ActualsSnapshot daily =
                new ActualsSnapshot(
                        List.of(
                                "機械名",
                                "依頼NO",
                                "工程名",
                                "加工日付",
                                "実製品出来高",
                                "受注数量",
                                "製品加工終了時間分"),
                        List.of(
                                List.of(
                                        "EC機",
                                        "W7-23",
                                        "EC",
                                        "2025/03/03",
                                        "3",
                                        "0.00",
                                        "10:00"),
                                List.of(
                                        "EC機",
                                        "W7-23",
                                        "EC",
                                        "2025/03/07",
                                        "3",
                                        "0",
                                        "11:00"),
                                List.of(
                                        "熱転写機",
                                        "W7-23-1",
                                        "検反",
                                        "2025/03/17",
                                        "300",
                                        "300.00",
                                        "15:00")));
        Filter filter =
                new Filter(
                        LocalDate.of(2025, 3, 1),
                        LocalDate.of(2025, 3, 31),
                        ProcessingTrendAggregator.ActualSource.DAILY_REPORT,
                        PlanSource.ALADDIN,
                        null,
                        null,
                        7);
        List<QuantityLine> actual =
                ProcessingFeeTrendQuantityExtractor.extractActual(daily, null, filter);
        assertEquals(1, actual.size());
        assertEquals("W7-23-1", actual.get(0).requestNo());
        assertEquals(300.0, actual.get(0).meters(), 1e-9);
    }

    @Test
    void extractActual_parsesEndTimeVariants() {
        assertEquals(
                LocalDateTime.of(2026, 7, 15, 9, 5),
                ProcessingFeeTrendQuantityExtractor.composeFinishedAt(
                        LocalDate.of(2026, 7, 15), "9:05"));
        assertEquals(
                LocalDateTime.of(2026, 7, 15, 16, 0),
                ProcessingFeeTrendQuantityExtractor.composeFinishedAt(
                        LocalDate.of(2026, 7, 15), "1600"));
        assertEquals(
                LocalDateTime.of(2026, 7, 15, 18, 0),
                ProcessingFeeTrendQuantityExtractor.composeFinishedAt(
                        LocalDate.of(2026, 7, 15), "18"));
    }

    @Test
    void extractActual_unboundedIncludesOutsidePeriod() {
        LocalDate in = LocalDate.of(2026, 7, 10);
        LocalDate out = LocalDate.of(2024, 7, 15);
        ActualsSnapshot daily =
                new ActualsSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "加工日付", "実製品出来高"),
                        List.of(
                                List.of("M1", "C7-10", "SEC", "2026/07/10", "100"),
                                List.of("M1", "C7-10", "SEC", "2024/07/15", "2000"),
                                List.of("M1", "C7-10", "SEC", "2024/07/15", "2000")));
        Filter filter =
                new Filter(
                        LocalDate.of(2026, 7, 1),
                        LocalDate.of(2026, 7, 31),
                        ProcessingTrendAggregator.ActualSource.DAILY_REPORT,
                        PlanSource.ALADDIN,
                        null,
                        null,
                        7);
        List<QuantityLine> bounded =
                ProcessingFeeTrendQuantityExtractor.extractActual(daily, null, filter, true);
        List<QuantityLine> unbounded =
                ProcessingFeeTrendQuantityExtractor.extractActual(daily, null, filter, false);
        assertEquals(1, bounded.size());
        assertEquals(100.0, bounded.get(0).meters(), 1e-9);
        assertEquals(3, unbounded.size());
        assertEquals(
                4_100.0, unbounded.stream().mapToDouble(QuantityLine::meters).sum(), 1e-9);
        assertEquals(in, bounded.get(0).date());
        assertTrue(unbounded.stream().anyMatch(q -> out.equals(q.date())));
    }

    @Test
    void extractDispatchNormalizesLegacyDuplicateRows() {
        LocalDate today = LocalDate.of(2026, 9, 1);
        DispatchSnapshot dispatch =
                new DispatchSnapshot(
                        List.of(
                                ResultDispatchSchema.COL_MACHINE,
                                ResultDispatchSchema.COL_PROCESS,
                                ResultDispatchSchema.COL_ORDER_NO,
                                ResultDispatchSchema.COL_DISPATCH_DATE,
                                ResultDispatchSchema.COL_DISPATCH_QTY,
                                ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL,
                                "依頼NO"),
                        List.of(
                                List.of("M1", "P1", "O1", "2026/09/05", "100", "0", "A"),
                                List.of("M1", "P1", "O1", "2026/09/05", "100", "80", "A")));
        Filter filter =
                new Filter(
                        LocalDate.of(2026, 9, 1),
                        LocalDate.of(2026, 9, 30),
                        ProcessingTrendAggregator.ActualSource.DAILY_REPORT,
                        PlanSource.DISPATCH,
                        null,
                        null,
                        7);
        List<QuantityLine> plan =
                ProcessingFeeTrendQuantityExtractor.extractPlan(null, dispatch, filter, today);
        assertEquals(1, plan.size());
        assertTrue(plan.get(0).meters() <= 100.0 + 1e-9);
        assertEquals(80.0, plan.get(0).meters(), 1e-9);
    }
}
