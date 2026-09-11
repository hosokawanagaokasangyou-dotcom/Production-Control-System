package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
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
                        List.of("機械名", "依頼NO", "工程名", "加工日付", "実加工量", "実製品出来高"),
                        List.of(List.of("M1", "Y7-2", "最終", "2026/09/10", "400", "1600")));
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
