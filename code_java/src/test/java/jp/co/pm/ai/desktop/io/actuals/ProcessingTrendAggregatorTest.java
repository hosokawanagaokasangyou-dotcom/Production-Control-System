package jp.co.pm.ai.desktop.io.actuals;

import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.DayPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Filter;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.PlanSource;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Result;

class ProcessingTrendAggregatorTest {

    private static final LocalDate TODAY = LocalDate.of(2026, 9, 3);
    private static final LocalDate FROM = LocalDate.of(2026, 9, 1);
    private static final LocalDate TO = LocalDate.of(2026, 9, 4);

    private static final List<String> ACT_HEADERS =
            List.of("機械名", "依頼NO", "工程名", "加工日", "加工開始日時", "換算数量", "実加工数");

    private static ActualsSnapshot actuals() {
        return new ActualsSnapshot(
                ACT_HEADERS,
                List.of(
                        List.of("W9-1", "R1", "スリット", "2026/09/01", "2026/09/01 09:00", "500", "100"),
                        List.of("W9-1", "R2", "スリット", "2026/09/02", "2026/09/02 09:00", "500", "50"),
                        List.of("EC-2", "R3", "EC", "2026/09/02", "2026/09/02 13:00", "300", "30"),
                        // 期間外
                        List.of("W9-1", "R4", "スリット", "2026/08/31", "2026/08/31 09:00", "100", "999"),
                        // 当日（見込では予定側を使う）
                        List.of("W9-1", "R5", "スリット", "2026/09/03", "2026/09/03 09:00", "100", "10")));
    }

    private static AladdinSnapshot aladdin() {
        return new AladdinSnapshot(
                List.of("機械名", "依頼NO", "工程名", "2026/09/01", "2026/09/02", "2026/09/03", "2026/09/04"),
                List.of(
                        List.of("W9-1", "R1", "スリット", "120", "0", "600", "0"),
                        List.of("EC-2", "R3", "EC", "0", "40", "0", "80")));
    }

    private static DispatchSnapshot dispatch() {
        return new DispatchSnapshot(
                List.of("機械名", "依頼NO", "工程名", "配台日", "当日配台数量"),
                List.of(
                        List.of("W9-1", "R1", "スリット", "2026/09/01", "111"),
                        List.of("W9-1", "R1", "スリット", "2026-09-03 08:00", "222"),
                        List.of("EC-2", "R3", "EC", "2026/09/04", "333"),
                        List.of("EC-2", "R9", "EC", "2026/09/30", "9999")));
    }

    @Test
    void aggregate_fillsEveryDayAndBuildsCumulatives() {
        Result r =
                ProcessingTrendAggregator.aggregate(
                        actuals(), aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);

        Assertions.assertEquals(4, r.days().size());
        DayPoint d1 = r.days().get(0);
        DayPoint d2 = r.days().get(1);
        DayPoint d3 = r.days().get(2);
        DayPoint d4 = r.days().get(3);

        Assertions.assertEquals(100, d1.actualM(), 1e-9);
        Assertions.assertEquals(120, d1.planM(), 1e-9);
        Assertions.assertEquals(80, d2.actualM(), 1e-9);
        Assertions.assertEquals(40, d2.planM(), 1e-9);
        Assertions.assertEquals(10, d3.actualM(), 1e-9);
        Assertions.assertEquals(600, d3.planM(), 1e-9);
        Assertions.assertEquals(0, d4.actualM(), 1e-9);
        Assertions.assertEquals(80, d4.planM(), 1e-9);

        Assertions.assertEquals(190, d4.actualCumM(), 1e-9);
        Assertions.assertEquals(840, d4.planCumM(), 1e-9);

        // 見込: 9/1, 9/2 は実績、9/3 以降は予定
        Assertions.assertFalse(d2.usesPlanForProjection());
        Assertions.assertTrue(d3.usesPlanForProjection());
        Assertions.assertEquals(100 + 80 + 600 + 80, d4.projectedCumM(), 1e-9);

        Assertions.assertEquals(190, r.actualTotalM(), 1e-9);
        Assertions.assertEquals(840, r.planTotalM(), 1e-9);
        Assertions.assertEquals(180, r.actualToDateM(), 1e-9);
        Assertions.assertEquals(160, r.planToDateM(), 1e-9);
        Assertions.assertEquals(680, r.remainingPlanM(), 1e-9);
        Assertions.assertEquals(860, r.projectedTotalM(), 1e-9);
        Assertions.assertEquals(112.5, r.progressPct(), 1e-9);
        Assertions.assertEquals(20, r.projectedDiffM(), 1e-9);
        Assertions.assertEquals(4, r.actualRowsCounted());
        Assertions.assertEquals(2, r.planRowsCounted());
        Assertions.assertEquals(LocalDate.of(2026, 8, 31), r.actualMinDate());
        Assertions.assertEquals(LocalDate.of(2026, 9, 3), r.actualMaxDate());
        Assertions.assertTrue(r.warnings().isEmpty());
        Assertions.assertFalse(r.periodStartsBeforeActualSource());
    }

    @Test
    void aggregate_todayUsesMaxOfActualAndPlanForProjection() {
        ActualsSnapshot act =
                new ActualsSnapshot(
                        ACT_HEADERS,
                        List.of(
                                List.of("W9-1", "R1", "スリット", "2026/09/02", "", "500", "100"),
                                // 当日: 実績 900 > 予定 600
                                List.of("W9-1", "R5", "スリット", "2026/09/03", "", "1000", "900")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        act, aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        DayPoint d3 = r.days().get(2);
        Assertions.assertEquals(900, d3.actualM(), 1e-9);
        Assertions.assertEquals(600, d3.planM(), 1e-9);
        Assertions.assertTrue(d3.usesPlanForProjection());
        // 9/1 実績 0, 9/2 実績 100, 9/3 max(900,600)=900, 9/4 予定 80
        Assertions.assertEquals(100 + 900 + 80, r.days().get(3).projectedCumM(), 1e-9);
        Assertions.assertEquals(900 + 80, r.remainingPlanM(), 1e-9);
        Assertions.assertEquals(100 + 900 + 80, r.projectedTotalM(), 1e-9);
    }

    @Test
    void aggregate_skipsAladdinTotalRow() {
        AladdinSnapshot withTotal =
                new AladdinSnapshot(
                        List.of("倉庫", "機械名", "依頼NO", "工程名", "2026/09/01", "2026/09/02", "2026/09/03", "2026/09/04"),
                        List.of(
                                List.of("511101", "W9-1", "R1", "スリット", "120", "0", "600", "0"),
                                List.of("511101", "EC-2", "R3", "EC", "0", "40", "0", "80"),
                                List.of("[合計]", "", "", "", "120", "40", "600", "80")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        actuals(), withTotal, dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        Assertions.assertEquals(840, r.planTotalM(), 1e-9);
        Assertions.assertEquals(2, r.planRowsCounted());

        // 倉庫列が無くても、機械名・依頼NO とも空の行は合計行として除外する
        AladdinSnapshot noWarehouse =
                new AladdinSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "2026/09/01", "2026/09/02", "2026/09/03", "2026/09/04"),
                        List.of(
                                List.of("W9-1", "R1", "スリット", "120", "0", "600", "0"),
                                List.of("", "", "", "120", "0", "600", "0")));
        Result r2 =
                ProcessingTrendAggregator.aggregate(
                        actuals(), noWarehouse, dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        Assertions.assertEquals(720, r2.planTotalM(), 1e-9);
    }

    @Test
    void aggregate_capsRemainingAladdinPlanByUnprocessed() {
        List<String> headers =
                List.of(
                        "機械名", "依頼NO", "工程名", "加工完了区分", "換算数量", "実加工数", "未加工",
                        "2026/09/01", "2026/09/03", "2026/09/04", "2026/09/07");
        AladdinSnapshot a =
                new AladdinSnapshot(
                        headers,
                        List.of(
                                // 既加工 3800 が日付列に残っている: 当日以降 2000+2400+800=5200 > 未加工 4200 → 遅い日から削る
                                List.of("W9-1", "T1", "スリット", "0:未完", "8000", "3800", "4200", "0", "2000", "2400", "800"),
                                // 全数未加工ルール: 換算>0・実加工=0・未加工=0 → 上限は換算 1000（超過なし）
                                List.of("W9-1", "T2", "スリット", "0:未完", "1000", "0", "0", "0", "400", "600", "0"),
                                // 完了行: 当日以降は 0、過去日は残す
                                List.of("W9-1", "T3", "スリット", "1:完了", "500", "500", "0", "300", "200", "0", "0"),
                                // 実加工>0・未加工=0（完了扱い）: 当日以降 0
                                List.of("W9-1", "T4", "スリット", "0:未完", "700", "700", "0", "0", "700", "0", "0")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        new ActualsSnapshot(ACT_HEADERS, List.of()), a, dispatch(),
                        new Filter(FROM, LocalDate.of(2026, 9, 7), PlanSource.ALADDIN, null, null), TODAY);
        double plan0901 = r.days().get(0).planM();
        double plan0903 = r.days().get(2).planM();
        double plan0904 = r.days().get(3).planM();
        double plan0907 = r.days().get(6).planM();
        Assertions.assertEquals(300, plan0901, 1e-9);
        // T1: 9/7 の 800 を全部削り、9/4 を 2400→2200。T2: 400。T3/T4: 0
        Assertions.assertEquals(2000 + 400, plan0903, 1e-9);
        Assertions.assertEquals(2200 + 600, plan0904, 1e-9);
        Assertions.assertEquals(0, plan0907, 1e-9);
        Assertions.assertEquals(3, r.planRowsCounted());
    }

    @Test
    void aggregate_capUsesFutureColumnsOutsidePeriod() {
        // 期間外（9/10）の当日以降予定も行合計に含めて上限を判定し、超過は遅い日（期間外）から削る
        AladdinSnapshot a =
                new AladdinSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "換算数量", "実加工数", "未加工", "2026/09/03", "2026/09/10"),
                        List.of(List.of("W9-1", "T1", "スリット", "1000", "0", "600", "500", "500")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        new ActualsSnapshot(ACT_HEADERS, List.of()), a, dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        Assertions.assertEquals(500, r.days().get(2).planM(), 1e-9);
    }

    @Test
    void aggregate_missingActualQtyColumn_warnsInsteadOfFallingBack() {
        ActualsSnapshot noQty =
                new ActualsSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "加工日", "換算数量", "累積実績"),
                        List.of(List.of("W9-1", "R1", "スリット", "2026/09/01", "500", "400")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        noQty, aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        Assertions.assertEquals(0, r.actualTotalM(), 1e-9);
        Assertions.assertEquals(0, r.actualRowsCounted());
        Assertions.assertEquals(List.of(ProcessingTrendAggregator.WARN_ACTUAL_QTY_COLUMN_MISSING), r.warnings());
        Assertions.assertNull(r.actualMinDate());
    }

    @Test
    void aggregate_actualDateHeadersToleratePaddingAndFallBackToKakouDate() {
        ActualsSnapshot padded =
                new ActualsSnapshot(
                        List.of("機械名 ", "工程名", " 加工日", "加工開始日時 ", "実加工数 "),
                        List.of(
                                // 加工開始日時が空 → 加工日にフォールバック
                                List.of("W9-1", "スリット", "2026/09/01", "", "100"),
                                // 加工開始日時（時刻付き）を優先
                                List.of("W9-1", "スリット", "2026/09/09", "2026/09/02 08:00", "50")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        padded, aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, "W9-1", null), TODAY);
        Assertions.assertEquals(100, r.days().get(0).actualM(), 1e-9);
        Assertions.assertEquals(50, r.days().get(1).actualM(), 1e-9);
        Assertions.assertEquals(2, r.actualRowsCounted());
    }

    @Test
    void aggregate_zeroActualRowsAreNotCounted() {
        ActualsSnapshot act =
                new ActualsSnapshot(
                        ACT_HEADERS,
                        List.of(
                                List.of("W9-1", "R1", "スリット", "2026/09/01", "", "500", "0"),
                                List.of("W9-1", "R2", "スリット", "2026/09/01", "", "500", "25")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        act, aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        Assertions.assertEquals(1, r.actualRowsCounted());
        Assertions.assertEquals(25, r.actualTotalM(), 1e-9);
    }

    @Test
    void result_progressDenominatorSufficient() {
        Result r =
                ProcessingTrendAggregator.aggregate(
                        actuals(), aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        // 前日まで 160 / 期間合計 840 = 19% ≥ 10%
        Assertions.assertTrue(r.progressDenominatorSufficient());

        AladdinSnapshot mostlyFuture =
                new AladdinSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "2026/09/01", "2026/09/03"),
                        List.of(List.of("W9-1", "R1", "スリット", "50", "950")));
        Result r2 =
                ProcessingTrendAggregator.aggregate(
                        actuals(), mostlyFuture, dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        Assertions.assertFalse(r2.progressDenominatorSufficient());
        Assertions.assertFalse(Double.isNaN(r2.progressPct()));
    }

    @Test
    void aggregate_machineFilterUsesNormalizedKey() {
        Result r =
                ProcessingTrendAggregator.aggregate(
                        actuals(), aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, " ＥＣ-2 ", null), TODAY);
        Assertions.assertEquals(0, r.days().get(0).actualM(), 1e-9);
        Assertions.assertEquals(30, r.days().get(1).actualM(), 1e-9);
        Assertions.assertEquals(40, r.days().get(1).planM(), 1e-9);
        Assertions.assertEquals(80, r.days().get(3).planM(), 1e-9);
        Assertions.assertEquals(1, r.actualRowsCounted());
        Assertions.assertEquals(1, r.planRowsCounted());
    }

    @Test
    void aggregate_processFilter() {
        Result r =
                ProcessingTrendAggregator.aggregate(
                        actuals(), aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, "EC"), TODAY);
        Assertions.assertEquals(30, r.actualTotalM(), 1e-9);
        Assertions.assertEquals(120, r.planTotalM(), 1e-9);
    }

    @Test
    void aggregate_dispatchSourceParsesDateVariants() {
        Result r =
                ProcessingTrendAggregator.aggregate(
                        actuals(), aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.DISPATCH, null, null), TODAY);
        Assertions.assertEquals(111, r.days().get(0).planM(), 1e-9);
        Assertions.assertEquals(0, r.days().get(1).planM(), 1e-9);
        Assertions.assertEquals(222, r.days().get(2).planM(), 1e-9);
        Assertions.assertEquals(333, r.days().get(3).planM(), 1e-9);
        Assertions.assertEquals(666, r.planTotalM(), 1e-9);
        Assertions.assertEquals(3, r.planRowsCounted());
    }

    @Test
    void aggregate_legacyDispatchWithActualQtyColumn_doesNotDoubleCount() {
        // 旧 段階3 JSON: 目標行（当日配台数量のみ）とタイムライン行（実配台数量・加工開始日時）が同一 (依頼,工程,機械) に共存
        DispatchSnapshot legacy =
                new DispatchSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "配台日", "当日配台数量", "実配台数量", "加工開始日時"),
                        List.of(
                                // 孤立目標行（時刻なし）: 実績行と共存するので除外される
                                List.of("W9-1", "R1", "スリット", "2026/09/03", "500", "", ""),
                                // タイムライン行: 実配台 320 を主数量に
                                List.of("W9-1", "R1", "スリット", "2026/09/03", "0", "320", "2026/09/03 08:00"),
                                // 実配台 0 の別暦日タイムライン行は、目標が孤立のみの場合は残る（当日配台 0 なので合算に影響なし）
                                List.of("W9-1", "R1", "スリット", "2026/09/04", "0", "0", "2026/09/04 08:00")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        new ActualsSnapshot(ACT_HEADERS, List.of()), aladdin(), legacy,
                        new Filter(FROM, TO, PlanSource.DISPATCH, null, null), TODAY);
        Assertions.assertEquals(320, r.days().get(2).planM(), 1e-9);
        Assertions.assertEquals(0, r.days().get(3).planM(), 1e-9);
        Assertions.assertEquals(320, r.planTotalM(), 1e-9);
    }

    @Test
    void aggregate_noPlanBeforeToday_progressIsNaN() {
        Result r =
                ProcessingTrendAggregator.aggregate(
                        actuals(), new AladdinSnapshot(List.of(), List.of()), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        Assertions.assertTrue(Double.isNaN(r.progressPct()));
        Assertions.assertEquals(0, r.planRowsCounted());
        Assertions.assertFalse(r.isEmpty());
    }

    @Test
    void aggregate_swapsReversedRangeAndClampsLength() {
        Result r =
                ProcessingTrendAggregator.aggregate(
                        actuals(), aladdin(), dispatch(),
                        new Filter(TO, FROM, PlanSource.ALADDIN, null, null), TODAY);
        Assertions.assertEquals(FROM, r.days().get(0).date());
        Assertions.assertEquals(TO, r.days().get(3).date());

        Result huge =
                ProcessingTrendAggregator.aggregate(
                        actuals(), aladdin(), dispatch(),
                        new Filter(FROM, FROM.plusYears(20), PlanSource.ALADDIN, null, null), TODAY);
        Assertions.assertEquals(ProcessingTrendAggregator.MAX_DAYS, huge.days().size());
    }

    @Test
    void machineAndProcessNames_unionAcrossSources() {
        DispatchSnapshot dis =
                new DispatchSnapshot(
                        List.of("機械名", "工程名", "配台日", "当日配台数量"),
                        List.of(List.of("Z-9", "ラミ", "2026/09/01", "1")));
        List<String> machines = ProcessingTrendAggregator.machineNames(actuals(), aladdin(), dis);
        Assertions.assertEquals(List.of("EC-2", "W9-1", "Z-9"), machines);
        List<String> processes = ProcessingTrendAggregator.processNames(actuals(), aladdin(), dis);
        Assertions.assertEquals(3, processes.size());
        Assertions.assertTrue(processes.containsAll(List.of("EC", "スリット", "ラミ")));
    }

    @Test
    void parseDate_variants() {
        Assertions.assertEquals(LocalDate.of(2026, 9, 3), ProcessingTrendAggregator.parseDate("2026/09/03"));
        Assertions.assertEquals(LocalDate.of(2026, 9, 3), ProcessingTrendAggregator.parseDate("2026-9-3"));
        Assertions.assertEquals(
                LocalDate.of(2026, 9, 3), ProcessingTrendAggregator.parseDate("2026/09/03 10:00"));
        Assertions.assertEquals(
                LocalDate.of(2026, 9, 3), ProcessingTrendAggregator.parseDate("2026-09-03T10:00:00"));
        Assertions.assertNull(ProcessingTrendAggregator.parseDate("abc"));
        Assertions.assertNull(ProcessingTrendAggregator.parseDate(""));
        // 存在しない日付は丸めない・2 桁年は受け付けない
        Assertions.assertNull(ProcessingTrendAggregator.parseDate("2026/02/30"));
        Assertions.assertNull(ProcessingTrendAggregator.parseDate("26/9/3"));
    }

    @Test
    void normKey_foldsWidthDashesAndSpaces() {
        String expected = ProcessingTrendAggregator.normKey("W9-1");
        Assertions.assertEquals(expected, ProcessingTrendAggregator.normKey("Ｗ９－１"));
        Assertions.assertEquals(expected, ProcessingTrendAggregator.normKey("W9\u20101"));
        Assertions.assertEquals(expected, ProcessingTrendAggregator.normKey("W9\u22121"));
        Assertions.assertEquals(expected, ProcessingTrendAggregator.normKey("\u200bW9-1\u3000"));
        Assertions.assertEquals(
                ProcessingTrendAggregator.normKey("SEC機 湖南"), ProcessingTrendAggregator.normKey("SEC機\u3000\u3000湖南"));
        Assertions.assertEquals("", ProcessingTrendAggregator.normKey("  "));
    }
}
