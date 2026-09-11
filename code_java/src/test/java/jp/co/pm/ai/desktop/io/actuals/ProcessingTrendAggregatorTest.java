package jp.co.pm.ai.desktop.io.actuals;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.ActualSource;
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
                        // 当日（見込でも実績を使い、先端で接続）
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

        // 見込: 9/1〜9/3 は実績、9/4 以降は予定（当日先端で接続）
        Assertions.assertFalse(d2.usesPlanForProjection());
        Assertions.assertFalse(d3.usesPlanForProjection());
        Assertions.assertTrue(d4.usesPlanForProjection());
        Assertions.assertEquals(100 + 80 + 10 + 80, d4.projectedCumM(), 1e-9);
        // 当日の見込累計 = 実績累計（先端接続）
        Assertions.assertEquals(d3.actualCumM(), d3.projectedCumM(), 1e-9);

        Assertions.assertEquals(190, r.actualTotalM(), 1e-9);
        Assertions.assertEquals(840, r.planTotalM(), 1e-9);
        Assertions.assertEquals(190, r.actualToDateM(), 1e-9);
        Assertions.assertEquals(760, r.planToDateM(), 1e-9);
        Assertions.assertEquals(80, r.remainingPlanM(), 1e-9);
        Assertions.assertEquals(270, r.projectedTotalM(), 1e-9);
        Assertions.assertEquals(190.0 / 760.0 * 100.0, r.progressPct(), 1e-9);
        Assertions.assertEquals(270 - 840, r.projectedDiffM(), 1e-9);
        Assertions.assertEquals(4, r.actualRowsCounted());
        Assertions.assertEquals(2, r.planRowsCounted());
        Assertions.assertEquals(LocalDate.of(2026, 8, 31), r.actualMinDate());
        Assertions.assertEquals(LocalDate.of(2026, 9, 3), r.actualMaxDate());
        Assertions.assertTrue(r.warnings().isEmpty());
        Assertions.assertFalse(r.periodStartsBeforeActualSource());
    }

    @Test
    void aggregate_resetsCumulativeAtMonthBoundary() {
        LocalDate from = LocalDate.of(2026, 8, 30);
        LocalDate to = LocalDate.of(2026, 9, 2);
        LocalDate today = LocalDate.of(2026, 9, 2);
        ActualsSnapshot act =
                new ActualsSnapshot(
                        ACT_HEADERS,
                        List.of(
                                List.of("W9-1", "R1", "スリット", "2026/08/30", "", "100", "100"),
                                List.of("W9-1", "R2", "スリット", "2026/08/31", "", "100", "50"),
                                List.of("W9-1", "R3", "スリット", "2026/09/01", "", "100", "40"),
                                List.of("W9-1", "R4", "スリット", "2026/09/02", "", "100", "10")));
        AladdinSnapshot plan =
                new AladdinSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "2026/08/30", "2026/08/31", "2026/09/01", "2026/09/02"),
                        List.of(List.of("W9-1", "R1", "スリット", "10", "20", "30", "40")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        act, plan, dispatch(),
                        new Filter(from, to, PlanSource.ALADDIN, null, null), today);
        Assertions.assertEquals(4, r.days().size());
        DayPoint aug30 = r.days().get(0);
        DayPoint aug31 = r.days().get(1);
        DayPoint sep01 = r.days().get(2);
        DayPoint sep02 = r.days().get(3);
        Assertions.assertEquals(100, aug30.actualCumM(), 1e-9);
        Assertions.assertEquals(150, aug31.actualCumM(), 1e-9);
        Assertions.assertEquals(10, aug30.planCumM(), 1e-9);
        Assertions.assertEquals(30, aug31.planCumM(), 1e-9);
        // 9/1 で累計リセット
        Assertions.assertEquals(40, sep01.actualCumM(), 1e-9);
        Assertions.assertEquals(30, sep01.planCumM(), 1e-9);
        Assertions.assertEquals(50, sep02.actualCumM(), 1e-9);
        Assertions.assertEquals(70, sep02.planCumM(), 1e-9);
        Assertions.assertEquals(sep02.actualCumM(), sep02.projectedCumM(), 1e-9);
        // 期間合計は通し（リセットしない）
        Assertions.assertEquals(200, r.actualTotalM(), 1e-9);
        Assertions.assertEquals(100, r.planTotalM(), 1e-9);
    }

    @Test
    void aggregate_todayUsesActualForProjectionToConnectAtTip() {
        ActualsSnapshot act =
                new ActualsSnapshot(
                        ACT_HEADERS,
                        List.of(
                                List.of("W9-1", "R1", "スリット", "2026/09/02", "", "500", "100"),
                                // 当日: 予定 600 > 実績 10 でも見込は実績（先端接続）
                                List.of("W9-1", "R5", "スリット", "2026/09/03", "", "100", "10")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        act, aladdin(), dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        DayPoint d3 = r.days().get(2);
        Assertions.assertEquals(10, d3.actualM(), 1e-9);
        Assertions.assertEquals(600, d3.planM(), 1e-9);
        Assertions.assertFalse(d3.usesPlanForProjection());
        Assertions.assertEquals(d3.actualCumM(), d3.projectedCumM(), 1e-9);
        // 9/1 実績 0, 9/2 実績 100, 9/3 実績 10, 9/4 予定 80
        Assertions.assertEquals(100 + 10 + 80, r.days().get(3).projectedCumM(), 1e-9);
        Assertions.assertEquals(80, r.remainingPlanM(), 1e-9);
        Assertions.assertEquals(100 + 10 + 80, r.projectedTotalM(), 1e-9);
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
        // 当日まで 760 / 期間合計 840 ≥ 10%
        Assertions.assertTrue(r.progressDenominatorSufficient());

        AladdinSnapshot mostlyFuture =
                new AladdinSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "2026/09/01", "2026/09/04"),
                        List.of(List.of("W9-1", "R1", "スリット", "50", "950")));
        Result r2 =
                ProcessingTrendAggregator.aggregate(
                        actuals(), mostlyFuture, dispatch(),
                        new Filter(FROM, TO, PlanSource.ALADDIN, null, null), TODAY);
        // 当日まで 50 / 期間合計 1000 = 5% < 10%
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
        // 当日以前の配台は予定に含めない（実績との二重計上防止）
        Assertions.assertEquals(0, r.days().get(0).planM(), 1e-9);
        Assertions.assertEquals(0, r.days().get(1).planM(), 1e-9);
        Assertions.assertEquals(0, r.days().get(2).planM(), 1e-9);
        Assertions.assertEquals(333, r.days().get(3).planM(), 1e-9);
        Assertions.assertEquals(333, r.planTotalM(), 1e-9);
        Assertions.assertEquals(1, r.planRowsCounted());
    }

    @Test
    void aggregate_legacyDispatchWithActualQtyColumn_doesNotDoubleCount() {
        // 旧 段階3 JSON: 目標行（当日配台数量のみ）とタイムライン行（実配台数量・加工開始日時）が同一 (依頼,工程,機械) に共存
        // 配台日は翌日以降（当日以前は予定スキップ）に置き、統合後の数量が二重にならないことだけ検証する
        DispatchSnapshot legacy =
                new DispatchSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "配台日", "当日配台数量", "実配台数量", "加工開始日時"),
                        List.of(
                                List.of("W9-1", "R1", "スリット", "2026/09/04", "500", "", ""),
                                List.of("W9-1", "R1", "スリット", "2026/09/04", "0", "320", "2026/09/04 08:00"),
                                List.of("W9-1", "R1", "スリット", "2026/09/05", "0", "0", "2026/09/05 08:00")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        new ActualsSnapshot(ACT_HEADERS, List.of()), aladdin(), legacy,
                        new Filter(FROM, TO, PlanSource.DISPATCH, null, null), TODAY);
        Assertions.assertEquals(0, r.days().get(2).planM(), 1e-9);
        Assertions.assertEquals(320, r.days().get(3).planM(), 1e-9);
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
    void machineNames_skipsJoinedNonMachineLabels() {
        String junk =
                "(原反4mm→1mm3mmスライス),東レ株式会社,自動車材料事業部,2026/06/08,2026/06/08,1:完了";
        String junk2 =
                "【他ユーザー向け製品は要確認】,東レ株式会社,自動車材料事業部,2026/03/31,1:完了,宮島 剛";
        String junk3 = "欠点数合計:接続点数含まず";
        String junk4 = "難燃品種(FR4)";
        AladdinSnapshot al =
                new AladdinSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "2026/09/01"),
                        List.of(
                                List.of("エンボス 湖南", "A-1", "エンボス", "100"),
                                List.of(junk, "B-1", "スライス", "50"),
                                List.of(junk2, "C-1", "スリット", "10"),
                                List.of(junk3, "D-1", "EC", "1"),
                                List.of(junk4, "E-1", "SEC", "1")));
        List<String> machines = ProcessingTrendAggregator.machineNames(null, null, al, null);
        Assertions.assertEquals(List.of("エンボス 湖南"), machines);
    }

    @Test
    void isPlausibleMachineLabel_acceptsFactoryMachineNames() {
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleMachineLabel("EC機 湖南"));
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleMachineLabel("スライス機1 湘南"));
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleMachineLabel("SEC機\u3000湖南"));
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleMachineLabel("エンボス 湖南"));
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleMachineLabel("W9-1"));
        Assertions.assertFalse(
                ProcessingTrendAggregator.isPlausibleMachineLabel(
                        "(原反4mm→1mm3mmスライス),東レ株式会社,自動車材料事業部,2026/06/08"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleMachineLabel("東レ株式会社"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleMachineLabel("2026/06/08"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleMachineLabel("欠点数合計:接続点数含まず"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleMachineLabel("難燃品種(FR4)"));
    }

    @Test
    void processNames_skipsNumericAndNonProcessLabels() {
        AladdinSnapshot al =
                new AladdinSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "2026/09/01"),
                        List.of(
                                List.of("EC機 湖南", "A-1", "エンボス", "100"),
                                List.of("SEC機 湖南", "B-1", "000,2", "50"),
                                List.of("スライス機1 湖南", "C-1", "750,1", "10"),
                                List.of("スリット機1 湖南", "D-1", "EC", "1"),
                                List.of("熱融着機 湖南", "E-1", "接続", "1"),
                                List.of("エンボス 湖南", "F-1", "難燃品種(FR4)", "1")));
        List<String> processes = ProcessingTrendAggregator.processNames(null, null, al, null);
        Assertions.assertEquals(List.of("EC", "エンボス", "接続"), processes);
    }

    @Test
    void isPlausibleProcessLabel_acceptsFactoryProcessNames() {
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleProcessLabel("EC"));
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleProcessLabel("SEC"));
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleProcessLabel("エンボス"));
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleProcessLabel("スライス"));
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleProcessLabel("接続"));
        Assertions.assertTrue(ProcessingTrendAggregator.isPlausibleProcessLabel("欠点表示"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleProcessLabel("000,2"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleProcessLabel("250,1"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleProcessLabel("600,1"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleProcessLabel("750,1"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleProcessLabel("難燃品種(FR4)"));
        Assertions.assertFalse(ProcessingTrendAggregator.isPlausibleProcessLabel("EC機 湖南"));
    }

    @Test
    void progressPct_isCappedAt100() {
        ActualsSnapshot act =
                new ActualsSnapshot(
                        ACT_HEADERS,
                        List.of(
                                List.of("W9-1", "R1", "スリット", "2026/09/01", "", "100", "500"),
                                List.of("W9-1", "R2", "スリット", "2026/09/02", "", "100", "500")));
        AladdinSnapshot al =
                new AladdinSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "2026/09/01", "2026/09/02", "2026/09/03"),
                        List.of(List.of("W9-1", "R1", "スリット", "100", "50", "200")));
        Result r =
                ProcessingTrendAggregator.aggregate(
                        act,
                        al,
                        null,
                        new Filter(
                                LocalDate.of(2026, 9, 1),
                                LocalDate.of(2026, 9, 3),
                                PlanSource.ALADDIN,
                                null,
                                null),
                        LocalDate.of(2026, 9, 3));
        Assertions.assertTrue(r.actualToDateM() > r.planToDateM());
        Assertions.assertEquals(100.0, r.progressPct(), 1e-9);
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

    @Test
    void rollUpMonthly_aggregatesDaysIntoMonthsAcrossYears() {
        LocalDate start = LocalDate.of(2025, 12, 15);
        LocalDate end = LocalDate.of(2026, 2, 10);
        Filter f = new Filter(start, end, PlanSource.ALADDIN, null, null);
        LocalDate today = LocalDate.of(2026, 1, 10);

        // ダミー日次データ: 12月(17日), 1月(31日), 2月(10日) — 累計は月単位リセット
        List<DayPoint> days = new java.util.ArrayList<>();
        double actTotal = 0;
        double planTotal = 0;
        double actualCum = 0;
        double planCum = 0;
        double projCum = 0;
        java.time.YearMonth cumYm = null;
        for (LocalDate d = start; !d.isAfter(end); d = d.plusDays(1)) {
            java.time.YearMonth ym = java.time.YearMonth.from(d);
            if (cumYm == null || !ym.equals(cumYm)) {
                actualCum = 0;
                planCum = 0;
                projCum = 0;
                cumYm = ym;
            }
            double act = d.getMonthValue() == 12 ? 10.0 : (d.getMonthValue() == 1 ? 20.0 : 5.0);
            double pl = 15.0;
            actTotal += act;
            planTotal += pl;
            actualCum += act;
            planCum += pl;
            boolean usePlan = d.isAfter(today);
            projCum += usePlan ? pl : act;
            days.add(new DayPoint(d, act, pl, actualCum, planCum, projCum, usePlan));
        }

        Result daily =
                new Result(
                        days, actTotal, planTotal, 500.0, 400.0, 300.0, actTotal + 100.0,
                        today, 10, 10, start, end, List.of("テスト注意"));

        ProcessingTrendAggregator.MonthlyResult mr =
                ProcessingTrendAggregator.rollUpMonthly(daily, f, today);

        Assertions.assertEquals(3, mr.months().size());

        ProcessingTrendAggregator.MonthPoint m1 = mr.months().get(0); // 2025-12
        Assertions.assertEquals(java.time.YearMonth.of(2025, 12), m1.month());
        Assertions.assertEquals(17, m1.daysInBucket());
        Assertions.assertTrue(m1.incomplete());
        Assertions.assertFalse(m1.isCurrentMonth());
        Assertions.assertFalse(m1.usesPlanForProjection());
        Assertions.assertEquals(170.0, m1.actualM(), 1e-9);
        Assertions.assertEquals(17 * 15.0, m1.planM(), 1e-9);
        Assertions.assertEquals(170.0, m1.actualCumM(), 1e-9);
        Assertions.assertEquals(17 * 15.0, m1.planCumM(), 1e-9);
        Assertions.assertEquals(170.0 - (17 * 15.0), m1.diffM(), 1e-9);

        ProcessingTrendAggregator.MonthPoint m2 = mr.months().get(1); // 2026-01
        Assertions.assertEquals(java.time.YearMonth.of(2026, 1), m2.month());
        Assertions.assertEquals(31, m2.daysInBucket());
        Assertions.assertFalse(m2.incomplete());
        Assertions.assertTrue(m2.isCurrentMonth());
        Assertions.assertTrue(m2.usesPlanForProjection());
        Assertions.assertEquals(31 * 20.0, m2.actualM(), 1e-9);
        Assertions.assertEquals(31 * 20.0, m2.actualCumM(), 1e-9);

        ProcessingTrendAggregator.MonthPoint m3 = mr.months().get(2); // 2026-02
        Assertions.assertEquals(java.time.YearMonth.of(2026, 2), m3.month());
        Assertions.assertEquals(10, m3.daysInBucket());
        Assertions.assertTrue(m3.incomplete());
        Assertions.assertFalse(m3.isCurrentMonth());
        Assertions.assertTrue(m3.usesPlanForProjection());
        Assertions.assertEquals(50.0, m3.actualM(), 1e-9);
        Assertions.assertEquals(10 * 15.0, m3.planM(), 1e-9);

        // 月末累計は月内合計（通し累計ではない）
        Assertions.assertEquals(daily.actualCumAtEnd(), m3.actualCumM(), 1e-9);
        Assertions.assertEquals(m3.planM(), m3.planCumM(), 1e-9);
        Assertions.assertEquals(daily.actualTotalM(), mr.actualTotalM(), 1e-9);
        Assertions.assertEquals(daily.planTotalM(), mr.planTotalM(), 1e-9);
        Assertions.assertEquals(daily.projectedTotalM(), mr.projectedTotalM(), 1e-9);
        Assertions.assertEquals(start, mr.periodFrom());
        Assertions.assertEquals(end, mr.periodTo());
    }

    @Test
    void aggregate_switchesBetweenDailyReportAndDetailWithComparison() {
        ActualsSnapshot daily =
                new ActualsSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "加工日", "実加工量"),
                        List.of(
                                List.of("W9-1", "R1", "スリット", "2026/09/01", "120"),
                                List.of("W9-1", "R2", "スリット", "2026/09/02", "150")));
        ActualsSnapshot detail =
                new ActualsSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "加工開始日時", "実加工数"),
                        List.of(
                                List.of("W9-1", "R1", "スリット", "2026/09/01 08:00:00", "100"),
                                List.of("W9-1", "R2", "スリット", "2026/09/02 08:00:00", "130")));

        Filter fDaily =
                new Filter(
                        FROM,
                        TO,
                        ProcessingTrendAggregator.ActualSource.DAILY_REPORT,
                        PlanSource.ALADDIN,
                        "W9-1",
                        null);
        Result rDaily =
                ProcessingTrendAggregator.aggregate(
                        daily, detail, aladdin(), dispatch(), fDaily, TODAY);

        // 主実績 = 日報 (120, 150 -> 270)
        Assertions.assertEquals(270.0, rDaily.actualTotalM(), 1e-9);
        Assertions.assertEquals(120.0, rDaily.days().get(0).actualM(), 1e-9);
        // 比較実績 = 実績明細 (100, 130 -> 230)
        Assertions.assertEquals(230.0, rDaily.compareActualTotalM(), 1e-9);
        Assertions.assertEquals(100.0, rDaily.days().get(0).compareActualM(), 1e-9);
        // 明細差 = 120 - 100 = 20, 150 - 130 = 20 -> 40
        Assertions.assertEquals(20.0, rDaily.days().get(0).actualCompareDiffM(), 1e-9);
        Assertions.assertEquals(40.0, rDaily.actualCompareDiffTotalM(), 1e-9);
        Assertions.assertEquals("実績明細", rDaily.compareSourceLabel());

        // 逆に DETAIL を主実績にした場合
        Filter fDetail =
                new Filter(
                        FROM,
                        TO,
                        ProcessingTrendAggregator.ActualSource.DETAIL,
                        PlanSource.ALADDIN,
                        "W9-1",
                        null);
        Result rDetail =
                ProcessingTrendAggregator.aggregate(
                        daily, detail, aladdin(), dispatch(), fDetail, TODAY);

        Assertions.assertEquals(230.0, rDetail.actualTotalM(), 1e-9);
        Assertions.assertEquals(270.0, rDetail.compareActualTotalM(), 1e-9);
        Assertions.assertEquals(-20.0, rDetail.days().get(0).actualCompareDiffM(), 1e-9);
        Assertions.assertEquals(-40.0, rDetail.actualCompareDiffTotalM(), 1e-9);
        Assertions.assertEquals("日報", rDetail.compareSourceLabel());
    }

    @Test
    void aggregate_calculates7DayMovingAverageWithPriorDays() {
        // 8/26 〜 9/03 の日次実績（各日 70m）
        List<List<String>> rows = new ArrayList<>();
        LocalDate start = LocalDate.of(2026, 8, 26);
        for (int i = 0; i < 9; i++) {
            LocalDate d = start.plusDays(i);
            rows.add(List.of("W9-1", "R" + i, "スリット", d.format(DateTimeFormatter.ofPattern("yyyy/MM/dd")), "70"));
        }
        ActualsSnapshot daily =
                new ActualsSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "加工日", "実加工量"),
                        rows);

        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 5);
        Filter filter =
                new Filter(
                        from,
                        to,
                        ActualSource.DAILY_REPORT,
                        PlanSource.ALADDIN,
                        "W9-1",
                        null,
                        7);

        Result r = ProcessingTrendAggregator.aggregate(
                daily, null, null, null, filter, LocalDate.of(2026, 9, 3));

        // 期間内の初日（9/1）: 8/26〜9/1 の 7 日間（すべて 70m）の平均 = 70.0m
        Assertions.assertEquals(70.0, r.days().get(0).actualMaM(), 1e-9);
        // 9/2: 8/27〜9/2 の 7 日間の平均 = 70.0m
        Assertions.assertEquals(70.0, r.days().get(1).actualMaM(), 1e-9);
        // 9/3: 8/28〜9/3 の 7 日間の平均 = 70.0m
        Assertions.assertEquals(70.0, r.days().get(2).actualMaM(), 1e-9);
        // 9/4: 8/29〜9/4（9/4 は実績 0 なので 6 日分 420m / 7 = 60.0m）
        Assertions.assertEquals(60.0, r.days().get(3).actualMaM(), 1e-9);
    }

    @Test
    void aggregate_defaultMovingAverageWindowIs30Days() {
        List<List<String>> rows = new ArrayList<>();
        // 8/3 〜 9/3: 各日 100m（期間 9/1〜9/5、窓 30 の初日は 8/3〜9/1）
        LocalDate dataStart = LocalDate.of(2026, 8, 3);
        for (int i = 0; i < 32; i++) {
            LocalDate d = dataStart.plusDays(i);
            rows.add(
                    List.of(
                            "W9-1",
                            "R" + i,
                            "スリット",
                            d.format(DateTimeFormatter.ofPattern("yyyy/MM/dd")),
                            "100"));
        }
        ActualsSnapshot daily =
                new ActualsSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "加工日", "実加工量"), rows);

        LocalDate from = LocalDate.of(2026, 9, 1);
        LocalDate to = LocalDate.of(2026, 9, 5);
        // 窓幅未指定 → 既定 30
        Filter filter = new Filter(from, to, PlanSource.ALADDIN, "W9-1", null);
        Assertions.assertEquals(30, filter.movingAverageDays());

        Result r =
                ProcessingTrendAggregator.aggregate(
                        daily, null, null, null, filter, LocalDate.of(2026, 9, 3));

        // 9/1: 8/3〜9/1 の 30 日すべて 100m → 平均 100
        Assertions.assertEquals(100.0, r.days().get(0).actualMaM(), 1e-9);
        // 9/4: 8/6〜9/4（9/4 は実績 0）→ 29*100/30
        Assertions.assertEquals(100.0 * 29 / 30.0, r.days().get(3).actualMaM(), 1e-9);
    }

    @Test
    void aggregate_calculates14DayMovingAverage() {
        List<List<String>> rows = new ArrayList<>();
        LocalDate start = LocalDate.of(2026, 8, 19);
        for (int i = 0; i < 17; i++) {
            LocalDate d = start.plusDays(i);
            rows.add(
                    List.of(
                            "W9-1",
                            "R" + i,
                            "スリット",
                            d.format(DateTimeFormatter.ofPattern("yyyy/MM/dd")),
                            "50"));
        }
        ActualsSnapshot daily =
                new ActualsSnapshot(
                        List.of("機械名", "依頼NO", "工程名", "加工日", "実加工量"), rows);

        Filter filter =
                new Filter(
                        LocalDate.of(2026, 9, 1),
                        LocalDate.of(2026, 9, 3),
                        ActualSource.DAILY_REPORT,
                        PlanSource.ALADDIN,
                        "W9-1",
                        null,
                        14);

        Result r =
                ProcessingTrendAggregator.aggregate(
                        daily, null, null, null, filter, LocalDate.of(2026, 9, 3));

        // 9/1: 8/19〜9/1 の 14 日すべて 50m
        Assertions.assertEquals(50.0, r.days().get(0).actualMaM(), 1e-9);
    }
}
