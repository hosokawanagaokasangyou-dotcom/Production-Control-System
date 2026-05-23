package jp.co.pm.ai.desktop.io.actuals;

import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardBuilder.ActualsSnapshot;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardBuilder.AladdinSnapshot;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardBuilder.DispatchSnapshot;

class EquipmentStatusDashboardBuilderTest {

    private static final LocalDate TODAY = LocalDate.of(2026, 5, 23);
    private static final LocalDate YESTERDAY = TODAY.minusDays(1);
    private static final LocalDate TOMORROW = TODAY.plusDays(1);

    @Test
    void build_yesterdayOnlyActual_notOnTodayActualDateUnion() {
        List<String> headers =
                List.of("機械名", "依頼NO", "工程名", "加工日", "加工開始日時", "換算数量", "実加工数");
        List<List<String>> rows =
                List.of(
                        List.of(
                                "M1",
                                "R1",
                                "P1",
                                "2026/05/22",
                                "2026/05/22 10:00",
                                "100",
                                "50"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(headers, rows),
                        new AladdinSnapshot(List.of(), List.of()),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals(0, out.size());
    }

    @Test
    void build_planOnlyMachine_showsStoppedOnTodayActual() {
        List<String> alHeaders = List.of("機械名", "依頼NO", "工程名", "2026/05/23");
        List<List<String>> alRows = List.of(List.of("M1", "A1", "AP1", "50"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(List.of(), List.of()),
                        new AladdinSnapshot(alHeaders, alRows),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals(1, out.size());
        Assertions.assertEquals(EquipmentMachineStatus.Status.STOPPED, out.get(0).status());
        Assertions.assertEquals(1, out.get(0).aladdinPlans().size());
    }

    @Test
    void build_todayRunningWhenActualBelowAladdinPlan() {
        List<String> headers =
                List.of(
                        "機械名",
                        "依頼NO",
                        "工程名",
                        "加工日",
                        "加工開始日時",
                        "換算数量",
                        "累積完了率",
                        "メンバー名");
        List<List<String>> rows =
                List.of(
                        List.of(
                                "M-run",
                                "R1",
                                "P1",
                                "2026/05/23",
                                "2026/05/23 09:00",
                                "100",
                                "45%",
                                "山田太郎"),
                        List.of(
                                "M-done",
                                "R2",
                                "P2",
                                "2026/05/23",
                                "2026/05/23 08:00",
                                "200",
                                "100%",
                                "佐藤花子"));
        List<String> alHeaders = List.of("機械名", "依頼NO", "工程名", "2026/05/23");
        List<List<String>> alRows =
                List.of(
                        List.of("M-run", "R1", "P1", "200"),
                        List.of("M-done", "R2", "P2", "200"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(headers, rows),
                        new AladdinSnapshot(alHeaders, alRows),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY,
                        TODAY);
        EquipmentMachineStatus running =
                out.stream().filter(s -> "M-run".equals(s.machineName())).findFirst().orElseThrow();
        EquipmentMachineStatus done =
                out.stream().filter(s -> "M-done".equals(s.machineName())).findFirst().orElseThrow();
        Assertions.assertEquals(EquipmentMachineStatus.Status.RUNNING, running.status());
        Assertions.assertEquals(22.5, running.actualTask().orElseThrow().completionPct(), 0.01);
        Assertions.assertEquals(EquipmentMachineStatus.Status.COMPLETED, done.status());
        Assertions.assertEquals(100.0, done.actualTask().orElseThrow().completionPct(), 0.01);
    }

    @Test
    void build_todayNotRunningWhenActualReachesAladdinPlan() {
        List<String> headers =
                List.of(
                        "機械名",
                        "依頼NO",
                        "工程名",
                        "加工開始日時",
                        "換算数量",
                        "実加工数",
                        "累積完了率");
        List<List<String>> rows =
                List.of(
                        List.of(
                                "M1",
                                "R1",
                                "P1",
                                "2026/05/23 09:00",
                                "100",
                                "80",
                                "80%"));
        List<String> alHeaders = List.of("機械名", "依頼NO", "工程名", "2026/05/23");
        List<List<String>> alRows = List.of(List.of("M1", "R1", "P1", "80"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(headers, rows),
                        new AladdinSnapshot(alHeaders, alRows),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals(EquipmentMachineStatus.Status.COMPLETED, out.get(0).status());
        Assertions.assertEquals(100.0, out.get(0).actualTask().orElseThrow().completionPct(), 0.01);
    }

    @Test
    void completionPctFromActualAndPlan_dividesActualByAladdinPlan() {
        Assertions.assertEquals(
                22.5,
                EquipmentStatusDashboardBuilder.completionPctFromActualAndPlan(45, 200),
                0.01);
        Assertions.assertEquals(
                0.0,
                EquipmentStatusDashboardBuilder.completionPctFromActualAndPlan(50, 0),
                0.01);
        Assertions.assertEquals(
                100.0,
                EquipmentStatusDashboardBuilder.completionPctFromActualAndPlan(200, 200),
                0.01);
    }

    @Test
    void build_todayNotRunningWhenNoAladdinPlanButHasActual() {
        List<String> headers =
                List.of(
                        "機械名",
                        "依頼NO",
                        "工程名",
                        "加工開始日時",
                        "換算数量",
                        "累積完了率");
        List<List<String>> rows =
                List.of(
                        List.of(
                                "M1",
                                "R1",
                                "P1",
                                "2026/05/23 09:00",
                                "100",
                                "45%"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(headers, rows),
                        new AladdinSnapshot(List.of(), List.of()),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals(EquipmentMachineStatus.Status.COMPLETED, out.get(0).status());
        Assertions.assertEquals(0.0, out.get(0).actualTask().orElseThrow().completionPct(), 0.01);
    }

    @Test
    void build_picksLatestStartAmongTodayRows() {
        List<String> headers =
                List.of("機械名", "依頼NO", "工程名", "加工開始日時", "換算数量", "累積完了率");
        List<List<String>> rows =
                List.of(
                        List.of("M1", "OLD", "P-old", "2026/05/23 08:00", "100", "10%"),
                        List.of("M1", "NEW", "P-new", "2026/05/23 14:00", "100", "20%"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(headers, rows),
                        new AladdinSnapshot(List.of(), List.of()),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals("NEW", out.get(0).actualTask().orElseThrow().requestNo());
    }

    @Test
    void build_aladdinAndDispatchPlansOnPlanDate() {
        List<String> alHeaders = List.of("機械名", "依頼NO", "工程名", "2026/05/24");
        List<List<String>> alRows =
                List.of(List.of("M1", "A1", "AP1", "120.5"));
        List<String> disHeaders =
                List.of(
                        ResultDispatchSchema.COL_MACHINE,
                        "依頼NO",
                        ResultDispatchSchema.COL_PROCESS,
                        ResultDispatchSchema.COL_DISPATCH_DATE,
                        ResultDispatchSchema.COL_DISPATCH_QTY);
        List<List<String>> disRows =
                List.of(List.of("M1", "D1", "DP1", "2026/05/24", "80"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(List.of(), List.of()),
                        new AladdinSnapshot(alHeaders, alRows),
                        new DispatchSnapshot(disHeaders, disRows),
                        TODAY,
                        TOMORROW,
                        TODAY);
        Assertions.assertEquals(1, out.size());
        Assertions.assertEquals(EquipmentMachineStatus.Status.STOPPED, out.get(0).status());
        Assertions.assertEquals(1, out.get(0).aladdinPlans().size());
        Assertions.assertEquals("A1", out.get(0).aladdinPlans().get(0).requestNo());
        Assertions.assertEquals(1, out.get(0).dispatchPlans().size());
        Assertions.assertEquals("D1", out.get(0).dispatchPlans().get(0).requestNo());
    }

    @Test
    void build_memberFromAladdinWhenActualsLackMemberColumn() {
        List<String> actHeaders =
                List.of(
                        "機械名",
                        "依頼NO",
                        "工程名",
                        "加工開始日時",
                        "換算数量",
                        "累積完了率");
        List<List<String>> actRows =
                List.of(
                        List.of(
                                "M1",
                                "R1",
                                "スリット",
                                "2026/05/23 09:00",
                                "100",
                                "45%"));
        List<String> alHeaders =
                List.of("機械名", "依頼NO", "工程名", "担当OP_指定", "2026/05/23");
        List<List<String>> alRows =
                List.of(List.of("M1", "R1", "スリット", "田中一郎", "50"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(actHeaders, actRows),
                        new AladdinSnapshot(alHeaders, alRows),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals("田中一郎", out.get(0).actualTask().orElseThrow().memberRaw());
        Assertions.assertEquals(EquipmentMachineStatus.Status.RUNNING, out.get(0).status());
        Assertions.assertEquals(90.0, out.get(0).actualTask().orElseThrow().completionPct(), 0.01);
    }

    @Test
    void build_memberFromDispatchWhenAladdinEmpty() {
        List<String> actHeaders =
                List.of(
                        "機械名",
                        "依頼NO",
                        "工程名",
                        "加工開始日時",
                        "換算数量",
                        "累積完了率");
        List<List<String>> actRows =
                List.of(
                        List.of(
                                "M1",
                                "R1",
                                "P1",
                                "2026/05/23 10:00",
                                "100",
                                "30%"));
        List<String> disHeaders =
                List.of(
                        ResultDispatchSchema.COL_MACHINE,
                        "依頼NO",
                        ResultDispatchSchema.COL_PROCESS,
                        "メンバー名",
                        ResultDispatchSchema.COL_DISPATCH_DATE,
                        ResultDispatchSchema.COL_DISPATCH_QTY);
        List<List<String>> disRows =
                List.of(List.of("M1", "R1", "P1", "鈴木次郎", "2026/05/23", "80"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(actHeaders, actRows),
                        new AladdinSnapshot(List.of(), List.of()),
                        new DispatchSnapshot(disHeaders, disRows),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals("鈴木次郎", out.get(0).actualTask().orElseThrow().memberRaw());
        Assertions.assertEquals(0.0, out.get(0).actualTask().orElseThrow().completionPct(), 0.01);
    }

    @Test
    void build_memberFromAladdinWhenProcessNameDiffers() {
        List<String> actHeaders =
                List.of(
                        "機械名",
                        "依頼NO",
                        "工程名",
                        "加工開始日時",
                        "換算数量",
                        "累積完了率");
        List<List<String>> actRows =
                List.of(
                        List.of(
                                "M1",
                                "R1",
                                "スリット",
                                "2026/05/23 09:00",
                                "100",
                                "45%"));
        List<String> alHeaders =
                List.of("機械名", "依頼NO", "工程名", "担当OP指定", "2026/05/23");
        List<List<String>> alRows =
                List.of(List.of("M1", "R1", "カット", "高橋三郎", "50"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(actHeaders, actRows),
                        new AladdinSnapshot(alHeaders, alRows),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals("高橋三郎", out.get(0).actualTask().orElseThrow().memberRaw());
    }

    @Test
    void build_memberFromDispatchWithoutDateWhenActualDateDiffers() {
        List<String> actHeaders =
                List.of(
                        "機械名",
                        "依頼NO",
                        "工程名",
                        "加工開始日時",
                        "換算数量",
                        "累積完了率");
        List<List<String>> actRows =
                List.of(
                        List.of(
                                "M1",
                                "R1",
                                "P1",
                                "2026/05/22 10:00",
                                "100",
                                "30%"));
        List<String> disHeaders =
                List.of(
                        ResultDispatchSchema.COL_MACHINE,
                        "依頼NO",
                        ResultDispatchSchema.COL_PROCESS,
                        "メンバー名",
                        ResultDispatchSchema.COL_DISPATCH_DATE,
                        ResultDispatchSchema.COL_DISPATCH_QTY);
        List<List<String>> disRows =
                List.of(List.of("M1", "R1", "P1", "伊藤四郎", "2026/05/23", "80"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(actHeaders, actRows),
                        new AladdinSnapshot(List.of(), List.of()),
                        new DispatchSnapshot(disHeaders, disRows),
                        YESTERDAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals("伊藤四郎", out.get(0).actualTask().orElseThrow().memberRaw());
    }

    @Test
    void parseCompletionPct_fractionAndRatioFallbacks() {
        List<String> headers = List.of("累積完了率", "累積実績", "換算数量", "実加工数");
        Assertions.assertEquals(
                50.0,
                EquipmentStatusDashboardBuilder.parseCompletionPct(
                        headers, List.of("0.5", "", "", "")),
                0.01);
        Assertions.assertEquals(
                25.0,
                EquipmentStatusDashboardBuilder.parseCompletionPct(
                        headers, List.of("", "25", "100", "")),
                0.01);
        Assertions.assertEquals(
                30.0,
                EquipmentStatusDashboardBuilder.parseCompletionPct(
                        List.of("実加工数", "換算数量"),
                        List.of("30", "100")),
                0.01);
    }

    @Test
    void build_actualDateYesterday_usesActualOverAladdinForCompletionPct() {
        List<String> headers =
                List.of("機械名", "依頼NO", "工程名", "加工開始日時", "換算数量", "累積完了率");
        List<List<String>> rows =
                List.of(
                        List.of(
                                "M1",
                                "R1",
                                "P1",
                                "2026/05/22 11:00",
                                "100",
                                "80%"));
        List<String> alHeaders = List.of("機械名", "依頼NO", "工程名", "2026/05/22");
        List<List<String>> alRows = List.of(List.of("M1", "R1", "P1", "100"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(headers, rows),
                        new AladdinSnapshot(alHeaders, alRows),
                        new DispatchSnapshot(List.of(), List.of()),
                        YESTERDAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals(EquipmentMachineStatus.Status.RUNNING, out.get(0).status());
        Assertions.assertEquals(80.0, out.get(0).actualTask().orElseThrow().completionPct(), 0.01);
    }

    @Test
    void build_aladdinZeroQtyOnPlanDate_stillShowsStoppedMachine() {
        List<String> alHeaders = List.of("機械名", "依頼NO", "工程名", "2026/05/23");
        List<List<String>> alRows = List.of(List.of("M-idle", "A1", "AP1", "0"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(List.of(), List.of()),
                        new AladdinSnapshot(alHeaders, alRows),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals(1, out.size());
        Assertions.assertEquals("M-idle", out.get(0).machineName());
        Assertions.assertEquals(EquipmentMachineStatus.Status.STOPPED, out.get(0).status());
        Assertions.assertTrue(out.get(0).aladdinPlans().isEmpty());
        Assertions.assertTrue(out.get(0).actualTask().isEmpty());
    }

    @Test
    void build_dispatchZeroQtyOnPlanDate_stillShowsStoppedMachine() {
        List<String> disHeaders =
                List.of(
                        ResultDispatchSchema.COL_MACHINE,
                        "依頼NO",
                        ResultDispatchSchema.COL_PROCESS,
                        ResultDispatchSchema.COL_DISPATCH_DATE,
                        ResultDispatchSchema.COL_DISPATCH_QTY);
        List<List<String>> disRows =
                List.of(List.of("M-idle", "D1", "DP1", "2026/05/23", "0"));
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(List.of(), List.of()),
                        new AladdinSnapshot(List.of(), List.of()),
                        new DispatchSnapshot(disHeaders, disRows),
                        TODAY,
                        TODAY,
                        TODAY);
        Assertions.assertEquals(1, out.size());
        Assertions.assertEquals("M-idle", out.get(0).machineName());
        Assertions.assertEquals(EquipmentMachineStatus.Status.STOPPED, out.get(0).status());
        Assertions.assertTrue(out.get(0).dispatchPlans().isEmpty());
    }

    @Test
    void sumActualQtyM_and_sumAladdinPlanQtyM() {
        List<String> actHeaders =
                List.of("実加工数", "累積実績", "換算数量", "累積完了率");
        List<List<String>> actRows =
                List.of(
                        List.of("30", "", "100", "30%"),
                        List.of("20", "", "100", "20%"));
        Assertions.assertEquals(
                50.0,
                EquipmentStatusDashboardBuilder.sumActualQtyM(actHeaders, actRows),
                0.01);

        List<String> alHeaders = List.of("機械名", "2026/05/23");
        List<List<String>> alRows =
                List.of(List.of("M1", "100"), List.of("M1", "50"), List.of("M2", "40"));
        Assertions.assertEquals(
                150.0,
                EquipmentStatusDashboardBuilder.sumAladdinPlanQtyM(
                        alHeaders, alRows, "M1", "2026/05/23"),
                0.01);
    }
}
