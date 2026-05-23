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
                        TODAY);
        Assertions.assertEquals(1, out.size());
        Assertions.assertEquals(EquipmentMachineStatus.Status.STOPPED, out.get(0).status());
        Assertions.assertEquals(1, out.get(0).aladdinPlans().size());
    }

    @Test
    void build_todayRunningAndCompleted() {
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
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(headers, rows),
                        new AladdinSnapshot(List.of(), List.of()),
                        new DispatchSnapshot(List.of(), List.of()),
                        TODAY,
                        TODAY);
        EquipmentMachineStatus running =
                out.stream().filter(s -> "M-run".equals(s.machineName())).findFirst().orElseThrow();
        EquipmentMachineStatus done =
                out.stream().filter(s -> "M-done".equals(s.machineName())).findFirst().orElseThrow();
        Assertions.assertEquals(EquipmentMachineStatus.Status.RUNNING, running.status());
        Assertions.assertEquals(45.0, running.actualTask().orElseThrow().completionPct(), 0.01);
        Assertions.assertEquals(EquipmentMachineStatus.Status.COMPLETED, done.status());
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
                        TOMORROW);
        Assertions.assertEquals(1, out.size());
        Assertions.assertEquals(EquipmentMachineStatus.Status.STOPPED, out.get(0).status());
        Assertions.assertEquals(1, out.get(0).aladdinPlans().size());
        Assertions.assertEquals("A1", out.get(0).aladdinPlans().get(0).requestNo());
        Assertions.assertEquals(1, out.get(0).dispatchPlans().size());
        Assertions.assertEquals("D1", out.get(0).dispatchPlans().get(0).requestNo());
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
    void build_actualDateYesterday() {
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
        List<EquipmentMachineStatus> out =
                EquipmentStatusDashboardBuilder.build(
                        new ActualsSnapshot(headers, rows),
                        new AladdinSnapshot(List.of(), List.of()),
                        new DispatchSnapshot(List.of(), List.of()),
                        YESTERDAY,
                        TODAY);
        Assertions.assertEquals(EquipmentMachineStatus.Status.RUNNING, out.get(0).status());
    }
}
