package jp.co.pm.ai.desktop.ui;

import java.util.EnumSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus.Status;
import jp.co.pm.ai.desktop.ui.EquipmentStatusDashboardFilter.SortOrder;
import jp.co.pm.ai.desktop.ui.EquipmentStatusDashboardFilter.StatusCounts;

class EquipmentStatusDashboardFilterTest {

    private static EquipmentMachineStatus machine(String name, Status status, double completionPct) {
        Optional<EquipmentMachineStatus.ActualTaskRow> task =
                status == Status.STOPPED
                        ? Optional.empty()
                        : Optional.of(
                                new EquipmentMachineStatus.ActualTaskRow(
                                        "R-1", "切断", 100, completionPct, "山田", "", ""));
        return new EquipmentMachineStatus(name, status, task, List.of(), List.of());
    }

    private static List<EquipmentMachineStatus> sample() {
        return List.of(
                machine("MC-3", Status.COMPLETED, 120),
                machine("MC-1", Status.RUNNING, 40),
                machine("MC-2", Status.STOPPED, 0),
                machine("MC-4", Status.RUNNING, 10));
    }

    @Test
    void countByStatus_countsEachStatus() {
        StatusCounts counts = EquipmentStatusDashboardFilter.countByStatus(sample());
        Assertions.assertEquals(2, counts.running());
        Assertions.assertEquals(1, counts.stopped());
        Assertions.assertEquals(1, counts.completed());
        Assertions.assertEquals(4, counts.total());
        Assertions.assertEquals(1, counts.of(Status.STOPPED));
    }

    @Test
    void countByStatus_toleratesNullList() {
        StatusCounts counts = EquipmentStatusDashboardFilter.countByStatus(null);
        Assertions.assertEquals(0, counts.total());
    }

    @Test
    void apply_withoutFilters_sortsByMachineName() {
        List<EquipmentMachineStatus> result =
                EquipmentStatusDashboardFilter.apply(sample(), Set.of(), "", SortOrder.MACHINE_NAME);
        Assertions.assertEquals(
                List.of("MC-1", "MC-2", "MC-3", "MC-4"), result.stream().map(EquipmentMachineStatus::machineName).toList());
    }

    @Test
    void apply_stoppedFirst_putsStoppedBeforeRunningAndCompleted() {
        List<EquipmentMachineStatus> result =
                EquipmentStatusDashboardFilter.apply(sample(), null, null, SortOrder.STOPPED_FIRST);
        Assertions.assertEquals(
                List.of("MC-2", "MC-1", "MC-4", "MC-3"),
                result.stream().map(EquipmentMachineStatus::machineName).toList());
    }

    @Test
    void apply_completionAsc_putsStoppedFirstThenLowestRate() {
        List<EquipmentMachineStatus> result =
                EquipmentStatusDashboardFilter.apply(sample(), null, null, SortOrder.COMPLETION_ASC);
        Assertions.assertEquals(
                List.of("MC-2", "MC-4", "MC-1", "MC-3"),
                result.stream().map(EquipmentMachineStatus::machineName).toList());
    }

    @Test
    void apply_statusFilter_keepsOnlySelectedStatuses() {
        List<EquipmentMachineStatus> result =
                EquipmentStatusDashboardFilter.apply(
                        sample(), EnumSet.of(Status.STOPPED), "", SortOrder.MACHINE_NAME);
        Assertions.assertEquals(1, result.size());
        Assertions.assertEquals("MC-2", result.get(0).machineName());
    }

    @Test
    void apply_keyword_matchesIgnoringCaseWidthAndSpaces() {
        List<EquipmentMachineStatus> result =
                EquipmentStatusDashboardFilter.apply(
                        List.of(machine("ＭＣ-10", Status.RUNNING, 50), machine("LC-2", Status.RUNNING, 50)),
                        null,
                        " mc-1 ",
                        SortOrder.MACHINE_NAME);
        Assertions.assertEquals(1, result.size());
        Assertions.assertEquals("ＭＣ-10", result.get(0).machineName());
    }

    @Test
    void apply_emptyInputReturnsEmptyList() {
        Assertions.assertTrue(
                EquipmentStatusDashboardFilter.apply(null, null, "MC", SortOrder.MACHINE_NAME).isEmpty());
        Assertions.assertTrue(
                EquipmentStatusDashboardFilter.apply(List.of(), null, "", SortOrder.MACHINE_NAME).isEmpty());
    }

    @Test
    void sortOrder_fromLabel_fallsBackToMachineName() {
        Assertions.assertEquals(SortOrder.STOPPED_FIRST, SortOrder.fromLabel("停機を先頭"));
        Assertions.assertEquals(SortOrder.COMPLETION_ASC, SortOrder.fromLabel(" 達成率が低い順 "));
        Assertions.assertEquals(SortOrder.MACHINE_NAME, SortOrder.fromLabel("存在しない"));
        Assertions.assertEquals(SortOrder.MACHINE_NAME, SortOrder.fromLabel(null));
    }

    @Test
    void normalizeKeyword_stripsWidthCaseAndSpaces() {
        Assertions.assertEquals("mc-1", EquipmentStatusDashboardFilter.normalizeKeyword(" ＭC-1 "));
        Assertions.assertEquals("mc1", EquipmentStatusDashboardFilter.normalizeKeyword("MC\u30001"));
        Assertions.assertEquals("", EquipmentStatusDashboardFilter.normalizeKeyword(null));
    }

    @Test
    void sortCompletionPct_returnsSentinelWhenNoActual() {
        Assertions.assertEquals(
                EquipmentStatusDashboardFilter.NO_ACTUAL_COMPLETION_PCT,
                EquipmentStatusDashboardFilter.sortCompletionPct(machine("MC-9", Status.STOPPED, 0)),
                0.001);
    }
}
