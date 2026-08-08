package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus.PlanLine;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus.Status;

class EquipmentStatusCardFactoryTest {

    @Test
    void statusLabel_coversAllStatuses() {
        Assertions.assertEquals("停機", EquipmentStatusCardFactory.statusLabel(Status.STOPPED));
        Assertions.assertEquals("稼働中", EquipmentStatusCardFactory.statusLabel(Status.RUNNING));
        Assertions.assertEquals("予定達成", EquipmentStatusCardFactory.statusLabel(Status.COMPLETED));
    }

    @Test
    void cardStateStyleClass_mapsStatusToAccentClass() {
        Assertions.assertEquals(
                "pm-equipment-status-card-stopped", EquipmentStatusCardFactory.cardStateStyleClass(Status.STOPPED));
        Assertions.assertEquals(
                "pm-equipment-status-card-completed",
                EquipmentStatusCardFactory.cardStateStyleClass(Status.COMPLETED));
        Assertions.assertEquals(
                "pm-equipment-status-card-running", EquipmentStatusCardFactory.cardStateStyleClass(null));
    }

    @Test
    void shouldShowStatusChip_hidesRunningChipOnlyWhenRequested() {
        EquipmentMachineStatus running =
                new EquipmentMachineStatus("MC-1", Status.RUNNING, Optional.empty(), List.of(), List.of());
        EquipmentMachineStatus stopped =
                new EquipmentMachineStatus("MC-2", Status.STOPPED, Optional.empty(), List.of(), List.of());
        EquipmentStatusCardFactory.DisplayOptions hideRunning =
                new EquipmentStatusCardFactory.DisplayOptions(true, true, "", "", false);
        Assertions.assertFalse(EquipmentStatusCardFactory.shouldShowStatusChip(running, hideRunning));
        Assertions.assertTrue(EquipmentStatusCardFactory.shouldShowStatusChip(stopped, hideRunning));
        Assertions.assertTrue(EquipmentStatusCardFactory.shouldShowStatusChip(running, null));
    }

    @Test
    void planLineText_joinsRequestProcessAndQty() {
        Assertions.assertEquals(
                "R-1 · 切断 · 120m", EquipmentStatusCardFactory.planLineText(new PlanLine("R-1", "切断", "120")));
    }

    @Test
    void remainingPlanLinesText_listsOnlyOverflowLines() {
        List<PlanLine> lines =
                List.of(
                        new PlanLine("R-1", "切断", "10"),
                        new PlanLine("R-2", "切断", "20"),
                        new PlanLine("R-3", "切断", "30"));
        Assertions.assertEquals(
                "R-2 · 切断 · 20m\nR-3 · 切断 · 30m",
                EquipmentStatusCardFactory.remainingPlanLinesText(lines, 1));
        Assertions.assertEquals("", EquipmentStatusCardFactory.remainingPlanLinesText(lines, 3));
    }

    @Test
    void cardAccessibleText_includesMachineStatusAndRate() {
        EquipmentMachineStatus status =
                new EquipmentMachineStatus(
                        "MC-1",
                        Status.RUNNING,
                        Optional.of(
                                new EquipmentMachineStatus.ActualTaskRow(
                                        "R-1", "切断", 100, 42.4, "山田", "", "")),
                        List.of(),
                        List.of());
        Assertions.assertEquals(
                "MC-1、稼働中、アラジン達成率 42パーセント",
                EquipmentStatusCardFactory.cardAccessibleText(status));
    }

    @Test
    void nz_replacesBlankWithDash() {
        Assertions.assertEquals("—", EquipmentStatusCardFactory.nz(null));
        Assertions.assertEquals("—", EquipmentStatusCardFactory.nz("   "));
        Assertions.assertEquals("2026/08/09", EquipmentStatusCardFactory.nz(" 2026/08/09 "));
    }
}
