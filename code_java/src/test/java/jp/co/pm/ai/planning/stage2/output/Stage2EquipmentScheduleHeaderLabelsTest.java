package jp.co.pm.ai.planning.stage2.output;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

class Stage2EquipmentScheduleHeaderLabelsTest {

    @Test
    void duplicateMachineAddsProcessQualifier() {
        List<String> eq = List.of("スライス+機A", "スリット+機A");
        assertEquals(List.of("機A（スライス）", "機A（スリット）"), Stage2EquipmentScheduleHeaderLabels.fromEquipmentCombos(eq));
    }
}
