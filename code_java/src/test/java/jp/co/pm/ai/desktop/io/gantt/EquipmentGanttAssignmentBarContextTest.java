package jp.co.pm.ai.desktop.io.gantt;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;

import org.junit.jupiter.api.Test;

class EquipmentGanttAssignmentBarContextTest {

    private static final Path FIXTURE =
            Path.of("src/test/resources/gantt/assignment_dnd_manual_equipment_gantt_contract.json");

    @Test
    void resolve_fromFixtureTableRow_returnsMachineName() throws Exception {
        EquipmentGanttSheetBundle bundle =
                EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(FIXTURE);
        String barId = bundle.assignmentMetadata().barUnits().getFirst().barId();

        var ctx =
                EquipmentGanttAssignmentBarContext.resolve(
                        bundle.assignmentMetadata(), bundle.table(), barId);

        assertTrue(ctx.isPresent());
        assertEquals("EC機　湖南", ctx.get().machineName());
    }

    @Test
    void resolve_unknownBar_returnsEmpty() throws Exception {
        EquipmentGanttSheetBundle bundle =
                EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(FIXTURE);

        assertFalse(
                EquipmentGanttAssignmentBarContext.resolve(
                                bundle.assignmentMetadata(), bundle.table(), "ev:missing")
                        .isPresent());
    }
}
