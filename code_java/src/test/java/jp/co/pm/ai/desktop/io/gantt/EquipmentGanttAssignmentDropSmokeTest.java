package jp.co.pm.ai.desktop.io.gantt;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentEditModel.Failure;

/**
 * 担当割当 DnD の手動テスト前スモーク（GUI なし）。{@code assignment_dnd_manual_equipment_gantt_contract.json} と同データ想定。
 */
class EquipmentGanttAssignmentDropSmokeTest {

    private static final Path FIXTURE =
            Path.of("src/test/resources/gantt/assignment_dnd_manual_equipment_gantt_contract.json");

    private EquipmentGanttSheetBundle bundle;
    private EquipmentGanttAssignmentEditModel model;
    private List<List<String>> badgeRows;

    @BeforeEach
    void setUp() throws Exception {
        bundle = EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(FIXTURE);
        model = new EquipmentGanttAssignmentEditModel(bundle.assignmentMetadata());
        badgeRows = deepCopy(bundle.badgeSlotRows());
    }

    @Test
    void fixture_hasThreeBars_secondHasOnePerson_thirdIsEmpty() {
        assertEquals(3, bundle.assignmentMetadata().barUnits().size());
        assertEquals("ev:0", bundle.assignmentMetadata().barUnits().get(0).barId());
        assertEquals(2, model.personsOnBar("ev:0").size());
        assertEquals(1, model.personsOnBar("ev:1").size());
        assertTrue(model.personsOnBar("ev:2").isEmpty());
    }

    @Test
    void fixture_slotBindings_resolveOnDataRow() {
        var bindings = bundle.assignmentMetadata().slotBindings();
        assertFalse(
                bindings.isEmpty(),
                "slotBindings empty — resolveBarIdForRun failed during metadata build");
        assertTrue(
                bindings.stream().anyMatch(b -> b.tableRowIndex() == 1),
                "expected data row index 1, got: " + bindings);
        assertTrue(
                bundle.assignmentMetadata().resolveBarId(1, 0, 5).isPresent(),
                "resolveBarId(1,0,5) should match first bar slots");
    }

    @Test
    void move_toEmptyBar_updatesBadgeGrid() {
        String memberKey = model.personsOnBar("ev:0").get(1).memberKey();
        Optional<Failure> failure = model.movePerson("ev:0", "ev:2", memberKey);
        assertTrue(failure.isEmpty(), failure.map(Failure::name).orElse("ok"));

        EquipmentGanttAssignmentBadgeGridUpdater.applyToBadgeRows(
                badgeRows, bundle.assignmentMetadata(), model.snapshotPersonsByBarId());

        assertEquals(1, model.personsOnBar("ev:0").size());
        assertEquals(1, model.personsOnBar("ev:2").size());
        assertEquals("山田 太郎", model.personsOnBar("ev:0").get(0).fullName());
        assertEquals("佐藤 花子", model.personsOnBar("ev:2").get(0).fullName());
        assertTrue(badgeCellContains(badgeRows, "佐藤"));
        assertTrue(badgeCellContains(badgeRows, "山田"));
    }

    @Test
    void swap_onOccupiedBar_exchangesOperators() {
        String fromKey = model.personsOnBar("ev:0").get(0).memberKey();
        String toKey = model.personsOnBar("ev:1").get(0).memberKey();
        Optional<Failure> failure = model.swapPerson("ev:0", "ev:1", fromKey, toKey);
        assertTrue(failure.isEmpty(), failure.map(Failure::name).orElse("ok"));

        EquipmentGanttAssignmentBadgeGridUpdater.applyToBadgeRows(
                badgeRows, bundle.assignmentMetadata(), model.snapshotPersonsByBarId());

        assertTrue(
                model.personsOnBar("ev:0").stream()
                        .anyMatch(p -> "鈴木 一郎".equals(p.fullName())));
        assertTrue(
                model.personsOnBar("ev:1").stream()
                        .anyMatch(p -> "山田 太郎".equals(p.fullName())));
        assertTrue(badgeCellContains(badgeRows, "鈴木"));
        assertTrue(badgeCellContains(badgeRows, "山田"));
    }

    @Test
    void addPerson_toEmptyBar_updatesBadgeGrid() {
        EquipmentGanttAssignmentPerson extra =
                EquipmentGanttAssignmentPerson.fromRawName(
                        "田中 次郎", EquipmentGanttAssignmentRole.SUB);
        assertTrue(model.addPerson("ev:2", extra).isEmpty());

        EquipmentGanttAssignmentBadgeGridUpdater.applyToBadgeRows(
                badgeRows, bundle.assignmentMetadata(), model.snapshotPersonsByBarId());

        assertEquals(1, model.personsOnBar("ev:2").size());
        assertEquals("田中 次郎", model.personsOnBar("ev:2").get(0).fullName());
        assertTrue(badgeCellContains(badgeRows, "田中"));
    }

    @Test
    void removePerson_promotesSubToPrimary() {
        String yamada = model.personsOnBar("ev:0").get(0).memberKey();
        assertTrue(model.removePerson("ev:0", yamada).isEmpty());

        EquipmentGanttAssignmentBadgeGridUpdater.applyToBadgeRows(
                badgeRows, bundle.assignmentMetadata(), model.snapshotPersonsByBarId());

        assertEquals(1, model.personsOnBar("ev:0").size());
        assertEquals("佐藤 花子", model.personsOnBar("ev:0").get(0).fullName());
        assertFalse(badgeCellContains(badgeRows, "山田"));
        assertTrue(badgeCellContains(badgeRows, "佐藤"));
    }

    private static boolean badgeCellContains(List<List<String>> rows, String fragment) {
        for (List<String> row : rows) {
            if (row == null) {
                continue;
            }
            for (String cell : row) {
                if (cell != null && cell.contains(fragment)) {
                    return true;
                }
            }
        }
        return false;
    }

    private static List<List<String>> deepCopy(List<List<String>> src) {
        List<List<String>> out = new ArrayList<>();
        for (List<String> row : src) {
            out.add(row != null ? new ArrayList<>(row) : new ArrayList<>());
        }
        return out;
    }
}
