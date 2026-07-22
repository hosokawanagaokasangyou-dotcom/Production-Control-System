package jp.co.pm.ai.desktop.io.gantt;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttContractSheetTableBuilder.TimelineEvent;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentEditModel.Failure;

class EquipmentGanttAssignmentEditModelTest {

    private EquipmentGanttAssignmentEditModel model;
    private String barA;
    private String barB;

    @BeforeEach
    void setUp() {
        List<TimelineEvent> events =
                List.of(
                        machining("Y1-1", "山田 太郎", "佐藤 花子"),
                        machining("Y2-1", "", ""));
        List<EquipmentGanttAssignmentBarUnit> units =
                EquipmentGanttAssignmentMetadataBuilder.buildBarUnits(events);
        barA = units.get(0).barId();
        barB = units.get(1).barId();
        model =
                new EquipmentGanttAssignmentEditModel(
                        new EquipmentGanttAssignmentMetadata(units, List.of()));
    }

    @Test
    void movePerson_toEmptyBar_movesAndPromotesPrimaryOnSource() {
        String yamada = model.personsOnBar(barA).get(0).memberKey();

        Optional<Failure> err = model.movePerson(barA, barB, yamada);

        assertTrue(err.isEmpty(), () -> "unexpected failure: " + err.orElse(null));
        assertEquals(1, model.personsOnBar(barA).size());
        assertEquals(EquipmentGanttAssignmentRole.PRIMARY, model.personsOnBar(barA).get(0).role());
        assertEquals(1, model.personsOnBar(barB).size());
        assertEquals(yamada, model.personsOnBar(barB).get(0).memberKey());
    }

    @Test
    void swapPerson_exchangesBetweenBars() {
        List<TimelineEvent> events =
                List.of(
                        machining("Y1-1", "山田 太郎", ""),
                        machining("Y2-1", "田中 一郎", ""));
        List<EquipmentGanttAssignmentBarUnit> units =
                EquipmentGanttAssignmentMetadataBuilder.buildBarUnits(events);
        EquipmentGanttAssignmentEditModel swapModel =
                new EquipmentGanttAssignmentEditModel(
                        new EquipmentGanttAssignmentMetadata(units, List.of()));
        String bar1 = units.get(0).barId();
        String bar2 = units.get(1).barId();
        String yamada = swapModel.personsOnBar(bar1).get(0).memberKey();
        String tanaka = swapModel.personsOnBar(bar2).get(0).memberKey();

        Optional<Failure> err = swapModel.swapPerson(bar1, bar2, yamada, tanaka);

        assertTrue(err.isEmpty());
        assertEquals(tanaka, swapModel.personsOnBar(bar1).get(0).memberKey());
        assertEquals(yamada, swapModel.personsOnBar(bar2).get(0).memberKey());
    }

    @Test
    void removePerson_lastPersonForbidden() {
        List<TimelineEvent> events = List.of(machining("Y2-1", "田中 一郎", ""));
        List<EquipmentGanttAssignmentBarUnit> units =
                EquipmentGanttAssignmentMetadataBuilder.buildBarUnits(events);
        EquipmentGanttAssignmentEditModel soloModel =
                new EquipmentGanttAssignmentEditModel(
                        new EquipmentGanttAssignmentMetadata(units, List.of()));
        String soloBar = units.get(0).barId();
        String tanaka = soloModel.personsOnBar(soloBar).get(0).memberKey();

        Optional<Failure> err = soloModel.removePerson(soloBar, tanaka);

        assertEquals(Optional.of(Failure.EMPTY_BAR_FORBIDDEN), err);
        assertEquals(1, soloModel.personsOnBar(soloBar).size());
    }

    @Test
    void removePerson_primaryPromotesFirstSub() {
        String yamada = model.personsOnBar(barA).get(0).memberKey();
        String sato = model.personsOnBar(barA).get(1).memberKey();

        assertTrue(model.removePerson(barA, yamada).isEmpty());

        assertEquals(1, model.personsOnBar(barA).size());
        assertEquals(sato, model.personsOnBar(barA).get(0).memberKey());
        assertEquals(EquipmentGanttAssignmentRole.PRIMARY, model.personsOnBar(barA).get(0).role());
    }

    @Test
    void addPerson_rejectsDuplicate() {
        EquipmentGanttAssignmentPerson extra =
                EquipmentGanttAssignmentPerson.fromRawName(
                        "鈴木 花子", EquipmentGanttAssignmentRole.SUB);
        String yamada = model.personsOnBar(barA).get(0).memberKey();

        assertTrue(model.addPerson(barA, extra).isEmpty());
        assertEquals(Optional.of(Failure.DUPLICATE_PERSON), model.addPerson(barA, extra));
        assertTrue(model.addPerson(barA, model.personsOnBar(barA).get(0)).isPresent());
        assertEquals(yamada, model.personsOnBar(barA).get(0).memberKey());
    }

    private static TimelineEvent machining(String taskId, String op, String sub) {
        return new TimelineEvent(
                LocalDate.of(2026, 5, 14),
                "EC機　湖南",
                taskId,
                "machining",
                LocalDateTime.of(2026, 5, 14, 8, 0),
                LocalDateTime.of(2026, 5, 14, 9, 0),
                100.0,
                1.0,
                null,
                null,
                false,
                op,
                sub,
                List.of(),
                -1,
                0,
                Double.NaN);
    }
}
