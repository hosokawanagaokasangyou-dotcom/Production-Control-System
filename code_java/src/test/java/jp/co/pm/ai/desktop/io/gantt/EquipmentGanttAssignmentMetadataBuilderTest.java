package jp.co.pm.ai.desktop.io.gantt;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttContractSheetTableBuilder.TimelineEvent;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentEditModel.Failure;

class EquipmentGanttAssignmentMetadataBuilderTest {

    @Test
    void buildBarUnits_groupsGapSplitMachiningEvents() {
        List<TimelineEvent> events =
                List.of(
                        ev(
                                0,
                                "Y1-1",
                                "machining",
                                LocalDateTime.of(2026, 5, 14, 8, 0),
                                LocalDateTime.of(2026, 5, 14, 10, 0),
                                "山田 太郎",
                                "佐藤 花子"),
                        ev(
                                1,
                                "Y1-1",
                                "machining",
                                LocalDateTime.of(2026, 5, 14, 11, 0),
                                LocalDateTime.of(2026, 5, 14, 12, 0),
                                "山田 太郎",
                                "佐藤 花子"));

        List<EquipmentGanttAssignmentBarUnit> units =
                EquipmentGanttAssignmentMetadataBuilder.buildBarUnits(events);

        assertEquals(2, units.size());
        assertEquals(List.of(0), units.get(0).timelineEventIndices());
        assertEquals(List.of(1), units.get(1).timelineEventIndices());
        assertEquals("ev:0", units.get(0).barId());
        assertEquals("ev:1", units.get(1).barId());
    }

    @Test
    void buildBarUnits_sameSegmentCombinesIndices() {
        List<TimelineEvent> events =
                List.of(
                        ev(
                                0,
                                "Y1-1",
                                "machining",
                                LocalDateTime.of(2026, 5, 14, 8, 0),
                                LocalDateTime.of(2026, 5, 14, 8, 30),
                                "山田 太郎",
                                ""),
                        ev(
                                1,
                                "Y1-1",
                                "machining",
                                LocalDateTime.of(2026, 5, 14, 8, 30),
                                LocalDateTime.of(2026, 5, 14, 9, 0),
                                "山田 太郎",
                                ""));

        List<EquipmentGanttAssignmentBarUnit> units =
                EquipmentGanttAssignmentMetadataBuilder.buildBarUnits(events);

        assertEquals(1, units.size());
        assertEquals(List.of(0, 1), units.get(0).timelineEventIndices());
        assertEquals("ev:0+1", units.get(0).barId());
    }

    @Test
    void personsFromEvent_primaryThenSub_sameSurnameDistinctKeys() {
        TimelineEvent e =
                ev(
                        0,
                        "Y1-1",
                        "machining",
                        LocalDateTime.of(2026, 5, 14, 8, 0),
                        LocalDateTime.of(2026, 5, 14, 9, 0),
                        "山田 太郎",
                        "山田 次郎");

        List<EquipmentGanttAssignmentPerson> persons =
                EquipmentGanttAssignmentMetadataBuilder.personsFromEvent(e);

        assertEquals(2, persons.size());
        assertEquals(EquipmentGanttAssignmentRole.PRIMARY, persons.get(0).role());
        assertEquals(EquipmentGanttAssignmentRole.SUB, persons.get(1).role());
        assertEquals("山田", persons.get(0).badgeLabel());
        assertEquals("山田", persons.get(1).badgeLabel());
        assertFalse(persons.get(0).memberKey().equals(persons.get(1).memberKey()));
    }

    @Test
    void buildBundleFromContract_includesAssignmentMetadata() throws Exception {
        java.nio.file.Path contract =
                java.nio.file.Files.createTempFile("assign-meta", ".json");
        String json =
                """
                {
                  "schema_version": 1,
                  "kind": "equipment_gantt",
                  "fn": "_write_results_equipment_gantt_sheet",
                  "kwargs_packed": {
                    "timeline_events": [
                      {
                        "date": {"__t": "date", "v": "2026-05-14"},
                        "machine": "EC機　湖南",
                        "task_id": "Y1-1",
                        "event_kind": "machining",
                        "start_dt": {"__t": "datetime", "v": "2026-05-14T08:05:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-05-14T08:07:00"},
                        "unit_m": 100.0,
                        "units_done": 1.0,
                        "op": "山田 太郎",
                        "sub": "佐藤 花子"
                      }
                    ],
                    "equipment_list": ["EC機　湖南"],
                    "sorted_dates": [{"__t": "date", "v": "2026-05-14"}]
                  }
                }
                """;
        java.nio.file.Files.writeString(contract, json);

        EquipmentGanttSheetBundle bundle =
                EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(contract);

        assertFalse(bundle.assignmentMetadata().barUnits().isEmpty());
        EquipmentGanttAssignmentBarUnit unit = bundle.assignmentMetadata().barUnits().get(0);
        assertEquals("ev:0", unit.barId());
        assertEquals(2, unit.persons().size());
    }

    @Test
    void slotBindings_useTimeColumnsOnly_whenProgressColumnsBetweenTaskAndTimeline() {
        List<TimelineEvent> events =
                List.of(
                        ev(
                                0,
                                "Y1-1",
                                "machining",
                                LocalDateTime.of(2026, 5, 14, 8, 0),
                                LocalDateTime.of(2026, 5, 14, 8, 20),
                                "宮島 太郎",
                                ""));
        List<String> columns =
                List.of("日付", "機械名", "工程名", "タスク", "進度1", "08:00", "08:10", "08:20");
        Map<String, String> section = new java.util.LinkedHashMap<>();
        section.put("日付", "【2026/05/14】");
        section.put("機械名", "");
        Map<String, String> data = new java.util.LinkedHashMap<>();
        data.put("日付", "");
        data.put("機械名", "EC機　湖南");
        data.put("工程名", "工程");
        data.put("タスク", "—");
        data.put("進度1", "50%");
        data.put("08:00", "依頼NO Y1-1 100m");
        data.put("08:10", "依頼NO Y1-1 100m");
        data.put("08:20", "");
        jp.co.pm.ai.desktop.io.JsonTableIo.SheetTable table =
                new jp.co.pm.ai.desktop.io.JsonTableIo.SheetTable(
                        columns, List.of(section, data));
        List<List<String>> badgeRows =
                List.of(
                        List.of("", "", ""),
                        List.of("宮島", "", ""));
        List<java.time.LocalTime> slotStarts =
                List.of(
                        java.time.LocalTime.of(8, 0),
                        java.time.LocalTime.of(8, 10),
                        java.time.LocalTime.of(8, 20));
        EquipmentGanttAssignmentMetadata meta =
                EquipmentGanttAssignmentMetadataBuilder.build(
                        events, table, badgeRows, slotStarts);
        assertFalse(meta.slotBindings().isEmpty(), "slotBindings must not be empty");
        Optional<String> barId = meta.resolveBarId(1, 0, 1);
        assertTrue(barId.isPresent(), "UI slot 0-1 must resolve barId");
        assertEquals("ev:0", barId.get());
    }

    @Test
    void equipmentColumnMatchesEventMachine_acceptsDisplayLabelVsProcessPlusMachineKey() {
        assertTrue(
                EquipmentGanttContractSheetTableBuilder.equipmentColumnMatchesEventMachine(
                        "EC機　湖南", "熱融着+EC機　湖南"));
        assertTrue(
                EquipmentGanttContractSheetTableBuilder.equipmentColumnMatchesEventMachine(
                        "EC機　湖南（熱融着）", "熱融着+EC機　湖南"));
    }

    @Test
    void resolveBarIdForBadgeRun_matchesStartupRunText() {
        EquipmentGanttAssignmentBarUnit unit =
                new EquipmentGanttAssignmentBarUnit(
                        "ev:0",
                        List.of(0),
                        LocalDate.of(2026, 5, 14),
                        "熱融着+EC機　湖南",
                        "",
                        "machine_daily_startup",
                        List.of(
                                EquipmentGanttAssignmentPerson.fromRawName(
                                        "宮島 太郎", EquipmentGanttAssignmentRole.PRIMARY)));
        EquipmentGanttAssignmentMetadata meta =
                new EquipmentGanttAssignmentMetadata(List.of(unit), List.of());
        Optional<String> barId =
                meta.resolveBarIdForBadgeRun(
                        1, 30, 34, "EC機　湖南", List.of("宮島"), "日次始業準備");
        assertTrue(barId.isPresent(), "startup run text must resolve barId");
        assertEquals("ev:0", barId.get());
    }

    @Test
    void resolveBarIdForBadgeRun_matchesGapSplitMachiningRunText() {
        EquipmentGanttAssignmentBarUnit unit =
                new EquipmentGanttAssignmentBarUnit(
                        "ev:12",
                        List.of(12),
                        LocalDate.of(2026, 5, 14),
                        "熱融着+EC機　湖南",
                        "W7-5",
                        "machining",
                        List.of(
                                EquipmentGanttAssignmentPerson.fromRawName(
                                        "宮島 太郎", EquipmentGanttAssignmentRole.PRIMARY),
                                EquipmentGanttAssignmentPerson.fromRawName(
                                        "竹内 次郎", EquipmentGanttAssignmentRole.SUB)));
        EquipmentGanttAssignmentMetadata meta =
                new EquipmentGanttAssignmentMetadata(List.of(unit), List.of());
        Optional<String> barId =
                meta.resolveBarIdForBadgeRun(
                        1, 35, 39, "EC機　湖南", List.of("宮島", "竹内"), "W7-5 休憩前 900m");
        assertTrue(barId.isPresent(), "gap-split machining run must resolve barId");
        assertEquals("ev:12", barId.get());
    }

    @Test
    void resolveBarIdForBadgeRun_fallsBackToRunTextWhenSlotBindingMissing() {
        EquipmentGanttAssignmentBarUnit unit =
                new EquipmentGanttAssignmentBarUnit(
                        "ev:5",
                        List.of(10),
                        LocalDate.of(2026, 5, 14),
                        "EC機　湖南",
                        "Y1-1",
                        "machining",
                        List.of(
                                EquipmentGanttAssignmentPerson.fromRawName(
                                        "山田 太郎", EquipmentGanttAssignmentRole.PRIMARY)));
        EquipmentGanttAssignmentMetadata meta =
                new EquipmentGanttAssignmentMetadata(List.of(unit), List.of());
        Optional<String> barId =
                meta.resolveBarIdForBadgeRun(
                        1, 3, 8, "EC機　湖南", List.of("宮島"), "依頼NO Y1-1 100m");
        assertTrue(barId.isPresent(), "run text + machine must resolve barId");
        assertEquals("ev:5", barId.get());
    }

    @Test
    void resolveBarIdForBadgeRun_fallsBackToPersonLabelWhenSlotBindingMissing() {
        EquipmentGanttAssignmentBarUnit unit =
                new EquipmentGanttAssignmentBarUnit(
                        "ev:0",
                        List.of(0),
                        LocalDate.of(2026, 5, 14),
                        "EC機　湖南",
                        "Y1-1",
                        "machining",
                        List.of(
                                EquipmentGanttAssignmentPerson.fromRawName(
                                        "宮島 太郎", EquipmentGanttAssignmentRole.PRIMARY)));
        EquipmentGanttAssignmentMetadata meta =
                new EquipmentGanttAssignmentMetadata(
                        List.of(unit),
                        List.of(
                                new EquipmentGanttAssignmentSlotBinding(1, 0, 1, "ev:other")));
        Optional<String> barId =
                meta.resolveBarIdForBadgeRun(
                        1, 3, 8, "EC機　湖南", List.of("宮島"));
        assertTrue(barId.isPresent(), "person+machine fallback must resolve barId");
        assertEquals("ev:0", barId.get());
    }

    private static TimelineEvent ev(
            int ignored,
            String taskId,
            String kind,
            LocalDateTime start,
            LocalDateTime end,
            String op,
            String sub) {
        return new TimelineEvent(
                LocalDate.of(2026, 5, 14),
                "EC機　湖南",
                taskId,
                kind,
                start,
                end,
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
