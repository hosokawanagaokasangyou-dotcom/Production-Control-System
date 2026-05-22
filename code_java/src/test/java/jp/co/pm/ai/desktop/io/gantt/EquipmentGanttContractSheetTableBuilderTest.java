package jp.co.pm.ai.desktop.io.gantt;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.io.JsonTableIo;

class EquipmentGanttContractSheetTableBuilderTest {

    @Test
    void shortMachining_fillsExactlyOneTenMinuteSlot(@TempDir Path tmp) throws Exception {
        Path contract = tmp.resolve("shortMach設.json");
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
                        "units_done": 1.0
                      }
                    ],
                    "equipment_list": ["EC機　湖南"],
                    "sorted_dates": [{"__t": "date", "v": "2026-05-14"}]
                  }
                }
                """;
        Files.writeString(contract, json, StandardCharsets.UTF_8);

        EquipmentGanttSheetBundle bundle =
                EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(contract);
        JsonTableIo.SheetTable table = bundle.table();
        List<String> cols = table.columns();
        int idx800 = cols.indexOf("8:00");
        int idx810 = cols.indexOf("8:10");
        int idx820 = cols.indexOf("8:20");
        assertTrue(idx800 >= 0, "8:00 column");
        assertTrue(idx810 >= 0, "8:10 column");
        assertTrue(idx820 >= 0, "8:20 column");

        Map<String, String> dataRow =
                table.rows().stream()
                        .filter(r -> r.getOrDefault("機械名", "").contains("EC機　湖南"))
                        .findFirst()
                        .orElseThrow();

        assertTrue(
                dataRow.getOrDefault("8:00", "").contains("Y1-1"),
                "8:00 slot should show short machining");
        assertEquals("", dataRow.getOrDefault("8:10", "").strip(), "8:10 slot should be empty");
        assertEquals("", dataRow.getOrDefault("8:20", "").strip(), "8:20 slot should be empty");
    }

    @Test
    void overlappingPrepAndMachining_prefersMachiningInSlot(@TempDir Path tmp) throws Exception {
        Path contract = tmp.resolve("prepVsMach.json");
        String json =
                """
                {
                  "schema_version": 1,
                  "kind": "equipment_gantt",
                  "fn": "_write_results_equipment_gantt_sheet",
                  "kwargs_packed": {
                    "timeline_events": [
                      {
                        "date": {"__t": "date", "v": "2026-05-25"},
                        "machine": "SEC機　湖南",
                        "task_id": "",
                        "event_kind": "request_switch_prep",
                        "start_dt": {"__t": "datetime", "v": "2026-05-25T14:55:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-05-25T15:35:00"}
                      },
                      {
                        "date": {"__t": "date", "v": "2026-05-25"},
                        "machine": "SEC機　湖南",
                        "task_id": "Y5-27",
                        "event_kind": "machining",
                        "start_dt": {"__t": "datetime", "v": "2026-05-25T15:00:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-05-25T15:12:00"},
                        "unit_m": 100.0,
                        "units_done": 1.0
                      }
                    ],
                    "equipment_list": ["SEC機　湖南"],
                    "sorted_dates": [{"__t": "date", "v": "2026-05-25"}]
                  }
                }
                """;
        Files.writeString(contract, json, StandardCharsets.UTF_8);

        EquipmentGanttSheetBundle bundle =
                EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(contract);
        Map<String, String> dataRow =
                bundle.table().rows().stream()
                        .filter(r -> r.getOrDefault("機械名", "").contains("SEC機"))
                        .findFirst()
                        .orElseThrow();

        assertTrue(
                dataRow.getOrDefault("15:00", "").contains("Y5-27"),
                "15:00 slot should show machining, not request_switch_prep");
        assertTrue(
                !dataRow.getOrDefault("15:00", "").contains("依頼切替準備"),
                "15:00 slot must not be prep-only label");
    }

    @Test
    void overlappingBreakResumePrepAndMachining_prefersBreakResumeInSlot(@TempDir Path tmp)
            throws Exception {
        Path contract = tmp.resolve("breakResumeVsMach.json");
        String json =
                """
                {
                  "schema_version": 1,
                  "kind": "equipment_gantt",
                  "fn": "_write_results_equipment_gantt_sheet",
                  "kwargs_packed": {
                    "timeline_events": [
                      {
                        "date": {"__t": "date", "v": "2026-05-25"},
                        "machine": "SEC機　湖南",
                        "task_id": "",
                        "event_kind": "break_resume_prep",
                        "start_dt": {"__t": "datetime", "v": "2026-05-25T12:45:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-05-25T13:05:00"}
                      },
                      {
                        "date": {"__t": "date", "v": "2026-05-25"},
                        "machine": "SEC機　湖南",
                        "task_id": "Y5-27",
                        "event_kind": "machining",
                        "start_dt": {"__t": "datetime", "v": "2026-05-25T13:05:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-05-25T14:35:00"},
                        "unit_m": 100.0,
                        "units_done": 1.0
                      }
                    ],
                    "equipment_list": ["SEC機　湖南"],
                    "sorted_dates": [{"__t": "date", "v": "2026-05-25"}]
                  }
                }
                """;
        Files.writeString(contract, json, StandardCharsets.UTF_8);

        EquipmentGanttSheetBundle bundle =
                EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(contract);
        Map<String, String> dataRow =
                bundle.table().rows().stream()
                        .filter(r -> r.getOrDefault("機械名", "").contains("SEC機"))
                        .findFirst()
                        .orElseThrow();

        assertTrue(
                dataRow.getOrDefault("13:00", "").contains("休憩再開準備"),
                "13:00 slot should show break_resume_prep, not machining that starts at 13:05");
        assertTrue(
                dataRow.getOrDefault("12:50", "").contains("休憩再開準備"),
                "12:50 slot should show break_resume_prep");
    }

    @Test
    void slotOverlapRange_expandsSubTenMinuteMachiningToOneSlot() {
        EquipmentGanttContractSheetTableBuilder.TimelineEvent ev =
                new EquipmentGanttContractSheetTableBuilder.TimelineEvent(
                        LocalDate.of(2026, 5, 14),
                        "EC機　湖南",
                        "Y1-1",
                        "machining",
                        LocalDateTime.of(2026, 5, 14, 8, 5),
                        LocalDateTime.of(2026, 5, 14, 8, 7),
                        100.0,
                        1.0,
                        null,
                        null,
                        false,
                        "",
                        "",
                        List.of(),
                        -1,
                        0,
                        Double.NaN);

        EquipmentGanttContractSheetTableBuilder.SlotOverlapRange range =
                EquipmentGanttContractSheetTableBuilder.slotOverlapRangeForDisplay(ev);

        assertEquals(LocalDateTime.of(2026, 5, 14, 8, 0), range.start());
        assertEquals(LocalDateTime.of(2026, 5, 14, 8, 10), range.end());
    }
}
