package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class DispatchTimelineCalendarMetersIndexTest {

    @TempDir Path tempDir;

    @Test
    void splitsY64SecAcrossCalendarDays() throws Exception {
        String json =
                """
                {
                  "kwargs_packed": {
                    "timeline_events": [
                      {
                        "date": {"__t": "date", "v": "2026-06-10"},
                        "task_id": "Y6-4",
                        "machine": "SEC+SEC機　湖南",
                        "event_kind": "machining",
                        "units_done": 1,
                        "unit_m": 200.0
                      },
                      {
                        "date": {"__t": "date", "v": "2026-06-10"},
                        "task_id": "Y6-4",
                        "machine": "SEC+SEC機　湖南",
                        "event_kind": "machining",
                        "units_done": 4,
                        "unit_m": 200.0
                      },
                      {
                        "date": {"__t": "date", "v": "2026-06-11"},
                        "task_id": "Y6-4",
                        "machine": "SEC+SEC機　湖南",
                        "event_kind": "machining",
                        "units_done": 3,
                        "unit_m": 200.0
                      }
                    ]
                  }
                }
                """;
        Path contract = tempDir.resolve("計画test設.json");
        Files.writeString(contract, json);
        DispatchTimelineCalendarMetersIndex idx =
                DispatchTimelineCalendarMetersIndex.loadFromContractPath(contract);
        assertTrue(idx.isLoaded());
        assertEquals(
                1000.0,
                idx.metersForTaskProfile("Y6-4", "SEC", "SEC機　湖南", LocalDate.of(2026, 6, 10))
                        .orElseThrow(),
                1e-9);
        assertEquals(
                600.0,
                idx.metersForTaskProfile("Y6-4", "SEC", "SEC機　湖南", LocalDate.of(2026, 6, 11))
                        .orElseThrow(),
                1e-9);
        assertEquals(
                0.0,
                idx.metersForTaskProfile("Y6-4", "SEC", "SEC機　湖南", LocalDate.of(2026, 6, 9))
                        .orElseThrow(),
                1e-9);
    }

    @Test
    void aggregatesBranchTaskIdUnderParentRequestNo() throws Exception {
        String json =
                """
                {
                  "kwargs_packed": {
                    "timeline_events": [
                      {
                        "date": {"__t": "date", "v": "2026-06-12"},
                        "task_id": "V6-2-01",
                        "machine": "分割+スリット機1　湖南",
                        "event_kind": "machining",
                        "units_done": 72,
                        "unit_m": 100.0
                      },
                      {
                        "date": {"__t": "date", "v": "2026-06-15"},
                        "task_id": "V6-2-01",
                        "machine": "分割+スリット機1　湖南",
                        "event_kind": "machining",
                        "units_done": 28,
                        "unit_m": 100.0
                      }
                    ]
                  }
                }
                """;
        Path contract = tempDir.resolve("v62_branch.json");
        Files.writeString(contract, json);
        DispatchTimelineCalendarMetersIndex idx =
                DispatchTimelineCalendarMetersIndex.loadFromContractPath(contract);
        assertEquals(
                7200.0,
                idx.metersForTaskProfile(
                                "V6-2", "分割", "スリット機1　湖南", LocalDate.of(2026, 6, 12))
                        .orElseThrow(),
                1e-6);
        assertEquals(
                2800.0,
                idx.metersForTaskProfile(
                                "V6-2", "分割", "スリット機1　湖南", LocalDate.of(2026, 6, 15))
                        .orElseThrow(),
                1e-6);
        assertEquals(
                0.0,
                idx.metersForTaskProfile(
                                "V6-2", "分割", "スリット機1　湖南", LocalDate.of(2026, 6, 11))
                        .orElseThrow(),
                1e-6);
    }
}
